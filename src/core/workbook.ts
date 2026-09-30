import { readFile } from 'node:fs/promises';
import ExcelJS from 'exceljs';
import { atomicWrite } from './atomic';
import { escapeCell, hasLoneSurrogate } from './cellEscape';
import { MAX_CELL_CHARS } from './constants';
import { patchXlsx, type PatchOp } from './xlsxPatch';

export { escapeCell, hasLoneSurrogate } from './cellEscape';

export interface CellIssue {
  row: number;
  /** 検出した型（number / boolean / date / formula / error） */
  type: string;
}

export interface SheetData {
  name: string;
  lines: string[];
  /** 文字列以外に自動変換されたセル（3.8） */
  issues: CellIssue[];
  /** B列以降に値がある（コードとしては扱わない） */
  hasExtraColumns: boolean;
}

const CODE_FONT = { name: 'Consolas', size: 10 };

function cellText(value: ExcelJS.CellValue): { text: string; type?: string } {
  if (value === null || value === undefined) return { text: '' };
  if (typeof value === 'string') return { text: value };
  if (typeof value === 'number') return { text: String(value), type: 'number' };
  if (typeof value === 'boolean') return { text: String(value), type: 'boolean' };
  if (value instanceof Date) return { text: value.toISOString(), type: 'date' };
  if (typeof value === 'object') {
    if ('richText' in value) return { text: value.richText.map((r) => r.text).join('') };
    if ('formula' in value || 'sharedFormula' in value) {
      const f = 'formula' in value ? value.formula : value.sharedFormula;
      return { text: `=${f}`, type: 'formula' };
    }
    if ('error' in value) return { text: String(value.error), type: 'error' };
    if ('text' in value) {
      const t = (value as { text: string | ExcelJS.CellRichTextValue }).text;
      return { text: typeof t === 'string' ? t : t.richText.map((r) => r.text).join('') };
    }
  }
  return { text: String(value), type: typeof value };
}

/** .xlcode.xlsx の読み書き。コードは A 列に 1行 = 1セルで格納する（3.6） */
export class Book {
  /** 読み込んだ時点（または最後に保存した時点）のファイルの中身。保存はこれを部分的に書き換える */
  private source: Uint8Array | null;
  /** 保存時に適用する変更（コードのシートの書き込み・削除・並べ替え） */
  private ops: PatchOp[] = [];

  private constructor(
    readonly wb: ExcelJS.Workbook,
    source: Uint8Array | null,
  ) {
    this.source = source;
  }

  static create(): Book {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'xlCode';
    return new Book(wb, null);
  }

  static async load(file: string): Promise<Book> {
    // ExcelJS の readFile はパスの扱いが環境依存なので Buffer 経由で読む
    let buf: Buffer;
    try {
      buf = await readFile(file);
    } catch (e) {
      const code = (e as NodeJS.ErrnoException).code;
      if (code === 'ENOENT') throw e;
      // OneDrive の「オンラインのみ」のファイルは、読むときにダウンロードされる。オフライン等で失敗すると EIO などになる
      throw new Error(
        `ブックを読み込めません（${code ?? e}）: ${file}。OneDrive の「オンラインのみ」のファイルの場合は、エクスプローラーで「このデバイス上で常に保持する」を選んでください`,
        { cause: e },
      );
    }
    return Book.fromBuffer(buf, file);
  }

  /** 読み込み済みのファイルの中身から開く（file はエラー表示用） */
  static async fromBuffer(buf: Uint8Array, file: string): Promise<Book> {
    const wb = new ExcelJS.Workbook();
    try {
      await wb.xlsx.load(buf as unknown as ArrayBuffer);
    } catch (e) {
      throw new Error(`ブックが壊れているか、xlsx ではありません: ${file}（${e instanceof Error ? e.message : e}）`, {
        cause: e,
      });
    }
    return new Book(wb, buf);
  }

  /**
   * 保存する。ブック全体を書き直すのではなく、変更したシートの XML だけを差し替える。
   * UI 用シートの図形・ボタン・グラフや VBA は、そのまま残る。
   */
  async save(file: string): Promise<void> {
    const base = this.source ?? new Uint8Array((await new ExcelJS.Workbook().xlsx.writeBuffer()) as ArrayBuffer);
    const out = await patchXlsx(base, this.ops);
    await atomicWrite(file, out);
    this.source = out;
    this.ops = [];
  }

  /** シートを先頭へ移動する */
  moveToFront(name: string): void {
    const sheets = this.wb.worksheets as unknown as { name: string; orderNo: number }[];
    const ws = sheets.find((s) => s.name === name);
    if (!ws) return;
    const min = Math.min(...sheets.map((s) => s.orderNo));
    if (ws.orderNo === min && sheets.filter((s) => s.orderNo === min).length === 1) return;
    ws.orderNo = min - 1;
    this.ops.push({ kind: 'front', name });
  }

  sheetNames(): string[] {
    return this.wb.worksheets.map((ws) => ws.name);
  }

  hasSheet(name: string): boolean {
    return this.wb.getWorksheet(name) !== undefined;
  }

  readSheet(name: string): SheetData {
    const ws = this.wb.getWorksheet(name);
    if (!ws) throw new Error(`シート「${name}」がありません`);
    const lines: string[] = [];
    const issues: CellIssue[] = [];
    let hasExtraColumns = false;
    for (let r = 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const { text, type } = cellText(row.getCell(1).value);
      lines.push(text);
      if (type) issues.push({ row: r, type });
      if (!hasExtraColumns) {
        row.eachCell({ includeEmpty: false }, (cell, col) => {
          if (col > 1 && cell.value !== null && cell.value !== '') hasExtraColumns = true;
        });
      }
    }
    return { name, lines, issues, hasExtraColumns };
  }

  /** 指定列の行配列を読む（衝突シート用） */
  readColumn(name: string, col: number, fromRow: number): string[] {
    const ws = this.wb.getWorksheet(name);
    if (!ws) return [];
    const out: string[] = [];
    for (let r = fromRow; r <= ws.rowCount; r++) out.push(cellText(ws.getRow(r).getCell(col).value).text);
    return out;
  }

  /**
   * A 列に行を書き込む。シートが無ければ末尾に作成する。
   * 全セルを文字列書式（@）にして、数式化・日付変換などを防ぐ（3.8）。
   */
  writeLines(name: string, lines: readonly string[], columns: readonly (readonly string[])[] = []): void {
    for (const [i, l] of lines.entries()) {
      if (escapeCell(l).length > MAX_CELL_CHARS) {
        throw new Error(`「${name}」の ${i + 1} 行目が ${MAX_CELL_CHARS} 文字を超えています`);
      }
      if (hasLoneSurrogate(l)) throw new Error(`「${name}」の ${i + 1} 行目に壊れた文字（孤立サロゲート）があります`);
    }
    let ws = this.wb.getWorksheet(name);
    if (!ws) {
      ws = this.wb.addWorksheet(name);
      ws.getColumn(1).width = 120;
    }
    const cols = [lines, ...columns];
    this.ops.push({ kind: 'write', name, columns: cols.map((c) => [...c]) });
    for (let c = 1; c <= cols.length; c++) {
      const col = ws.getColumn(c);
      col.numFmt = '@';
      col.font = CODE_FONT;
    }
    const rowCount = Math.max(ws.rowCount, ...cols.map((c) => c.length));
    for (let r = 1; r <= rowCount; r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= cols.length; c++) {
        const cell = row.getCell(c);
        const v = cols[c - 1][r - 1];
        cell.value = v === undefined || v === '' ? null : v;
        cell.numFmt = '@';
        cell.font = CODE_FONT;
      }
    }
  }

  deleteSheet(name: string): void {
    const ws = this.wb.getWorksheet(name);
    if (!ws) return;
    this.wb.removeWorksheet(ws.id);
    this.ops.push({ kind: 'delete', name });
  }
}
