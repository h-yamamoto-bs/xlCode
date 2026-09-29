import { readFile } from 'node:fs/promises';
import ExcelJS from 'exceljs';
import { atomicWrite } from './atomic';
import { MAX_CELL_CHARS } from './constants';

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

/**
 * xlsx（XML）に保存できる形へ変換する。読み込み時は ExcelJS が元に戻す。
 * - 文字列中の「_xHHHH_」は xlsx では文字コードとして解釈されるため、先頭の「_」を _x005F_ にする
 * - XML に書けない制御文字は _xHHHH_ で表す（そのまま書くと読み込めないブックになる）
 */
export function escapeCell(text: string): string {
  return text
    .replace(/_(x[0-9A-Fa-f]{4}_)/g, '_x005F_$1')
    .replace(XML_INVALID, (c) => `_x${c.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}_`);
}

/** XML 1.0 に書けない文字（タブ・改行・復帰以外の制御文字など） */
// eslint-disable-next-line no-control-regex
const XML_INVALID = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g;

/** 対になっていないサロゲート（壊れた UTF-16）を含むか。xlsx に保存できない */
export function hasLoneSurrogate(text: string): boolean {
  return /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(text);
}

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
  private constructor(readonly wb: ExcelJS.Workbook) {}

  static create(): Book {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'xlCode';
    return new Book(wb);
  }

  static async load(file: string): Promise<Book> {
    const wb = new ExcelJS.Workbook();
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
    try {
      await wb.xlsx.load(buf as unknown as ArrayBuffer);
    } catch (e) {
      throw new Error(`ブックが壊れているか、xlsx ではありません: ${file}（${e instanceof Error ? e.message : e}）`, {
        cause: e,
      });
    }
    return new Book(wb);
  }

  async save(file: string): Promise<void> {
    const buf = await this.wb.xlsx.writeBuffer();
    await atomicWrite(file, Buffer.from(buf as ArrayBuffer));
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
        cell.value = v === undefined || v === '' ? null : escapeCell(v);
        cell.numFmt = '@';
        cell.font = CODE_FONT;
      }
    }
  }

  deleteSheet(name: string): void {
    const ws = this.wb.getWorksheet(name);
    if (ws) this.wb.removeWorksheet(ws.id);
  }
}
