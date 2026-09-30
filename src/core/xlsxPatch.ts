import JSZip from 'jszip';
import { escapeCell } from './cellEscape';

/**
 * 既存の xlsx / xlsm を「変更したシートだけ」書き換える。
 *
 * ExcelJS でブック全体を書き直すと、UI 用シートの図形・ボタン・グラフ・フォームコントロールや
 * VBA（vbaProject.bin）が失われる。そこで zip の中の該当シートの XML だけを差し替え、
 * それ以外のパーツはそのまま残す。
 */
export type PatchOp =
  | { kind: 'write'; name: string; columns: readonly (readonly string[])[] }
  | { kind: 'delete'; name: string }
  | { kind: 'front'; name: string };

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_WORKSHEET = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
const REL_CALCCHAIN = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain';
const CT_WORKSHEET = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
const REL_SST = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
const CT_SST = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml';

function attrs(tag: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const m of tag.matchAll(/([\w:]+)="([^"]*)"/g)) out[m[1]] = m[2];
  return out;
}

function xmlEscape(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function xmlUnescape(s: string): string {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&amp;/g, '&');
}

/** 列番号（1 始まり）→ A, B, … */
function colName(n: number): string {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

/**
 * 共有文字列表（sharedStrings.xml）への追加。既存の項目は変えず、末尾に足すだけにする
 * （UI 用シートが参照している番号を変えないため。使われなくなった項目は Excel が保存時に整理する）
 */
class SharedStrings {
  private added: string[] = [];
  private index = new Map<string, number>();
  constructor(private readonly base: number) {}
  add(text: string): number {
    let i = this.index.get(text);
    if (i === undefined) {
      i = this.base + this.added.length;
      this.added.push(text);
      this.index.set(text, i);
    }
    return i;
  }
  get items(): readonly string[] {
    return this.added;
  }
}

/** ワークシートの XML を作る */
function worksheetXml(columns: readonly (readonly string[])[], style: number | null, sst: SharedStrings): string {
  const rows = Math.max(0, ...columns.map((c) => c.length));
  const s = style === null ? '' : ` s="${style}"`;
  const body: string[] = [];
  for (let r = 1; r <= rows; r++) {
    const cells: string[] = [];
    columns.forEach((col, i) => {
      const v = col[r - 1];
      if (v === undefined || v === '') return;
      cells.push(`<c r="${colName(i + 1)}${r}"${s} t="s"><v>${sst.add(v)}</v></c>`);
    });
    if (cells.length > 0) body.push(`<row r="${r}">${cells.join('')}</row>`);
  }
  const cols = columns
    .map(
      (_, i) =>
        `<col min="${i + 1}" max="${i + 1}" width="120" customWidth="1"${style === null ? '' : ` style="${style}"`}/>`,
    )
    .join('');
  const dim = rows === 0 ? 'A1' : `A1:${colName(Math.max(1, columns.length))}${rows}`;
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">` +
    `<dimension ref="${dim}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>` +
    `<sheetFormatPr defaultRowHeight="15"/>` +
    (cols ? `<cols>${cols}</cols>` : '') +
    `<sheetData>${body.join('')}</sheetData>` +
    `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>` +
    `</worksheet>`
  );
}

/** 文字列書式（@ = numFmtId 49）・等幅フォントのセルスタイルを探し、無ければ追加する */
function ensureTextStyle(styles: string): { styles: string; index: number | null } {
  const xfsMatch = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles);
  const fontsMatch = /<fonts\b[^>]*>([\s\S]*?)<\/fonts>/.exec(styles);
  if (!xfsMatch || !fontsMatch) return { styles, index: null };
  const xfs = [...xfsMatch[1].matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].map((m) => m[0]);
  const fonts = [...fontsMatch[1].matchAll(/<font\b[^>]*?(?:\/>|>[\s\S]*?<\/font>)/g)].map((m) => m[0]);
  const found = xfs.findIndex((x) => {
    const a = attrs(x.slice(0, x.indexOf('>') + 1));
    return a.numFmtId === '49' && /Consolas/.test(fonts[Number(a.fontId)] ?? '');
  });
  if (found >= 0) return { styles, index: found };
  const fontId = fonts.length;
  const xfId = xfs.length;
  let out = styles.replace(
    fontsMatch[0],
    fontsMatch[0]
      .replace(/<fonts\b([^>]*?)count="\d+"/, `<fonts$1count="${fontId + 1}"`)
      .replace('</fonts>', '<font><sz val="10"/><name val="Consolas"/><family val="3"/></font></fonts>'),
  );
  const xfsBlock = /<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/.exec(out)![0];
  out = out.replace(
    xfsBlock,
    xfsBlock
      .replace(/<cellXfs\b([^>]*?)count="\d+"/, `<cellXfs$1count="${xfId + 1}"`)
      .replace(
        '</cellXfs>',
        `<xf numFmtId="49" fontId="${fontId}" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/></cellXfs>`,
      ),
  );
  return { styles: out, index: xfId };
}

interface SheetEntry {
  tag: string;
  name: string;
  sheetId: number;
  rid: string;
}

function readSheets(workbook: string): SheetEntry[] {
  const block = /<sheets>([\s\S]*?)<\/sheets>/.exec(workbook)?.[1] ?? '';
  return [...block.matchAll(/<sheet\b[^>]*\/>/g)].map((m) => {
    const a = attrs(m[0]);
    return { tag: m[0], name: xmlUnescape(a.name ?? ''), sheetId: Number(a.sheetId), rid: a['r:id'] };
  });
}

/** シートの並びが変わったときに、位置で参照している値（名前の定義・開いているタブ）を合わせる */
function remapPositions(workbook: string, map: (oldIndex: number) => number | null, count: number): string {
  let out = workbook.replace(/<definedName\b[^>]*>[\s\S]*?<\/definedName>/g, (dn) => {
    const m = /localSheetId="(\d+)"/.exec(dn);
    if (!m) return dn;
    const next = map(Number(m[1]));
    return next === null ? '' : dn.replace(m[0], `localSheetId="${next}"`);
  });
  out = out.replace(/<definedNames>\s*<\/definedNames>/, '');
  out = out.replace(/(<workbookView\b[^>]*?)\b(activeTab|firstSheet)="(\d+)"/g, (all, pre, key, v) => {
    const next = map(Number(v));
    return `${pre}${key}="${Math.min(Math.max(next ?? 0, 0), Math.max(count - 1, 0))}"`;
  });
  return out;
}

/**
 * シート一覧（<sheets>）が無い・空要素のブック（シート 0 枚の新規ブックなど）に、
 * スキーマの順序どおりの位置へ <bookViews> と <sheets> を補う
 */
export function ensureSheetsElement(workbook: string): string {
  let wb = workbook.replace(/<sheets\s*\/>/, '<sheets></sheets>');
  if (/<sheets>/.test(wb)) return wb;
  const before = [
    '<functionGroups',
    '<externalReferences',
    '<definedNames',
    '<calcPr',
    '<oleSize',
    '<customWorkbookViews',
    '<pivotCaches',
    '<extLst',
    '</workbook>',
  ]
    .map((tag) => wb.indexOf(tag))
    .filter((i) => i >= 0);
  const at = Math.min(...before);
  const views = /<bookViews>/.test(wb) ? '' : '<bookViews><workbookView/></bookViews>';
  wb = wb.slice(0, at) + views + '<sheets></sheets>' + wb.slice(at);
  return wb;
}

function resolveTarget(target: string): string {
  return target.startsWith('/') ? target.slice(1) : `xl/${target}`;
}

export async function patchXlsx(original: Uint8Array, ops: readonly PatchOp[]): Promise<Buffer> {
  const zip = await JSZip.loadAsync(original);
  const read = async (p: string) => (await zip.file(p)?.async('string')) ?? null;
  let workbook = ensureSheetsElement((await read('xl/workbook.xml'))!);
  let rels = (await read('xl/_rels/workbook.xml.rels'))!;
  let types = (await read('[Content_Types].xml'))!;
  let styles = await read('xl/styles.xml');
  if (!workbook || !rels || !types) throw new Error('xlsx の構成が想定と違います（workbook.xml などがありません）');

  // 共有文字列表（無ければ作る）
  const SST_PART = 'xl/sharedStrings.xml';
  let sstXml = await read(SST_PART);
  const sstCount = sstXml ? (sstXml.match(/<si[\s>]/g) ?? []).length : 0;
  const sst = new SharedStrings(sstCount);

  let styleIndex: number | null = null;
  if (styles && ops.some((o) => o.kind === 'write')) {
    const r = ensureTextStyle(styles);
    styles = r.styles;
    styleIndex = r.index;
  }

  const relTarget = (rid: string): string | null => {
    const tag = [...rels.matchAll(/<Relationship\b[^>]*\/>/g)].map((m) => m[0]).find((t) => attrs(t).Id === rid);
    return tag ? resolveTarget(attrs(tag).Target) : null;
  };

  for (const op of ops) {
    const sheets = readSheets(workbook);
    const index = sheets.findIndex((s) => s.name.toLowerCase() === op.name.toLowerCase());
    if (op.kind === 'write') {
      const xml = worksheetXml(op.columns, styleIndex, sst);
      if (index >= 0) {
        const part = relTarget(sheets[index].rid);
        if (!part) throw new Error(`シート「${op.name}」の実体が見つかりません`);
        zip.file(part, xml);
        continue;
      }
      // 新しいシート: パーツ・リレーション・コンテンツタイプ・シート一覧に追加
      let n = 1;
      while (zip.file(`xl/worksheets/sheet${n}.xml`)) n++;
      const part = `xl/worksheets/sheet${n}.xml`;
      zip.file(part, xml);
      const ids = [...rels.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]));
      const rid = `rId${Math.max(0, ...ids) + 1}`;
      rels = rels.replace(
        '</Relationships>',
        `<Relationship Id="${rid}" Type="${REL_WORKSHEET}" Target="worksheets/sheet${n}.xml"/></Relationships>`,
      );
      if (!types.includes(`PartName="/${part}"`)) {
        types = types.replace('</Types>', `<Override PartName="/${part}" ContentType="${CT_WORKSHEET}"/></Types>`);
      }
      const sheetId = Math.max(0, ...sheets.map((s) => s.sheetId)) + 1;
      workbook = workbook.replace(
        '</sheets>',
        `<sheet name="${xmlEscape(op.name)}" sheetId="${sheetId}" r:id="${rid}"/></sheets>`,
      );
    } else if (op.kind === 'delete') {
      if (index < 0) continue;
      const entry = sheets[index];
      const part = relTarget(entry.rid);
      workbook = workbook.replace(entry.tag, '');
      workbook = remapPositions(workbook, (i) => (i === index ? null : i > index ? i - 1 : i), sheets.length - 1);
      rels = rels.replace(new RegExp(`<Relationship\\b[^>]*Id="${entry.rid}"[^>]*/>`), '');
      if (part) {
        zip.remove(part);
        zip.remove(part.replace(/([^/]+)$/, '_rels/$1.rels'));
        types = types.replace(new RegExp(`<Override\\b[^>]*PartName="/${part.replace(/[.]/g, '\\.')}"[^>]*/>`), '');
      }
    } else if (op.kind === 'front') {
      if (index <= 0) continue;
      const entry = sheets[index];
      workbook = workbook.replace(entry.tag, '').replace(/<sheets>/, `<sheets>${entry.tag}`);
      workbook = remapPositions(workbook, (i) => (i === index ? 0 : i < index ? i + 1 : i), sheets.length);
    }
  }

  // 計算チェーンは古いセル位置を指していると修復の対象になる。無くても Excel が作り直す
  if (zip.file('xl/calcChain.xml')) {
    zip.remove('xl/calcChain.xml');
    rels = rels.replace(new RegExp(`<Relationship\\b[^>]*Type="${REL_CALCCHAIN}"[^>]*/>`), '');
    types = types.replace(/<Override\b[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/, '');
  }

  if (sst.items.length > 0) {
    const si = sst.items.map((s) => `<si><t xml:space="preserve">${xmlEscape(escapeCell(s))}</t></si>`).join('');
    if (!sstXml) {
      sstXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${NS_MAIN}" count="0" uniqueCount="0"></sst>`;
      const ids = [...rels.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]));
      rels = rels.replace(
        '</Relationships>',
        `<Relationship Id="rId${Math.max(0, ...ids) + 1}" Type="${REL_SST}" Target="sharedStrings.xml"/></Relationships>`,
      );
      if (!types.includes('PartName="/xl/sharedStrings.xml"')) {
        types = types.replace('</Types>', `<Override PartName="/${SST_PART}" ContentType="${CT_SST}"/></Types>`);
      }
    }
    const total = sstCount + sst.items.length;
    sstXml = sstXml
      .replace(/<sst\b([^>]*?)\/>/, `<sst$1></sst>`)
      .replace(/(<sst\b[^>]*?)\buniqueCount="\d+"/, `$1uniqueCount="${total}"`)
      .replace(/(<sst\b[^>]*?)\bcount="(\d+)"/, (_, pre, c) => `${pre}count="${Math.max(Number(c), total)}"`)
      .replace('</sst>', `${si}</sst>`);
    zip.file(SST_PART, sstXml);
  }

  zip.file('xl/workbook.xml', workbook);
  zip.file('xl/_rels/workbook.xml.rels', rels);
  zip.file('[Content_Types].xml', types);
  if (styles) zip.file('xl/styles.xml', styles);
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}
