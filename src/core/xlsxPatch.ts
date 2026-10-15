import JSZip from 'jszip';
import { escapeCell } from './cellEscape';
import { THEME_FONT, THEMES, type BookTheme, type SheetRole } from './theme';

/**
 * 既存の xlsx / xlsm を「変更したシートだけ」書き換える。
 *
 * ExcelJS でブック全体を書き直すと、UI 用シートの図形・ボタン・グラフ・フォームコントロールや
 * VBA（vbaProject.bin）が失われる。そこで zip の中の該当シートの XML だけを差し替え、
 * それ以外のパーツはそのまま残す。
 */
export type PatchOp =
  | { kind: 'write'; name: string; columns: readonly (readonly string[])[]; role?: SheetRole }
  /** 中身は変えずに、テーマ（塗りつぶし・文字色・枠線・タブの色）だけを当て直す */
  | { kind: 'restyle'; name: string; role?: SheetRole }
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

/** シートに当てるスタイル（styles.xml の cellXfs の番号など） */
interface SheetStyle {
  /** 文字列書式（@）で、フォント・塗りつぶしをテーマに合わせたセルスタイル */
  text: number | null;
  /** データの無い列（B 列以降）・行を塗りつぶすための既定スタイル（書式は標準） */
  fill: number | null;
  /** タブの色（RRGGBB）。テーマの目印にもなる */
  tab: string | null;
  /** 枠線（目盛線）を消す */
  hideGrid: boolean;
}

/** 列の最大番号（xlsx は 16384 列 = XFD） */
const MAX_COL = 16384;
/** 列幅の既定値（Excel の標準の幅） */
const FILL_COL_WIDTH = '9.140625';

/** ワークシートの XML を作る */
function worksheetXml(columns: readonly (readonly string[])[], style: SheetStyle, sst: SharedStrings): string {
  const rows = Math.max(0, ...columns.map((c) => c.length));
  const s = style.text === null ? '' : ` s="${style.text}"`;
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
        `<col min="${i + 1}" max="${i + 1}" width="120" customWidth="1"${style.text === null ? '' : ` style="${style.text}"`}/>`,
    )
    // 残りの列をすべて同じ背景にする（データの無い範囲まで塗りつぶす）
    .concat(
      style.fill !== null && columns.length < MAX_COL
        ? [`<col min="${columns.length + 1}" max="${MAX_COL}" width="${FILL_COL_WIDTH}" style="${style.fill}"/>`]
        : [],
    )
    .join('');
  const dim = rows === 0 ? 'A1' : `A1:${colName(Math.max(1, columns.length))}${rows}`;
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">` +
    (style.tab ? `<sheetPr><tabColor rgb="FF${style.tab}"/></sheetPr>` : '') +
    `<dimension ref="${dim}"/><sheetViews><sheetView${style.hideGrid ? ' showGridLines="0"' : ''} workbookViewId="0"/></sheetViews>` +
    `<sheetFormatPr defaultRowHeight="15"/>` +
    (cols ? `<cols>${cols}</cols>` : '') +
    `<sheetData>${body.join('')}</sheetData>` +
    `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>` +
    `</worksheet>`
  );
}

/**
 * 既存のシート XML に、テーマだけを当て直す（中身のセルは変えない）。
 * - タブの色（sheetPr）・枠線なし（sheetView）
 * - 列の既定スタイル（cols）。16384 列までの残りも塗りつぶす
 * - すべてのセルのスタイル
 */
export function restyleSheetXml(xml: string, style: SheetStyle): string {
  let out = xml;
  // タブの色
  if (style.tab) {
    const tab = `<tabColor rgb="FF${style.tab}"/>`;
    if (/<sheetPr\b[^>]*\/>/.test(out)) {
      out = out.replace(/<sheetPr\b([^>]*?)\s*\/>/, `<sheetPr$1>${tab}</sheetPr>`);
    } else if (/<sheetPr\b/.test(out)) {
      out = /<tabColor\b[^>]*\/>/.test(out)
        ? out.replace(/<tabColor\b[^>]*\/>/, tab)
        : out.replace(/(<sheetPr\b[^>]*>)/, `$1${tab}`);
    } else {
      out = out.replace(/(<worksheet\b[^>]*>)/, `$1<sheetPr>${tab}</sheetPr>`);
    }
  }
  // 枠線（sheetViews が無いシート（ExcelJS が保存したものなど）には、スキーマの順序の位置へ足す）
  if (/<sheetView\b/.test(out)) {
    out = out.replace(/<sheetView\b([^>]*?)(\/?)>/, (_all, attrsText: string, close: string) => {
      const rest = attrsText.replace(/\s*showGridLines="[^"]*"/, '');
      return `<sheetView${style.hideGrid ? ' showGridLines="0"' : ''}${rest}${close}>`;
    });
  } else if (style.hideGrid) {
    const views = '<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>';
    const after = /<dimension\b[^>]*\/>/.exec(out) ?? /<\/sheetPr>/.exec(out) ?? /<sheetPr\b[^>]*\/>/.exec(out);
    if (after) out = out.replace(after[0], `${after[0]}${views}`);
    else out = out.replace(/(<worksheet\b[^>]*>)/, `$1${views}`);
  }
  // 列の既定スタイル
  if (style.text !== null) {
    const cols: string[] = [];
    let last = 0;
    const colsBlock = /<cols>([\s\S]*?)<\/cols>/.exec(out);
    for (const m of (colsBlock?.[1] ?? '').matchAll(/<col\b[^>]*\/>/g)) {
      const a = attrs(m[0]);
      const max = Number(a.max ?? a.min ?? '1');
      last = Math.max(last, max);
      const id = max >= MAX_COL ? style.fill : style.text;
      const bare = m[0].replace(/\s*style="[^"]*"/, '').replace(/\s*\/>$/, '');
      cols.push(`${bare}${id === null ? '' : ` style="${id}"`}/>`);
    }
    if (cols.length === 0) {
      cols.push(`<col min="1" max="1" width="120" customWidth="1" style="${style.text}"/>`);
      last = 1;
    }
    if (last < MAX_COL && style.fill !== null) {
      cols.push(`<col min="${last + 1}" max="${MAX_COL}" width="${FILL_COL_WIDTH}" style="${style.fill}"/>`);
    }
    const block = `<cols>${cols.join('')}</cols>`;
    out = colsBlock ? out.replace(colsBlock[0], block) : out.replace(/<sheetData\b/, `${block}<sheetData`);
    // セル
    out = out.replace(/<c(\s[^>]*?)?(\/?)>/g, (_all, a: string | undefined, close: string) => {
      const attrsText = a ?? '';
      // 属性の並びは保ち、s だけを差し替える（無ければ足す）
      const next = /\ss="[^"]*"/.test(attrsText)
        ? attrsText.replace(/(\s)s="[^"]*"/, `$1s="${style.text}"`)
        : `${attrsText} s="${style.text}"`;
      return `<c${next}${close}>`;
    });
  }
  return out;
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

/** テーマ用のフォント・塗りつぶし・セルスタイル（cellXfs）を探し、無ければ styles.xml に追加する */
function ensureThemeStyles(
  styles: string,
  theme: BookTheme,
): { styles: string; text: number | null; fill: number | null } {
  const palette = THEMES[theme];
  const fontsMatch = /<fonts\b[^>]*>([\s\S]*?)<\/fonts>/.exec(styles);
  const fillsMatch = /<fills\b[^>]*>([\s\S]*?)<\/fills>/.exec(styles);
  const xfsMatch = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles);
  if (!fontsMatch || !fillsMatch || !xfsMatch) return { styles, text: null, fill: null };
  const items = (block: string, tag: string) =>
    [...block.matchAll(new RegExp(`<${tag}\\b[^>]*?(?:\\/>|>[\\s\\S]*?<\\/${tag}>)`, 'g'))].map((m) => m[0]);
  const fonts = items(fontsMatch[1], 'font');
  const fills = items(fillsMatch[1], 'fill');
  const xfs = items(xfsMatch[1], 'xf');
  let out = styles;

  const add = (kind: 'fonts' | 'fills' | 'cellXfs', count: number, xml: string) => {
    const block = new RegExp(`<${kind}\\b[^>]*>[\\s\\S]*?<\\/${kind}>`).exec(out)![0];
    out = out.replace(
      block,
      block
        .replace(new RegExp(`<${kind}\\b([^>]*?)count="\\d+"`), `<${kind}$1count="${count + 1}"`)
        .replace(`</${kind}>`, `${xml}</${kind}>`),
    );
  };

  let fontId = fonts.findIndex(
    (f) =>
      f.includes(`val="${THEME_FONT.name}"`) &&
      f.includes(`<sz val="${THEME_FONT.size}"/>`) &&
      f.includes(`rgb="FF${palette.text}"`),
  );
  if (fontId < 0) {
    fontId = fonts.length;
    add(
      'fonts',
      fonts.length,
      `<font><sz val="${THEME_FONT.size}"/><color rgb="FF${palette.text}"/><name val="${THEME_FONT.name}"/><family val="3"/></font>`,
    );
  }
  let fillId = fills.findIndex(
    (f) => f.includes('patternType="solid"') && f.includes(`<fgColor rgb="FF${palette.fill}"/>`),
  );
  if (fillId < 0) {
    fillId = fills.length;
    add(
      'fills',
      fills.length,
      `<fill><patternFill patternType="solid"><fgColor rgb="FF${palette.fill}"/><bgColor indexed="64"/></patternFill></fill>`,
    );
  }
  const xfIndex = (numFmtId: number): number => {
    const found = xfs.findIndex((x) => {
      const a = attrs(x.slice(0, x.indexOf('>') + 1));
      return (
        a.numFmtId === String(numFmtId) &&
        a.fontId === String(fontId) &&
        a.fillId === String(fillId) &&
        a.borderId === '0'
      );
    });
    if (found >= 0) return found;
    const index = xfs.length;
    xfs.push(`<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}"/>`);
    add(
      'cellXfs',
      index,
      `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="0" xfId="0"${numFmtId === 0 ? '' : ' applyNumberFormat="1"'} applyFont="1" applyFill="1"/>`,
    );
    return index;
  };
  // 文字列書式（@ = 49）と、データの無い範囲用の標準書式（0）
  const text = xfIndex(49);
  const fill = xfIndex(0);
  return { styles: out, text, fill };
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

export async function patchXlsx(original: Uint8Array, ops: readonly PatchOp[], theme?: BookTheme): Promise<Buffer> {
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

  // セルのスタイル。テーマが決まっていれば塗りつぶし・文字色・枠線なし・タブの色まで、無ければ従来の（等幅・文字列書式だけ）
  let textStyle: number | null = null;
  let fillStyle: number | null = null;
  if (styles && ops.some((o) => o.kind === 'write' || o.kind === 'restyle')) {
    if (theme) {
      const r = ensureThemeStyles(styles, theme);
      styles = r.styles;
      textStyle = r.text;
      fillStyle = r.fill;
    } else {
      const r = ensureTextStyle(styles);
      styles = r.styles;
      textStyle = r.index;
    }
  }
  const styleFor = (role: SheetRole | undefined): SheetStyle => ({
    text: textStyle,
    fill: fillStyle,
    tab: theme ? THEMES[theme].tab[role ?? 'code'] : null,
    hideGrid: theme !== undefined,
  });

  const relTarget = (rid: string): string | null => {
    const tag = [...rels.matchAll(/<Relationship\b[^>]*\/>/g)].map((m) => m[0]).find((t) => attrs(t).Id === rid);
    return tag ? resolveTarget(attrs(tag).Target) : null;
  };

  for (const op of ops) {
    const sheets = readSheets(workbook);
    const index = sheets.findIndex((s) => s.name.toLowerCase() === op.name.toLowerCase());
    if (op.kind === 'write') {
      const xml = worksheetXml(op.columns, styleFor(op.role), sst);
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
    } else if (op.kind === 'restyle') {
      if (index < 0 || !theme) continue;
      const part = relTarget(sheets[index].rid);
      if (!part) continue;
      const current = await read(part);
      if (current === null) continue;
      zip.file(part, restyleSheetXml(current, styleFor(op.role)));
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
