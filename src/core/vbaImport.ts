import { copyFile, mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { LOCAL_AGENTS_SHEET, REFS_SHEET } from './constants';
import { excelPathError } from './fsutil';
import { bookPathFor, bookRef, bookRootOf, booksInDir, exists, openProject } from './project';
import { classifySheet, validateFileName } from './sheetName';
import { bookState, saveState } from './state';
import { applyTreeToBook, computeTree, readLocalAgents, readRootAgents } from './tree';
import { VBOM_HELP, type VbaRunResult } from './vba';
import { FORM_CONTROLS } from './vbaForm';
import { renderRefs, type VbaReference } from './vbaRefs';
import { Book } from './workbook';
import type { CreateBookResult } from './init';

/**
 * 既存の Excel ツール（.xlsm など）から編集用ブックを作る（VBA モード）。
 *
 * 画面に出さない Excel でツールのコピーを開き、
 * - シート（UI・データ）は .xlsx として保存し直したものをそのまま使う（図形・ボタン・書式を残す）
 * - VBA のモジュールはシート（Module1.bas など）にする。フォームは配置を Begin 〜 End の形に書き起こす
 * - 参照設定は #refs シートにする
 * 元のツールは変更しない。
 */

export type ImportValue = string | number | boolean;

export interface ImportedProp {
  name: string;
  value: ImportValue;
}

export interface ImportedControl {
  /** TypeName（CommandButton など。MultiPage のページは Page） */
  type: string;
  name: string;
  /** 親（フォーム・Frame・Page）の名前 */
  parent: string;
  props: ImportedProp[];
  /** 画像があった（取り込めない） */
  picture?: boolean;
}

export interface ImportedForm {
  props: ImportedProp[];
  controls: ImportedControl[];
  picture?: boolean;
}

export interface ImportedModule {
  name: string;
  /** 1: 標準 2: クラス 3: フォーム 100: ブック・シート */
  type: number;
  code: string[];
  /** ブックのコード */
  workbook?: boolean;
  /** シートのコードの場合のシート名 */
  sheet?: string;
  form?: ImportedForm;
}

export interface ImportJob {
  /** 取り込むツール（のコピー） */
  input: string;
  /** シートだけを残した .xlsx の保存先 */
  output: string;
}

export interface ImportResult extends VbaRunResult {
  modules?: ImportedModule[];
  references?: (VbaReference & { name?: string })[];
  warnings?: string[];
}

export type VbaImporter = (job: ImportJob) => Promise<ImportResult>;

/** Excel のツールとして取り込めるファイル */
export const TOOL_EXTENSIONS = ['xlsm', 'xlsb', 'xls', 'xlsx', 'xltm'];

const COLOR_PROPS = new Set(['ForeColor', 'BackColor', 'BorderColor']);

function formatValue(name: string, v: ImportValue): string {
  if (typeof v === 'boolean') return v ? 'True' : 'False';
  if (typeof v === 'number') {
    if (COLOR_PROPS.has(name) && Number.isInteger(v))
      return `&H${(v >>> 0).toString(16).toUpperCase().padStart(8, '0')}&`;
    return String(v);
  }
  return `"${v.replace(/"/g, '""')}"`;
}

/** フォームの配置を、シートに書く形（Begin 〜 End）にする */
export function renderForm(name: string, form: ImportedForm): { lines: string[]; warnings: string[] } {
  const warnings: string[] = [];
  const known = new Set([...Object.keys(FORM_CONTROLS), 'Page']);
  const props = (list: ImportedProp[], indent: string, owner: string): string[] =>
    list.flatMap((p) => {
      let v = p.value;
      if (typeof v === 'string' && /[\r\n]/.test(v)) {
        warnings.push(`${name} の ${owner}.${p.name} の改行は空白にしました`);
        v = v.replace(/\r\n|\r|\n/g, ' ');
      }
      return [`${indent}${p.name} = ${formatValue(p.name, v)}`];
    });
  const names = new Set(form.controls.map((c) => c.name.toLowerCase()));
  const children = new Map<string, ImportedControl[]>();
  for (const c of form.controls) {
    // 親が見つからないものはフォームの直下に置く
    const parent = names.has(c.parent.toLowerCase()) ? c.parent.toLowerCase() : '';
    if (!children.has(parent)) children.set(parent, []);
    children.get(parent)!.push(c);
  }
  const render = (c: ImportedControl, depth: number): string[] => {
    const indent = '   '.repeat(depth);
    if (!known.has(c.type)) {
      warnings.push(`${name} の ${c.name}（${c.type}）は取り込めないコントロールのため省きました`);
      return [`${indent}' 取り込めないコントロール: ${c.name}（${c.type}）`];
    }
    if (c.picture) warnings.push(`${name} の ${c.name} の画像は取り込めません。Build 後に VBE で設定し直してください`);
    return [
      `${indent}Begin ${c.type} ${c.name}`,
      ...props(c.props, `${indent}   `, c.name),
      ...(children.get(c.name.toLowerCase()) ?? []).flatMap((k) => render(k, depth + 1)),
      `${indent}End`,
    ];
  };
  if (form.picture) warnings.push(`${name} の背景画像は取り込めません`);
  const lines = [
    `Begin UserForm ${name}`,
    ...props(form.props, '   ', name),
    ...(children.get('') ?? []).flatMap((c) => render(c, 1)),
    'End',
  ];
  return { lines, warnings };
}

function isEmptyCode(code: readonly string[]): boolean {
  return code.every((l) => l.trim() === '' || /^option explicit$/i.test(l.trim()));
}

/** 取り込んだモジュールをシートにする。シート名と内容の組を返す */
export function modulesToSheets(
  modules: readonly ImportedModule[],
  uiSheets: readonly string[],
): { sheets: { name: string; lines: string[] }[]; warnings: string[] } {
  const sheets: { name: string; lines: string[] }[] = [];
  const warnings: string[] = [];
  const trim = (code: string[]) => {
    const out = [...code];
    while (out.length > 0 && out[out.length - 1].trim() === '') out.pop();
    return out;
  };
  for (const m of modules) {
    let sheet: string;
    let lines = trim(m.code);
    if (m.type === 1) sheet = `${m.name}.bas`;
    else if (m.type === 2) sheet = `${m.name}.cls`;
    else if (m.type === 3) {
      sheet = `${m.name}.frm`;
      const form = renderForm(m.name, m.form ?? { props: [], controls: [] });
      warnings.push(...form.warnings);
      lines = [...form.lines, '', ...lines];
    } else if (m.type === 100) {
      // 中身の無いシート・ブックのコードはシートにしない
      if (isEmptyCode(lines)) continue;
      if (m.workbook) sheet = 'ThisWorkbook.cls';
      else if (m.sheet && uiSheets.includes(m.sheet) && !validateFileName(`${m.sheet}.cls`)) sheet = `${m.sheet}.cls`;
      else sheet = `${m.name}.cls`;
    } else {
      warnings.push(`${m.name} は取り込めない種類のモジュールです（種類 ${m.type}）`);
      continue;
    }
    const err = validateFileName(sheet);
    if (err) {
      warnings.push(`${m.name} を取り込めません: ${err}。VBE で名前を短くしてから取り込み直してください`);
      continue;
    }
    if (sheets.some((s) => s.name.toLowerCase() === sheet.toLowerCase()) || uiSheets.includes(sheet)) {
      warnings.push(`${sheet} という名前が重複するため ${m.name} を取り込めません`);
      continue;
    }
    sheets.push({ name: sheet, lines });
  }
  return { sheets, warnings };
}

/**
 * 既存のツールから、ディレクトリの編集用ブックを作る。
 * importer は Excel を操作する処理（Windows 以外では null）
 */
export async function createBookFromTool(
  root: string,
  dirAbs: string,
  toolFile: string,
  importer: VbaImporter | null,
): Promise<CreateBookResult> {
  const ctx = await openProject(root);
  if (ctx.config.mode !== 'vba') throw new Error('Excel ツールの取り込みは VBA モードのプロジェクトで使えます');
  const ext = path.extname(toolFile).slice(1).toLowerCase();
  if (!TOOL_EXTENSIONS.includes(ext)) {
    throw new Error(`取り込めるのは ${TOOL_EXTENSIONS.map((e) => `.${e}`).join(' / ')} です: ${toolFile}`);
  }
  if (!(await exists(toolFile))) throw new Error(`ファイルがありません: ${toolFile}`);
  const bookRoot = bookRootOf(root, ctx.config);
  const bookAbs = bookPathFor(root, bookRoot, dirAbs);
  const existing = await booksInDir(path.dirname(bookAbs));
  if (existing.length > 0) throw new Error(`既にブックがあります: ${existing.join(', ')}`);
  const longPath = excelPathError(bookAbs);
  if (longPath) throw new Error(longPath);
  if (!importer) throw new Error('Excel ツールの取り込みには Windows とデスクトップ版 Excel が必要です');

  const work = await mkdtemp(path.join(tmpdir(), 'xlcode-import-'));
  try {
    // 元のツールには触らない（開いていても読めるよう、コピーしてから開く）
    const input = path.join(work, `tool.${ext}`);
    await copyFile(toolFile, input);
    const output = path.join(work, 'sheets.xlsx');
    const res = await importer({ input, output });
    if (!res.ok) {
      const msg = [`Excel での読み取りに失敗しました: ${res.error ?? '原因不明'}`];
      if (res.errorKind === 'vbom') msg.push(VBOM_HELP);
      throw new Error(msg.join('\n'));
    }
    const bytes = await readFile(output);
    const book = await Book.fromBuffer(bytes, toolFile);
    const skipped = [...(res.warnings ?? [])];
    const ui = book.sheetNames();
    for (const s of ui) {
      if (classifySheet(s, ctx.config.extraCodeNames).kind !== 'other') {
        skipped.push(
          `シート「${s}」は拡張子付き（または # で始まる）の名前のため、Build でビルド結果から除かれます。名前を変えてください`,
        );
      }
    }
    const conv = modulesToSheets(res.modules ?? [], ui);
    skipped.push(...conv.warnings);

    await mkdir(path.dirname(bookAbs), { recursive: true });
    const ref = bookRef(root, bookAbs, bookRoot);
    const localAgents = path.join(dirAbs, LOCAL_AGENTS_SHEET);
    if (!(await exists(localAgents))) {
      await writeFile(
        localAgents,
        `# LocalAgents.md\n\n- ${path.basename(toolFile)} から取り込んだツール\n- UI シートの構成（どのシートの何のセルに何があるか、ボタンに登録するマクロ名など）をここに書く\n`,
      );
    }
    // 取り込むブック自身も #tree に載るよう、先に保存してからツリーを取る
    await book.save(bookAbs);
    applyTreeToBook(book, await computeTree(root, ctx.ig), await readRootAgents(root), await readLocalAgents(dirAbs));
    book.writeLines(REFS_SHEET, renderRefs(res.references ?? []));
    for (const s of conv.sheets) book.writeLines(s.name, s.lines);
    await book.save(bookAbs);

    const bs = bookState(ctx.state, ref.rel);
    bs.files = {};
    delete bs.outputHash;
    await saveState(root, ctx.state);
    return { book: ref.rel, sheets: conv.sheets.map((s) => s.name), skipped };
  } finally {
    await rm(work, { recursive: true, force: true });
  }
}
