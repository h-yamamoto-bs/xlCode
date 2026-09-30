import path from 'node:path';
import type { Canon } from './canonical';
import type { TextFormat } from './encoding';
import { AGENTS_SHEET } from './constants';
import { readdir } from 'node:fs/promises';
import { excelPathError, isBookCopy, isIgnored, listSourceFiles, readTextFile } from './fsutil';
import { normalizeText } from './normalize';
import { booksInDir, exists, type BookRef, type ProjectContext } from './project';
import { classifySheet, isCodeName, sheetKey, validateFileName } from './sheetName';
import { bookState, type FileState } from './state';
import { Book } from './workbook';
import { readFile } from 'node:fs/promises';

/**
 * ファイルごとの状態。前回値（state.json）とソース側・Excel側のハッシュを比較して決める。
 *
 * | 状態            | 意味                                         |
 * |-----------------|----------------------------------------------|
 * | clean           | 両側一致                                     |
 * | excel-changed   | Excel 側だけ変更（未 Build）                 |
 * | excel-new       | Excel 側にだけ新しいシートがある             |
 * | source-changed  | ソース側だけ変更（未 Sync）                  |
 * | source-new      | ソース側にだけ新しいファイルがある           |
 * | source-deleted  | ソース側で削除され、Excel 側は未変更         |
 * | sheet-missing   | Excel 側でシートが削除された（Sync で復元）  |
 * | conflict        | 両側とも変更                                 |
 * | gone            | 両側とも存在しない（state から除去）         |
 */
export type FileStatus =
  | 'clean'
  | 'excel-changed'
  | 'excel-new'
  | 'source-changed'
  | 'source-new'
  | 'source-deleted'
  | 'sheet-missing'
  | 'conflict'
  | 'gone'
  /** VBA モード: 前回 Build したシートが無くなった（次の Build でモジュールを削除） */
  | 'removed';

export interface FileEntry {
  name: string;
  status: FileStatus;
  srcAbs: string;
  src?: Canon;
  xl?: Canon;
  /** シートの生の行（整形結果と異なればシートを書き直す） */
  xlRaw?: string[];
  sheetName?: string;
  prev?: FileState;
  /** ソースファイルの形式（無ければ新規。書き出すときは拡張子の標準を使う） */
  format?: TextFormat;
}

export interface Scan {
  book: Book;
  ref: BookRef;
  errors: string[];
  warnings: string[];
  entries: FileEntry[];
  /** DEL_ 付きシート */
  deletes: { sheet: string; fileName: string }[];
  conflictSheets: string[];
}

export function decideStatus(prev: FileState | undefined, src: Canon | undefined, xl: Canon | undefined): FileStatus {
  if (!src && !xl) return 'gone';
  if (src && xl && src.hash === xl.hash) return 'clean';
  if (!prev) {
    if (src && xl) return 'conflict';
    return src ? 'source-new' : 'excel-new';
  }
  if (src && !xl) return 'sheet-missing';
  if (!src && xl) return xl.hash === prev.hash ? 'source-deleted' : 'conflict';
  const srcChanged = src!.hash !== prev.hash;
  const xlChanged = xl!.hash !== prev.hash;
  if (srcChanged && xlChanged) return 'conflict';
  return srcChanged ? 'source-changed' : 'excel-changed';
}

export async function scanBook(ctx: ProjectContext, ref: BookRef): Promise<Scan> {
  const errors: string[] = [];
  const warnings: string[] = [];
  const books = await booksInDir(path.dirname(ref.abs));
  if (books.length > 1) errors.push(`同一ディレクトリに複数の .xlcode.xlsx があります: ${books.join(', ')}`);
  if (!(await exists(ref.dirAbs))) {
    errors.push(
      `ブックに対応するソースのフォルダがありません: ${ref.dirAbs}（ブックの置き場所の設定を確認してください）`,
    );
    return { book: await Book.load(ref.abs), ref, errors, warnings, entries: [], deletes: [], conflictSheets: [] };
  }
  const copies = (await readdir(path.dirname(ref.abs))).filter(isBookCopy);
  if (copies.length > 0) {
    warnings.push(
      `ブックの複製があります: ${copies.join(', ')}。OneDrive の同期がぶつかった可能性があります。` +
        '必要な編集が複製側に入っていないか確認し、不要なら削除してください',
    );
  }
  const longPath = excelPathError(ref.abs);
  if (longPath) warnings.push(longPath);

  const book = await Book.load(ref.abs);
  const deletes: Scan['deletes'] = [];
  const conflictSheets: string[] = [];
  const sheets = new Map<string, { name: string; lines: string[]; canon: Canon }>();
  const seen = new Map<string, string>();

  for (const name of book.sheetNames()) {
    const info = classifySheet(name, ctx.config.extraCodeNames);
    if (info.kind === 'conflict') conflictSheets.push(name);
    if (info.kind === 'unknown-reserved') warnings.push(`「${name}」は予約接頭辞「#」で始まるため無視します`);
    if (info.kind === 'agents') {
      const rootAgents = await readFile(path.join(ctx.root, AGENTS_SHEET), 'utf8').catch(() => null);
      const opts = { trimTrailingWhitespace: ctx.config.trimTrailingWhitespace };
      const sheetText = normalizeText(book.readSheet(name).lines.join('\n'), name, opts);
      if (rootAgents !== null && normalizeText(rootAgents, name, opts) !== sheetText) {
        warnings.push('Agents.md シートの変更は反映されません。ルートの Agents.md をエディタで編集してください');
      }
    }
    if (info.kind !== 'code' && info.kind !== 'delete') continue;

    const fileName = info.fileName!;
    const err = validateFileName(fileName);
    if (err) {
      errors.push(`シート名エラー: ${err}`);
      continue;
    }
    const key = sheetKey(fileName) + (info.kind === 'delete' ? '\0del' : '');
    if (seen.has(key)) {
      errors.push(`生成先が重複するシートがあります: 「${seen.get(key)}」と「${name}」`);
      continue;
    }
    seen.set(key, name);
    if (info.kind === 'delete') {
      deletes.push({ sheet: name, fileName });
      continue;
    }
    if (isIgnored(ctx.ig, ref.dirRel ? `${ref.dirRel}/${fileName}` : fileName, false)) {
      warnings.push(`「${name}」は .gitignore の対象なのでスキップします`);
      continue;
    }
    const data = book.readSheet(name);
    for (const is of data.issues) {
      errors.push(`「${name}」${is.row} 行目が文字列ではありません（${is.type}）。Excel の自動変換の可能性があります`);
    }
    if (data.hasExtraColumns) warnings.push(`「${name}」の B 列以降の内容は無視します`);
    const canon = await ctx.canon.canonical(fileName, path.join(ref.dirAbs, fileName), data.lines.join('\n'));
    if (canon.formatError) warnings.push(`「${name}」を整形できませんでした（Excel 側）: ${canon.formatError}`);
    sheets.set(sheetKey(fileName), { name, lines: data.lines, canon });
  }

  for (const d of deletes) {
    if (sheets.has(sheetKey(d.fileName))) {
      errors.push(`「${d.sheet}」と「${d.fileName}」が両方あります。どちらかを削除してください`);
    }
  }

  const sources = new Map<string, { name: string; abs: string; canon: Canon; format: TextFormat }>();
  for (const f of await listSourceFiles(ctx.root, ref.dirAbs, ctx.ig)) {
    if (!isCodeName(f.name, ctx.config.extraCodeNames)) {
      warnings.push(
        `拡張子のないファイル「${f.name}」は対象外です。コードとして扱うには設定（extraCodeNames）に追加してください`,
      );
      continue;
    }
    const err = validateFileName(f.name);
    if (err) {
      errors.push(`ファイル名エラー: ${err}`);
      continue;
    }
    const key = sheetKey(f.name);
    if (sources.has(key)) {
      errors.push(`大文字小文字だけが異なるファイルがあります: ${sources.get(key)!.name}, ${f.name}`);
      continue;
    }
    const read = await readTextFile(f.abs);
    if (read.kind === 'binary') {
      warnings.push(`「${f.name}」は扱えないためスキップします（${read.reason}）`);
      continue;
    }
    const canon = await ctx.canon.canonical(f.name, f.abs, read.text);
    if (canon.formatError) warnings.push(`「${f.name}」を整形できませんでした（ソース側）: ${canon.formatError}`);
    sources.set(key, { name: f.name, abs: f.abs, canon, format: read.format });
  }

  const prevFiles = bookState(ctx.state, ref.rel).files;
  const prevByKey = new Map(Object.entries(prevFiles).map(([n, s]) => [sheetKey(n), { name: n, s }]));
  const keys = new Set([...sheets.keys(), ...sources.keys(), ...prevByKey.keys()]);
  const deleteKeys = new Set(deletes.map((d) => sheetKey(d.fileName)));

  const entries: FileEntry[] = [];
  for (const key of [...keys].sort()) {
    if (deleteKeys.has(key)) continue;
    const s = sheets.get(key);
    const f = sources.get(key);
    const p = prevByKey.get(key);
    const name = f?.name ?? s?.name ?? p!.name;
    entries.push({
      name,
      status: decideStatus(p?.s, f?.canon, s?.canon),
      srcAbs: path.join(ref.dirAbs, name),
      src: f?.canon,
      xl: s?.canon,
      xlRaw: s?.lines,
      sheetName: s?.name,
      prev: p?.s,
      format: f?.format,
    });
  }

  return { book, ref, errors, warnings, entries, deletes, conflictSheets };
}
