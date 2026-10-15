import { TREE_SHEET } from './constants';
import { diffLines, summarize, toHunks, type Hunk } from './diff';
import { textToLines } from './normalize';
import { defaultFormat, formatLabel } from './encoding';
import { readTextFile } from './fsutil';
import { classifySheet } from './sheetName';
import path from 'node:path';
import { bookRef, bookRootOf, openProject } from './project';
import { parseTreeVersion } from './tree';
import { scanBook, type FileStatus } from './scan';
import { undoInfo, type UndoInfo } from './undo';
import { themeState, type ThemeState } from './applyTheme';
import { modulesFromBook, vbaOutputInfo, type VbaBookInfo } from './vba';
import { Book } from './workbook';

export interface BookStatus {
  book: string;
  errors: string[];
  warnings: string[];
  files: { name: string; status: FileStatus; format?: string }[];
  conflictSheets: string[];
  deletes: string[];
  /** #tree 1行目に記録されたバージョン */
  treeVersion: string | null;
  /** 元に戻せる直前の Build / Sync */
  undo: UndoInfo | null;
  /** VBA モードのビルド結果 */
  vba?: VbaBookInfo;
  /** ブックのテーマ（背景・文字色・枠線）が今の設定どおりか */
  theme: ThemeState;
}

/** 読み取り専用で各ファイルの状態を返す（GUI の「編集中ファイル一覧」用、5.5-3） */
export async function bookStatus(root: string, bookAbs: string): Promise<BookStatus> {
  const ctx = await openProject(root);
  const ref = bookRef(root, bookAbs, bookRootOf(root, ctx.config));
  const scan = await scanBook(ctx, ref);
  const errors = [...scan.errors];
  const warnings = [...scan.warnings];
  let vba: VbaBookInfo | undefined;
  if (ctx.config.mode === 'vba') {
    // フォームの書き方などの間違いを、Build の前に知らせる
    const m = modulesFromBook(scan.book, ctx.config.extraCodeNames);
    errors.push(...m.errors);
    warnings.push(...m.warnings);
    vba = await vbaOutputInfo(ctx, ref);
  }
  return {
    book: ref.rel,
    errors,
    warnings,
    vba,
    files: scan.entries
      .filter((e) => e.status !== 'gone')
      .map((e) => ({ name: e.name, status: e.status, format: formatLabel(e.format ?? defaultFormat(e.name)) })),
    conflictSheets: scan.conflictSheets,
    deletes: scan.deletes.map((d) => d.sheet),
    treeVersion: scan.book.hasSheet(TREE_SHEET) ? parseTreeVersion(scan.book.readSheet(TREE_SHEET).lines[0]) : null,
    undo: await undoInfo(root, ref.rel),
    theme: themeState(scan.book, ctx.bookTheme, ctx.config.extraCodeNames),
  };
}

export interface FileDiff {
  name: string;
  status: FileStatus;
  /** 差分の向き。old 側が現在の反映先、new 側がこれから反映される内容 */
  from: 'excel' | 'source';
  to: 'excel' | 'source';
  hunks: Hunk[];
  added: number;
  removed: number;
  /** 全体の行数（old / new） */
  oldLines: number;
  newLines: number;
}

/**
 * 変更のあるファイルごとの差分（画面の「変更を確認」用）。
 * Excel 側の変更は Build でソースがどう変わるか、ソース側の変更は Sync でシートがどう変わるかを示す。
 * 両側変更は「ソース側 → Excel 側」の向きで、どこが食い違うかを示す。
 */
export async function bookDiff(root: string, bookAbs: string): Promise<FileDiff[]> {
  const ctx = await openProject(root);
  const ref = bookRef(root, bookAbs, bookRootOf(root, ctx.config));
  const scan = await scanBook(ctx, ref);
  const out: FileDiff[] = [];
  for (const e of scan.entries) {
    if (e.status === 'clean' || e.status === 'gone') continue;
    const excelFirst = EXCEL_TO_SOURCE.includes(e.status);
    const [from, to]: ['excel' | 'source', 'excel' | 'source'] = excelFirst ? ['source', 'excel'] : ['excel', 'source'];
    const oldText = (from === 'excel' ? e.xl : e.src)?.text;
    const newText = (to === 'excel' ? e.xl : e.src)?.text;
    const oldLines = oldText === undefined ? [] : textToLines(oldText);
    const newLines = newText === undefined ? [] : textToLines(newText);
    const lines = diffLines(oldLines, newLines);
    out.push({
      name: e.name,
      status: e.status,
      from,
      to,
      hunks: toHunks(lines),
      ...summarize(lines),
      oldLines: oldLines.length,
      newLines: newLines.length,
    });
  }
  return out;
}

/** 差分を「ソース側（現在）→ Excel 側（これから）」の向きで見せる状態 */
const EXCEL_TO_SOURCE: FileStatus[] = ['excel-changed', 'excel-new', 'conflict'];

export interface FileContents {
  name: string;
  /** Excel 側（シートの A 列）。シートが無ければ null */
  sheet: string[] | null;
  /** ソース側（ファイルの内容の行）。ファイルが無ければ null */
  source: string[] | null;
  /** ソースファイルの形式（無ければ拡張子の標準） */
  format: string;
}

/**
 * 1 ファイルの Excel 側とソース側の内容をそのまま返す（右のパネルの「内容」用。整形はしない）。
 * コードのシート・ファイルだけが対象（UI 用シートなどは読まない）
 */
export async function fileContents(root: string, bookAbs: string, name: string): Promise<FileContents> {
  const ctx = await openProject(root);
  const ref = bookRef(root, bookAbs, bookRootOf(root, ctx.config));
  const book = await Book.load(ref.abs);
  const sheetName = book.sheetNames().find((n) => {
    const info = classifySheet(n, ctx.config.extraCodeNames);
    return info.kind === 'code' && info.fileName!.toLowerCase() === name.toLowerCase();
  });
  const sheet = sheetName ? book.readSheet(sheetName).lines : null;
  let source: string[] | null = null;
  let format = formatLabel(defaultFormat(name));
  if (/[\\/]/.test(name)) throw new Error(`ファイル名が不正です: ${name}`);
  try {
    const read = await readTextFile(path.join(ref.dirAbs, name));
    if (read.kind !== 'binary') {
      source = textToLines(read.text.replace(/\r\n?/g, '\n'));
      format = formatLabel(read.format);
    }
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code !== 'ENOENT') throw e;
  }
  return { name, sheet, source, format };
}
