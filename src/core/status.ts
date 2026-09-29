import { TREE_SHEET } from './constants';
import { diffLines, summarize, toHunks, type Hunk } from './diff';
import { textToLines } from './normalize';
import { defaultFormat, formatLabel } from './encoding';
import { bookRef, bookRootOf, openProject } from './project';
import { parseTreeVersion } from './tree';
import { scanBook, type FileStatus } from './scan';
import { undoInfo, type UndoInfo } from './undo';

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
}

/** 読み取り専用で各ファイルの状態を返す（GUI の「編集中ファイル一覧」用、5.5-3） */
export async function bookStatus(root: string, bookAbs: string): Promise<BookStatus> {
  const ctx = await openProject(root);
  const ref = bookRef(root, bookAbs, bookRootOf(root, ctx.config));
  const scan = await scanBook(ctx, ref);
  return {
    book: ref.rel,
    errors: scan.errors,
    warnings: scan.warnings,
    files: scan.entries
      .filter((e) => e.status !== 'gone')
      .map((e) => ({ name: e.name, status: e.status, format: formatLabel(e.format ?? defaultFormat(e.name)) })),
    conflictSheets: scan.conflictSheets,
    deletes: scan.deletes.map((d) => d.sheet),
    treeVersion: scan.book.hasSheet(TREE_SHEET) ? parseTreeVersion(scan.book.readSheet(TREE_SHEET).lines[0]) : null,
    undo: await undoInfo(root, ref.rel),
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
