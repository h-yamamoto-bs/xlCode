import { TREE_SHEET } from './constants';
import { defaultFormat, formatLabel } from './encoding';
import { bookRef, bookRootOf, openProject } from './project';
import { parseTreeVersion } from './tree';
import { scanBook, type FileStatus } from './scan';
import { vbaFileStatus, vbaOutputInfo, type VbaBookInfo } from './vba';
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
  /** VBA モードのビルド結果 */
  vba?: VbaBookInfo;
}

/** 読み取り専用で各ファイルの状態を返す（GUI の「編集中ファイル一覧」用、5.5-3） */
export async function bookStatus(root: string, bookAbs: string): Promise<BookStatus> {
  const ctx = await openProject(root);
  const ref = bookRef(root, bookAbs, bookRootOf(root, ctx.config));
  if (ctx.config.mode === 'vba') {
    const book = await Book.load(ref.abs);
    const { collect, files } = await vbaFileStatus(ctx, ref, book);
    return {
      book: ref.rel,
      errors: collect.errors,
      warnings: collect.warnings,
      files,
      conflictSheets: [],
      deletes: [],
      treeVersion: book.hasSheet(TREE_SHEET) ? parseTreeVersion(book.readSheet(TREE_SHEET).lines[0]) : null,
      vba: await vbaOutputInfo(ctx, ref),
    };
  }
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
  };
}
