import { bookRef, openProject } from './project';
import { scanBook, type FileStatus } from './scan';

export interface BookStatus {
  book: string;
  errors: string[];
  warnings: string[];
  files: { name: string; status: FileStatus }[];
  conflictSheets: string[];
  deletes: string[];
}

/** 読み取り専用で各ファイルの状態を返す（GUI の「編集中ファイル一覧」用、5.5-3） */
export async function bookStatus(root: string, bookAbs: string): Promise<BookStatus> {
  const ctx = await openProject(root);
  const ref = bookRef(root, bookAbs);
  const scan = await scanBook(ctx, ref);
  return {
    book: ref.rel,
    errors: scan.errors,
    warnings: scan.warnings,
    files: scan.entries.filter((e) => e.status !== 'gone').map((e) => ({ name: e.name, status: e.status })),
    conflictSheets: scan.conflictSheets,
    deletes: scan.deletes.map((d) => d.sheet),
  };
}
