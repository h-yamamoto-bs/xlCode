import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { STATE_FILE, XLCODE_DIR } from './constants';

/** 前回 Build / Sync 完了時点の正準テキストの情報（No.10: 保存はこの2箇所のみ） */
export interface FileState {
  hash: string;
  lines: number;
  chars: number;
}

export interface BookState {
  files: Record<string, FileState>;
  lastBuildAt?: string;
  lastSyncAt?: string;
}

/** .xlcode/state.json（3.9）。books のキーはプロジェクトルートからのブックの相対パス */
export interface State {
  version: 1;
  books: Record<string, BookState>;
}

export async function loadState(root: string): Promise<State> {
  try {
    const s = JSON.parse(await readFile(path.join(root, XLCODE_DIR, STATE_FILE), 'utf8')) as State;
    if (s.version === 1 && s.books) return s;
  } catch {
    // 未作成
  }
  return { version: 1, books: {} };
}

export async function saveState(root: string, state: State): Promise<void> {
  await mkdir(path.join(root, XLCODE_DIR), { recursive: true });
  await writeFile(path.join(root, XLCODE_DIR, STATE_FILE), JSON.stringify(state, null, 2) + '\n');
}

export function bookState(state: State, bookRel: string): BookState {
  return (state.books[bookRel] ??= { files: {} });
}
