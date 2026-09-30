import { mkdir, readFile } from 'node:fs/promises';
import path from 'node:path';
import { atomicWrite, backup } from './atomic';
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
  /** VBA モード: 前回 Build で生成した .xlsm の SHA-256（その後に直接変更されたかの判定に使う） */
  outputHash?: string;
}

/** .xlcode/state.json（3.9）。books のキーはプロジェクトルートからのブックの相対パス */
export interface State {
  version: 1;
  books: Record<string, BookState>;
}

export class StateError extends Error {}

/**
 * state.json を読む。無ければ空の状態を返す。
 * 壊れている場合は例外にする。黙って初期化すると前回値が失われ、全ファイルが衝突扱いになるため。
 */
export async function loadState(root: string): Promise<State> {
  const file = path.join(root, XLCODE_DIR, STATE_FILE);
  let raw: string;
  try {
    raw = await readFile(file, 'utf8');
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code === 'ENOENT') return { version: 1, books: {} };
    throw e;
  }
  const broken = (why: string) =>
    new StateError(
      `${XLCODE_DIR}/${STATE_FILE} が壊れています（${why}）。` +
        `${XLCODE_DIR}/${STATE_FILE}.bak があれば ${STATE_FILE} にコピーして戻してください。` +
        `無ければ ${STATE_FILE} を削除すると初期化できますが、次の Build / Sync で両側に差があるファイルは衝突扱いになります。`,
    );
  let s: State;
  try {
    s = JSON.parse(raw) as State;
  } catch {
    throw broken('JSON として読めません');
  }
  if (s?.version !== 1 || typeof s.books !== 'object' || s.books === null) throw broken('形式が違います');
  for (const [book, b] of Object.entries(s.books)) {
    if (typeof b?.files !== 'object' || b.files === null) throw broken(`${book} の形式が違います`);
    for (const [name, f] of Object.entries(b.files)) {
      if (typeof f?.hash !== 'string' || typeof f.lines !== 'number' || typeof f.chars !== 'number') {
        throw broken(`${book} の ${name} の形式が違います`);
      }
    }
  }
  return s;
}

export async function saveState(root: string, state: State): Promise<void> {
  await mkdir(path.join(root, XLCODE_DIR), { recursive: true });
  const file = path.join(root, XLCODE_DIR, STATE_FILE);
  await backup(file);
  await atomicWrite(file, JSON.stringify(state, null, 2) + '\n');
}

export function bookState(state: State, bookRel: string): BookState {
  return (state.books[bookRel] ??= { files: {} });
}
