import { createHash } from 'node:crypto';
import { copyFile, mkdir, readFile, rm } from 'node:fs/promises';
import path from 'node:path';
import { atomicWrite } from './atomic';
import { XLCODE_DIR } from './constants';
import { checkBookOpen } from './lock';
import { exists, type BookRef } from './project';
import { bookState, loadState, saveState, type BookState } from './state';

export const UNDO_DIR = 'undo';

/** 画面に出す「元に戻せる操作」の情報 */
export interface UndoInfo {
  label: 'Build' | 'Sync';
  /** 操作した時刻（ISO 8601） */
  at: string;
  /** 元に戻すソースファイル（Build が書き出した・削除したもの） */
  files: string[];
  /** VBA モード: 元に戻すビルド結果（.xlsm）のファイル名 */
  output?: string;
}

interface Manifest extends UndoInfo {
  version: 1;
  book: string;
  /** 各ファイルが操作前に存在したか（false なら元に戻すときに削除する） */
  existed: Record<string, boolean>;
  /** 操作前の state.json のこのブックの項目（無ければ null） */
  state: BookState | null;
  /** VBA モード: ビルド結果（ルートからの相対パス）と、操作前に存在したか */
  outputFile?: { rel: string; existed: boolean };
}

export interface SnapshotOptions {
  /** VBA モード: Build が作り直すビルド結果（.xlsm）の絶対パス。これも控えて元に戻せるようにする */
  output?: string;
}

export interface UndoResult {
  label: UndoInfo['label'];
  at: string;
  /** 書き戻したファイル */
  restored: string[];
  /** 削除したファイル（操作前には無かったもの） */
  removed: string[];
}

function undoDir(root: string, bookRel: string): string {
  const key = createHash('sha1').update(bookRel).digest('hex').slice(0, 12);
  return path.join(root, XLCODE_DIR, UNDO_DIR, key);
}

async function readManifest(dir: string): Promise<Manifest | null> {
  try {
    const m = JSON.parse(await readFile(path.join(dir, 'manifest.json'), 'utf8')) as Manifest;
    return m?.version === 1 ? m : null;
  } catch {
    return null;
  }
}

/**
 * Build / Sync が書き込む直前に、書き換える前の状態を保存する（直前の操作 1 回分だけ）。
 * ブック・書き換えるソースファイル・state.json のこのブックの項目を控える。
 * manifest.json を最後に書くので、途中で失敗した保存は「元に戻せる操作」として扱われない。
 */
export async function takeSnapshot(
  root: string,
  ref: BookRef,
  label: UndoInfo['label'],
  stateBefore: BookState | null,
  fileNames: readonly string[],
  opts: SnapshotOptions = {},
): Promise<void> {
  const dir = undoDir(root, ref.rel);
  await rm(dir, { recursive: true, force: true });
  await mkdir(path.join(dir, 'files'), { recursive: true });
  await copyFile(ref.abs, path.join(dir, 'book.xlsx'));
  const existed: Record<string, boolean> = {};
  for (const name of fileNames) {
    const abs = path.join(ref.dirAbs, name);
    existed[name] = await exists(abs);
    if (existed[name]) await copyFile(abs, path.join(dir, 'files', name));
  }
  let outputFile: Manifest['outputFile'];
  if (opts.output) {
    const outExists = await exists(opts.output);
    if (outExists) await copyFile(opts.output, path.join(dir, 'output.bin'));
    outputFile = { rel: path.relative(root, opts.output).split(path.sep).join('/'), existed: outExists };
  }
  const manifest: Manifest = {
    version: 1,
    book: ref.rel,
    label,
    at: new Date().toISOString(),
    files: [...fileNames],
    existed,
    state: stateBefore,
    ...(outputFile ? { output: path.basename(outputFile.rel), outputFile } : {}),
  };
  await atomicWrite(path.join(dir, 'manifest.json'), JSON.stringify(manifest, null, 2) + '\n');
}

/** 元に戻せる操作があれば、その情報を返す */
export async function undoInfo(root: string, bookRel: string): Promise<UndoInfo | null> {
  const m = await readManifest(undoDir(root, bookRel));
  if (!m || m.book !== bookRel) return null;
  return { label: m.label, at: m.at, files: m.files, ...(m.output ? { output: m.output } : {}) };
}

/**
 * 直前の Build / Sync を元に戻す。ソースファイル・ブック・state.json をその操作の直前の状態に戻し、
 * 控えを削除する（戻せるのは 1 回だけ）。操作の後に加えた編集も、対象ファイルについては失われる。
 */
export async function undoLast(root: string, ref: BookRef): Promise<UndoResult> {
  const dir = undoDir(root, ref.rel);
  const m = await readManifest(dir);
  if (!m || m.book !== ref.rel) throw new Error(`元に戻せる操作がありません: ${ref.rel}`);
  const open = await checkBookOpen(ref.abs);
  if (open.open) throw new Error(`${ref.rel} が開かれています。Excel を閉じてから元に戻してください（${open.reason}）`);
  const outAbs = m.outputFile ? path.join(root, ...m.outputFile.rel.split('/')) : null;
  if (outAbs) {
    const outOpen = await checkBookOpen(outAbs);
    if (outOpen.open) {
      throw new Error(
        `ビルド結果 ${path.basename(outAbs)} が開かれています。Excel を閉じてから元に戻してください（${outOpen.reason}）`,
      );
    }
  }

  const restored: string[] = [];
  const removed: string[] = [];
  if (outAbs && m.outputFile) {
    const name = path.basename(outAbs);
    if (m.outputFile.existed) {
      await atomicWrite(outAbs, await readFile(path.join(dir, 'output.bin')));
      restored.push(name);
    } else {
      await rm(outAbs, { force: true });
      removed.push(name);
    }
  }
  for (const name of m.files) {
    const abs = path.join(ref.dirAbs, name);
    if (m.existed[name]) {
      await atomicWrite(abs, await readFile(path.join(dir, 'files', name)));
      restored.push(name);
    } else {
      await rm(abs, { force: true });
      removed.push(name);
    }
  }
  await atomicWrite(ref.abs, await readFile(path.join(dir, 'book.xlsx')));
  const state = await loadState(root);
  if (m.state) state.books[ref.rel] = m.state;
  else delete state.books[ref.rel];
  // 戻した後の state を確実に持つ（bookState は無ければ作る）
  bookState(state, ref.rel);
  await saveState(root, state);
  await rm(dir, { recursive: true, force: true });
  return { label: m.label, at: m.at, restored, removed };
}
