import { copyFile, mkdir, rm, stat } from 'node:fs/promises';
import path from 'node:path';
import { loadConfig, saveConfig } from './config';
import { findBooks, loadIgnore, toPosixRel } from './fsutil';
import { checkBookOpen } from './lock';
import { bookRootOf, exists } from './project';

export interface RelocateResult {
  /** 移動したブック（置き場所からの相対パス） */
  moved: string[];
  bookRoot: string | null;
}

export interface RelocateOptions {
  /**
   * 既存のブックを新しい置き場所へ移動するか（既定: true）。
   * false なら設定だけを変える（利用者が自分でブックを移した・OneDrive のフォルダ名を変えたとき用。
   * 元の場所のブックはそのまま残る）
   */
  move?: boolean;
  /** 一時的なロック（OneDrive やウイルス対策が掴んでいる）を待つ合計時間（ミリ秒） */
  retryMs?: number;
}

/** 移動の計画を立てる前に、置き場所として使えるかを確かめる。使えなければ理由を返す */
export async function validateBookRoot(root: string, newRoot: string | null): Promise<string | null> {
  const target = newRoot?.trim() ? path.resolve(newRoot.trim()) : null;
  if (!target) return null;
  if (!path.isAbsolute(newRoot!.trim())) return 'ブックの置き場所はフルパスで指定してください';
  const inRoot = path.relative(path.resolve(root), target);
  if (inRoot !== '' && !inRoot.startsWith('..') && !path.isAbsolute(inRoot)) {
    return 'ブックの置き場所をソースのフォルダの中にはできません。OneDrive 内の専用フォルダを指定してください';
  }
  if ((await exists(target)) && !(await stat(target)).isDirectory()) return `フォルダではありません: ${target}`;
  return null;
}

/**
 * ブックが開かれていないかを、短く再試行しながら確かめる。
 * Excel のロックファイルがあれば本当に開かれているので再試行しない。
 * 書き込み用に開けないだけ（EBUSY など）は、OneDrive の同期やウイルス対策が一瞬掴んでいることがあるため待つ
 */
async function waitClosed(abs: string, retryMs: number): Promise<{ open: boolean; reason?: string }> {
  const deadline = Date.now() + retryMs;
  for (;;) {
    const open = await checkBookOpen(abs);
    if (!open.open) return open;
    if (open.reason?.includes('ロックファイル') || Date.now() >= deadline) return open;
    await new Promise((r) => setTimeout(r, 250));
  }
}

/**
 * ブックの置き場所を変える。既存のブックは同じフォルダ構成のまま新しい置き場所へ移動する。
 * newRoot が空ならソースの中に戻す。
 * 開いているブックや、移動先に同名のファイルがあるときは、何も動かさずに止める。
 * 移動先に同じブックが既にあり、元の場所に無いもの（移動済み）は、そのまま採用する。
 */
export async function relocateBooks(
  root: string,
  newRoot: string | null,
  opts: RelocateOptions = {},
): Promise<RelocateResult> {
  const config = await loadConfig(root);
  const from = bookRootOf(root, config);
  const target = newRoot?.trim() ? path.resolve(newRoot.trim()) : null;
  const to = target ?? path.resolve(root);
  const problem = await validateBookRoot(root, newRoot);
  if (problem) throw new Error(problem);

  const move = opts.move ?? true;
  const books =
    !move || path.resolve(from) === path.resolve(to) ? [] : await findBooks(root, await loadIgnore(root), from);
  const plan = books.map((abs) => ({ abs, rel: toPosixRel(from, abs), dest: path.join(to, path.relative(from, abs)) }));
  const problems: string[] = [];
  for (const p of plan) {
    const open = await waitClosed(p.abs, opts.retryMs ?? 2000);
    if (open.open) problems.push(`${p.rel} が開かれています。Excel を閉じてから実行してください（${open.reason}）`);
    else if (await exists(p.dest)) problems.push(`移動先に同じ名前のファイルがあります: ${p.dest}`);
  }
  if (problems.length > 0) {
    throw new Error(
      [
        '何も動かしていません。次の問題を解決してからもう一度実行してください。',
        ...problems,
        '（自分で移動を済ませてある場合は「移動せずに変更」を選んでください）',
      ].join('\n'),
    );
  }
  const moved: string[] = [];
  for (const p of plan) {
    await mkdir(path.dirname(p.dest), { recursive: true });
    // ドライブをまたいでも動くよう、コピーしてから消す
    await copyFile(p.abs, p.dest);
    await rm(p.abs);
    moved.push(p.rel);
  }
  if (target) await mkdir(target, { recursive: true });
  await saveConfig(root, { ...config, bookRoot: target ?? undefined });
  return { moved, bookRoot: target };
}
