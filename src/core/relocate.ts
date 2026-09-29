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

/**
 * ブックの置き場所を変える。既存のブックは同じフォルダ構成のまま新しい置き場所へ移動する。
 * newRoot が空ならソースの中に戻す。
 * 開いているブックや、移動先に同名のファイルがあるときは、何も動かさずに止める。
 */
export async function relocateBooks(root: string, newRoot: string | null): Promise<RelocateResult> {
  const config = await loadConfig(root);
  const from = bookRootOf(root, config);
  const target = newRoot?.trim() ? path.resolve(newRoot.trim()) : null;
  const to = target ?? path.resolve(root);
  if (target) {
    if (!path.isAbsolute(newRoot!.trim())) throw new Error('ブックの置き場所はフルパスで指定してください');
    const inRoot = path.relative(path.resolve(root), target);
    if (inRoot !== '' && !inRoot.startsWith('..') && !path.isAbsolute(inRoot)) {
      throw new Error('ブックの置き場所をソースのフォルダの中にはできません。OneDrive 内の専用フォルダを指定してください');
    }
    if ((await exists(target)) && !(await stat(target)).isDirectory()) throw new Error(`フォルダではありません: ${target}`);
  }

  const books = path.resolve(from) === path.resolve(to) ? [] : await findBooks(root, await loadIgnore(root), from);
  const plan = books.map((abs) => ({ abs, rel: toPosixRel(from, abs), dest: path.join(to, path.relative(from, abs)) }));
  for (const p of plan) {
    const open = await checkBookOpen(p.abs);
    if (open.open) throw new Error(`${p.rel} が開かれています。Excel を閉じてから実行してください（${open.reason}）`);
    if (await exists(p.dest)) throw new Error(`移動先に同じ名前のファイルがあります: ${p.dest}`);
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
