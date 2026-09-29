import { access, readdir } from 'node:fs/promises';
import path from 'node:path';
import type { Ignore } from 'ignore';
import { Canonicalizer } from './canonical';
import { loadConfig, type XlcodeConfig } from './config';
import { BOOK_SUFFIX } from './constants';
import { isBookFile, loadIgnore, toPosixRel } from './fsutil';
import { loadState, type State } from './state';

export interface ProjectContext {
  root: string;
  config: XlcodeConfig;
  ig: Ignore;
  state: State;
  canon: Canonicalizer;
}

export async function openProject(root: string): Promise<ProjectContext> {
  const config = await loadConfig(root);
  const state = await loadState(root);
  const known = new Set<string>();
  for (const b of Object.values(state.books)) for (const f of Object.values(b.files)) known.add(f.hash);
  return { root, config, ig: await loadIgnore(root), state, canon: new Canonicalizer(config, known) };
}

export interface BookRef {
  /** ブックの絶対パス */
  abs: string;
  /** ブックの置き場所からの相対パス（state.json のキー。置き場所を変えても同じ） */
  rel: string;
  /** 担当するソースのディレクトリ */
  dirAbs: string;
  dirRel: string;
}

/** ブックの置き場所（絶対パス）。未設定ならソースのルート */
export function bookRootOf(root: string, config: Pick<XlcodeConfig, 'bookRoot'>): string {
  return path.resolve(root, config.bookRoot?.trim() || '.');
}

/**
 * ブックとソースの対応はフォルダの位置だけで決まる。
 *   <bookRoot>/app/app.xlcode.xlsx ↔ <root>/app/
 */
export function bookRef(root: string, bookAbs: string, bookRoot: string = root): BookRef {
  const abs = path.resolve(bookRoot, bookAbs);
  const dirRel = toPosixRel(bookRoot, path.dirname(abs));
  return { abs, rel: toPosixRel(bookRoot, abs), dirAbs: path.join(root, ...dirRel.split('/')), dirRel };
}

/** ソースのディレクトリに対応するブックの絶対パス */
export function bookPathFor(root: string, bookRoot: string, dirAbs: string): string {
  const dirRel = path.relative(root, dirAbs);
  const name = path.basename(path.resolve(dirAbs));
  return path.join(bookRoot, dirRel, `${name}${BOOK_SUFFIX}`);
}

/** ブックの相対パス（置き場所から）を絶対パスにする。置き場所の外を指していたら例外 */
export function resolveBook(bookRoot: string, rel: string): string {
  const abs = path.resolve(bookRoot, rel);
  const r = path.relative(bookRoot, abs);
  if (r.startsWith('..') || path.isAbsolute(r)) throw new Error(`ブックの置き場所の外のパスです: ${rel}`);
  return abs;
}

export async function exists(p: string): Promise<boolean> {
  try {
    await access(p);
    return true;
  } catch {
    return false;
  }
}

/** 同一ディレクトリ内のブック一覧（2つ以上ならエラー）。ディレクトリが無ければ空 */
export async function booksInDir(dirAbs: string): Promise<string[]> {
  if (!(await exists(dirAbs))) return [];
  return (await readdir(dirAbs)).filter(isBookFile);
}
