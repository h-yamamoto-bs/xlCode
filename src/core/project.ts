import { readdir } from 'node:fs/promises';
import path from 'node:path';
import type { Ignore } from 'ignore';
import { Canonicalizer } from './canonical';
import { loadConfig, type XlcodeConfig } from './config';
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
  /** ルートからの相対パス（state.json のキー） */
  rel: string;
  /** 担当ディレクトリ */
  dirAbs: string;
  dirRel: string;
}

export function bookRef(root: string, bookAbs: string): BookRef {
  const abs = path.resolve(root, bookAbs);
  const dirAbs = path.dirname(abs);
  return { abs, rel: toPosixRel(root, abs), dirAbs, dirRel: toPosixRel(root, dirAbs) };
}

/** 同一ディレクトリ内のブック一覧（2つ以上ならエラー） */
export async function booksInDir(dirAbs: string): Promise<string[]> {
  return (await readdir(dirAbs)).filter(isBookFile);
}
