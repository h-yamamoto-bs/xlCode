import { mkdir, readFile } from 'node:fs/promises';
import { atomicWrite } from './atomic';
import path from 'node:path';
import type { Ignore } from 'ignore';
import { AGENTS_SHEET, LOCAL_AGENTS_SHEET, TREE_FILE, TREE_SHEET, TREE_VERSION_PREFIX, XLCODE_DIR } from './constants';
import { loadConfig } from './config';
import { findBooks, loadIgnore, toPosixRel, walkTree, type TreeNode } from './fsutil';
import { bookRef, bookRootOf } from './project';
import { checkBookOpen } from './lock';
import { normalizeText, sha256, textToLines } from './normalize';
import { Book } from './workbook';

export function renderTree(node: TreeNode, depth = 0): string[] {
  const indent = '  '.repeat(depth);
  if (!node.dir) return [`${indent}${node.name}`];
  return [`${indent}${node.name}/`, ...node.children.flatMap((c) => renderTree(c, depth + 1))];
}

export interface TreeSnapshot {
  lines: string[];
  /** ツリー内容のハッシュ（先頭12桁） */
  version: string;
  /** #tree の1行目 */
  header: string;
}

export async function computeTree(root: string, ig?: Ignore): Promise<TreeSnapshot> {
  const lines = renderTree(await walkTree(root, ig ?? (await loadIgnore(root))));
  const version = sha256(lines.join('\n')).slice(0, 12);
  return { lines, version, header: `${TREE_VERSION_PREFIX}${version} (${new Date().toISOString()})` };
}

export function parseTreeVersion(firstLine: string | undefined): string | null {
  if (!firstLine?.startsWith(TREE_VERSION_PREFIX)) return null;
  return firstLine.slice(TREE_VERSION_PREFIX.length).split(' ')[0] || null;
}

export interface BookTreeResult {
  book: string;
  ok: boolean;
  /** 既に最新だったので保存しなかった */
  unchanged?: boolean;
  error?: string;
}

function sheetIs(book: Book, name: string, lines: string[]): boolean {
  if (!book.hasSheet(name)) return false;
  const current = book.readSheet(name).lines;
  while (current.length > 0 && current[current.length - 1] === '') current.pop();
  return current.length === lines.length && current.every((l, i) => l === lines[i]);
}

/**
 * #tree と Agents.md シート（VBA モードでは LocalAgents.md シートも）が既に最新か。
 * 最新なら保存しない（Web 版で編集中のブックとの同期の衝突を避ける）
 */
export function treeUpToDate(
  book: Book,
  tree: TreeSnapshot,
  agentsLines: string[] | null,
  localAgentsLines: string[] | null = null,
): boolean {
  if (!book.hasSheet(TREE_SHEET)) return false;
  const lines = book.readSheet(TREE_SHEET).lines;
  if (parseTreeVersion(lines[0]) !== tree.version) return false;
  if (agentsLines && !sheetIs(book, AGENTS_SHEET, agentsLines)) return false;
  if (localAgentsLines && !sheetIs(book, LOCAL_AGENTS_SHEET, localAgentsLines)) return false;
  return true;
}

export interface RefreshResult {
  version: string;
  books: BookTreeResult[];
  /** 一部だけ更新された（#tree が不整合） */
  partial: boolean;
}

/** #tree と Agents.md シート（VBA モードでは LocalAgents.md シートも）を書き込む（ブックの保存は呼び出し側） */
export function applyTreeToBook(
  book: Book,
  tree: TreeSnapshot,
  agentsLines: string[] | null,
  localAgentsLines: string[] | null = null,
): void {
  book.writeLines(TREE_SHEET, [tree.header, ...tree.lines]);
  book.moveToFront(TREE_SHEET);
  // No.21 暫定案: ルートの Agents.md を全ブックへ一方向配布する
  if (agentsLines) book.writeLines(AGENTS_SHEET, agentsLines);
  // VBA モードには Sync が無いため、LocalAgents.md もここで配布する
  if (localAgentsLines) book.writeLines(LOCAL_AGENTS_SHEET, localAgentsLines);
}

async function readRule(file: string): Promise<string[] | null> {
  try {
    const text = await readFile(file, 'utf8');
    return textToLines(normalizeText(text, AGENTS_SHEET, { trimTrailingWhitespace: false }));
  } catch {
    return null;
  }
}

export async function readRootAgents(root: string): Promise<string[] | null> {
  return readRule(path.join(root, AGENTS_SHEET));
}

/** ディレクトリの LocalAgents.md（VBA モード用） */
export async function readLocalAgents(dirAbs: string): Promise<string[] | null> {
  return readRule(path.join(dirAbs, LOCAL_AGENTS_SHEET));
}

/** Refresh Tree（6章）: 全ブックの #tree を同一内容に更新する */
export async function refreshTree(root: string): Promise<RefreshResult> {
  const ig = await loadIgnore(root);
  const tree = await computeTree(root, ig);
  await mkdir(path.join(root, XLCODE_DIR), { recursive: true });
  await atomicWrite(path.join(root, XLCODE_DIR, TREE_FILE), [tree.header, ...tree.lines].join('\n') + '\n');
  const agents = await readRootAgents(root);

  const config = await loadConfig(root);
  const bookRoot = bookRootOf(root, config);
  const books: BookTreeResult[] = [];
  for (const abs of await findBooks(root, ig, bookRoot)) {
    const rel = toPosixRel(bookRoot, abs);
    const local = config.mode === 'vba' ? await readLocalAgents(bookRef(root, abs, bookRoot).dirAbs) : null;
    try {
      const open = await checkBookOpen(abs);
      if (open.open) {
        books.push({ book: rel, ok: false, error: `開かれています（${open.reason}）` });
        continue;
      }
      const book = await Book.load(abs);
      if (treeUpToDate(book, tree, agents, local)) {
        books.push({ book: rel, ok: true, unchanged: true });
        continue;
      }
      applyTreeToBook(book, tree, agents, local);
      const again = await checkBookOpen(abs);
      if (again.open) {
        books.push({ book: rel, ok: false, error: `開かれています（${again.reason}）` });
        continue;
      }
      await book.save(abs);
      books.push({ book: rel, ok: true });
    } catch (e) {
      books.push({ book: rel, ok: false, error: e instanceof Error ? e.message : String(e) });
    }
  }
  const failed = books.filter((b) => !b.ok).length;
  return { version: tree.version, books, partial: failed > 0 && failed < books.length };
}

export interface TreeCheck {
  current: string;
  stored: string | null;
  upToDate: boolean;
}

/** ブックの #tree が最新か確認する（6.6） */
export async function checkTree(root: string, bookAbs: string): Promise<TreeCheck> {
  const tree = await computeTree(root);
  const book = await Book.load(bookAbs);
  const stored = book.hasSheet(TREE_SHEET) ? parseTreeVersion(book.readSheet(TREE_SHEET).lines[0]) : null;
  return { current: tree.version, stored, upToDate: stored === tree.version };
}
