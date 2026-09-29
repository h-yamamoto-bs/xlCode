import { readdir, readFile } from 'node:fs/promises';
import path from 'node:path';
import ignore, { type Ignore } from 'ignore';
import { AGENTS_SHEET, BOOK_SUFFIX, EXCEL_MAX_PATH, XLCODE_DIR } from './constants';
import { decodeFile, type Decoded } from './encoding';

/** .gitignore に関係なく常に除外するもの */
const ALWAYS_IGNORED = ['.git/', `${XLCODE_DIR}/`, '~$*'];

/** プロジェクトルートの .gitignore を読み込む（7章。自作パーサは作らない） */
export async function loadIgnore(root: string): Promise<Ignore> {
  const ig = ignore().add(ALWAYS_IGNORED);
  try {
    ig.add(await readFile(path.join(root, '.gitignore'), 'utf8'));
  } catch {
    // .gitignore が無ければ既定の除外のみ
  }
  return ig;
}

export function toPosixRel(root: string, abs: string): string {
  return path.relative(root, abs).split(path.sep).join('/');
}

export function isIgnored(ig: Ignore, relPosix: string, isDir: boolean): boolean {
  if (relPosix === '' || relPosix === '.') return false;
  return ig.ignores(isDir ? `${relPosix}/` : relPosix);
}

export function isBookFile(name: string): boolean {
  return name.endsWith(BOOK_SUFFIX) && !name.startsWith('~$');
}

/**
 * OneDrive の同期がぶつかったときの複製（例: app.xlcode-PC名.xlsx）や、
 * コピーでできたブック（例: app.xlcode (1).xlsx）。管理対象ではないが、編集内容が紛れている可能性がある
 */
export function isBookCopy(name: string): boolean {
  const n = name.toLowerCase();
  return !name.startsWith('~$') && n.includes('.xlcode') && n.endsWith('.xlsx') && !isBookFile(name);
}

/** Excel が開けないほど長いパスならメッセージを返す */
export function excelPathError(abs: string): string | null {
  if (abs.length <= EXCEL_MAX_PATH) return null;
  return `ブックのパスが ${abs.length} 文字あり、Excel が開ける ${EXCEL_MAX_PATH} 文字を超えています。フォルダを浅くするか名前を短くしてください: ${abs}`;
}

/** テキストとして読み、文字コード・改行コードを判定する（UTF-8 / UTF-8 BOM / UTF-16 LE / Shift_JIS） */
export async function readTextFile(abs: string): Promise<Decoded> {
  return decodeFile(await readFile(abs), path.basename(abs));
}

export interface DirFile {
  name: string;
  abs: string;
}

/**
 * ブックが担当するディレクトリ直下のソースファイルを列挙する。
 * 除外: .gitignore 対象、ブック自身、ルートの Agents.md（Agents.md シートで別管理）
 */
export async function listSourceFiles(root: string, dirAbs: string, ig: Ignore): Promise<DirFile[]> {
  const entries = await readdir(dirAbs, { withFileTypes: true });
  const isRoot = path.resolve(dirAbs) === path.resolve(root);
  const files: DirFile[] = [];
  for (const e of entries) {
    if (!e.isFile()) continue;
    if (isBookFile(e.name) || e.name.startsWith('~$')) continue;
    if (isRoot && e.name === AGENTS_SHEET) continue;
    const abs = path.join(dirAbs, e.name);
    if (isIgnored(ig, toPosixRel(root, abs), false)) continue;
    files.push({ name: e.name, abs });
  }
  return files.sort((a, b) => a.name.localeCompare(b.name));
}

export interface TreeNode {
  name: string;
  dir: boolean;
  children: TreeNode[];
}

function sortNodes(nodes: TreeNode[]): TreeNode[] {
  return nodes.sort((a, b) => (a.dir === b.dir ? a.name.localeCompare(b.name) : a.dir ? -1 : 1));
}

/**
 * プロジェクト全体を走査する。.gitignore 対象は除外するが、
 * *.xlcode.xlsx は .gitignore に含まれていても掲載する（Git 管理外にする運用のため）。
 */
export async function walkTree(root: string, ig: Ignore): Promise<TreeNode> {
  async function walk(dirAbs: string, name: string): Promise<TreeNode> {
    const entries = await readdir(dirAbs, { withFileTypes: true });
    const children: TreeNode[] = [];
    for (const e of entries) {
      const abs = path.join(dirAbs, e.name);
      const rel = toPosixRel(root, abs);
      if (e.isDirectory()) {
        if (isIgnored(ig, rel, true)) continue;
        children.push(await walk(abs, e.name));
      } else if (e.isFile()) {
        if (e.name.startsWith('~$')) continue;
        if (!isBookFile(e.name) && isIgnored(ig, rel, false)) continue;
        children.push({ name: e.name, dir: false, children: [] });
      }
    }
    return { name, dir: true, children: sortNodes(children) };
  }
  return walk(root, path.basename(path.resolve(root)));
}

/** プロジェクト配下の全ブックを検出する（絶対パス） */
export async function findBooks(root: string, ig: Ignore): Promise<string[]> {
  const out: string[] = [];
  async function walk(dirAbs: string): Promise<void> {
    const entries = await readdir(dirAbs, { withFileTypes: true });
    for (const e of entries) {
      const abs = path.join(dirAbs, e.name);
      if (e.isDirectory()) {
        if (!isIgnored(ig, toPosixRel(root, abs), true)) await walk(abs);
      } else if (e.isFile() && isBookFile(e.name)) {
        out.push(abs);
      }
    }
  }
  await walk(root);
  return out.sort();
}
