import { access, readdir, readFile } from 'node:fs/promises';
import path from 'node:path';
import ignore, { type Ignore } from 'ignore';
import { AGENTS_SHEET, BOOK_SUFFIX, EXCEL_MAX_PATH, IGNORE_FILE, XLCODE_DIR } from './constants';
import { decodeFile, type Decoded } from './encoding';

/** .gitignore に関係なく常に除外するもの */
const ALWAYS_IGNORED = ['.git/', `${XLCODE_DIR}/`, '~$*', IGNORE_FILE];

/** 除外の設定ファイル（どちらも .gitignore の書き方。先に読んだものを後のもので上書きできる） */
export const IGNORE_SOURCES: readonly string[] = ['.gitignore', IGNORE_FILE];

/**
 * プロジェクトルートの .gitignore と .xlcodeignore を読み込む（7章。自作パーサは作らない）。
 * .gitignore は Git 管理外のもの（node_modules など）、.xlcodeignore は Git 管理はするがブックにはしないもの
 * （ドキュメント、生成ファイル、ロックファイルなど）を書く。
 */
export async function loadIgnore(root: string): Promise<Ignore> {
  const ig = ignore().add(ALWAYS_IGNORED);
  for (const name of IGNORE_SOURCES) {
    try {
      ig.add(await readFile(path.join(root, name), 'utf8'));
    } catch {
      // 無ければ既定の除外のみ
    }
  }
  return ig;
}

/** .xlcodeignore の雛形（initProject で無ければ作る） */
export const IGNORE_TEMPLATE = `# xlCode でブック（シート）にしないファイル・フォルダ
# .gitignore と同じ書き方。ここに書いたものは、Git では管理したまま、シートにも #tree にも載りません。
# 変更したら、ブック画面の「Sync」（既にシートになっているものはシートから消えます）と Refresh Tree を実行します。
#
# 例:
# docs/
# *.min.js
# package-lock.json
`;

/**
 * ブックを持たないディレクトリ（ブック作成の候補。ルートからの相対パス、'' はルート）。
 * 除外（.gitignore / .xlcodeignore）対象のディレクトリは候補にしない。
 */
export async function listDirsWithoutBook(root: string, ig: Ignore, bookRoot: string = root): Promise<string[]> {
  const separate = path.resolve(bookRoot) !== path.resolve(root);
  const out: string[] = [];
  async function walk(dirAbs: string): Promise<void> {
    const entries = await readdir(dirAbs, { withFileTypes: true });
    const hasBook = separate
      ? await exists(bookPathFor(root, bookRoot, dirAbs))
      : entries.some((e) => e.isFile() && isBookFile(e.name));
    if (!hasBook) out.push(toPosixRel(root, dirAbs));
    for (const e of entries) {
      const abs = path.join(dirAbs, e.name);
      if (e.isDirectory() && !isIgnored(ig, toPosixRel(root, abs), true)) await walk(abs);
    }
  }
  await walk(root);
  return out.sort();
}

async function exists(p: string): Promise<boolean> {
  try {
    await access(p);
    return true;
  } catch {
    return false;
  }
}

/** ソースのディレクトリに対応するブックの絶対パス */
export function bookPathFor(root: string, bookRoot: string, dirAbs: string): string {
  const dirRel = path.relative(root, dirAbs);
  const name = path.basename(path.resolve(dirAbs));
  return path.join(bookRoot, dirRel, `${name}${BOOK_SUFFIX}`);
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

/** 除外ファイル（.gitignore / .xlcodeignore）の有効な行（空行・コメント以外） */
export function ignoreRules(text: string): string[] {
  return text
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => l !== '' && !l.startsWith('#'));
}

export interface DirFilesPreview {
  /** シートになるファイル */
  sheets: string[];
  /** .gitignore / .xlcodeignore で除外されるファイル */
  ignored: string[];
  /** 拡張子が無いなどでコードとして扱わないファイル */
  nonCode: string[];
}

/**
 * ブックを作る前に、ディレクトリ直下のファイルがどう扱われるかを見る（ブック作成の確認用）。
 * バイナリかどうかは中身を読まないと分からないので、ここでは見ない
 */
export async function previewDirFiles(
  root: string,
  dirAbs: string,
  ig: Ignore,
  isCode: (name: string) => boolean,
): Promise<DirFilesPreview> {
  const out: DirFilesPreview = { sheets: [], ignored: [], nonCode: [] };
  const isRoot = path.resolve(dirAbs) === path.resolve(root);
  for (const e of await readdir(dirAbs, { withFileTypes: true })) {
    if (!e.isFile() || isBookFile(e.name) || e.name.startsWith('~$')) continue;
    if (isRoot && e.name === AGENTS_SHEET) continue;
    const rel = toPosixRel(root, path.join(dirAbs, e.name));
    if (isIgnored(ig, rel, false)) out.ignored.push(e.name);
    else if (!isCode(e.name)) out.nonCode.push(e.name);
    else out.sheets.push(e.name);
  }
  for (const k of ['sheets', 'ignored', 'nonCode'] as const) out[k].sort((a, b) => a.localeCompare(b));
  return out;
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

/**
 * 全ブックを検出する（絶対パス）。
 * ブックの置き場所がソースと別なら、置き場所の下を探す（.gitignore は適用しない）
 */
export async function findBooks(root: string, ig: Ignore, bookRoot: string = root): Promise<string[]> {
  if (path.resolve(bookRoot) !== path.resolve(root)) {
    const out: string[] = [];
    const walk = async (dirAbs: string): Promise<void> => {
      for (const e of await readdir(dirAbs, { withFileTypes: true })) {
        const abs = path.join(dirAbs, e.name);
        if (e.isDirectory() && !e.name.startsWith('.')) await walk(abs);
        else if (e.isFile() && isBookFile(e.name)) out.push(abs);
      }
    };
    try {
      await walk(bookRoot);
    } catch (e) {
      if ((e as NodeJS.ErrnoException).code !== 'ENOENT') throw e;
    }
    return out.sort();
  }
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
