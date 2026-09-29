import { access, appendFile, mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { AGENTS_SHEET, LOCAL_AGENTS_SHEET, XLCODE_DIR } from './constants';
import { excelPathError, listSourceFiles, readTextFile } from './fsutil';
import { textToLines } from './normalize';
import { bookPathFor, bookRef, bookRootOf, booksInDir, openProject } from './project';
import { bookState, saveState } from './state';
import { applyTreeToBook, computeTree, readRootAgents } from './tree';
import { isCodeName, validateFileName } from './sheetName';
import { Book } from './workbook';

const GITIGNORE_ENTRIES = ['*.xlcode.xlsx', '~$*', `${XLCODE_DIR}/`];

const AGENTS_TEMPLATE = `# Agents.md

- TypeScript の strict モードを使う
- コメントは日本語で書く
- 1シート = 1ファイル。A列に1行ずつ書く
- シート名はファイル名（拡張子込み・31文字以内）
- 「#」で始まるシートは編集しない
`;

const LOCAL_AGENTS_TEMPLATE = `# LocalAgents.md

- このディレクトリ固有のルールを書く
`;

async function exists(p: string): Promise<boolean> {
  try {
    await access(p);
    return true;
  } catch {
    return false;
  }
}

/**
 * プロジェクトの初期設定。
 * - .gitignore に xlCode 用の除外（ブック・ロックファイル・.xlcode/）を追記（8.2）
 * - ルートに Agents.md が無ければ雛形を作成
 */
export async function initProject(root: string): Promise<string[]> {
  const done: string[] = [];
  const gi = path.join(root, '.gitignore');
  const current = (await readFile(gi, 'utf8').catch(() => '')).split(/\r?\n/).map((l) => l.trim());
  const missing = GITIGNORE_ENTRIES.filter((e) => !current.includes(e));
  if (missing.length > 0) {
    const prefix = current.join('') === '' || current[current.length - 1] === '' ? '' : '\n';
    await appendFile(gi, `${prefix}# xlCode\n${missing.join('\n')}\n`);
    done.push(`.gitignore に追記: ${missing.join(', ')}`);
  }
  const agents = path.join(root, AGENTS_SHEET);
  if (!(await exists(agents))) {
    await writeFile(agents, AGENTS_TEMPLATE);
    done.push('Agents.md を作成');
  }
  return done;
}

export interface CreateBookResult {
  book: string;
  sheets: string[];
  skipped: string[];
}

/** ディレクトリに <dir_name>.xlcode.xlsx を作成し、既存ファイルをシートとして取り込む */
export async function createBook(root: string, dirAbs: string): Promise<CreateBookResult> {
  const ctx = await openProject(root);
  const bookRoot = bookRootOf(root, ctx.config);
  const bookAbs = bookPathFor(root, bookRoot, dirAbs);
  const existing = await booksInDir(path.dirname(bookAbs));
  if (existing.length > 0) throw new Error(`既にブックがあります: ${existing.join(', ')}`);
  const longPath = excelPathError(bookAbs);
  if (longPath) throw new Error(longPath);
  await mkdir(path.dirname(bookAbs), { recursive: true });
  const ref = bookRef(root, bookAbs, bookRoot);

  const localAgents = path.join(dirAbs, LOCAL_AGENTS_SHEET);
  if (!(await exists(localAgents))) await writeFile(localAgents, LOCAL_AGENTS_TEMPLATE);

  // 作成するブック自身も #tree に載るよう、先に空のブックを保存してからツリーを取る
  const book = Book.create();
  await book.save(bookAbs);
  applyTreeToBook(book, await computeTree(root, ctx.ig), await readRootAgents(root));
  const bs = bookState(ctx.state, ref.rel);
  const sheets: string[] = [];
  const skipped: string[] = [];
  const files = await listSourceFiles(root, dirAbs, ctx.ig);
  // LocalAgents.md を先頭に
  files.sort((a, b) => Number(b.name === LOCAL_AGENTS_SHEET) - Number(a.name === LOCAL_AGENTS_SHEET));
  for (const f of files) {
    if (!isCodeName(f.name, ctx.config.extraCodeNames)) {
      skipped.push(`${f.name}（拡張子がないため対象外）`);
      continue;
    }
    const err = validateFileName(f.name);
    if (err) throw new Error(err);
    const read = await readTextFile(f.abs);
    if (read.kind === 'binary') {
      skipped.push(`${f.name}（${read.reason}）`);
      continue;
    }
    const c = await ctx.canon.canonical(f.name, f.abs, read.text);
    book.writeLines(f.name, textToLines(c.text));
    bs.files[f.name] = { hash: c.hash, lines: c.lines, chars: c.chars };
    sheets.push(f.name);
  }
  await book.save(bookAbs);
  bs.lastSyncAt = new Date().toISOString();
  await saveState(root, ctx.state);
  return { book: ref.rel, sheets, skipped };
}
