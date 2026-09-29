import { spawn } from 'node:child_process';
import { readdir, readFile } from 'node:fs/promises';
import path from 'node:path';
import { BrowserWindow, dialog, ipcMain, shell } from 'electron';
import {
  AGENTS_SHEET,
  LOCAL_AGENTS_SHEET,
  bookStatus,
  build,
  checkBookOpen,
  computeTree,
  createBook,
  findBooks,
  gitSummary,
  initProject,
  bookPathFor,
  bookRef,
  bookRootOf,
  loadConfig,
  relocateBooks,
  resolveBook,
  loadIgnore,
  refreshTree,
  saveConfig,
  sync,
  type XlcodeConfig,
} from '../core';
import { atomicWrite } from '../core/atomic';
import { excelPathError, isBookFile, isIgnored, toPosixRel } from '../core/fsutil';
import { exists } from '../core/project';
import type { BookSummary, OpenVia, ProjectInfo, Result } from '../shared/api';
import { joinUrl, readSyncRoots, toWebUrl } from './onedrive';
import { bookStamp, querySyncStatus } from './syncStatus';

async function wrap<T>(fn: () => Promise<T>): Promise<Result<T>> {
  try {
    return { ok: true, value: await fn() };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : String(e) };
  }
}

/** レンダラーから渡された相対パスがプロジェクト外を指していないか確認する */
function inside(root: string, rel: string): string {
  const abs = path.resolve(root, rel);
  const r = path.relative(path.resolve(root), abs);
  if (r.startsWith('..') || path.isAbsolute(r)) throw new Error(`プロジェクト外のパスです: ${rel}`);
  return abs;
}

/** ブックの置き場所 */
async function bookRootFor(root: string): Promise<string> {
  return bookRootOf(root, await loadConfig(root));
}

/** レンダラーから渡されたブックの相対パス（置き場所から）を絶対パスにする */
async function bookAbsOf(root: string, rel: string): Promise<string> {
  return resolveBook(await bookRootFor(root), rel);
}

async function dirsWithoutBook(root: string, bookRoot: string): Promise<string[]> {
  const ig = await loadIgnore(root);
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

async function loadProject(root: string): Promise<ProjectInfo> {
  const ig = await loadIgnore(root);
  const tree = await computeTree(root, ig);
  const config = await loadConfig(root);
  const bookRoot = bookRootOf(root, config);
  const books: BookSummary[] = [];
  for (const abs of await findBooks(root, ig, bookRoot)) {
    const { rel, dirRel } = bookRef(root, abs, bookRoot);
    const open = await checkBookOpen(abs, { quick: true });
    try {
      books.push({ ...(await bookStatus(root, abs)), rel, dirRel, open: open.open, openReason: open.reason });
    } catch (e) {
      books.push({
        rel,
        dirRel,
        open: open.open,
        openReason: open.reason,
        loadError: e instanceof Error ? e.message : String(e),
      });
    }
  }
  const hasAgents = await readFile(path.join(root, AGENTS_SHEET)).then(
    () => true,
    () => false,
  );
  return {
    root,
    name: path.basename(root),
    books,
    treeVersion: tree.version,
    git: await gitSummary(root),
    hasAgents,
    dirsWithoutBook: await dirsWithoutBook(root, bookRoot),
    bookRoot: config.bookRoot ? bookRoot : null,
  };
}

/** OS の標準ターミナルを対象ディレクトリで起動する（8.4。コマンド実行機能は持たない） */
function openTerminal(dir: string): void {
  const opts = { cwd: dir, detached: true, stdio: 'ignore' as const };
  if (process.platform === 'win32') {
    // Windows Terminal があれば優先し、無ければ cmd
    const wt = spawn('wt.exe', ['-d', dir], opts);
    wt.on('error', () => spawn('cmd.exe', ['/c', 'start', 'cmd.exe'], { ...opts, shell: false }).unref());
    wt.unref();
  } else if (process.platform === 'darwin') {
    spawn('open', ['-a', 'Terminal', dir], opts).unref();
  } else {
    spawn('x-terminal-emulator', [], opts).unref();
  }
}

/** Web 版 Excel で開く URL。設定（webUrlBase）優先、無ければ OneDrive の同期設定から求める */
async function webUrl(root: string, abs: string): Promise<string> {
  const { webUrlBase } = await loadConfig(root);
  if (webUrlBase) return joinUrl(webUrlBase, toPosixRel(await bookRootFor(root), abs).split('/'));
  const url = toWebUrl(abs, await readSyncRoots());
  if (url) return url;
  throw new Error(
    'Web 版の URL を特定できません。プロジェクトが OneDrive / SharePoint の同期フォルダ内にあるか確認するか、設定の「Web 版の URL」にプロジェクトルートの URL を入力してください',
  );
}

const RULE_FILES = new Set([AGENTS_SHEET, LOCAL_AGENTS_SHEET]);

function ruleFile(root: string, rel: string): string {
  if (!RULE_FILES.has(path.posix.basename(rel))) throw new Error(`編集できないファイルです: ${rel}`);
  return inside(root, rel);
}

export function registerIpc(): void {
  ipcMain.handle('pickProject', async (e) => {
    const win = BrowserWindow.fromWebContents(e.sender)!;
    const r = await dialog.showOpenDialog(win, { title: 'プロジェクトを開く', properties: ['openDirectory'] });
    return r.canceled ? null : r.filePaths[0];
  });
  ipcMain.handle('loadProject', (_e, root: string) => wrap(() => loadProject(root)));
  ipcMain.handle('initProject', (_e, root: string) => wrap(() => initProject(root)));
  ipcMain.handle('refreshTree', (_e, root: string) => wrap(() => refreshTree(root)));
  ipcMain.handle('createBook', (_e, root: string, dirRel: string) =>
    wrap(() => createBook(root, inside(root, dirRel))),
  );
  ipcMain.handle('build', (_e, root: string, book: string, opts) =>
    wrap(async () => build(root, await bookAbsOf(root, book), opts)),
  );
  ipcMain.handle('sync', (_e, root: string, book: string, opts) =>
    wrap(async () => sync(root, await bookAbsOf(root, book), opts)),
  );
  ipcMain.handle('pickFolder', async (e, title: string) => {
    const win = BrowserWindow.fromWebContents(e.sender)!;
    const r = await dialog.showOpenDialog(win, { title, properties: ['openDirectory', 'createDirectory'] });
    return r.canceled ? null : r.filePaths[0];
  });
  ipcMain.handle('relocateBooks', (_e, root: string, newRoot: string | null) =>
    wrap(() => relocateBooks(root, newRoot)),
  );
  ipcMain.handle('openInExcel', (_e, root: string, book: string, via: OpenVia) =>
    wrap(async () => {
      const abs = await bookAbsOf(root, book);
      const longPath = excelPathError(abs);
      if (longPath) throw new Error(longPath);
      if (via === 'desktop') {
        // 関連付けられたアプリ（通常はデスクトップ版 Excel）で開く
        const err = await shell.openPath(abs);
        if (err) throw new Error(err);
        return null;
      }
      const url = await webUrl(root, abs);
      await shell.openExternal(url);
      return url;
    }),
  );
  ipcMain.handle('bookLocks', (_e, root: string, books: string[]) =>
    wrap(async () => {
      const out: Record<string, boolean> = {};
      const bookRoot = await bookRootFor(root);
      for (const b of books) out[b] = (await checkBookOpen(resolveBook(bookRoot, b), { quick: true })).open;
      return out;
    }),
  );
  ipcMain.handle('syncStatus', (_e, root: string, books: string[]) =>
    wrap(async () => {
      const bookRoot = await bookRootFor(root);
      return querySyncStatus(Object.fromEntries(books.map((b) => [b, resolveBook(bookRoot, b)])));
    }),
  );
  ipcMain.handle('bookStamp', (_e, root: string, book: string) =>
    wrap(async () => bookStamp(await bookAbsOf(root, book))),
  );
  ipcMain.handle('readConfig', (_e, root: string) => wrap(() => loadConfig(root)));
  ipcMain.handle('writeConfig', (_e, root: string, config: XlcodeConfig) => wrap(() => saveConfig(root, config)));
  ipcMain.handle('openTerminal', (_e, root: string, dirRel: string) =>
    wrap(async () => openTerminal(inside(root, dirRel))),
  );
  ipcMain.handle('revealInFolder', (_e, root: string, book: string) =>
    wrap(async () => shell.showItemInFolder(await bookAbsOf(root, book))),
  );
  ipcMain.handle('readRuleFile', (_e, root: string, rel: string) =>
    wrap(() => readFile(ruleFile(root, rel), 'utf8').catch(() => null)),
  );
  ipcMain.handle('writeRuleFile', (_e, root: string, rel: string, text: string) =>
    wrap(() => atomicWrite(ruleFile(root, rel), text)),
  );
}
