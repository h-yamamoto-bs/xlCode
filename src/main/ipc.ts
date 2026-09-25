import { spawn } from 'node:child_process';
import { readdir, readFile, writeFile } from 'node:fs/promises';
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
  loadIgnore,
  refreshTree,
  sync,
} from '../core';
import { isBookFile, isIgnored, toPosixRel } from '../core/fsutil';
import type { BookSummary, ProjectInfo, Result } from '../shared/api';

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

async function dirsWithoutBook(root: string): Promise<string[]> {
  const ig = await loadIgnore(root);
  const out: string[] = [];
  async function walk(dirAbs: string): Promise<void> {
    const entries = await readdir(dirAbs, { withFileTypes: true });
    if (!entries.some((e) => e.isFile() && isBookFile(e.name))) out.push(toPosixRel(root, dirAbs));
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
  const books: BookSummary[] = [];
  for (const abs of await findBooks(root, ig)) {
    const rel = toPosixRel(root, abs);
    const dirRel = toPosixRel(root, path.dirname(abs));
    const open = await checkBookOpen(abs);
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
    dirsWithoutBook: await dirsWithoutBook(root),
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
  ipcMain.handle('build', (_e, root: string, book: string, opts) => wrap(() => build(root, inside(root, book), opts)));
  ipcMain.handle('sync', (_e, root: string, book: string, opts) => wrap(() => sync(root, inside(root, book), opts)));
  ipcMain.handle('openInExcel', (_e, root: string, book: string) =>
    wrap(async () => {
      const err = await shell.openPath(inside(root, book));
      if (err) throw new Error(err);
    }),
  );
  ipcMain.handle('openTerminal', (_e, root: string, dirRel: string) =>
    wrap(async () => openTerminal(inside(root, dirRel))),
  );
  ipcMain.handle('revealInFolder', (_e, root: string, rel: string) =>
    wrap(async () => shell.showItemInFolder(inside(root, rel))),
  );
  ipcMain.handle('readRuleFile', (_e, root: string, rel: string) =>
    wrap(() => readFile(ruleFile(root, rel), 'utf8').catch(() => null)),
  );
  ipcMain.handle('writeRuleFile', (_e, root: string, rel: string, text: string) =>
    wrap(() => writeFile(ruleFile(root, rel), text)),
  );
}
