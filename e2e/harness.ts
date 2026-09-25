import { mkdtemp, rm } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { _electron, type ElectronApplication, type Page } from 'playwright';
import { createBook, initProject } from '../src/core';
import { fixture, type Fixture } from '../tests/helpers';

const require = createRequire(import.meta.url);
const electronPath = require('electron') as unknown as string;
const appDir = path.resolve(import.meta.dirname, '..');

export const APP_BOOK = 'app/app.xlcode.xlsx';
export const API_BOOK = 'api/api.xlcode.xlsx';

/** app/ と api/ にブックを持つ Git 管理のプロジェクト */
export async function project(): Promise<Fixture> {
  const f = await fixture(
    {
      '.gitignore': 'node_modules/\n',
      'app/App.tsx': 'export const App = () => <p>hello</p>;\n',
      'app/util.ts': 'export const one = 1;\n',
      'api/routes.ts': "export const routes = ['/items'];\n",
    },
    { git: true },
  );
  await initProject(f.root);
  await createBook(f.root, f.file('app'));
  await createBook(f.root, f.file('api'));
  f.git('add', '-A');
  f.git('commit', '-q', '-m', 'setup');
  return f;
}

export interface Running {
  app: ElectronApplication;
  win: Page;
  close: () => Promise<void>;
}

/**
 * アプリを起動してプロジェクトを開く。
 * excelMode を渡すと初回の確認を省く（null なら確認が出る）。
 */
export async function launch(root: string, excelMode: string | null = 'both'): Promise<Running> {
  const userData = await mkdtemp(path.join(tmpdir(), 'xlcode-e2e-'));
  const args = [appDir];
  if (process.platform === 'linux' && process.getuid?.() === 0) args.unshift('--no-sandbox');
  const app = await _electron.launch({
    executablePath: electronPath,
    args,
    env: { ...process.env, XLCODE_USER_DATA: userData },
  });
  const win = await app.firstWindow();
  win.on('pageerror', (e) => {
    throw e;
  });
  await win.waitForLoadState('domcontentloaded');
  await win.evaluate(
    ([p, m]) => {
      localStorage.setItem('recent', JSON.stringify([p]));
      if (m) localStorage.setItem('excelMode', JSON.stringify(m));
    },
    [root, excelMode] as const,
  );
  await win.reload();
  await win
    .getByRole('button', { name: path.basename(root) })
    .first()
    .click();
  return {
    app,
    win,
    close: async () => {
      await app.close();
      await rm(userData, { recursive: true, force: true });
    },
  };
}

/** 左のブック一覧でブックを選ぶ */
export async function selectBook(win: Page, rel: string): Promise<void> {
  await win
    .getByRole('treeitem', { name: new RegExp(rel.split('/').pop()!.replace(/\./g, '\\.')) })
    .first()
    .locator('div')
    .first()
    .click();
}

/** 出力パネルに指定の文字列が出るまで待つ */
export async function waitLog(win: Page, text: string | RegExp): Promise<void> {
  await win.getByText(text).last().waitFor();
}
