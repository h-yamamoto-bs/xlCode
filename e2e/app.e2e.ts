import { rm, writeFile } from 'node:fs/promises';
import { afterEach, describe, expect, it } from 'vitest';
import { bookStatus } from '../src/core';
import { API_BOOK, APP_BOOK, launch, project, selectBook, waitLog, type Running } from './harness';

let running: Running | null = null;
afterEach(async () => {
  await running?.close();
  running = null;
});

describe('xlCode GUI', () => {
  it('初回に使う Excel を聞き、「両方」なら開くボタンが2つ出る', async () => {
    const f = await project();
    running = await launch(f.root, null);
    const { win } = running;
    await win.getByText('どの Excel でブックを編集しますか？').waitFor();
    await win.getByRole('button', { name: '両方', exact: true }).click();
    await selectBook(win, APP_BOOK);
    await expect(win.getByRole('button', { name: 'デスクトップで開く' }).isVisible()).resolves.toBe(true);
    await expect(win.getByRole('button', { name: 'Webで開く' }).isVisible()).resolves.toBe(true);
    await expect(win.getByText('Excel: 両方').isVisible()).resolves.toBe(true);
  });

  it('Excel 側の変更を Build でファイルへ出力する', async () => {
    const f = await project();
    await f.editBook(APP_BOOK, (b) => b.writeLines('util.ts', ['export const one = 1;', 'export const two = 2;']));
    running = await launch(f.root);
    const { win } = running;
    await selectBook(win, APP_BOOK);
    await win.getByText('Excel側で編集中', { exact: true }).first().waitFor();
    await win.getByRole('button', { name: 'Build', exact: true }).click();
    await win.getByRole('button', { name: '閉じたので続行' }).click();
    await waitLog(win, 'Build: 完了');
    expect(await f.read('app/util.ts')).toBe('export const one = 1;\nexport const two = 2;\n');
    expect(f.git('log', '-1', '--format=%s').trim()).toBe('setup');
  });

  it('Sync で未 Build の変更があれば選ばせ、「先に Build」なら両方反映する', async () => {
    const f = await project();
    await f.editBook(APP_BOOK, (b) => b.writeLines('util.ts', ['export const one = 11;']));
    await f.write('app/App.tsx', 'export const App = () => <p>changed</p>;\n');
    running = await launch(f.root);
    const { win } = running;
    await selectBook(win, APP_BOOK);
    await win.getByRole('button', { name: 'Sync', exact: true }).click();
    await win.getByRole('button', { name: '閉じたので続行' }).click();
    await win.getByText('Excel 側に Build していない変更があります').waitFor();
    await win.getByRole('button', { name: '先に Build してから Sync' }).click();
    // Build 前の未コミット確認（App.tsx）は出るが、Build 後の Sync では出ない
    await win.getByRole('button', { name: '続行', exact: true }).click();
    await waitLog(win, 'Sync: 完了');
    expect(await f.read('app/util.ts')).toBe('export const one = 11;\n');
    const st = await bookStatus(f.root, f.file(APP_BOOK));
    expect(st.files.every((x) => x.status === 'clean')).toBe(true);
  });

  it('両側で変更されたら衝突シートを作り、案内を出す', async () => {
    const f = await project();
    await f.editBook(API_BOOK, (b) => b.writeLines('routes.ts', ["export const routes = ['/items', '/excel'];"]));
    await f.write('api/routes.ts', "export const routes = ['/items', '/editor'];\n");
    running = await launch(f.root);
    const { win } = running;
    await selectBook(win, API_BOOK);
    await win.getByRole('button', { name: 'Build', exact: true }).click();
    await win.getByRole('button', { name: '閉じたので続行' }).click();
    await win.getByText('両側で変更されたファイルがあります').waitFor();
    await expect(win.getByText('#conflict_01 ← routes.ts').isVisible()).resolves.toBe(true);
    await win.getByRole('button', { name: 'OK' }).click();
    await win.getByText('未解決の衝突シートがあります').waitFor();
    await expect(win.getByRole('button', { name: 'Build', exact: true }).isEnabled()).resolves.toBe(true);
    expect((await f.sheet(API_BOOK, '#conflict_01'))[0]).toBe('api/routes.ts');
  });

  it('デスクトップ版で開かれたら操作を止め、閉じたら自動で解除する', async () => {
    const f = await project();
    running = await launch(f.root);
    const { win } = running;
    await selectBook(win, APP_BOOK);
    const lock = f.file('app/~$app.xlcode.xlsx');
    await writeFile(lock, '');
    await win.getByText('デスクトップ版 Excel で開かれています').waitFor({ timeout: 10_000 });
    await expect(win.getByRole('button', { name: 'Build', exact: true }).isDisabled()).resolves.toBe(true);
    await rm(lock);
    await win.getByText('デスクトップ版 Excel で開かれています').waitFor({ state: 'detached', timeout: 10_000 });
    await expect(win.getByRole('button', { name: 'Build', exact: true }).isEnabled()).resolves.toBe(true);
    await waitLog(win, 'Excel が閉じられました: app/app.xlcode.xlsx');
  });

  it('保存していない Markdown はファイルを切り替えても残り、保存で書き込まれる', async () => {
    const f = await project();
    running = await launch(f.root);
    const { win } = running;
    await win.getByRole('button', { name: 'ルール（Agents.md）' }).click();
    const editor = win.getByRole('textbox', { name: 'Agents.md' });
    await editor.click();
    await editor.press('Control+End');
    await editor.pressSequentially('\n- 追加したルール');
    await win.getByText('app', { exact: true }).click();
    await win.getByRole('textbox', { name: 'app/LocalAgents.md' }).waitFor();
    await win.getByText('Agents.md', { exact: true }).first().click();
    await expect(editor.inputValue()).resolves.toContain('- 追加したルール');
    await expect(win.getByTitle('保存していない変更があります').count()).resolves.toBe(1);
    await editor.press('Control+s');
    await waitLog(win, '保存しました: Agents.md');
    expect(await f.read('Agents.md')).toContain('- 追加したルール');
    await expect(win.getByTitle('保存していない変更があります').count()).resolves.toBe(0);
  });

  it('編集中にファイルがほかで変更されたら知らせる', async () => {
    const f = await project();
    running = await launch(f.root);
    const { win } = running;
    await win.getByRole('button', { name: 'ルール（Agents.md）' }).click();
    const editor = win.getByRole('textbox', { name: 'Agents.md' });
    await editor.click();
    await editor.pressSequentially('x');
    await f.write('Agents.md', '# 別の場所で編集\n');
    // ウィンドウに戻ったことにする
    await win.evaluate(() => window.dispatchEvent(new Event('focus')));
    await win.getByText('ほかの場所で変更されました').waitFor();
    await win.getByRole('button', { name: '編集を破棄して読み込み直す' }).click();
    await expect(editor.inputValue()).resolves.toBe('# 別の場所で編集\n');
  });

  it('state.json が壊れていたら Build を止めて理由を出す', async () => {
    const f = await project();
    running = await launch(f.root);
    const { win } = running;
    await selectBook(win, APP_BOOK);
    await writeFile(f.file('.xlcode/state.json'), '{ broken');
    await win.getByRole('button', { name: 'Build', exact: true }).click();
    await win.getByRole('button', { name: '閉じたので続行' }).click();
    await waitLog(win, /state\.json が壊れています/);
    expect(await f.read('app/util.ts')).toBe('export const one = 1;\n');
  });
});
