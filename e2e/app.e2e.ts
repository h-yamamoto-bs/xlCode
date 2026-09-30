import { rm, writeFile } from 'node:fs/promises';
import { afterEach, describe, expect, it } from 'vitest';
import { bookStatus, createBook, refreshTree } from '../src/core';
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

  it('ソースの変更が無ければ、「開く」で確認も Sync もせずにそのまま開く', async () => {
    const f = await project();
    running = await launch(f.root, 'web');
    const { win } = running;
    await selectBook(win, APP_BOOK);
    await win.getByRole('button', { name: 'Excelで開く' }).click();
    // Linux には OneDrive が無いので URL を求められず、開けなかった旨が出る（その前に確認は出ない）
    await win.getByText('Web 版で開けませんでした').waitFor();
    await expect(win.getByRole('button', { name: '閉じたので続行' }).count()).resolves.toBe(0);
    await expect(win.getByText('Sync 開始').count()).resolves.toBe(0);
  });

  it('ブックをディレクトリの階層で表示する', async () => {
    const f = await project();
    await f.write('pkg/web/src/main.ts', 'export {};\n');
    await createBook(f.root, f.file('pkg/web/src'));
    running = await launch(f.root);
    const { win } = running;
    const tree = win.getByRole('tree', { name: 'ブック', exact: true });
    await tree.getByRole('treeitem', { name: 'pkg/web/src', exact: true }).waitFor();
    await tree.getByRole('treeitem', { name: 'src.xlcode.xlsx' }).waitFor();
    // 閉じると中のブックが隠れる
    await tree.getByRole('treeitem', { name: 'pkg/web/src', exact: true }).locator('div').first().click();
    await expect(tree.getByRole('treeitem', { name: 'src.xlcode.xlsx' }).count()).resolves.toBe(0);
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

describe('Web 版での編集', () => {
  it('Web 版で開いてからブックが変わっていなければ、Build 前に知らせる', async () => {
    const f = await project();
    // 開いたときの Refresh Tree でブックが書き換わらないよう、先に最新にしておく
    await refreshTree(f.root);
    const { stat } = await import('node:fs/promises');
    const s = await stat(f.file(APP_BOOK));
    running = await launch(f.root, 'both', {
      lastOpened: { [`${f.root}|${APP_BOOK}`]: { via: 'web', stamp: `${s.mtimeMs}:${s.size}` } },
    });
    const { win } = running;
    await selectBook(win, APP_BOOK);
    await win.getByRole('button', { name: 'Build', exact: true }).click();
    await win.getByRole('button', { name: '閉じたので続行' }).click();
    await win.getByText('Web 版での編集が、まだこの PC に届いていない可能性があります').waitFor();
    await win.getByRole('button', { name: '編集していないのでBuild' }).click();
    await waitLog(win, /Build: (完了|変更なし)/);
  });

  it('ブックが更新されていれば知らせない', async () => {
    const f = await project();
    running = await launch(f.root, 'both', {
      lastOpened: { [`${f.root}|${APP_BOOK}`]: { via: 'web', stamp: '0:0' } },
    });
    const { win } = running;
    await selectBook(win, APP_BOOK);
    await win.getByRole('button', { name: 'Build', exact: true }).click();
    await win.getByRole('button', { name: '閉じたので続行' }).click();
    await waitLog(win, /Build: (完了|変更なし)/);
    await expect(win.getByText('まだこの PC に届いていない').count()).resolves.toBe(0);
  });
});

describe('ブックの置き場所', () => {
  it('設定画面で置き場所を変えると既存のブックが移動し、そのまま Build できる', async () => {
    const f = await project();
    const { mkdtemp, readdir } = await import('node:fs/promises');
    const { tmpdir } = await import('node:os');
    const path = await import('node:path');
    const xl = await mkdtemp(path.join(tmpdir(), 'xlfolder-'));
    running = await launch(f.root);
    const { app, win } = running;
    // OS のフォルダ選択ダイアログを差し替える
    await app.evaluate(({ dialog }, dir) => {
      dialog.showOpenDialog = (async () => ({ canceled: false, filePaths: [dir] })) as typeof dialog.showOpenDialog;
    }, xl);
    await win.getByRole('button', { name: '設定' }).click();
    await win.getByRole('button', { name: 'フォルダを選ぶ…' }).click();
    await win.getByText('既存のブック 2 冊を、同じフォルダ構成のまま移動します').waitFor();
    await win.getByRole('button', { name: '移動して変更' }).click();
    await waitLog(win, /ブックの置き場所を変更しました/);
    expect(await readdir(path.join(xl, 'app'))).toEqual(['app.xlcode.xlsx']);
    expect((await readdir(f.file('app'))).filter((n) => n.endsWith('.xlsx'))).toEqual([]);

    await f.editBook(path.relative(f.root, path.join(xl, 'app', 'app.xlcode.xlsx')), (b) =>
      b.writeLines('util.ts', ['export const one = 1;', 'export const moved = 2;']),
    );
    await win.getByRole('button', { name: 'エクスプローラー' }).click();
    await selectBook(win, APP_BOOK);
    await win.getByRole('button', { name: 'Build', exact: true }).click();
    await win.getByRole('button', { name: '閉じたので続行' }).click();
    await waitLog(win, 'Build: 完了');
    expect(await f.read('app/util.ts')).toBe('export const one = 1;\nexport const moved = 2;\n');
  });

  it('最初のブックを作るときにプロジェクトの種類を選び、VBA ならソースと .xlsm の案内が出る', async () => {
    const { fixture } = await import('../tests/helpers');
    const f = await fixture({ 'README.txt': 'ツール\n' }, { git: true });
    running = await launch(f.root);
    const { win } = running;
    await win.getByTitle('ルート にブックを作成').click();
    await win.getByText('プロジェクトの種類を選んでください').waitFor();
    await win.getByRole('button', { name: 'VBA', exact: true }).click();
    await win.getByRole('button', { name: 'プロジェクトの中に置く' }).click();
    await win.getByText(/Build 結果は .*\.xlsm です/).waitFor();
    await win.getByRole('button', { name: '空のブックを作成', exact: true }).click();
    await waitLog(win, /ブックを作成しました/);
    const path = await import('node:path');
    const name = path.basename(f.root);
    await win.getByText('まだ Build していません', { exact: false }).waitFor();
    await expect(win.getByRole('button', { name: 'Sync', exact: true }).count()).resolves.toBe(1);
    await expect(win.getByText('Module1.bas').first().isVisible()).resolves.toBe(true);
    expect(await f.read('Agents.md')).toContain('Begin UserForm');
    const { readdir } = await import('node:fs/promises');
    expect(await readdir(f.root)).toEqual(expect.arrayContaining(['Module1.bas', 'References.refs']));

    // 作成の結果はブック画面に出る
    const status = win.getByRole('status');
    await status.getByText(/ブックを作成しました（\d+ シート）/).waitFor();
    // この環境（Windows 以外）では Excel が無いため、「次の操作」で先に伝え、Build はエラーになる
    await win.getByRole('note').getByText('VBA の Build は Windows のデスクトップ版 Excel が必要です').waitFor();
    await win.getByRole('button', { name: 'Build', exact: true }).click();
    await win.getByRole('button', { name: '閉じたので続行' }).click();
    await waitLog(win, /Windows とデスクトップ版 Excel が必要です/);
    await status.getByText('Build を中断しました（何も書き換えていません）').waitFor();
    expect(await readdir(f.root)).not.toContain(`${name}.xlsm`);

    // VBA モードでも Sync は使える（エディタで直したソースをシートへ）
    await f.write('Module1.bas', 'Option Explicit\r\nSub FromEditor()\r\nEnd Sub\r\n');
    await win.evaluate(() => window.dispatchEvent(new Event('focus')));
    await win.getByRole('note').getByText('Sync でエディタ側の変更 1 ファイルをシートへ反映').waitFor();
    await win.getByRole('button', { name: 'Sync', exact: true }).click();
    await win.getByRole('button', { name: '閉じたので続行' }).click();
    await win.getByRole('button', { name: '続行', exact: true }).click();
    await waitLog(win, 'Sync: 完了');
    expect(await f.sheet(`${name}.xlcode.xlsx`, 'Module1.bas')).toEqual([
      'Option Explicit',
      'Sub FromEditor()',
      'End Sub',
    ]);
  });

  it('VBA: 既存の Excel ツールから作成を選ぶと、ファイルを選んで取り込む（Windows 以外ではエラー）', async () => {
    const { fixture } = await import('../tests/helpers');
    const { saveConfig } = await import('../src/core');
    const { DEFAULT_CONFIG } = await import('../src/core/config');
    const f = await fixture({ 'tool.xlsm': 'x' }, { git: true });
    await saveConfig(f.root, { ...DEFAULT_CONFIG, mode: 'vba' });
    running = await launch(f.root);
    const { app, win } = running;
    await app.evaluate(({ dialog }, file) => {
      dialog.showOpenDialog = (async () => ({ canceled: false, filePaths: [file] })) as typeof dialog.showOpenDialog;
    }, f.file('tool.xlsm'));
    await win.getByTitle('ルート にブックを作成').click();
    await win.getByRole('button', { name: 'プロジェクトの中に置く' }).click();
    await win.getByRole('button', { name: '既存の Excel ツールから作成…' }).click();
    // 失敗の理由はその場（ダイアログ）に出る
    const dialog = win.getByRole('dialog', { name: '取り込めませんでした' });
    await dialog.getByText(/Windows とデスクトップ版 Excel が必要です/).waitFor();
    await expect(dialog.getByText('元のファイルは変更していません', { exact: false }).isVisible()).resolves.toBe(true);
    await dialog.getByRole('button', { name: '閉じる' }).click();
    await waitLog(win, /Windows とデスクトップ版 Excel が必要です/);
  });

  it('ソースコード: すべてのディレクトリにまとめてブックを作成する', async () => {
    const { fixture } = await import('../tests/helpers');
    const f = await fixture({ 'a/x.ts': 'export const x = 1;\n', 'b/y.ts': 'export const y = 1;\n' }, { git: true });
    running = await launch(f.root);
    const { win } = running;
    await win.getByTitle('すべてのディレクトリにブックを作成').click();
    await win.getByRole('button', { name: 'ソースコード', exact: true }).click();
    await win.getByRole('button', { name: 'プロジェクトの中に置く' }).click();
    await win.getByText('3 個のディレクトリにブックを作成しますか？').waitFor();
    await win.getByRole('button', { name: 'すべて作成' }).click();
    await waitLog(win, 'ブックを 3 / 3 冊作成しました');
    expect((await f.sheet('a/a.xlcode.xlsx', 'x.ts'))[0]).toBe('export const x = 1;');
  });

  it('最初のブックを作るときに OneDrive のフォルダを選ぶと、編集用ブックはそこに作られる', async () => {
    const { fixture } = await import('../tests/helpers');
    const { mkdtemp, readdir } = await import('node:fs/promises');
    const { tmpdir } = await import('node:os');
    const path = await import('node:path');
    const f = await fixture({ 'a.ts': 'export const a = 1;\n' }, { git: true });
    const onedrive = await mkdtemp(path.join(tmpdir(), 'onedrive-'));
    running = await launch(f.root);
    const { app, win } = running;
    await app.evaluate(({ dialog }, dir) => {
      dialog.showOpenDialog = (async () => ({ canceled: false, filePaths: [dir] })) as typeof dialog.showOpenDialog;
    }, onedrive);
    await win.getByTitle('ルート にブックを作成').click();
    await win.getByRole('button', { name: 'ソースコード', exact: true }).click();
    await win.getByText('編集用ブックをどこに置きますか？').waitFor();
    await win.getByRole('button', { name: 'OneDrive のフォルダを選ぶ…' }).click();
    await win.getByText(onedrive, { exact: false }).first().waitFor();
    await win.getByRole('button', { name: '作成', exact: true }).click();
    await waitLog(win, /ブックを作成しました/);
    const name = `${path.basename(f.root)}.xlcode.xlsx`;
    expect(await readdir(onedrive)).toEqual([name]);
    expect((await readdir(f.root)).filter((n) => n.endsWith('.xlsx'))).toEqual([]);
  });
});

describe('迷わないための案内', () => {
  it('次の操作を 1 つ示し、対応するボタンだけを強調する', async () => {
    const f = await project();
    await f.editBook(APP_BOOK, (b) => b.writeLines('util.ts', ['export const one = 1;', 'export const two = 2;']));
    running = await launch(f.root);
    const { win } = running;
    await selectBook(win, APP_BOOK);
    const note = win.getByRole('note');
    await note.getByText('Build で Excel 側の変更 1 ファイルをソースコードへ出力').waitFor();
    // ツールバーの Build が primary（アクセント色）、Sync は通常
    const build = win.getByRole('button', { name: 'Build', exact: true });
    await expect(build.evaluate((el) => el.className.includes('bg-accent'))).resolves.toBe(true);
    const syncBtn = win.getByRole('button', { name: 'Sync', exact: true });
    await expect(syncBtn.evaluate((el) => el.className.includes('bg-accent'))).resolves.toBe(false);
    // 案内の中のボタンからも実行できる
    await note.getByRole('button', { name: 'Build' }).click();
    await win.getByRole('button', { name: '閉じたので続行' }).click();
    await waitLog(win, 'Build: 完了');
    await note.getByText('すべて同期済み').waitFor();
  });

  it('押せないボタンには理由が出る', async () => {
    const f = await project();
    running = await launch(f.root);
    const { win } = running;
    await selectBook(win, APP_BOOK);
    const lock = f.file('app/~$app.xlcode.xlsx');
    await writeFile(lock, '');
    const build = win.getByRole('button', { name: 'Build', exact: true });
    await build.waitFor();
    await win.getByRole('note').getByText('Excel を閉じると Build / Sync できます').waitFor({ timeout: 10_000 });
    await expect(build.getAttribute('title')).resolves.toContain('開かれているため実行できません');
    await rm(lock);
  });
});

describe('変更の確認と元に戻す', () => {
  it('Build 前に差分を確認でき、Build 後は結果が画面に出て元に戻せる', async () => {
    const f = await project();
    await f.editBook(APP_BOOK, (b) => b.writeLines('util.ts', ['export const one = 1;', 'export const two = 2;']));
    running = await launch(f.root);
    const { win } = running;
    await selectBook(win, APP_BOOK);
    await win.getByRole('button', { name: '変更を確認' }).click();
    await win.getByText('+1', { exact: true }).waitFor();
    await win.getByText('Build 後のソースファイル', { exact: false }).waitFor();
    await expect(win.getByLabel('util.ts の差分').getByText('export const two = 2;').isVisible()).resolves.toBe(true);

    await win.getByRole('button', { name: 'Build', exact: true }).click();
    await win.getByRole('button', { name: '閉じたので続行' }).click();
    await waitLog(win, 'Build: 完了');
    const status = win.getByRole('status');
    await status.getByText('Build 完了').waitFor();
    await expect(status.getByText('ファイル出力 1').isVisible()).resolves.toBe(true);
    expect(await f.read('app/util.ts')).toBe('export const one = 1;\nexport const two = 2;\n');

    await win.getByRole('button', { name: 'Build を元に戻す' }).click();
    await win.getByText('を元に戻しますか？').waitFor();
    await expect(win.getByRole('dialog').getByText('util.ts').isVisible()).resolves.toBe(true);
    await win.getByRole('button', { name: '元に戻す', exact: true }).click();
    await waitLog(win, 'Build を元に戻しました: app/app.xlcode.xlsx');
    expect(await f.read('app/util.ts')).toBe('export const one = 1;\n');
    await win.getByRole('note').getByText('Build で Excel 側の変更 1 ファイル').waitFor();
    await expect(win.getByRole('button', { name: 'Build を元に戻す' }).count()).resolves.toBe(0);
  });
});

describe('ヘルプ', () => {
  it('F1 やブック画面の案内からヘルプを開き、目次と検索で探せる', async () => {
    const f = await project();
    running = await launch(f.root);
    const { win } = running;
    await selectBook(win, APP_BOOK);
    // 次の操作の「使い方」から、該当する見出しを開く
    await win.getByRole('note').getByRole('button', { name: '使い方' }).click();
    const article = win.locator('article');
    await article.getByRole('heading', { name: '基本の流れ' }).waitFor();
    // 目次から移動する
    const toc = win.getByRole('complementary', { name: '目次' });
    await toc.getByRole('button', { name: '衝突の解決', exact: true }).click();
    await expect(article.getByRole('heading', { name: '衝突の解決' }).isVisible()).resolves.toBe(true);
    // 検索で章を絞り込み、語を強調する
    await win.getByLabel('ヘルプを検索').fill('DEL_');
    await toc.getByText(/件の章が見つかりました/).waitFor();
    await expect(article.locator('mark').first().textContent()).resolves.toBe('DEL_');
    await expect(article.getByRole('heading', { name: 'OneDrive との付き合い方' }).count()).resolves.toBe(0);
    // 本文中のリンクは、検索で隠れた見出しでも移動できる
    await article.getByRole('link', { name: '衝突シート' }).first().click();
    await article.getByRole('heading', { name: '衝突の解決' }).waitFor();
    await expect(win.getByLabel('ヘルプを検索').inputValue()).resolves.toBe('');
    // エクスプローラーに戻り、F1 でまた開く
    await win.getByRole('button', { name: 'エクスプローラー' }).click();
    await win.getByRole('button', { name: 'Build', exact: true }).waitFor();
    await win.keyboard.press('F1');
    await article.getByRole('heading', { name: 'xlCode の使い方' }).waitFor();
  });
});
