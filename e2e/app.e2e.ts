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
    // ルールはエクスプローラーの階層の中にある（ルートに Agents.md、各フォルダに LocalAgents.md）
    await win.getByRole('treeitem', { name: 'Agents.md', exact: true }).click();
    const editor = win.getByRole('textbox', { name: 'Agents.md' });
    await editor.click();
    await editor.press('Control+End');
    await editor.pressSequentially('\n- 追加したルール');
    await win.getByRole('treeitem', { name: 'app/LocalAgents.md', exact: true }).click();
    await win.getByRole('textbox', { name: 'app/LocalAgents.md' }).waitFor();
    // 編集中のタブは固定され、● が付く
    await expect(win.getByRole('tab', { name: 'Agents.md', exact: true }).count()).resolves.toBe(1);
    await win.getByRole('tab', { name: 'Agents.md', exact: true }).click();
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
    await win.getByRole('treeitem', { name: 'Agents.md', exact: true }).click();
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
    // サイドバーの「置き場所」からも設定を開ける
    await win.getByRole('button', { name: /^置き場所:/ }).click();
    await win.getByRole('button', { name: 'フォルダを選ぶ…' }).click();
    await expect.poll(() => win.getByLabel('ブックの置き場所').inputValue()).toBe(xl);
    await win.getByRole('button', { name: '変更してブック 2 冊を移動' }).click();
    await win.getByText('既存のブック 2 冊を、同じフォルダ構成のまま移動します').waitFor();
    await win.getByRole('button', { name: '移動して変更' }).click();
    await waitLog(win, /ブックの置き場所を変更しました/);
    expect(await readdir(path.join(xl, 'app'))).toEqual(['app.xlcode.xlsx']);
    expect((await readdir(f.file('app'))).filter((n) => n.endsWith('.xlsx'))).toEqual([]);

    await f.editBook(path.relative(f.root, path.join(xl, 'app', 'app.xlcode.xlsx')), (b) =>
      b.writeLines('util.ts', ['export const one = 1;', 'export const moved = 2;']),
    );
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
    // .xlcodeignore も .gitignore も無いので、全ファイルがシートになる旨の確認が出る
    await win.getByText('.gitignore もありません', { exact: false }).waitFor();
    await win.getByRole('button', { name: 'このまま作成' }).click();
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
    await win.getByRole('button', { name: 'このまま作成' }).click();
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
    await win.getByText('3 個のフォルダでシートになるファイル: 2 件').waitFor();
    await win.getByRole('button', { name: 'このまま作成' }).click();
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
    await win.getByRole('button', { name: 'このまま作成' }).click();
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
    // ヘルプのタブを閉じるとブックに戻り、F1 でまた開く
    await win.keyboard.press('Control+w');
    await win.getByRole('button', { name: 'Build', exact: true }).waitFor();
    await win.keyboard.press('F1');
    await article.getByRole('heading', { name: 'xlCode の使い方' }).waitFor();
  });
});

describe('タブ', () => {
  it('複数のブックとドキュメントをタブで開き、キーで切り替え・閉じる', async () => {
    const f = await project();
    running = await launch(f.root);
    const { win } = running;
    const tabs = win.getByRole('tablist', { name: '開いているタブ' });
    // 1 回押しはプレビュー（次のプレビューに置き換わる）。2 回押すと固定される
    await selectBook(win, APP_BOOK);
    await tabs.getByRole('tab', { name: 'app.xlcode.xlsx' }).waitFor();
    await selectBook(win, API_BOOK);
    await tabs.getByRole('tab', { name: 'api.xlcode.xlsx' }).waitFor();
    await expect(tabs.getByRole('tab').count()).resolves.toBe(1);
    await win.getByRole('treeitem', { name: 'api.xlcode.xlsx' }).locator('div').first().dblclick();
    await selectBook(win, APP_BOOK);
    await tabs.getByRole('tab', { name: 'app.xlcode.xlsx' }).waitFor();
    await expect(tabs.getByRole('tab').count()).resolves.toBe(2);
    await expect(tabs.getByRole('tab', { name: 'app.xlcode.xlsx' }).getAttribute('aria-selected')).resolves.toBe(
      'true',
    );
    // ドキュメントもタブ（app は 2 回押して固定しておく）
    await win.getByRole('treeitem', { name: 'app.xlcode.xlsx' }).locator('div').first().dblclick();
    await win.getByRole('treeitem', { name: 'app/LocalAgents.md', exact: true }).click();
    await win.getByRole('textbox', { name: 'app/LocalAgents.md' }).waitFor();
    await expect(tabs.getByRole('tab').count()).resolves.toBe(3);
    // Ctrl+Tab で次へ、Ctrl+W で閉じる
    await win.keyboard.press('Control+Tab');
    await expect(tabs.getByRole('tab', { name: 'api.xlcode.xlsx' }).getAttribute('aria-selected')).resolves.toBe(
      'true',
    );
    await win.keyboard.press('Control+w');
    await expect(tabs.getByRole('tab').count()).resolves.toBe(2);
    await expect(tabs.getByRole('tab', { name: 'app.xlcode.xlsx' }).getAttribute('aria-selected')).resolves.toBe(
      'true',
    );
    // Ctrl+Shift+T で開き直す。Alt+1 で 1 番目へ
    await win.keyboard.press('Control+Shift+t');
    await expect(tabs.getByRole('tab').count()).resolves.toBe(3);
    await win.keyboard.press('Alt+1');
    await expect(tabs.getByRole('tab').first().getAttribute('aria-selected')).resolves.toBe('true');
    // Ctrl+B でサイドバーを隠す
    await win.keyboard.press('Control+b');
    await expect(win.getByRole('tree', { name: 'ブック', exact: true }).count()).resolves.toBe(0);
    await win.keyboard.press('Control+b');
    await win.getByRole('tree', { name: 'ブック', exact: true }).waitFor();
  });

  it('保存していないドキュメントのタブを閉じるときは確認し、保存できる', async () => {
    const f = await project();
    running = await launch(f.root);
    const { win } = running;
    await win.getByRole('treeitem', { name: 'Agents.md', exact: true }).click();
    const editor = win.getByRole('textbox', { name: 'Agents.md' });
    await editor.click();
    await editor.press('Control+End');
    await editor.pressSequentially('\n- タブを閉じて保存');
    await win.keyboard.press('Control+w');
    await win.getByText('Agents.md を保存しますか？').waitFor();
    await win.getByRole('dialog').getByRole('button', { name: '保存', exact: true }).click();
    await waitLog(win, '保存しました: Agents.md');
    expect(await f.read('Agents.md')).toContain('- タブを閉じて保存');
    await expect(win.getByRole('tab', { name: 'Agents.md', exact: true }).count()).resolves.toBe(0);
  });
});

describe('コマンドパレットとショートカット', () => {
  it('Ctrl+P でブックを名前で開き、Ctrl+Shift+P でコマンドを実行する', async () => {
    const f = await project();
    running = await launch(f.root);
    const { win } = running;
    await win.keyboard.press('Control+p');
    const box = win.getByLabel('ブック・ドキュメントの名前で検索（先頭に > でコマンド）');
    await box.waitFor();
    await box.fill('api');
    await win.getByRole('option', { name: /api\.xlcode\.xlsx/ }).waitFor();
    await box.press('Enter');
    await win.getByRole('tablist', { name: '開いているタブ' }).getByRole('tab', { name: 'api.xlcode.xlsx' }).waitFor();
    // コマンドパレット: Build を実行すると、Web 版を閉じたかの確認が出る（= Build が始まった）
    await win.keyboard.press('Control+Shift+p');
    const cmd = win.getByRole('dialog', { name: 'コマンドパレット' });
    await cmd.waitFor();
    await expect(cmd.getByRole('textbox').inputValue()).resolves.toBe('>');
    await cmd.getByRole('textbox').pressSequentially('build');
    await cmd.getByRole('option', { name: /Build（Excel → ソース）/ }).waitFor();
    await cmd.getByRole('textbox').press('Enter');
    await win.getByText('Web 版 Excel でこのブックを閉じましたか？').waitFor();
    await win.getByRole('button', { name: 'キャンセル' }).click();
    // ショートカットでも同じ（キャンセル後の再読み込みが終わるのを待つ。処理中はキーが効かない）
    await expect.poll(() => win.getByText('Build 中', { exact: true }).count()).toBe(0);
    await win.keyboard.press('Control+Shift+b');
    await win.getByText('Web 版 Excel でこのブックを閉じましたか？').waitFor();
    await win.keyboard.press('Escape');
    // Ctrl+, で設定、Ctrl+Shift+M で問題パネル
    await win.keyboard.press('Control+,');
    await win.getByRole('tab', { name: '設定' }).waitFor();
    await win.keyboard.press('Control+Shift+m');
    await win.getByText('問題は検出されていません。').waitFor();
  });
});

describe('ブックにしないファイル（.xlcodeignore）', () => {
  it('エクスプローラーから .xlcodeignore を編集すると、そのファイルはシートの対象から外れる', async () => {
    const f = await project();
    running = await launch(f.root);
    const { win } = running;
    await win.getByRole('treeitem', { name: '.xlcodeignore', exact: true }).click();
    const editor = win.getByRole('textbox', { name: '.xlcodeignore' });
    await editor.waitFor();
    await win.getByText('ブックにしないファイル・フォルダ').first().waitFor();
    await editor.click();
    await editor.press('Control+End');
    await editor.pressSequentially('\napp/util.ts\n');
    await editor.press('Control+s');
    await waitLog(win, '保存しました: .xlcodeignore');
    await selectBook(win, APP_BOOK);
    // 問題パネルに警告として出る
    await win.keyboard.press('Control+Shift+m');
    await win
      .getByText('「util.ts」は .gitignore / .xlcodeignore の対象なのでスキップします', { exact: false })
      .first()
      .waitFor();
    const st = await bookStatus(f.root, f.file(APP_BOOK));
    expect(st.files.map((x) => x.name)).not.toContain('util.ts');
    expect(await f.read('.gitignore')).not.toContain('util.ts');
    // 設定画面にも反映される
    await win.keyboard.press('Control+,');
    await win.getByText('1 件: app/util.ts').waitFor();
  });

  it('置き場所は「移動しない」でも変えられる', async () => {
    const f = await project();
    const { mkdtemp, readdir } = await import('node:fs/promises');
    const { tmpdir } = await import('node:os');
    const path = await import('node:path');
    const xl = await mkdtemp(path.join(tmpdir(), 'xlfolder-'));
    running = await launch(f.root);
    const { win } = running;
    await win.keyboard.press('Control+,');
    const input = win.getByLabel('ブックの置き場所');
    await input.fill(xl);
    await win.getByLabel('既存のブックを移動しない', { exact: false }).check();
    await win.getByRole('button', { name: '変更', exact: true }).click();
    await win.getByText('ブックは移動しません。').waitFor();
    await win.getByRole('button', { name: '変更', exact: true }).last().click();
    await waitLog(win, /ブックの置き場所を変更しました/);
    expect(await readdir(xl)).toEqual([]);
    expect((await readdir(f.file('app'))).filter((n) => n.endsWith('.xlsx'))).toEqual(['app.xlcode.xlsx']);
    await win.getByText('ブックがありません').first().waitFor();
  });
});

describe('ブック作成前の除外の確認', () => {
  it('.xlcodeignore に指定が無ければ .gitignore に従ってよいか確認し、先に編集もできる', async () => {
    const f = await project();
    running = await launch(f.root);
    const { win } = running;
    // ルートにはまだブックが無い
    await win.getByText('ブック未作成のディレクトリ').click();
    await win.getByTitle('ルート にブックを作成').click();
    const dialog = win.getByRole('dialog', { name: 'ブックにしないファイルの指定がありません' });
    await dialog.waitFor();
    // 初期化で xlCode 用の除外も .gitignore に足されているので 4 件
    await dialog.getByText(/\.gitignore の除外（4 件: node_modules\//).waitFor();
    // ルート直下でシートになるのは .gitignore だけ（Agents.md と .xlcodeignore は常に対象外）
    await dialog.getByText('シートになるファイル: 1 件', { exact: false }).waitFor();
    // 先に編集: 雛形が入った .xlcodeignore のタブが開く
    await dialog.getByRole('button', { name: '先に .xlcodeignore を編集' }).click();
    const editor = win.getByRole('textbox', { name: '.xlcodeignore' });
    await editor.waitFor();
    await expect(editor.inputValue()).resolves.toContain('.gitignore と同じ書き方');
    await editor.click();
    await editor.press('Control+End');
    await editor.pressSequentially('\n.gitignore\n');
    await editor.press('Control+s');
    await waitLog(win, '保存しました: .xlcodeignore');
    // 指定があれば確認は出ず、そのまま（種類が未設定なのでその選択のあと）ブック作成の確認へ進む
    await win.getByTitle('ルート にブックを作成').click();
    await win.getByText('プロジェクトの種類を選んでください').waitFor();
    await expect(win.getByRole('dialog', { name: 'ブックにしないファイルの指定がありません' }).count()).resolves.toBe(
      0,
    );
    await win.getByRole('button', { name: 'ソースコード', exact: true }).click();
    await win.getByText('ブックを作成しますか？').waitFor();
    await win.getByRole('button', { name: '作成', exact: true }).click();
    await waitLog(win, /ブックを作成しました/);
    const { readdir } = await import('node:fs/promises');
    const path = await import('node:path');
    const name = `${path.basename(f.root)}.xlcode.xlsx`;
    expect(await readdir(f.root)).toContain(name);
    const { Book } = await import('../src/core/workbook');
    const b = await Book.load(f.file(name));
    expect(b.sheetNames()).not.toContain('.gitignore');
    expect(b.sheetNames()).not.toContain('.xlcodeignore');
  });

  it('「今後確認しない」を選ぶと、そのプロジェクトでは聞かれなくなる', async () => {
    const { fixture } = await import('../tests/helpers');
    const f = await fixture({ 'a/x.ts': 'export const x = 1;\n', 'b/y.ts': 'export const y = 1;\n' }, { git: true });
    running = await launch(f.root);
    const { win } = running;
    await win.getByTitle('a にブックを作成').click();
    const dialog = win.getByRole('dialog', { name: 'ブックにしないファイルの指定がありません' });
    await dialog.getByLabel('このプロジェクトでは今後確認しない').check();
    await dialog.getByRole('button', { name: 'このまま作成' }).click();
    await win.getByRole('button', { name: 'ソースコード', exact: true }).click();
    await win.getByRole('button', { name: 'プロジェクトの中に置く' }).click();
    await win.getByRole('button', { name: '作成', exact: true }).click();
    await waitLog(win, /ブックを作成しました/);
    await win.getByTitle('b にブックを作成').click();
    await win.getByText('ブックを作成しますか？').waitFor();
    await expect(dialog.count()).resolves.toBe(0);
    await win.getByRole('button', { name: 'キャンセル' }).click();
  });
});

describe('左右のサイドバー', () => {
  it('プロジェクトを開いていないときは、サイドバーが「プロジェクトを開く」になる', async () => {
    const f = await project();
    running = await launch(f.root, 'both', {}, { openProject: false });
    const { win } = running;
    const side = win.getByRole('complementary').first();
    await win.getByRole('button', { name: 'プロジェクトを開く', exact: true }).first().waitFor();
    // 最近使用したプロジェクトからも開ける（サイドバー側のボタン）
    await expect(side.getByText('最近使用したプロジェクト').count()).resolves.toBeGreaterThan(0);
    // 左のサイドバーはプロジェクトを開く前でも Ctrl+B で開閉できる
    await win.keyboard.press('Control+b');
    await expect(win.getByRole('button', { name: 'プロジェクトを開く', exact: true }).count()).resolves.toBe(1);
    await win.keyboard.press('Control+b');
    await win.getByRole('button', { name: 'プロジェクトを開く', exact: true }).first().waitFor();
  });

  it('タブバーのボタンで左右を開閉し、右のパネルで指示を組み立て、メモを書ける', async () => {
    const f = await project();
    running = await launch(f.root);
    const { win } = running;
    await selectBook(win, APP_BOOK);
    await win.getByRole('button', { name: /^左のサイドバーを隠す/ }).click();
    await expect(win.getByRole('tree', { name: 'ブック', exact: true }).count()).resolves.toBe(0);
    await win.getByRole('button', { name: /^左のサイドバーを表示/ }).click();
    await win.getByRole('tree', { name: 'ブック', exact: true }).waitFor();

    // 右のパネル: 指示。テンプレートを作成欄に足し、ファイルを選ぶと名前が埋まる
    await win.keyboard.press('Control+Alt+b');
    const right = win.getByRole('complementary', { name: '右のパネル' });
    await right.waitFor();
    await win
      .getByRole('row', { name: /util\.ts/ })
      .first()
      .click();
    await right.getByText('util.ts', { exact: true }).first().waitFor();
    // 「編集」のグループは閉じているので開く
    await right.getByRole('button', { name: /^編集/ }).click();
    await right.getByRole('button', { name: 'ファイルを直す を作成欄に足す' }).click();
    const box = right.getByLabel('指示の作成欄');
    await expect(box.inputValue()).resolves.toContain('util.ts シートの内容を次のように変更してください');
    // 開いているグループは 1 つだけなので「基本」を開き直す
    await right.getByRole('button', { name: /^基本/ }).click();
    await right.getByRole('button', { name: 'ルールを読ませる を作成欄に足す' }).click();
    await expect(box.inputValue()).resolves.toMatch(/変更してください[\s\S]*Agents\.md シート/);
    await right.getByRole('button', { name: 'クリア' }).click();
    await expect(box.inputValue()).resolves.toBe('');

    // メモ: 入力をやめると .xlcode/notes.md に保存される
    await right.getByRole('tab', { name: 'メモ' }).click();
    const memo = right.getByLabel('作業メモ');
    await memo.waitFor();
    await memo.fill('明日は api を直す');
    await expect.poll(() => f.read('.xlcode/notes.md').catch(() => ''), { timeout: 5000 }).toBe('明日は api を直す');
    // 閉じる
    await right.getByRole('button', { name: /右のパネルを閉じる/ }).click();
    await expect(right.count()).resolves.toBe(0);
  });

  it('左のサイドバーに Git と OneDrive の状態が出る', async () => {
    const f = await project();
    running = await launch(f.root);
    const { win } = running;
    const side = win.getByRole('complementary').first();
    await side.getByText('Git — main').click();
    await side.getByText('setup', { exact: true }).waitFor();
    await side.getByText('変更なし').waitFor();
    await f.write('app/App.tsx', 'export const App = () => <p>dirty</p>;\n');
    await win.evaluate(() => window.dispatchEvent(new Event('focus')));
    await side.getByText('未コミット 1').waitFor();
    await side.getByText('未コミット 1').click();
    await side.getByText('app/App.tsx', { exact: true }).waitFor();
    await side.getByText('OneDrive の同期状態').click();
    // Linux では Windows だけの機能である旨が出る
    await side.getByText('Windows でだけ表示します', { exact: false }).waitFor();
  });
});

describe('コードの閲覧', () => {
  it('エクスプローラーのシートを押すと閲覧のタブで開き、Markdown は編集できる', async () => {
    const f = await project();
    await f.editBook(APP_BOOK, (b) => b.writeLines('util.ts', ['export const one = 1;', 'export const two = 2;']));
    running = await launch(f.root);
    const { win } = running;
    await selectBook(win, APP_BOOK);
    await win.getByRole('treeitem', { name: 'util.ts', exact: true }).click();
    const tabs = win.getByRole('tablist', { name: '開いているタブ' });
    await tabs.getByRole('tab', { name: /util\.ts/ }).waitFor();
    // 開いていた階層（ブックのシート一覧）は閉じず、選んだシートの行が選択される
    await expect(
      win.getByRole('treeitem', { name: 'util.ts', exact: true }).getAttribute('aria-selected'),
    ).resolves.toBe('true');
    await expect(win.getByRole('treeitem', { name: 'App.tsx', exact: true }).count()).resolves.toBe(1);
    // LocalAgents.md（ブックの中のシート）を開いても同じ
    await win.getByRole('treeitem', { name: 'app/LocalAgents.md', exact: true }).click();
    await win.getByRole('textbox', { name: 'app/LocalAgents.md' }).waitFor();
    await expect(win.getByRole('treeitem', { name: 'App.tsx', exact: true }).count()).resolves.toBe(1);
    await win.getByRole('treeitem', { name: 'util.ts', exact: true }).click();
    // 既定はソース側。Excel 側に切り替えると Copilot の変更が見える
    const src = win.getByLabel('util.ts のソース側の内容');
    await src.waitFor();
    // 内容は読み込みが終わってから出る
    await expect.poll(() => src.innerText()).toContain('export const one = 1;');
    await expect(src.innerText()).resolves.not.toContain('two');
    await win.getByRole('tab', { name: 'Excel 側' }).click();
    await expect(win.getByLabel('util.ts のExcel 側の内容').innerText()).resolves.toContain('export const two = 2;');
    // 編集はできず、既定のアプリ / VS Code / フォルダ のボタンがある
    await expect(win.getByRole('button', { name: '既定のアプリで開く' }).isVisible()).resolves.toBe(true);
    await expect(win.getByRole('button', { name: 'VS Code で開く' }).isVisible()).resolves.toBe(true);
    await expect(win.getByRole('textbox', { name: 'util.ts' }).count()).resolves.toBe(0);
    // ブック画面のファイル名からも開ける（閲覧はプレビューなので、ブックを選び直すと置き換わる）。
    // Markdown（LocalAgents.md）は編集のタブ
    await selectBook(win, APP_BOOK);
    await tabs.getByRole('tab', { name: 'app.xlcode.xlsx' }).waitFor();
    await win
      .getByRole('row', { name: /LocalAgents\.md/ })
      .getByRole('button', { name: 'LocalAgents.md', exact: true })
      .click();
    await win.getByRole('textbox', { name: 'app/LocalAgents.md' }).waitFor();
  });
});

describe('ドキュメントの編集の細かい動き', () => {
  it('開いただけ（1 文字も入力していない）なら、閉じるときに保存を聞かない', async () => {
    const { fixture } = await import('../tests/helpers');
    // .xlcodeignore が無い古いプロジェクト
    const f = await fixture(
      { 'a.ts': 'export {};\n', 'Agents.md': '# Agents.md\r\n\r\n- CRLF のルール\r\n' },
      { git: true },
    );
    running = await launch(f.root);
    const { win } = running;
    const tabs = win.getByRole('tablist', { name: '開いているタブ' });
    // 雛形を見せている .xlcodeignore
    await win.getByRole('treeitem', { name: '.xlcodeignore', exact: true }).click();
    const ign = win.getByRole('textbox', { name: '.xlcodeignore' });
    await expect(ign.inputValue()).resolves.toContain('.gitignore と同じ書き方');
    await expect(win.getByRole('button', { name: '保存', exact: true }).isEnabled()).resolves.toBe(true);
    await win.keyboard.press('Control+w');
    await expect(win.getByRole('dialog').count()).resolves.toBe(0);
    await expect(tabs.getByRole('tab', { name: '.xlcodeignore' }).count()).resolves.toBe(0);
    // 既存の Agents.md（CRLF）も同じ。1 文字直して保存すると CRLF のまま
    await win.getByRole('treeitem', { name: 'Agents.md', exact: true }).click();
    const editor = win.getByRole('textbox', { name: 'Agents.md' });
    await editor.waitFor();
    await expect(win.getByRole('button', { name: '保存', exact: true }).isEnabled()).resolves.toBe(false);
    await win.keyboard.press('Control+w');
    await expect(win.getByRole('dialog').count()).resolves.toBe(0);
    await win.getByRole('treeitem', { name: 'Agents.md', exact: true }).click();
    await editor.click();
    await editor.press('Control+End');
    await editor.pressSequentially('- 追加');
    await editor.press('Control+s');
    await waitLog(win, '保存しました: Agents.md');
    const { readFile } = await import('node:fs/promises');
    const raw = (await readFile(f.file('Agents.md'))).toString('utf8');
    expect(raw).toBe('# Agents.md\r\n\r\n- CRLF のルール\r\n- 追加');
  });
});
