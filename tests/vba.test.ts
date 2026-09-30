import { execFileSync } from 'node:child_process';
import { readdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import ExcelJS from 'exceljs';
import iconv from 'iconv-lite';
import { describe, expect, it } from 'vitest';
import { bookStatus, createBook, initProject, refreshTree, saveConfig, type VbaJob, type VbaRunner } from '../src/core';
import { DEFAULT_CONFIG } from '../src/core/config';
import { loadState } from '../src/core/state';
import { collectModules, stripExportHeader, vbaBuild, VBOM_HELP } from '../src/core/vba';
import { parseFormSheet } from '../src/core/vbaForm';
import { Book } from '../src/core/workbook';
import { fixture, type Fixture } from './helpers';

// ---- フォームの書き方 ----

const FORM = [
  'Begin UserForm UserForm1',
  '   Caption = "顧客登録"   \' タイトル',
  '   Width = 300',
  '   Height = 200.5',
  '   BackColor = &H8000000F&',
  '   Begin Label lblName',
  '      Caption = "氏名 ""様"""',
  '      Left = 12',
  '   End',
  '   Begin Frame fraType',
  '      Caption = "区分"',
  '      Begin OptionButton optA',
  '         Caption = "個人"',
  '         Value = True',
  '      End',
  '   End',
  '   Begin MultiPage mpg',
  '      Begin Page pgA',
  '         Caption = "基本"',
  '         Begin TextBox txtA',
  '            Font.Size = 11',
  '         End',
  '      End',
  '   End',
  'End',
  '',
  'Private Sub btnOK_Click()',
  '    MsgBox "OK" \' コメント',
  'End Sub',
];

describe('フォームの配置の読み取り', () => {
  it('配置とコードを分け、値の型を読み分ける', () => {
    const r = parseFormSheet('UserForm1.frm', FORM);
    expect(r.errors).toEqual([]);
    expect(r.code).toEqual(['Private Sub btnOK_Click()', '    MsgBox "OK" \' コメント', 'End Sub']);
    const f = r.form!;
    expect(f.name).toBe('UserForm1');
    expect(f.props.map((p) => [p.name, p.value, p.type])).toEqual([
      ['Caption', '顧客登録', 'string'],
      ['Width', 300, 'int'],
      ['Height', 200.5, 'double'],
      ['BackColor', 0x8000000f - 0x100000000, 'int'],
    ]);
    expect(f.children.map((c) => [c.type, c.name, c.progId])).toEqual([
      ['Label', 'lblName', 'Forms.Label.1'],
      ['Frame', 'fraType', 'Forms.Frame.1'],
      ['MultiPage', 'mpg', 'Forms.MultiPage.1'],
    ]);
    expect(f.children[0].props[0].value).toBe('氏名 "様"');
    expect(f.children[1].children[0].props.map((p) => p.value)).toEqual(['個人', true]);
    const page = f.children[2].children[0];
    expect([page.type, page.name, page.progId]).toEqual(['Page', 'pgA', '']);
    expect(page.children[0].props[0]).toMatchObject({ name: 'Font.Size', value: 11, line: 21 });
  });

  it.each([
    [['Private Sub A()', 'End Sub'], '先頭に「Begin UserForm UserForm1」'],
    [['Begin UserForm UserForm1', '  Width = 10'], '「End」が足りません'],
    [['Begin UserForm Other', 'End'], 'シート名と違います'],
    [['Begin UserForm UserForm1', '  Begin WebBrowser web', '  End', 'End'], '「WebBrowser」は使えません'],
    [['Begin UserForm UserForm1', '  Begin Page p', '  End', 'End'], 'Page は MultiPage の中にだけ'],
    [
      ['Begin UserForm UserForm1', '  Begin MultiPage m', '    Begin Label l', '    End', '  End', 'End'],
      'Page を書き',
    ],
    [['Begin UserForm UserForm1', '  Begin Label l', '    Begin Label x', '    End', '  End', 'End'], '置けません'],
    [
      ['Begin UserForm UserForm1', '  Begin Label a', '  End', '  Begin TextBox A', '  End', 'End'],
      '2 行目でも使われています',
    ],
    [['Begin UserForm UserForm1', '  Begin Label 1a', '  End', 'End'], '名前「1a」は使えません'],
    [['Begin UserForm UserForm1', '  BorderStyle = fmBorderStyleSingle', 'End'], '数値で書いてください'],
    [['Begin UserForm UserForm1', '  Name = "x"', 'End'], 'Begin 種類 名前'],
    [['Begin UserForm UserForm1', '  これは何', 'End'], '読めない行です'],
  ])('間違いは行番号付きで知らせる %#', (lines, message) => {
    const r = parseFormSheet('UserForm1.frm', lines);
    expect(r.form).toBeNull();
    expect(r.errors.join('\n')).toContain(message);
  });
});

describe('エクスポート形式の見出し', () => {
  it('VERSION 〜 END と Attribute 行を除く', () => {
    expect(
      stripExportHeader([
        'VERSION 1.0 CLASS',
        'BEGIN',
        "  MultiUse = -1  'True",
        'END',
        'Attribute VB_Name = "Class1"',
        'Attribute VB_Exposed = False',
        'Option Explicit',
      ]),
    ).toEqual({ lines: ['Option Explicit'], stripped: true });
    expect(stripExportHeader(['Option Explicit'])).toEqual({ lines: ['Option Explicit'], stripped: false });
  });
});

// ---- シートの仕分け ----

function bookWith(sheets: Record<string, string[]>): Book {
  const b = Book.create();
  for (const [name, lines] of Object.entries(sheets)) b.writeLines(name, lines);
  return b;
}

describe('シートの仕分け', () => {
  it('拡張子の無いシートは残し、.bas / .cls / .frm を VBA にする', () => {
    const c = collectModules(
      bookWith({
        '#tree': ['# tree-version: x'],
        'Agents.md': ['# Agents.md'],
        入力画面: ['氏名'],
        マスタ: ['A'],
        'Module1.bas': ['Option Explicit'],
        'Class1.cls': ['Option Explicit'],
        '入力画面.cls': ['Private Sub Worksheet_Change(ByVal Target As Range)', 'End Sub'],
        'ThisWorkbook.cls': ['Private Sub Workbook_Open()', 'End Sub'],
        'UserForm1.frm': FORM,
        'memo.txt': ['メモ'],
      }),
      DEFAULT_CONFIG,
    );
    expect(c.errors).toEqual([]);
    expect(c.keep).toEqual(['入力画面', 'マスタ']);
    expect(c.remove).toEqual(expect.arrayContaining(['#tree', 'Agents.md', 'Module1.bas', 'memo.txt']));
    expect(c.modules.map((m) => [m.sheet, m.kind, m.name, m.targetSheet])).toEqual([
      ['Module1.bas', 'standard', 'Module1', undefined],
      ['Class1.cls', 'class', 'Class1', undefined],
      ['入力画面.cls', 'sheet', '入力画面', '入力画面'],
      ['ThisWorkbook.cls', 'workbook', 'ThisWorkbook', undefined],
      ['UserForm1.frm', 'form', 'UserForm1', undefined],
    ]);
    expect(c.modules[4].code[0]).toBe('Private Sub btnOK_Click()');
    expect(c.warnings.join('\n')).toContain('memo.txt');
    expect(c.warnings.join('\n')).not.toContain('Agents.md');
  });

  it('使えないモジュール名・重複・文字列以外のセルはエラー', () => {
    const c = collectModules(
      bookWith({ 'My Module.bas': ['x'], 'a.bas': ['x'], 'A.cls': ['x'], '1st.bas': ['x'] }),
      DEFAULT_CONFIG,
    );
    const msg = c.errors.join('\n');
    expect(msg).toContain('「My Module.bas」');
    expect(msg).toContain('「1st.bas」');
    expect(msg).toContain('「a.bas」と「A.cls」は同じモジュール');
  });

  it('Attribute 行は除いて書き込み、警告する', () => {
    const c = collectModules(bookWith({ 'M.bas': ['Attribute VB_Name = "M"', 'Option Explicit'] }), DEFAULT_CONFIG);
    expect(c.modules[0].code).toEqual(['Option Explicit']);
    expect(c.warnings.join()).toContain('Attribute');
  });
});

// ---- Build ----

interface Fake {
  runner: VbaRunner;
  jobs: VbaJob[];
  /** Excel に渡された .xlsx のシート名 */
  inputSheets: string[][];
}

/** Excel の代わり。渡された .xlsx をそのまま「.xlsm」として保存する */
function fakeExcel(result: { ok: boolean; error?: string; errorKind?: 'vbom' } = { ok: true }): Fake {
  const fake: Fake = { jobs: [], inputSheets: [], runner: async () => ({ ok: true }) };
  fake.runner = async (job) => {
    fake.jobs.push(job);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load((await readFile(job.input)) as unknown as ArrayBuffer);
    fake.inputSheets.push(wb.worksheets.map((w) => w.name));
    if (!result.ok) return result;
    await writeFile(job.output, Buffer.concat([await readFile(job.input), Buffer.from(JSON.stringify(job.modules))]));
    return { ok: true };
  };
  return fake;
}

async function vbaProject(opts: { git?: boolean } = {}): Promise<Fixture> {
  const f = await fixture({}, { git: opts.git ?? true });
  await saveConfig(f.root, { ...DEFAULT_CONFIG, mode: 'vba' });
  await initProject(f.root);
  await createBook(f.root, f.root);
  return f;
}

/** フィクスチャのフォルダ名から決まるブック・ビルド結果の名前 */
function names(f: Fixture) {
  const base = path.basename(f.root);
  return { book: `${base}.xlcode.xlsx`, out: `${base}.xlsm` };
}

async function build(f: Fixture, runner: VbaRunner | null, confirmed = true) {
  return vbaBuild(f.root, f.file(names(f).book), { confirmed }, runner);
}

describe('VBA モードの Build', () => {
  it('ブックの作成: ルールのシートと Module1.bas の雛形が入り、Agents.md は VBA 用', async () => {
    const f = await vbaProject();
    const b = await Book.load(f.file(names(f).book));
    expect(b.sheetNames()).toEqual(expect.arrayContaining(['#tree', 'Agents.md', 'LocalAgents.md', 'Module1.bas']));
    expect(b.readSheet('Module1.bas').lines[0]).toBe('Option Explicit');
    expect(await f.read('Agents.md')).toContain('Begin UserForm');
  });

  it('拡張子付きのシートを除いて Excel に渡し、.xlsm と控えを書き出す', async () => {
    const f = await vbaProject();
    const { book, out } = names(f);
    await f.editBook(book, (b) => {
      b.writeLines('入力画面', ['氏名', '住所']);
      b.writeLines('Module1.bas', ['Option Explicit', '', 'Sub Hello()', '    MsgBox "こんにちは"', 'End Sub']);
      b.writeLines('Class1.cls', ['Option Explicit', 'Public Name As String']);
      b.writeLines('入力画面.cls', ['Private Sub Worksheet_Change(ByVal Target As Range)', 'End Sub']);
      b.writeLines('UserForm1.frm', FORM);
    });
    const fake = fakeExcel();
    const r = await build(f, fake.runner);
    expect(r.errors).toEqual([]);
    expect(r.status).toBe('ok');

    // Excel に渡したブックには UI のシートだけ
    expect(fake.inputSheets[0]).toEqual(['入力画面']);
    const job = fake.jobs[0];
    expect(job.modules.map((m) => [m.name, m.kind])).toEqual([
      ['Module1', 'standard'],
      ['Class1', 'class'],
      ['入力画面', 'sheet'],
      ['UserForm1', 'form'],
    ]);
    expect(job.modules[0].code).toBe('Option Explicit\r\n\r\nSub Hello()\r\n    MsgBox "こんにちは"\r\nEnd Sub');
    expect(job.modules[3].form?.children).toHaveLength(3);

    // ビルド結果はプロジェクトのフォルダに
    expect(await readdir(f.root)).toContain(out);
    // 控えは Shift_JIS・CRLF、VBE でインポートできる見出し付き
    const bas = iconv.decode(await readFile(f.file('vba/Module1.bas')), 'cp932');
    expect(bas).toBe(
      'Attribute VB_Name = "Module1"\r\nOption Explicit\r\n\r\nSub Hello()\r\n    MsgBox "こんにちは"\r\nEnd Sub\r\n',
    );
    const cls = iconv.decode(await readFile(f.file('vba/Class1.cls')), 'cp932');
    expect(cls).toMatch(/^VERSION 1\.0 CLASS\r\nBEGIN\r\n/);
    expect(cls).toContain('Attribute VB_Name = "Class1"\r\n');
    expect(iconv.decode(await readFile(f.file('vba/UserForm1.frm.txt')), 'cp932')).toContain(
      'Begin UserForm UserForm1',
    );
    expect(await readdir(f.file('vba'))).toEqual(['Class1.cls', 'Module1.bas', 'UserForm1.frm.txt', '入力画面.cls']);

    // 編集用ブックは変更しない
    const b = await Book.load(f.file(book));
    expect(b.hasSheet('Module1.bas')).toBe(true);

    // 状態: すべてビルド済み、ビルド結果は変更なし
    const st = await bookStatus(f.root, f.file(book));
    expect(st.files.every((x) => x.status === 'clean')).toBe(true);
    expect(st.vba).toMatchObject({ exists: true, changed: false });

    // Git: Build 前の自動コミット
    expect(f.git('log', '--oneline')).toContain('Build 前の自動コミット');
  });

  it('変更が無ければ控えは書き直さない。シートの変更・削除は状態に出て、削除は確認のうえ控えも消す', async () => {
    const f = await vbaProject();
    const { book } = names(f);
    await f.editBook(book, (b) => b.writeLines('Module2.bas', ['Sub A()', 'End Sub']));
    await build(f, fakeExcel().runner);

    const again = await build(f, fakeExcel().runner);
    expect(again.changes.filter((c) => c.target.startsWith('vba/'))).toEqual([]);

    await f.editBook(book, (b) => {
      b.writeLines('Module1.bas', ['Option Explicit', "' 変更"]);
      b.deleteSheet('Module2.bas');
    });
    const st = await bookStatus(f.root, f.file(book));
    expect(st.files.map((x) => [x.name, x.status])).toEqual([
      ['Module1.bas', 'excel-changed'],
      ['Module2.bas', 'removed'],
    ]);
    const confirm = await build(f, fakeExcel().runner, false);
    expect(confirm.status).toBe('confirm');
    expect(confirm.confirmations.find((c) => c.kind === 'delete')?.files).toEqual(['Module2.bas']);
    const r = await build(f, fakeExcel().runner);
    expect(r.errors).toEqual([]);
    expect(await readdir(f.file('vba'))).toEqual(['Module1.bas']);
    expect(r.changes).toContainEqual({ action: 'delete-module', target: 'Module2.bas' });
  });

  it('ビルド結果が直接変更されていたら確認し、上書き前にバックアップを取る', async () => {
    const f = await vbaProject();
    const { book, out } = names(f);
    await build(f, fakeExcel().runner);
    await writeFile(f.file(out), 'データを直接入力した');
    expect((await bookStatus(f.root, f.file(book))).vba?.changed).toBe(true);

    const confirm = await build(f, fakeExcel().runner, false);
    expect(confirm.confirmations.map((c) => c.kind)).toContain('overwrite');
    const r = await build(f, fakeExcel().runner);
    expect(r.errors).toEqual([]);
    const backups = await readdir(f.file('.xlcode/backup/_root'));
    expect(backups).toHaveLength(1);
    expect(await readFile(f.file(`.xlcode/backup/_root/${backups[0]}`), 'utf8')).toBe('データを直接入力した');
    expect((await bookStatus(f.root, f.file(book))).vba?.changed).toBe(false);
  });

  it('前からある .xlsm を初めて上書きするときも確認する', async () => {
    const f = await vbaProject();
    await writeFile(f.file(names(f).out), '既存のツール');
    const r = await build(f, fakeExcel().runner, false);
    expect(r.confirmations.find((c) => c.kind === 'overwrite')?.message).toContain('既にある');
  });

  it('コードが大きく減ったら確認する', async () => {
    const f = await vbaProject();
    const { book } = names(f);
    const long = Array.from({ length: 20 }, (_, i) => `' ${i}`);
    await f.editBook(book, (b) => b.writeLines('Module1.bas', long));
    await build(f, fakeExcel().runner);
    await f.editBook(book, (b) => b.writeLines('Module1.bas', long.slice(0, 5)));
    const r = await build(f, fakeExcel().runner, false);
    expect(r.confirmations.map((c) => c.kind)).toContain('shrink');
  });

  it('Shift_JIS で表せない文字があれば、Excel を呼ばずに止める', async () => {
    const f = await vbaProject();
    await f.editBook(names(f).book, (b) => b.writeLines('Module1.bas', ['MsgBox "😀"']));
    const fake = fakeExcel();
    const r = await build(f, fake.runner);
    expect(r.status).toBe('error');
    expect(r.errors.join()).toContain('Shift_JIS');
    expect(fake.jobs).toHaveLength(0);
  });

  it('シートの間違い（フォームの書き方など）は、Excel を呼ばずに止める', async () => {
    const f = await vbaProject();
    await f.editBook(names(f).book, (b) => b.writeLines('UserForm1.frm', ['Begin UserForm UserForm1']));
    const fake = fakeExcel();
    const r = await build(f, fake.runner);
    expect(r.errors.join()).toContain('「End」が足りません');
    expect(fake.jobs).toHaveLength(0);
  });

  it('Excel で失敗したら、ビルド結果・控え・状態は変えない', async () => {
    const f = await vbaProject();
    const { book, out } = names(f);
    await build(f, fakeExcel().runner);
    const before = await readFile(f.file(out));
    const stateBefore = JSON.stringify(await loadState(f.root));
    await f.editBook(book, (b) => b.writeLines('Module1.bas', ['Sub Changed()', 'End Sub']));
    const r = await build(f, fakeExcel({ ok: false, error: 'アクセス拒否', errorKind: 'vbom' }).runner);
    expect(r.status).toBe('error');
    expect(r.errors).toContain(VBOM_HELP);
    expect(await readFile(f.file(out))).toEqual(before);
    expect(iconv.decode(await readFile(f.file('vba/Module1.bas')), 'cp932')).not.toContain('Changed');
    expect(JSON.stringify(await loadState(f.root))).toBe(stateBefore);
  });

  it('ソースコードモードの Build / Sync は動かない', async () => {
    const f = await vbaProject();
    const { build: sourceBuild, sync } = await import('../src/core');
    const r1 = await sourceBuild(f.root, f.file(names(f).book), { confirmed: true });
    const r2 = await sync(f.root, f.file(names(f).book), { confirmed: true });
    expect(r1.errors.join()).toContain('VBA モードのプロジェクトです');
    expect(r2.errors.join()).toContain('VBA モードのプロジェクトです');
  });

  it('Windows 以外（Excel が無い）では、確認の後にエラー', async () => {
    const f = await vbaProject();
    const r = await build(f, null);
    expect(r.status).toBe('error');
    expect(r.errors.join()).toContain('Windows とデスクトップ版 Excel');
  });

  it('編集用ブックがデスクトップ版で開かれていれば止める', async () => {
    const f = await vbaProject();
    await writeFile(f.file(`~$${names(f).book}`), '');
    const r = await build(f, fakeExcel().runner);
    expect(r.errors.join()).toContain('開かれています');
  });

  it('UI のシートが 1 枚も無ければ、白紙のシートを 1 枚置く', async () => {
    const f = await vbaProject();
    const fake = fakeExcel();
    await build(f, fake.runner);
    expect(fake.inputSheets[0]).toEqual(['Sheet1']);
  });

  it('Git: 控え（vba/）の変更も Build 前に自動コミットする', async () => {
    const f = await vbaProject();
    await build(f, fakeExcel().runner);
    await writeFile(f.file('vba/Module1.bas'), 'エディタで直した');
    const r = await build(f, fakeExcel().runner, false);
    expect(r.confirmations.find((c) => c.kind === 'uncommitted')?.files).toContain('vba/Module1.bas');
    await build(f, fakeExcel().runner);
    expect(f.git('log', '--format=%s', '-1', '--', 'vba/Module1.bas')).toContain('Build 前の自動コミット');
  });

  it('控え（vba/）があれば、ブックを作り直したときにシートへ戻す', async () => {
    const f = await vbaProject();
    const { book } = names(f);
    await f.editBook(book, (b) => {
      b.writeLines('Class1.cls', ['Public Name As String']);
      b.writeLines('UserForm1.frm', FORM);
    });
    await build(f, fakeExcel().runner);
    const { rm } = await import('node:fs/promises');
    await rm(f.file(book));
    await createBook(f.root, f.root);
    const b = await Book.load(f.file(book));
    expect(b.readSheet('Class1.cls').lines).toEqual(['Public Name As String']);
    expect(b.readSheet('UserForm1.frm').lines).toEqual(FORM.filter((_, i) => i < FORM.length));
    expect(b.readSheet('Module1.bas').lines).toEqual(['Option Explicit']);
  });

  it('Refresh Tree で LocalAgents.md もシートへ配布する', async () => {
    const f = await vbaProject();
    await f.write('LocalAgents.md', '# LocalAgents.md\n\n- 入力画面の B2 が氏名\n');
    await refreshTree(f.root);
    const b = await Book.load(f.file(names(f).book));
    expect(b.readSheet('LocalAgents.md').lines).toContain('- 入力画面の B2 が氏名');
    const again = await refreshTree(f.root);
    expect(again.books[0].unchanged).toBe(true);
  });
});

// PowerShell（pwsh）があれば、実際のスクリプトを偽の Excel で動かす
const pwsh = (() => {
  try {
    execFileSync('pwsh', ['-NoProfile', '-Command', 'exit 0'], { stdio: 'ignore' });
    return true;
  } catch {
    return false;
  }
})();

describe.skipIf(!pwsh)('Excel を操作する PowerShell スクリプト（偽の Excel で実行）', async () => {
  const { runVbaJob, VBA_SCRIPT } = await import('../src/main/excelCom');
  const prelude = path.resolve(import.meta.dirname, 'fixtures/fakeExcel.ps1');
  const script = VBA_SCRIPT.replace(
    "$ErrorActionPreference = 'Stop'",
    `. '${prelude}'\n$ErrorActionPreference = 'Stop'`,
  ).replace('New-Object -ComObject Excel.Application', 'New-FakeExcel');

  async function runJob(
    modules: VbaJob['modules'],
    env: Record<string, string> = {},
    timeout?: number,
    references: VbaJob['references'] = [],
  ) {
    const f = await fixture({ 'input.xlsx': 'x' });
    const log = f.file('excel.json');
    const saved = { ...process.env };
    Object.assign(process.env, { FAKE_EXCEL_SHEETS: '入力画面|マスタ', FAKE_EXCEL_LOG: log }, env);
    try {
      const r = await runVbaJob(
        { input: f.file('input.xlsx'), output: f.file('out.xlsm'), modules, references },
        { exe: 'pwsh', script, timeout },
      );
      const out = await readFile(f.file('out.xlsm'), 'utf8').catch(() => null);
      const excel = await readFile(log, 'utf8').catch(() => null);
      return { r, out: out && JSON.parse(out), excel: excel && JSON.parse(excel) };
    } finally {
      process.env = saved;
    }
  }

  const code = (lines: string[]) => lines.join('\r\n');

  it('モジュール・シートのコード・フォームを書き込み、.xlsm で保存して Excel を終了する', async () => {
    const form = parseFormSheet('UserForm1.frm', FORM);
    const { r, out, excel } = await runJob([
      {
        sheet: 'Module1.bas',
        name: 'Module1',
        kind: 'standard',
        code: code(['Option Explicit', 'Sub A()', 'End Sub']),
      },
      { sheet: 'Class1.cls', name: 'Class1', kind: 'class', code: code(['Public X As Long']) },
      { sheet: 'Sheet2.cls', name: 'Sheet2', kind: 'class', code: code(["' コード名で指定"]) },
      { sheet: '入力画面.cls', name: '入力画面', kind: 'sheet', targetSheet: '入力画面', code: code(["' シート"]) },
      { sheet: 'ThisWorkbook.cls', name: 'ThisWorkbook', kind: 'workbook', code: code(["' ブック"]) },
      { sheet: 'UserForm1.frm', name: 'UserForm1', kind: 'form', code: code(form.code), form: form.form! },
    ]);
    expect(r).toEqual({ ok: true });
    expect(out.format).toBe(52);
    expect(out.autoSave).toBe(false);
    expect(excel).toEqual({
      visible: false,
      displayAlerts: false,
      enableEvents: false,
      automationSecurity: 3,
      closed: true,
    });
    const comp = (name: string) => out.components.find((c: { name: string }) => c.name === name);
    expect(comp('Module1')).toMatchObject({ type: 1, code: ['Option Explicit', 'Sub A()', 'End Sub'] });
    expect(comp('Class1')).toMatchObject({ type: 2, code: ['Public X As Long'] });
    // 既にあるシートのコードモジュールの名前（コード名）なら、クラスを作らずそこに書く
    expect(comp('Sheet2')).toMatchObject({ type: 100, code: ["' コード名で指定"] });
    expect(comp('Sheet1')).toMatchObject({ type: 100, code: ["' シート"] });
    expect(comp('ThisWorkbook')).toMatchObject({ type: 100, code: ["' ブック"] });
    const uf = comp('UserForm1');
    expect(uf.type).toBe(3);
    expect(uf.props).toMatchObject({ Caption: '顧客登録', Width: 300, Height: 200.5 });
    expect(uf.designer.BackColor).toBe(0x8000000f - 0x100000000);
    expect(uf.code[0]).toBe('Private Sub btnOK_Click()');
    const [lbl, fra, mpg] = uf.designer.Controls;
    expect(lbl).toMatchObject({ ProgId: 'Forms.Label.1', Name: 'lblName', Caption: '氏名 "様"', Left: 12 });
    expect(fra.Controls[0]).toMatchObject({ Name: 'optA', Caption: '個人', Value: true });
    // MultiPage は最初からある 2 ページを使い回し、余りは消す
    expect(mpg.Pages).toHaveLength(1);
    expect(mpg.Pages[0]).toMatchObject({ Name: 'pgA', Caption: '基本' });
    expect(mpg.Pages[0].Controls[0]).toMatchObject({ Name: 'txtA', Font: { Size: 11 } });
  }, 60_000);

  it('VBA プロジェクトへのアクセスが信頼されていなければ vbom で失敗し、Excel は終了する', async () => {
    const { r, out, excel } = await runJob([], { FAKE_EXCEL_NO_VBOM: '1' });
    expect(r.ok).toBe(false);
    expect(r.errorKind).toBe('vbom');
    expect(out).toBeNull();
    expect(excel.closed).toBe(true);
  }, 60_000);

  it('Excel が無ければ no-excel', async () => {
    const { r } = await runJob([], { FAKE_EXCEL_MISSING: '1' });
    expect(r).toMatchObject({ ok: false, errorKind: 'no-excel' });
    expect(r.error).toContain('デスクトップ版 Excel を起動できません');
  }, 60_000);

  it('書き込めなかったら、どのシートの何行目かを知らせる', async () => {
    const { r } = await runJob([{ sheet: '無い.cls', name: '無い', kind: 'sheet', targetSheet: '無い', code: '' }]);
    expect(r.ok).toBe(false);
    expect(r.error).toContain('無い.cls');
    expect(r.error).toContain('コードモジュールが見つかりません');
  }, 60_000);

  it('時間切れなら中止する', async () => {
    const { r } = await runJob([], { FAKE_EXCEL_SLEEP: '30' }, 8_000);
    expect(r).toMatchObject({ ok: false, errorKind: 'timeout' });
  }, 60_000);
});

describe('PowerShell が無いとき', () => {
  it('実行できない旨を返す', async () => {
    const { runVbaJob } = await import('../src/main/excelCom');
    const f = await fixture();
    const r = await runVbaJob(
      { input: f.file('a.xlsx'), output: f.file('a.xlsm'), modules: [], references: [] },
      { exe: 'no-such-powershell-xlcode' },
    );
    expect(r).toMatchObject({ ok: false, errorKind: 'other' });
    expect(r.error).toContain('PowerShell を実行できません');
  });
});
