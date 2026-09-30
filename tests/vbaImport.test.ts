import { execFileSync } from 'node:child_process';
import { readdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import ExcelJS from 'exceljs';
import iconv from 'iconv-lite';
import { describe, expect, it } from 'vitest';
import { bookStatus, initProject, saveConfig, type VbaImporter, type VbaRunner } from '../src/core';
import { DEFAULT_CONFIG } from '../src/core/config';
import { modulesFromBook, vbaBuild } from '../src/core/vba';
import { parseFormSheet } from '../src/core/vbaForm';
import { createBookFromTool, modulesToSheets, renderForm, type ImportedModule } from '../src/core/vbaImport';
import { parseRefs, renderRefs } from '../src/core/vbaRefs';
import { Book } from '../src/core/workbook';
import { fixture, type Fixture } from './helpers';

const SCRIPTING = '{420B2830-E718-11CF-893D-00A0C9054228}';

describe('参照設定（#refs）', () => {
  it('GUID・バージョン・説明を読み、コメント・空行・重複は無視する', () => {
    const r = parseRefs('#refs', [
      "' コメント",
      '',
      `${SCRIPTING} 1.0 Microsoft Scripting Runtime`,
      '3f4daca7-160d-11d2-a8e9-00104b365c9f',
      `${SCRIPTING} 1.0 重複`,
    ]);
    expect(r.errors).toEqual([]);
    expect(r.refs).toEqual([
      { guid: SCRIPTING, major: 1, minor: 0, description: 'Microsoft Scripting Runtime' },
      { guid: '{3F4DACA7-160D-11D2-A8E9-00104B365C9F}', major: 0, minor: 0, description: '' },
    ]);
  });

  it('書き方が違えば行番号付きのエラー', () => {
    expect(parseRefs('#refs', ['Microsoft Scripting Runtime']).errors[0]).toContain('1 行目');
  });

  it('書き出したものを読み直すと同じになる（説明の行はコメント）', () => {
    const refs = [{ guid: SCRIPTING, major: 1, minor: 0, description: 'Microsoft Scripting Runtime' }];
    expect(parseRefs('#refs', renderRefs(refs)).refs).toEqual(refs);
    expect(parseRefs('#refs', renderRefs([])).refs).toEqual([]);
  });
});

describe('フォームの配置の書き起こし', () => {
  it('入れ子・値の書き方を保ち、読み直すと同じ構造になる', () => {
    const { lines, warnings } = renderForm('frmMain', {
      props: [
        { name: 'Caption', value: '顧客 "登録"' },
        { name: 'Width', value: 300 },
        { name: 'BackColor', value: -2147483633 },
      ],
      controls: [
        { type: 'Frame', name: 'fra', parent: 'frmMain', props: [{ name: 'Caption', value: '区分' }] },
        { type: 'OptionButton', name: 'optA', parent: 'fra', props: [{ name: 'Value', value: true }] },
        { type: 'MultiPage', name: 'mpg', parent: 'frmMain', props: [{ name: 'Height', value: 120.5 }] },
        { type: 'Page', name: 'pg1', parent: 'mpg', props: [{ name: 'Caption', value: '基本' }] },
        { type: 'TextBox', name: 'txt', parent: 'pg1', props: [{ name: 'Font.Size', value: 11 }] },
        { type: 'Label', name: 'lbl', parent: 'frmMain', props: [{ name: 'Caption', value: '1行目\r\n2行目' }] },
        { type: 'RefEdit', name: 'ref1', parent: 'frmMain', props: [] },
        { type: 'Image', name: 'img', parent: 'frmMain', props: [], picture: true },
      ],
      picture: true,
    });
    expect(lines).toContain('   BackColor = &H8000000F&');
    expect(lines).toContain('   Caption = "顧客 ""登録"""');
    expect(lines).toContain("   ' 取り込めないコントロール: ref1（RefEdit）");
    expect(warnings.join('\n')).toMatch(/改行は空白/);
    expect(warnings.join('\n')).toMatch(/ref1/);
    expect(warnings.join('\n')).toMatch(/img の画像/);
    expect(warnings.join('\n')).toMatch(/背景画像/);

    const parsed = parseFormSheet('frmMain.frm', lines);
    expect(parsed.errors).toEqual([]);
    const f = parsed.form!;
    expect(f.props.map((p) => [p.name, p.value])).toEqual([
      ['Caption', '顧客 "登録"'],
      ['Width', 300],
      ['BackColor', -2147483633],
    ]);
    expect(f.children.map((c) => c.name)).toEqual(['fra', 'mpg', 'lbl', 'img']);
    expect(f.children[0].children[0].props[0]).toMatchObject({ name: 'Value', value: true });
    expect(f.children[1].children[0].children[0]).toMatchObject({ name: 'txt', type: 'TextBox' });
    expect(f.children[1].props[0]).toMatchObject({ name: 'Height', value: 120.5 });
    expect(f.children[2].props[0].value).toBe('1行目 2行目');
  });
});

describe('モジュールからシートへ', () => {
  const mod = (m: Partial<ImportedModule> & { name: string; type: number }): ImportedModule => ({ code: [], ...m });

  it('種類ごとの名前にし、中身の無いシート・ブックのコードは省く', () => {
    const r = modulesToSheets(
      [
        mod({ name: 'Module1', type: 1, code: ['Option Explicit', 'Sub A()', 'End Sub', '', ''] }),
        mod({ name: 'Class1', type: 2, code: ['Public X'] }),
        mod({ name: 'ThisWorkbook', type: 100, workbook: true, code: ['Private Sub Workbook_Open()', 'End Sub'] }),
        mod({
          name: 'Sheet1',
          type: 100,
          sheet: '入力画面',
          code: ['Private Sub Worksheet_Change(ByVal T As Range)', 'End Sub'],
        }),
        mod({ name: 'Sheet2', type: 100, sheet: 'マスタ', code: ['Option Explicit', ''] }),
        mod({
          name: 'Sheet3',
          type: 100,
          sheet: 'とても長いシートの名前とても長いシートの名前ですよね本当に',
          code: ["' x"],
        }),
        mod({
          name: 'UserForm1',
          type: 3,
          code: ['Private Sub UserForm_Initialize()', 'End Sub'],
          form: { props: [], controls: [] },
        }),
        mod({ name: 'Designer1', type: 11 }),
      ],
      ['入力画面', 'マスタ', 'とても長いシートの名前とても長いシートの名前ですよね本当に'],
    );
    expect(r.sheets.map((s) => s.name)).toEqual([
      'Module1.bas',
      'Class1.cls',
      'ThisWorkbook.cls',
      '入力画面.cls',
      'Sheet3.cls',
      'UserForm1.frm',
    ]);
    expect(r.sheets[0].lines).toEqual(['Option Explicit', 'Sub A()', 'End Sub']);
    expect(r.sheets[5].lines).toEqual([
      'Begin UserForm UserForm1',
      'End',
      '',
      'Private Sub UserForm_Initialize()',
      'End Sub',
    ]);
    expect(r.warnings.join()).toContain('Designer1');
  });

  it('シート名にできない長さの名前は、理由を付けて省く', () => {
    const r = modulesToSheets([mod({ name: 'M'.repeat(30), type: 1, code: ['x'] })], []);
    expect(r.sheets).toEqual([]);
    expect(r.warnings[0]).toContain('名前を短く');
  });
});

// ---- 取り込み（Excel の代わりに偽の importer を使う） ----

async function uiXlsx(file: string, sheets: string[]): Promise<void> {
  const wb = new ExcelJS.Workbook();
  for (const s of sheets) wb.addWorksheet(s).getCell('A1').value = `${s} の画面`;
  await wb.xlsx.writeFile(file);
}

function fakeImporter(modules: ImportedModule[], sheets = ['入力画面', 'マスタ']): VbaImporter & { calls: string[] } {
  const calls: string[] = [];
  const fn = (async (job) => {
    calls.push(job.input);
    await uiXlsx(job.output, sheets);
    return {
      ok: true,
      modules,
      references: [{ guid: SCRIPTING, major: 1, minor: 0, description: 'Microsoft Scripting Runtime' }],
      warnings: ['参照設定「Other」はライブラリではない（ほかのブックへの参照）ため取り込めません'],
    };
  }) as VbaImporter & { calls: string[] };
  fn.calls = calls;
  return fn;
}

const TOOL_MODULES: ImportedModule[] = [
  { name: 'Module1', type: 1, code: ['Option Explicit', 'Sub Hello()', '    MsgBox "こんにちは"', 'End Sub'] },
  { name: 'Sheet1', type: 100, sheet: '入力画面', code: ['Private Sub Worksheet_Activate()', 'End Sub'] },
  {
    name: 'UserForm1',
    type: 3,
    code: ['Private Sub btnOK_Click()', 'End Sub'],
    form: {
      props: [{ name: 'Caption', value: '登録' }],
      controls: [
        { type: 'CommandButton', name: 'btnOK', parent: 'UserForm1', props: [{ name: 'Caption', value: 'OK' }] },
      ],
    },
  },
];

async function vbaProject(): Promise<Fixture> {
  const f = await fixture({}, { git: true });
  await saveConfig(f.root, { ...DEFAULT_CONFIG, mode: 'vba' });
  await initProject(f.root);
  return f;
}

const bookOf = (f: Fixture) => `${path.basename(f.root)}.xlcode.xlsx`;
const outOf = (f: Fixture) => `${path.basename(f.root)}.xlsm`;

describe('既存の Excel ツールから編集用ブックを作る', () => {
  it('シートはそのまま、モジュール・フォーム・参照設定をソースコードとシートにする。元のツールは変えない', async () => {
    const f = await vbaProject();
    await writeFile(f.file('tool.xlsm'), '元のツール');
    const imp = fakeImporter(TOOL_MODULES);
    const r = await createBookFromTool(f.root, f.root, f.file('tool.xlsm'), imp);
    // ルートのディレクトリなので .gitignore もシートになる（ソースコードモードと同じ）
    expect(r.sheets.filter((n) => n !== '.gitignore')).toEqual([
      'LocalAgents.md',
      'Module1.bas',
      'References.refs',
      'UserForm1.frm',
      '入力画面.cls',
    ]);
    expect(r.skipped.join()).toContain('ほかのブックへの参照');
    // 元のファイルではなくコピーを開く
    expect(imp.calls[0]).not.toBe(f.file('tool.xlsm'));
    expect(await readFile(f.file('tool.xlsm'), 'utf8')).toBe('元のツール');

    // ソースコード（Shift_JIS・CRLF）
    expect(iconv.decode(await readFile(f.file('Module1.bas')), 'cp932')).toBe(
      'Option Explicit\r\nSub Hello()\r\n    MsgBox "こんにちは"\r\nEnd Sub\r\n',
    );
    expect(iconv.decode(await readFile(f.file('References.refs')), 'cp932')).toContain(
      `${SCRIPTING} 1.0 Microsoft Scripting Runtime`,
    );
    expect(await f.read('LocalAgents.md')).toContain('tool.xlsm から取り込んだ');

    // 編集用ブック: 画面のシートはそのまま、ソースコードがシートになる
    const b = await Book.load(f.file(bookOf(f)));
    expect(b.sheetNames()).toEqual(
      expect.arrayContaining([
        '#tree',
        '入力画面',
        'マスタ',
        'Agents.md',
        'LocalAgents.md',
        'References.refs',
        'Module1.bas',
      ]),
    );
    expect(b.readSheet('入力画面').lines).toEqual(['入力画面 の画面']);
    expect(b.readSheet('UserForm1.frm').lines[0]).toBe('Begin UserForm UserForm1');
    const c = modulesFromBook(b);
    expect(c.errors).toEqual([]);
    expect(c.references.map((x) => x.guid)).toEqual([SCRIPTING]);
    expect(c.modules.map((m) => [m.file, m.kind])).toEqual([
      ['Module1.bas', 'standard'],
      ['UserForm1.frm', 'form'],
      ['入力画面.cls', 'sheet'],
    ]);
    // シートとソースは同じ（すべて同期済み）
    const st = await bookStatus(f.root, f.file(bookOf(f)));
    expect(st.errors).toEqual([]);
    expect(st.files.every((x) => x.status === 'clean')).toBe(true);
  });

  it('取り込んだツールがビルド結果の場所にあれば、最初の Build で上書きを確認する', async () => {
    const f = await vbaProject();
    await writeFile(f.file(outOf(f)), '元のツール');
    await createBookFromTool(f.root, f.root, f.file(outOf(f)), fakeImporter(TOOL_MODULES));
    const runner: VbaRunner = async (job) => {
      await writeFile(job.output, JSON.stringify(job));
      return { ok: true };
    };
    const confirm = await vbaBuild(f.root, f.file(bookOf(f)), {}, runner);
    expect(confirm.confirmations.find((x) => x.kind === 'overwrite')?.message).toContain('既にある');
    const r = await vbaBuild(f.root, f.file(bookOf(f)), { confirmed: true }, runner);
    expect(r.errors).toEqual([]);
    const job = JSON.parse(await readFile(f.file(outOf(f)), 'utf8'));
    expect(job.references).toEqual([
      { guid: SCRIPTING, major: 1, minor: 0, description: 'Microsoft Scripting Runtime' },
    ]);
    const backups = await readdir(f.file('.xlcode/backup/_root'));
    expect(await readFile(f.file(`.xlcode/backup/_root/${backups[0]}`), 'utf8')).toBe('元のツール');
  });

  it('取り込み先に同じ名前のソースがあれば、何も書かずに止める', async () => {
    const f = await vbaProject();
    await writeFile(f.file('tool.xlsm'), 'x');
    await f.write('Module1.bas', '既存');
    await expect(createBookFromTool(f.root, f.root, f.file('tool.xlsm'), fakeImporter(TOOL_MODULES))).rejects.toThrow(
      '同じ名前のファイル',
    );
    expect(await f.read('Module1.bas')).toBe('既存');
    expect(await readdir(f.root)).not.toContain('UserForm1.frm');
  });

  it('拡張子のように見える名前のシートは注意を出す', async () => {
    const f = await vbaProject();
    await writeFile(f.file('tool.xlsm'), 'x');
    const r = await createBookFromTool(f.root, f.root, f.file('tool.xlsm'), fakeImporter([], ['入力画面', 'data.csv']));
    expect(r.skipped.join()).toContain('シート「data.csv」');
  });

  it('ソースコードモード・Excel が無い環境・既にブックがある・対象外の拡張子はエラー', async () => {
    const f = await vbaProject();
    await writeFile(f.file('tool.xlsm'), 'x');
    await writeFile(f.file('tool.docx'), 'x');
    await expect(createBookFromTool(f.root, f.root, f.file('tool.xlsm'), null)).rejects.toThrow('Windows');
    await expect(createBookFromTool(f.root, f.root, f.file('tool.docx'), fakeImporter([]))).rejects.toThrow(
      '取り込めるのは',
    );
    await createBookFromTool(f.root, f.root, f.file('tool.xlsm'), fakeImporter([]));
    await expect(createBookFromTool(f.root, f.root, f.file('tool.xlsm'), fakeImporter([]))).rejects.toThrow(
      '既にブック',
    );

    const g = await fixture({}, { git: true });
    await writeFile(g.file('tool.xlsm'), 'x');
    await expect(createBookFromTool(g.root, g.root, g.file('tool.xlsm'), fakeImporter([]))).rejects.toThrow(
      'VBA モード',
    );
  });

  it('Excel で失敗したら、ブックを作らずに理由を出す', async () => {
    const f = await vbaProject();
    await writeFile(f.file('tool.xlsm'), 'x');
    const imp: VbaImporter = async () => ({ ok: false, error: 'アクセス拒否', errorKind: 'vbom' });
    await expect(createBookFromTool(f.root, f.root, f.file('tool.xlsm'), imp)).rejects.toThrow(
      /アクセス拒否[\s\S]*信頼する/,
    );
    expect(await readdir(f.root)).not.toContain(bookOf(f));
  });
});

// ---- 実際の PowerShell スクリプトを偽の Excel で動かす（取り込み → Build の往復） ----

const pwsh = (() => {
  try {
    execFileSync('pwsh', ['-NoProfile', '-Command', 'exit 0'], { stdio: 'ignore' });
    return true;
  } catch {
    return false;
  }
})();

describe.skipIf(!pwsh)('取り込み用の PowerShell スクリプト（偽の Excel で実行）', async () => {
  const { runImportJob, runVbaJob, IMPORT_SCRIPT, VBA_SCRIPT } = await import('../src/main/excelCom');
  const prelude = path.resolve(import.meta.dirname, 'fixtures/fakeExcel.ps1');
  const fake = (s: string) =>
    s
      .replace("$ErrorActionPreference = 'Stop'", `. '${prelude}'\n$ErrorActionPreference = 'Stop'`)
      .replace('New-Object -ComObject Excel.Application', 'New-FakeExcel');

  /** 既存のツールの VBA プロジェクト */
  const SPEC = {
    components: [
      { name: 'Module1', type: 1, code: ['Option Explicit', 'Sub Hello()', '    MsgBox "こんにちは"', 'End Sub'] },
      { name: 'Class1', type: 2, code: ['Public Name As String'] },
      { name: 'Sheet1', type: 100, code: ['Private Sub Worksheet_Activate()', 'End Sub'] },
      { name: 'Sheet2', type: 100, code: [] },
      { name: 'ThisWorkbook', type: 100, code: ['Private Sub Workbook_Open()', 'End Sub'] },
      {
        name: 'UserForm1',
        type: 3,
        code: ['Private Sub btnOK_Click()', '    Unload Me', 'End Sub'],
        form: {
          props: { Caption: '顧客登録', Width: 300, Height: 200 },
          controls: [
            {
              type: 'CommandButton',
              name: 'btnOK',
              parent: 'UserForm1',
              // ForeColor・BackColor・Font.Name は既定値と同じなので書き出さない
              props: {
                Caption: 'OK',
                Left: 150,
                Top: 150,
                Width: 72,
                Height: 24,
                ForeColor: -2147483630,
                BackColor: 255,
                Default: true,
                'Font.Name': 'MS UI Gothic',
                'Font.Size': 11,
              },
            },
            {
              type: 'Frame',
              name: 'fra',
              parent: 'UserForm1',
              props: { Caption: '区分', Left: 6, Top: 6, Width: 100, Height: 60 },
            },
            {
              type: 'OptionButton',
              name: 'optA',
              parent: 'fra',
              props: { Caption: '個人', Left: 6, Top: 6, Width: 60, Height: 18, Value: true },
            },
            {
              type: 'MultiPage',
              name: 'mpg',
              parent: 'UserForm1',
              props: { Left: 6, Top: 70, Width: 200, Height: 70 },
            },
            { type: 'Page', name: 'pgA', parent: 'mpg', props: { Caption: '基本' } },
            { type: 'TextBox', name: 'txtA', parent: 'pgA', props: { Left: 6, Top: 6, Width: 100, Height: 18 } },
            { type: 'RefEdit', name: 'ref1', parent: 'UserForm1', props: { Left: 0, Top: 0, Width: 10, Height: 10 } },
          ],
        },
      },
    ],
    references: [
      { guid: SCRIPTING, major: 1, minor: 0, type: 0, name: 'Scripting', description: 'Microsoft Scripting Runtime' },
      { guid: '', major: 0, minor: 0, type: 1, name: 'OtherBook', description: '' },
    ],
  };

  async function withEnv<T>(env: Record<string, string>, fn: () => Promise<T>): Promise<T> {
    const saved = { ...process.env };
    Object.assign(process.env, env);
    try {
      return await fn();
    } finally {
      process.env = saved;
    }
  }

  it('既存のツールを取り込み、Build すると同じ構成の VBA が書き込まれる', async () => {
    const f = await vbaProject();
    const tool = f.file('tool.xlsm');
    await uiXlsx(tool, ['入力画面', 'マスタ']);
    const specFile = f.file('spec.json');
    await writeFile(specFile, JSON.stringify(SPEC));

    // 取り込み
    const importer: VbaImporter = (job) =>
      withEnv({ FAKE_EXCEL_PROJECT: specFile, FAKE_EXCEL_SHEETS: '入力画面|マスタ' }, () =>
        runImportJob(job, { exe: 'pwsh', script: fake(IMPORT_SCRIPT) }),
      );
    const created = await createBookFromTool(f.root, f.root, tool, importer);
    expect(created.sheets.filter((n) => !['.gitignore', 'spec.json'].includes(n))).toEqual([
      'LocalAgents.md',
      'Class1.cls',
      'Module1.bas',
      'References.refs',
      'ThisWorkbook.cls',
      'UserForm1.frm',
      '入力画面.cls',
    ]);
    expect(created.skipped.join('\n')).toContain('OtherBook');
    expect(created.skipped.join('\n')).toContain('ref1（RefEdit）');

    const b = await Book.load(f.file(bookOf(f)));
    expect(b.readSheet('Module1.bas').lines).toEqual(SPEC.components[0].code);
    const frm = b.readSheet('UserForm1.frm').lines;
    expect(frm).toContain('   Caption = "顧客登録"');
    expect(frm).toContain('   Begin CommandButton btnOK');
    expect(frm).toContain('      BackColor = &H000000FF&');
    expect(frm).toContain('      Default = True');
    expect(frm).toContain('      Font.Size = 11');
    // 既定値と同じものは書かない
    expect(frm.join('\n')).not.toContain('ForeColor');
    expect(frm.join('\n')).not.toContain('Font.Name');
    // 入れ子
    expect(frm).toContain('      Begin OptionButton optA');
    expect(frm).toContain('      Begin Page pgA');
    expect(frm).toContain('         Begin TextBox txtA');
    expect(frm.slice(-3)).toEqual(SPEC.components[5].code);
    expect(parseFormSheet('UserForm1.frm', frm).errors).toEqual([]);
    expect(parseRefs('References.refs', b.readSheet('References.refs').lines).refs.map((r) => r.guid)).toEqual([
      SCRIPTING,
    ]);

    // Build（偽の Excel が書き込んだ内容を確かめる）
    const runner: VbaRunner = (job) =>
      withEnv({ FAKE_EXCEL_SHEETS: '入力画面|マスタ' }, () =>
        runVbaJob(job, { exe: 'pwsh', script: fake(VBA_SCRIPT) }),
      );
    const r = await vbaBuild(f.root, f.file(bookOf(f)), { confirmed: true }, runner);
    expect(r.errors).toEqual([]);
    const out = JSON.parse(await readFile(f.file(outOf(f)), 'utf8'));
    const comp = (name: string) => out.components.find((c: { name: string }) => c.name === name);
    expect(comp('Module1').code).toEqual(SPEC.components[0].code);
    expect(comp('Class1')).toMatchObject({ type: 2, code: ['Public Name As String'] });
    expect(comp('Sheet1').code).toEqual(SPEC.components[2].code);
    expect(comp('ThisWorkbook').code).toEqual(SPEC.components[4].code);
    const uf = comp('UserForm1');
    expect(uf.props).toMatchObject({ Caption: '顧客登録', Width: 300, Height: 200 });
    const [btn, fra, mpg] = uf.designer.Controls;
    expect(btn).toMatchObject({
      Name: 'btnOK',
      Caption: 'OK',
      Left: 150,
      BackColor: 255,
      Default: true,
      Font: { Size: 11 },
    });
    expect(fra.Controls[0]).toMatchObject({ Name: 'optA', Value: true });
    expect(mpg.Pages[0]).toMatchObject({ Name: 'pgA', Caption: '基本' });
    expect(mpg.Pages[0].Controls[0]).toMatchObject({ Name: 'txtA', Left: 6 });
    expect(out.references.map((x: { GUID: string }) => x.GUID)).toContain(SCRIPTING);
  }, 120_000);

  it('Build: この PC に無いライブラリの参照設定は、理由を付けて失敗する', async () => {
    const f = await fixture({ 'input.xlsx': 'x' });
    const r = await withEnv({ FAKE_EXCEL_SHEETS: 'Sheet1' }, () =>
      runVbaJob(
        {
          input: f.file('input.xlsx'),
          output: f.file('out.xlsm'),
          modules: [],
          references: [
            { guid: '{00000000-0000-0000-0000-000000000001}', major: 1, minor: 0, description: '無いライブラリ' },
          ],
        },
        { exe: 'pwsh', script: fake(VBA_SCRIPT) },
      ),
    );
    expect(r.ok).toBe(false);
    expect(r.error).toContain('参照設定「無いライブラリ');
  }, 60_000);
});
