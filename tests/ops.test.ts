import { access, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { bookStatus, build, checkTree, createBook, initProject, refreshTree, sync } from '../src/core';
import { fixture, type Fixture } from './helpers';

const BOOK = 'app/app.xlcode.xlsx';
const APP = `export function add(a: number, b: number): number {\n  return a + b;\n}\n`;

async function setup(extra: Record<string, string> = {}, git = false): Promise<Fixture> {
  const f = await fixture({ 'app/add.ts': APP, ...extra }, { git });
  await initProject(f.root);
  await createBook(f.root, f.file('app'));
  if (git) {
    f.git('add', '-A');
    f.git('commit', '-q', '-m', 'setup');
  }
  return f;
}

const exists = (p: string) =>
  access(p).then(
    () => true,
    () => false,
  );

describe('ブック作成', () => {
  it('既存ファイル・LocalAgents.md・#tree・Agents.md をシートにする', async () => {
    const f = await setup();
    const res = await bookStatus(f.root, f.file(BOOK));
    expect(res.errors).toEqual([]);
    expect(res.files.map((x) => [x.name, x.status])).toEqual([
      ['add.ts', 'clean'],
      ['LocalAgents.md', 'clean'],
    ]);
    const tree = await f.sheet(BOOK, '#tree');
    expect(tree[0]).toMatch(/^# tree-version: [0-9a-f]{12} /);
    expect(tree).toContain('    app.xlcode.xlsx');
    expect((await f.sheet(BOOK, 'Agents.md'))[0]).toBe('# Agents.md');
    expect(await f.read('.gitignore')).toContain('*.xlcode.xlsx');
  });
});

describe('Build', () => {
  it('Excel の変更をファイルへ出力し、整形結果をシートにも書き戻す', async () => {
    const f = await setup();
    await f.editBook(BOOK, (b) => b.writeLines('sub.ts', ['export const sub=(a:number,b:number)=>a-b']));
    const r = await build(f.root, f.file(BOOK), { confirmed: true });
    expect(r.errors).toEqual([]);
    expect(r.status).toBe('ok');
    expect(await f.read('app/sub.ts')).toBe('export const sub = (a: number, b: number) => a - b;\n');
    expect(await f.sheet(BOOK, 'sub.ts')).toEqual(['export const sub = (a: number, b: number) => a - b;']);
    // 往復で差分が出ない
    const s = await sync(f.root, f.file(BOOK), { confirmed: true });
    expect(s.status).toBe('ok');
    expect(s.changes).toEqual([]);
  });

  it('ソース側だけの変更は強制 Sync でシートへ取り込む', async () => {
    const f = await setup();
    await f.write('app/add.ts', APP + '\nexport const x = 1;\n');
    const r = await build(f.root, f.file(BOOK), { confirmed: true });
    expect(r.status).toBe('ok');
    expect(r.changes).toContainEqual({ action: 'write-sheet', target: 'add.ts' });
    expect((await f.sheet(BOOK, 'add.ts')).at(-1)).toBe('export const x = 1;');
  });

  it('シートが削除されてもファイルは削除しない（4.7）', async () => {
    const f = await setup();
    await f.editBook(BOOK, (b) => b.deleteSheet('add.ts'));
    const r = await build(f.root, f.file(BOOK), { confirmed: true });
    expect(r.status).toBe('ok');
    expect(await f.read('app/add.ts')).toBe(APP);
  });

  it('DEL_ 付きシートは確認後にファイルとシートを削除する', async () => {
    const f = await setup();
    await f.editBook(BOOK, (b) => {
      const lines = b.readSheet('add.ts').lines;
      b.deleteSheet('add.ts');
      b.writeLines('DEL_add.ts', lines);
    });
    const c = await build(f.root, f.file(BOOK));
    expect(c.status).toBe('confirm');
    expect(c.confirmations.find((x) => x.kind === 'delete')?.files).toEqual(['add.ts']);
    expect(await exists(f.file('app/add.ts'))).toBe(true);
    const r = await build(f.root, f.file(BOOK), { confirmed: true });
    expect(r.status).toBe('ok');
    expect(await exists(f.file('app/add.ts'))).toBe(false);
    expect((await bookStatus(f.root, f.file(BOOK))).files.map((x) => x.name)).toEqual(['LocalAgents.md']);
  });

  it('30% 以上の急減は確認を求める（省略検知）', async () => {
    const long = Array.from({ length: 40 }, (_, i) => `export const v${i} = ${i};`).join('\n') + '\n';
    const f = await setup({ 'app/long.ts': long });
    await f.editBook(BOOK, (b) => b.writeLines('long.ts', ['export const v0 = 0;', '// ...省略...']));
    const r = await build(f.root, f.file(BOOK));
    expect(r.status).toBe('confirm');
    expect(r.confirmations[0].kind).toBe('shrink');
    expect(r.confirmations[0].files[0]).toMatch(/long\.ts: 40→2 行/);
    expect(await f.read('app/long.ts')).toBe(long);
  });

  it('数値に変換されたセルはエラーで中断する', async () => {
    const f = await setup();
    await f.editBook(BOOK, (b) => {
      b.wb.getWorksheet('add.ts')!.getCell('A5').value = 42;
    });
    const r = await build(f.root, f.file(BOOK), { confirmed: true });
    expect(r.status).toBe('error');
    expect(r.errors.join('\n')).toMatch(/5 行目が文字列ではありません（number）/);
    expect(await f.read('app/add.ts')).toBe(APP);
  });

  it('.gitignore 対象のシートは警告してスキップする', async () => {
    const f = await setup();
    await f.write('.gitignore', (await f.read('.gitignore')) + '*.log\n');
    await f.editBook(BOOK, (b) => b.writeLines('debug.log', ['x']));
    const r = await build(f.root, f.file(BOOK), { confirmed: true });
    expect(r.status).toBe('ok');
    expect(r.warnings.join()).toMatch(/debug\.log.*スキップ/);
    expect(await exists(f.file('app/debug.log'))).toBe(false);
  });

  it('ブックが開かれている（ロックファイルがある）と実行しない', async () => {
    const f = await setup();
    await writeFile(path.join(f.root, 'app', '~$app.xlcode.xlsx'), '');
    const r = await build(f.root, f.file(BOOK), { confirmed: true });
    expect(r.status).toBe('error');
    expect(r.errors[0]).toMatch(/開かれています/);
  });

  it('同一ディレクトリに複数のブックがあるとエラー', async () => {
    const f = await setup();
    await writeFile(f.file('app/other.xlcode.xlsx'), await f.read(BOOK).catch(() => ''));
    const { copyFile } = await import('node:fs/promises');
    await copyFile(f.file(BOOK), f.file('app/other.xlcode.xlsx'));
    const r = await build(f.root, f.file(BOOK), { confirmed: true });
    expect(r.errors[0]).toMatch(/複数の \.xlcode\.xlsx/);
  });
});

describe('衝突（5.6）', () => {
  it('両側変更で衝突シートを作り、統合後の Build で Excel 側を採用する', async () => {
    const f = await setup();
    await f.write('app/add.ts', APP + 'export const fromEditor = 1;\n');
    await f.editBook(BOOK, (b) =>
      b.writeLines('add.ts', [...b.readSheet('add.ts').lines, 'export const fromCopilot = 2;']),
    );

    const r = await build(f.root, f.file(BOOK), { confirmed: true });
    expect(r.status).toBe('conflict');
    expect(r.conflicts).toEqual([{ sheet: '#conflict_01', file: 'add.ts' }]);
    const sheet = await f.sheet(BOOK, '#conflict_01');
    expect(sheet[0]).toBe('app/add.ts');
    expect(sheet[1]).toBe('Excel側');
    expect(sheet.at(-1)).toBe('export const fromCopilot = 2;');
    // ソースは変更されない
    expect(await f.read('app/add.ts')).toContain('fromEditor');

    // 衝突シートが残っている間は Build できない
    const blocked = await build(f.root, f.file(BOOK), { confirmed: true });
    expect(blocked.status).toBe('error');
    expect(blocked.errors[0]).toMatch(/#conflict_01/);

    // Copilot が統合して衝突シートを削除
    await f.editBook(BOOK, (b) => {
      b.writeLines('add.ts', [...b.readSheet('add.ts').lines, 'export const fromEditor = 1;']);
      b.deleteSheet('#conflict_01');
    });
    const ok = await build(f.root, f.file(BOOK), { confirmed: true });
    expect(ok.status).toBe('ok');
    const out = await f.read('app/add.ts');
    expect(out).toContain('fromCopilot');
    expect(out).toContain('fromEditor');
  });
});

describe('Sync', () => {
  it('ソースの変更・追加・削除をシートへ反映する', async () => {
    const f = await setup({ 'app/old.ts': 'export const old = 1;\n' });
    await f.write('app/add.ts', APP + 'export const y = 2;\n');
    await f.write('app/new.ts', 'export const n = 1;\n');
    const { rm } = await import('node:fs/promises');
    await rm(f.file('app/old.ts'));
    const r = await sync(f.root, f.file(BOOK), { confirmed: true });
    expect(r.status).toBe('ok');
    const st = await bookStatus(f.root, f.file(BOOK));
    expect(st.files.map((x) => [x.name, x.status])).toEqual([
      ['add.ts', 'clean'],
      ['LocalAgents.md', 'clean'],
      ['new.ts', 'clean'],
    ]);
  });

  it('未 Build のシート変更があれば中断し、破棄を選ぶとソースで上書きする（5.1）', async () => {
    const f = await setup();
    await f.editBook(BOOK, (b) => {
      b.writeLines('add.ts', ['export const changed = 1;']);
      b.writeLines('draft.ts', ['export const d = 1;']);
    });
    const r = await sync(f.root, f.file(BOOK), { confirmed: true });
    expect(r.status).toBe('needs-decision');
    expect(r.unbuilt).toEqual(['add.ts', 'draft.ts']);

    const d = await sync(f.root, f.file(BOOK), { confirmed: true, discardExcelChanges: true });
    expect(d.status).toBe('ok');
    expect(await f.sheet(BOOK, 'add.ts')).toEqual(APP.trimEnd().split('\n'));
    expect((await bookStatus(f.root, f.file(BOOK))).files.map((x) => x.name)).not.toContain('draft.ts');
  });

  it('Excel で削除されたシートを復元する', async () => {
    const f = await setup();
    await f.editBook(BOOK, (b) => b.deleteSheet('add.ts'));
    const r = await sync(f.root, f.file(BOOK), { confirmed: true });
    expect(r.status).toBe('ok');
    expect(await f.sheet(BOOK, 'add.ts')).toEqual(APP.trimEnd().split('\n'));
  });

  it('不正なファイル名があればエラーで中断する', async () => {
    const f = await setup();
    await f.write('app/[slug].tsx', 'x\n');
    const r = await sync(f.root, f.file(BOOK), { confirmed: true });
    expect(r.status).toBe('error');
    expect(r.errors[0]).toMatch(/\[slug\]\.tsx/);
  });

  it('バイナリファイルはスキップする', async () => {
    const f = await setup();
    await writeFile(f.file('app/logo.png'), Buffer.from([0x89, 0x50, 0x4e, 0x47, 0, 1, 2]));
    const r = await sync(f.root, f.file(BOOK), { confirmed: true });
    expect(r.status).toBe('ok');
    expect(r.warnings.join()).toMatch(/logo\.png/);
  });
});

describe('Git 連携', () => {
  it('未コミット変更を確認し、対象ディレクトリだけを自動コミットする', async () => {
    const f = await setup({ 'other/x.ts': 'export const x = 1;\n' }, true);
    await f.write('app/add.ts', APP + 'export const z = 3;\n');
    await f.write('other/x.ts', 'export const x = 2;\n');
    await f.editBook(BOOK, (b) => b.writeLines('mul.ts', ['export const mul = (a: number, b: number) => a * b;']));

    const c = await build(f.root, f.file(BOOK));
    expect(c.status).toBe('confirm');
    expect(c.confirmations.find((x) => x.kind === 'uncommitted')?.files).toEqual(['app/add.ts']);

    const r = await build(f.root, f.file(BOOK), { confirmed: true });
    expect(r.status).toBe('ok');
    expect(f.git('log', '-1', '--format=%s')).toMatch(/Build 前の自動コミット \(app\/app\.xlcode\.xlsx\)/);
    expect(f.git('show', '--name-only', '--format=', 'HEAD').trim()).toBe('app/add.ts');
    // 他ディレクトリの変更とブック自体はコミットしない
    expect(f.git('status', '--porcelain')).toContain('other/x.ts');
    expect(f.git('ls-files')).not.toContain('xlcode.xlsx');
  });

  it('Git リポジトリでなければ確認を求める', async () => {
    const f = await setup();
    await f.editBook(BOOK, (b) => b.writeLines('mul.ts', ['export const m = 1;']));
    const r = await build(f.root, f.file(BOOK));
    expect(r.confirmations.map((x) => x.kind)).toEqual(['no-git']);
  });
});

describe('Refresh Tree', () => {
  it('全ブックの #tree を同一内容に更新し、Agents.md を配布する', async () => {
    const f = await setup({
      'api/server.ts': 'export {};\n',
      'node_modules/x/index.js': 'x',
      '.gitignore': 'node_modules/\n',
    });
    await createBook(f.root, f.file('api'));
    await f.write('app/extra.ts', 'export {};\n');
    expect((await checkTree(f.root, f.file(BOOK))).upToDate).toBe(false);

    await f.write('Agents.md', '# Agents.md\n\n- 新しいルール\n');
    const r = await refreshTree(f.root);
    expect(r.books.map((b) => [b.book, b.ok])).toEqual([
      ['api/api.xlcode.xlsx', true],
      ['app/app.xlcode.xlsx', true],
    ]);
    const a = await f.sheet(BOOK, '#tree');
    const b = await f.sheet('api/api.xlcode.xlsx', '#tree');
    expect(a).toEqual(b);
    expect(a.join('\n')).not.toContain('node_modules');
    expect(a).toContain('    extra.ts');
    expect(await f.sheet(BOOK, 'Agents.md')).toContain('- 新しいルール');
    expect((await checkTree(f.root, f.file(BOOK))).upToDate).toBe(true);
    expect(await f.read('.xlcode/tree.txt')).toContain('extra.ts');
  });

  it('開いているブックは失敗として報告し、一部更新を警告する', async () => {
    const f = await setup({ 'api/server.ts': 'export {};\n' });
    await createBook(f.root, f.file('api'));
    await writeFile(f.file('api/~$api.xlcode.xlsx'), '');
    const r = await refreshTree(f.root);
    expect(r.partial).toBe(true);
    expect(r.books.find((b) => b.book === 'api/api.xlcode.xlsx')?.ok).toBe(false);
  });
});
