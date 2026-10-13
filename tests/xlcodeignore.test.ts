import { mkdtemp, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import {
  bookStatus,
  computeTree,
  createBook,
  initProject,
  listDirsWithoutBook,
  loadIgnore,
  relocateBooks,
  saveConfig,
  sync,
} from '../src/core';
import { DEFAULT_CONFIG } from '../src/core/config';
import { fixture } from './helpers';

describe('.xlcodeignore（ブックにしないファイル）', () => {
  it('initProject が雛形を作り、雛形自体はシートにならない', async () => {
    const f = await fixture({ 'a.ts': 'export {};\n' });
    const done = await initProject(f.root);
    expect(done.some((l) => l.includes('.xlcodeignore'))).toBe(true);
    expect(await f.read('.xlcodeignore')).toContain('.gitignore と同じ書き方');
    // 2 回目は作らない
    expect((await initProject(f.root)).some((l) => l.includes('.xlcodeignore'))).toBe(false);
    const r = await createBook(f.root, f.root);
    expect(r.sheets).not.toContain('.xlcodeignore');
    expect(r.sheets).toContain('a.ts');
  });

  it('書いたファイル・フォルダはシート・#tree・ブック作成の候補から外れる（Git 管理は影響しない）', async () => {
    const f = await fixture({
      'app/App.tsx': 'export const App = 1;\n',
      'app/App.min.js': 'x\n',
      'app/package-lock.json': '{}\n',
      'docs/readme.md': '# docs\n',
      'lib/x.ts': 'export {};\n',
    });
    await initProject(f.root);
    await f.write('.xlcodeignore', '# メモ\ndocs/\n*.min.js\npackage-lock.json\n');
    const ig = await loadIgnore(f.root);
    expect(await listDirsWithoutBook(f.root, ig)).toEqual(['', 'app', 'lib']);
    const tree = (await computeTree(f.root, ig)).lines.join('\n');
    expect(tree).not.toContain('docs/');
    expect(tree).not.toContain('App.min.js');
    expect(tree).toContain('App.tsx');
    const r = await createBook(f.root, f.file('app'));
    expect(r.sheets).toEqual(['LocalAgents.md', 'App.tsx']);
    // .gitignore には何も足していない
    expect(await f.read('.gitignore')).not.toContain('docs/');
  });

  it('あとから除外したファイルは、シートがあっても触らず、警告だけ出す', async () => {
    const f = await fixture({ 'app/App.tsx': 'export const App = 1;\n', 'app/gen.ts': 'export const g = 1;\n' });
    await initProject(f.root);
    await createBook(f.root, f.file('app'));
    await f.write('.xlcodeignore', 'app/gen.ts\n');
    const st = await bookStatus(f.root, f.file('app/app.xlcode.xlsx'));
    expect(st.files.map((x) => x.name)).not.toContain('gen.ts');
    expect(st.warnings.some((w) => w.includes('gen.ts') && w.includes('.xlcodeignore'))).toBe(true);
    const r = await sync(f.root, f.file('app/app.xlcode.xlsx'), { confirmed: true });
    expect(r.status).toBe('ok');
    // シートは残る（消したいときは Excel で削除する）
    expect(await f.sheet('app/app.xlcode.xlsx', 'gen.ts')).toEqual(['export const g = 1;']);
  });
});

describe('ブックの置き場所の変更（移動せずに変更）', () => {
  it('move: false なら設定だけ変え、ブックは動かさない（手で移動した・フォルダ名を変えたとき用）', async () => {
    const f = await fixture({ 'app/add.ts': 'export const a = 1;\n' });
    await initProject(f.root);
    await createBook(f.root, f.file('app'));
    const xl = await mkdtemp(path.join(tmpdir(), 'xlfolder-'));
    const r = await relocateBooks(f.root, xl, { move: false });
    expect(r.moved).toEqual([]);
    expect(r.bookRoot).toBe(xl);
    expect(await readdir(f.file('app'))).toContain('app.xlcode.xlsx');
    expect(await readdir(xl)).toEqual([]);
    // 置き場所を戻すのも同様
    const back = await relocateBooks(f.root, null, { move: false });
    expect(back.bookRoot).toBeNull();
    expect(await readdir(f.file('app'))).toContain('app.xlcode.xlsx');
  });

  it('移動先に同じブックが既にあり、元が無ければ（移動済み）そのまま採用する', async () => {
    const f = await fixture({ 'app/add.ts': 'export const a = 1;\n' });
    await initProject(f.root);
    await createBook(f.root, f.file('app'));
    const xl = await mkdtemp(path.join(tmpdir(), 'xlfolder-'));
    await relocateBooks(f.root, xl);
    const xl2 = await mkdtemp(path.join(tmpdir(), 'xlfolder-'));
    // 利用者が自分でフォルダごと移した状態
    const { mkdir, copyFile } = await import('node:fs/promises');
    await mkdir(path.join(xl2, 'app'), { recursive: true });
    await copyFile(path.join(xl, 'app', 'app.xlcode.xlsx'), path.join(xl2, 'app', 'app.xlcode.xlsx'));
    await rm(path.join(xl, 'app', 'app.xlcode.xlsx'));
    const r = await relocateBooks(f.root, xl2);
    expect(r.moved).toEqual([]);
    const st = await bookStatus(f.root, path.join(xl2, 'app', 'app.xlcode.xlsx'));
    expect(st.files.every((x) => x.status === 'clean')).toBe(true);
  });

  it('一時的なロック（EBUSY 相当）は短く再試行してから諦める', async () => {
    const f = await fixture({ 'app/add.ts': 'x\n' });
    await initProject(f.root);
    await createBook(f.root, f.file('app'));
    await saveConfig(f.root, { ...DEFAULT_CONFIG });
    const xl = await mkdtemp(path.join(tmpdir(), 'xlfolder-'));
    const lock = f.file('app/~$app.xlcode.xlsx');
    await writeFile(lock, '');
    // 本当に開かれている（ロックファイルがある）ときは、再試行しても止める
    await expect(relocateBooks(f.root, xl, { retryMs: 50 })).rejects.toThrow(/開かれています/);
    await rm(lock);
    expect((await relocateBooks(f.root, xl)).moved).toEqual(['app/app.xlcode.xlsx']);
  });
});

describe('ブック作成前の確認用の情報', () => {
  it('除外ファイルの有効な行を数える', async () => {
    const { ignoreRules } = await import('../src/core');
    expect(ignoreRules('# メモ\n\n docs/ \n*.min.js\r\n#x\n')).toEqual(['docs/', '*.min.js']);
    expect(ignoreRules('')).toEqual([]);
  });

  it('ディレクトリのファイルを、シート / 除外 / 対象外に分ける', async () => {
    const { previewDirFiles, isCodeName, loadIgnore } = await import('../src/core');
    const f = await fixture({
      '.gitignore': '*.log\n',
      'a.ts': '1\n',
      'b.log': '2\n',
      Makefile: '3\n',
      NOTES: '4\n',
      'Agents.md': '# x\n',
      'sub/c.ts': '5\n',
      'sub/c.min.js': '6\n',
    });
    await f.write('.xlcodeignore', '*.min.js\n');
    const ig = await loadIgnore(f.root);
    const root = await previewDirFiles(f.root, f.root, ig, (n) => isCodeName(n));
    expect(root).toEqual({
      sheets: ['.gitignore', 'a.ts', 'Makefile'],
      ignored: ['.xlcodeignore', 'b.log'],
      nonCode: ['NOTES'],
    });
    const sub = await previewDirFiles(f.root, f.file('sub'), ig, (n) => isCodeName(n));
    expect(sub).toEqual({ sheets: ['c.ts'], ignored: ['c.min.js'], nonCode: [] });
  });
});
