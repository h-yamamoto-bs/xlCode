import { mkdtemp, readdir, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { bookStatus, build, createBook, initProject, refreshTree, relocateBooks, saveConfig, sync } from '../src/core';
import { DEFAULT_CONFIG } from '../src/core/config';
import { bookPathFor, bookRef } from '../src/core/project';
import { fixture } from './helpers';

/** ソース（f.root）と、別の場所のブック置き場（xl）を用意する */
async function setup() {
  const f = await fixture({ 'app/add.ts': 'export const a = 1;\n', 'root.ts': 'export {};\n' });
  const xl = await mkdtemp(path.join(tmpdir(), 'xlfolder-'));
  await initProject(f.root);
  await saveConfig(f.root, { ...DEFAULT_CONFIG, bookRoot: xl });
  return { f, xl };
}

describe('ブックの置き場所をソースと分ける', () => {
  it('対応はフォルダの位置だけで決まる', () => {
    const root = path.resolve('/src/shop');
    const xl = path.resolve('/od/xl');
    expect(bookPathFor(root, xl, path.join(root, 'app'))).toBe(path.join(xl, 'app', 'app.xlcode.xlsx'));
    expect(bookPathFor(root, xl, root)).toBe(path.join(xl, 'shop.xlcode.xlsx'));
    const ref = bookRef(root, path.join(xl, 'app', 'app.xlcode.xlsx'), xl);
    expect(ref).toMatchObject({ rel: 'app/app.xlcode.xlsx', dirRel: 'app', dirAbs: path.join(root, 'app') });
  });

  it('ブックは置き場所に作られ、ソースのフォルダには作られない', async () => {
    const { f, xl } = await setup();
    const r = await createBook(f.root, f.file('app'));
    expect(r.book).toBe('app/app.xlcode.xlsx');
    expect(await readdir(path.join(xl, 'app'))).toEqual(['app.xlcode.xlsx']);
    expect((await readdir(f.file('app'))).filter((n) => n.endsWith('.xlsx'))).toEqual([]);
  });

  it('Build / Sync / Refresh Tree が動く', async () => {
    const { f, xl } = await setup();
    await createBook(f.root, f.file('app'));
    await createBook(f.root, f.root);
    const book = path.join(xl, 'app', 'app.xlcode.xlsx');
    const { Book } = await import('../src/core/workbook');
    const b = await Book.load(book);
    b.writeLines('sub.ts', ['export const s = 2;']);
    await b.save(book);

    const r = await build(f.root, book, { confirmed: true });
    expect(r.errors).toEqual([]);
    expect(await f.read('app/sub.ts')).toBe('export const s = 2;\n');

    await f.write('app/add.ts', 'export const a = 10;\n');
    expect((await sync(f.root, book, { confirmed: true })).changes).toContainEqual({
      action: 'write-sheet',
      target: 'add.ts',
    });

    const t = await refreshTree(f.root);
    expect(t.books.map((x) => x.book).sort()).toEqual(['app/app.xlcode.xlsx', path.basename(f.root) + '.xlcode.xlsx']);
    // ルートのブックはルートのファイルを担当する
    const st = await bookStatus(f.root, path.join(xl, path.basename(f.root) + '.xlcode.xlsx'));
    expect(st.files.map((x) => x.name)).toContain('root.ts');
  });

  it('対応するソースのフォルダが無ければエラー', async () => {
    const { f, xl } = await setup();
    await createBook(f.root, f.file('app'));
    const { rm } = await import('node:fs/promises');
    await rm(f.file('app'), { recursive: true });
    const r = await build(f.root, path.join(xl, 'app', 'app.xlcode.xlsx'), { confirmed: true });
    expect(r.status).toBe('error');
    expect(r.errors[0]).toMatch(/対応するソースのフォルダがありません/);
  });
});

describe('置き場所の変更（ブックの移動）', () => {
  it('ソースの中 → 別の場所 → ソースの中 と移動しても状態が保たれる', async () => {
    const f = await fixture({ 'app/add.ts': 'export const a = 1;\n' });
    await initProject(f.root);
    await createBook(f.root, f.file('app'));
    const xl = await mkdtemp(path.join(tmpdir(), 'xlfolder-'));

    const r = await relocateBooks(f.root, xl);
    expect(r.moved).toEqual(['app/app.xlcode.xlsx']);
    expect(await readdir(path.join(xl, 'app'))).toEqual(['app.xlcode.xlsx']);
    const st = await bookStatus(f.root, path.join(xl, 'app', 'app.xlcode.xlsx'));
    expect(st.files.every((x) => x.status === 'clean')).toBe(true);

    const back = await relocateBooks(f.root, null);
    expect(back.bookRoot).toBeNull();
    const st2 = await bookStatus(f.root, f.file('app/app.xlcode.xlsx'));
    expect(st2.files.every((x) => x.status === 'clean')).toBe(true);
  });

  it('ソースの中のフォルダ・相対パス・開いているブック・移動先の重複は拒否し、何も動かさない', async () => {
    const f = await fixture({ 'app/add.ts': 'x\n' });
    await initProject(f.root);
    await createBook(f.root, f.file('app'));
    await expect(relocateBooks(f.root, f.file('books'))).rejects.toThrow(/ソースのフォルダの中にはできません/);
    await expect(relocateBooks(f.root, 'rel/path')).rejects.toThrow(/フルパス/);

    const xl = await mkdtemp(path.join(tmpdir(), 'xlfolder-'));
    await writeFile(f.file('app/~$app.xlcode.xlsx'), '');
    await expect(relocateBooks(f.root, xl)).rejects.toThrow(/開かれています/);
    const { rm, mkdir } = await import('node:fs/promises');
    await rm(f.file('app/~$app.xlcode.xlsx'));

    await mkdir(path.join(xl, 'app'), { recursive: true });
    await writeFile(path.join(xl, 'app', 'app.xlcode.xlsx'), 'x');
    await expect(relocateBooks(f.root, xl)).rejects.toThrow(/同じ名前のファイルがあります/);
    expect(await readdir(f.file('app'))).toContain('app.xlcode.xlsx');
  });
});
