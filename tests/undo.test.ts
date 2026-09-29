import { access, writeFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { bookDiff, bookRef, bookStatus, build, createBook, initProject, sync, undoLast } from '../src/core';
import { fixture, type Fixture } from './helpers';

const BOOK = 'app/app.xlcode.xlsx';
const ADD = `export function add(a: number, b: number): number {\n  return a + b;\n}\n`;

async function setup(git = false): Promise<Fixture> {
  const f = await fixture({ 'app/add.ts': ADD }, { git });
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

describe('元に戻す', () => {
  it('Build が書き出したファイルと削除したファイルを操作前に戻し、状態も戻る', async () => {
    const f = await setup();
    expect((await bookStatus(f.root, f.file(BOOK))).undo).toBeNull();
    await f.editBook(BOOK, (b) => {
      b.writeLines('add.ts', ['export const add = 0;']);
      b.writeLines('new.ts', ['export const n = 1;']);
    });
    const r = await build(f.root, f.file(BOOK), { confirmed: true });
    expect(r.status).toBe('ok');
    expect(await f.read('app/add.ts')).toBe('export const add = 0;\n');
    expect(await exists(f.file('app/new.ts'))).toBe(true);

    const st = await bookStatus(f.root, f.file(BOOK));
    expect(st.undo?.label).toBe('Build');
    expect(st.undo?.files.sort()).toEqual(['add.ts', 'new.ts']);

    const u = await undoLast(f.root, bookRef(f.root, f.file(BOOK)));
    expect(u.restored).toEqual(['add.ts']);
    expect(u.removed).toEqual(['new.ts']);
    expect(await f.read('app/add.ts')).toBe(ADD);
    expect(await exists(f.file('app/new.ts'))).toBe(false);
    // ブックには Copilot の編集が残ったままなので、また「Excel側で編集中」になる
    const after = await bookStatus(f.root, f.file(BOOK));
    expect(after.undo).toBeNull();
    expect(after.files.map((x) => [x.name, x.status])).toEqual([
      ['add.ts', 'excel-changed'],
      ['LocalAgents.md', 'clean'],
      ['new.ts', 'excel-new'],
    ]);
    await expect(undoLast(f.root, bookRef(f.root, f.file(BOOK)))).rejects.toThrow('元に戻せる操作がありません');
  });

  it('Sync でシートに取り込んだ内容をブック側で戻す', async () => {
    const f = await setup();
    await f.write('app/add.ts', 'export const add = 2;\n');
    const r = await sync(f.root, f.file(BOOK), { confirmed: true });
    expect(r.status).toBe('ok');
    expect((await f.sheet(BOOK, 'add.ts')).join('\n').trimEnd()).toBe('export const add = 2;');
    expect((await bookStatus(f.root, f.file(BOOK))).undo?.label).toBe('Sync');
    const u = await undoLast(f.root, bookRef(f.root, f.file(BOOK)));
    expect(u.restored).toEqual([]);
    expect((await f.sheet(BOOK, 'add.ts')).join('\n')).toBe(ADD.trimEnd());
    // ソースは触っていないので、また「エディタ側で変更」になる
    const st = await bookStatus(f.root, f.file(BOOK));
    expect(st.files.find((x) => x.name === 'add.ts')?.status).toBe('source-changed');
  });

  it('変更のない Build / Sync は元に戻す対象にならない', async () => {
    const f = await setup();
    await build(f.root, f.file(BOOK), { confirmed: true });
    expect((await bookStatus(f.root, f.file(BOOK))).undo).toBeNull();
    await sync(f.root, f.file(BOOK), { confirmed: true });
    expect((await bookStatus(f.root, f.file(BOOK))).undo).toBeNull();
  });

  it('ブックが開かれていたら戻さない', async () => {
    const f = await setup();
    await f.editBook(BOOK, (b) => b.writeLines('add.ts', ['export const add = 0;']));
    await build(f.root, f.file(BOOK), { confirmed: true });
    await writeFile(f.file('app/~$app.xlcode.xlsx'), '');
    await expect(undoLast(f.root, bookRef(f.root, f.file(BOOK)))).rejects.toThrow('開かれています');
    expect(await f.read('app/add.ts')).toBe('export const add = 0;\n');
  });
});

describe('差分', () => {
  it('Excel 側の変更はソース→Excel、ソース側の変更は Excel→ソースの向きで返す', async () => {
    const f = await setup();
    await f.editBook(BOOK, (b) => b.writeLines('add.ts', ['export const add = 0;']));
    await f.write('app/LocalAgents.md', '# rules\n- new\n');
    const d = await bookDiff(f.root, f.file(BOOK));
    const add = d.find((x) => x.name === 'add.ts')!;
    expect([add.from, add.to]).toEqual(['source', 'excel']);
    expect(add.added).toBe(1);
    expect(add.removed).toBe(3);
    expect(add.hunks[0].lines.filter((l) => l.kind === 'add').map((l) => l.text)).toEqual(['export const add = 0;']);
    const md = d.find((x) => x.name === 'LocalAgents.md')!;
    expect([md.from, md.to]).toEqual(['excel', 'source']);
    expect(md.hunks.length).toBeGreaterThan(0);
  });
});
