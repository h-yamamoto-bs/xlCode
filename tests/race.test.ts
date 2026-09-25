import { describe, expect, it, vi } from 'vitest';

// 1回目（preflight）は閉じている、それ以降は開かれた状態を再現する
let calls = 0;
let openFrom = Infinity;
vi.mock('../src/core/lock', () => ({
  checkBookOpen: async () => {
    calls++;
    return calls >= openFrom ? { open: true, reason: 'テスト' } : { open: false };
  },
}));

const { build, sync, createBook, initProject } = await import('../src/core');
const { fixture } = await import('./helpers');

const BOOK = 'app/app.xlcode.xlsx';

async function setup() {
  const f = await fixture({ 'app/add.ts': 'export const a = 1;\n' });
  await initProject(f.root);
  await createBook(f.root, f.file('app'));
  return f;
}

describe('確認から保存までの間に開かれた場合', () => {
  it('Build: ソースを書く前に中断する', async () => {
    const f = await setup();
    await f.editBook(BOOK, (b) => b.writeLines('add.ts', ['export const a = 2;']));
    calls = 0;
    openFrom = 2;
    const r = await build(f.root, f.file(BOOK), { confirmed: true });
    expect(r.status).toBe('error');
    expect(r.errors[0]).toMatch(/途中で開かれた.*ソースは変更していません/);
    expect(await f.read('app/add.ts')).toBe('export const a = 1;\n');
  });

  it('Sync: ブックを保存せずに中断する', async () => {
    const f = await setup();
    await f.write('app/add.ts', 'export const a = 9;\n');
    const before = await f.sheet(BOOK, 'add.ts');
    calls = 0;
    openFrom = 2;
    const r = await sync(f.root, f.file(BOOK), { confirmed: true });
    expect(r.status).toBe('error');
    expect(r.errors[0]).toMatch(/ブックは変更していません/);
    expect(await f.sheet(BOOK, 'add.ts')).toEqual(before);
    openFrom = Infinity;
    // 閉じた後にやり直せば反映される
    expect((await sync(f.root, f.file(BOOK), { confirmed: true })).status).toBe('ok');
    expect(await f.sheet(BOOK, 'add.ts')).toEqual(['export const a = 9;']);
  });
});
