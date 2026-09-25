import { describe, expect, it } from 'vitest';
import { bookStatus, build, createBook, initProject, sync } from '../src/core';
import { fixture } from './helpers';

/** 13章 受け入れ基準: 1ブック 20ファイル × 各500行 */
function makeFile(seed: number): string {
  const out: string[] = [];
  for (let i = 0; out.length < 500; i++) {
    out.push(`// 関数 ${seed}-${i}`);
    out.push(`export function f${seed}_${i}(a: number, b: string): string {`);
    out.push(`  const v = a * ${i} + ${seed};`);
    out.push('  return `${b}:${v}`;');
    out.push('}');
    out.push('');
  }
  return out.slice(0, 500).join('\n') + '\n';
}

describe('受け入れ基準', () => {
  it('20ファイル×500行を全件変更しても Build が3秒以内、往復で差分が出ない', async () => {
    const files: Record<string, string> = {};
    for (let i = 0; i < 20; i++) files[`app/mod${i}.ts`] = makeFile(i);
    const f = await fixture(files);
    await initProject(f.root);
    await createBook(f.root, f.file('app'));
    const BOOK = 'app/app.xlcode.xlsx';

    // Copilot が全シートを書き換えた想定（末尾に1関数追加）
    await f.editBook(BOOK, (b) => {
      for (let i = 0; i < 20; i++) {
        const name = `mod${i}.ts`;
        b.writeLines(name, [...b.readSheet(name).lines, 'export const added = 1;']);
      }
    });

    const t0 = performance.now();
    const r = await build(f.root, f.file(BOOK), { confirmed: true });
    const ms = performance.now() - t0;
    expect(r.errors).toEqual([]);
    expect(r.changes.filter((c) => c.action === 'write-file')).toHaveLength(20);
    console.log(`Build 20ファイル×500行: ${Math.round(ms)}ms`);
    expect(ms).toBeLessThan(3000);

    const s = await sync(f.root, f.file(BOOK), { confirmed: true });
    expect(s.changes).toEqual([]);
    const st = await bookStatus(f.root, f.file(BOOK));
    expect(st.files.every((x) => x.status === 'clean')).toBe(true);
  }, 30000);
});
