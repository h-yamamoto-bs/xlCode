import { describe, expect, it } from 'vitest';
import { diffLines, summarize, toHunks } from '../src/core/diff';

const show = (old: string[], next: string[]) =>
  diffLines(old, next).map((l) => `${{ same: ' ', del: '-', add: '+' }[l.kind]}${l.text}`);

describe('diffLines', () => {
  it('同じなら全行 same', () => {
    expect(show(['a', 'b'], ['a', 'b'])).toEqual([' a', ' b']);
  });

  it('追加・削除・変更を最小の編集で表す', () => {
    expect(show(['a', 'b', 'c'], ['a', 'x', 'c', 'd'])).toEqual([' a', '-b', '+x', ' c', '+d']);
    expect(show([], ['a'])).toEqual(['+a']);
    expect(show(['a'], [])).toEqual(['-a']);
  });

  it('共通の先頭・末尾を除いても行番号がずれない', () => {
    const old = ['h1', 'h2', 'x', 'y', 't1'];
    const next = ['h1', 'h2', 'x', 'z', 'w', 't1'];
    const hunks = toHunks(diffLines(old, next), 1);
    expect(hunks).toHaveLength(1);
    expect(hunks[0].oldStart).toBe(3);
    expect(hunks[0].newStart).toBe(3);
    expect(hunks[0].lines.map((l) => l.kind)).toEqual(['same', 'del', 'add', 'add', 'same']);
  });

  it('離れた変更は別のハンクになる', () => {
    const old = Array.from({ length: 30 }, (_, i) => `l${i}`);
    const next = [...old];
    next[2] = 'changed';
    next[25] = 'changed2';
    const hunks = toHunks(diffLines(old, next), 2);
    expect(hunks).toHaveLength(2);
    expect(hunks[1].oldStart).toBe(24);
    expect(summarize(diffLines(old, next))).toEqual({ added: 2, removed: 2 });
  });

  it('数千行のファイルでも最小の編集を返す', () => {
    const old = Array.from({ length: 1500 }, (_, i) => `line ${i}`);
    const next = old.map((l, i) => (i % 7 === 0 ? `${l}!` : l));
    const d = diffLines(old, next);
    expect(summarize(d).added).toBe(Math.ceil(1500 / 7));
  });

  it('大きすぎる差分は全置換として扱い、止まらない', () => {
    const old = Array.from({ length: 20000 }, (_, i) => `a ${i}`);
    const next = Array.from({ length: 20000 }, (_, i) => `b ${i}`);
    const t = Date.now();
    expect(summarize(diffLines(old, next))).toEqual({ added: 20000, removed: 20000 });
    expect(Date.now() - t).toBeLessThan(2000);
  });
});
