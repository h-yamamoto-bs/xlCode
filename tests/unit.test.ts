import { describe, expect, it } from 'vitest';
import { normalizeText, linesToText, textToLines } from '../src/core/normalize';
import { classifySheet, validateFileName } from '../src/core/sheetName';
import { decideStatus } from '../src/core/scan';
import { Book } from '../src/core/workbook';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';

const opts = { trimTrailingWhitespace: true };

describe('正規化', () => {
  it('CRLF・BOM・末尾空行・行末空白を正規化する', () => {
    expect(normalizeText('﻿a  \r\nb\r\n\r\n\r\n', 'a.ts', opts)).toBe('a\nb\n');
    expect(normalizeText('a', 'a.ts', opts)).toBe('a\n');
    expect(normalizeText('', 'a.ts', opts)).toBe('');
    expect(normalizeText('\n\n', 'a.ts', opts)).toBe('');
  });
  it('途中の空行は保持する', () => {
    expect(normalizeText('a\n\n\nb', 'a.ts', opts)).toBe('a\n\n\nb\n');
  });
  it('Markdown の行末空白（改行記法）は保持する', () => {
    expect(normalizeText('a  \nb', 'x.md', opts)).toBe('a  \nb\n');
  });
  it('行配列と相互変換できる', () => {
    expect(textToLines('a\n\nb\n')).toEqual(['a', '', 'b']);
    expect(linesToText(['a', '', 'b'])).toBe('a\n\nb\n');
    expect(textToLines('')).toEqual([]);
  });
});

describe('シート名', () => {
  it('31文字以内・禁止記号を検証する', () => {
    expect(validateFileName('App.tsx')).toBeNull();
    expect(validateFileName('a'.repeat(28) + '.ts')).toBeNull();
    expect(validateFileName('a'.repeat(29) + '.ts')).toMatch(/32 文字/);
    expect(validateFileName('[id].tsx')).toMatch(/\[ \]/);
    expect(validateFileName('a:b')).toMatch(/:/);
    expect(validateFileName('#x.ts')).toMatch(/予約/);
  });
  it('予約シートを分類する', () => {
    expect(classifySheet('#tree').kind).toBe('tree');
    expect(classifySheet('#meta').kind).toBe('reserved');
    expect(classifySheet('#conflict_03').kind).toBe('conflict');
    expect(classifySheet('#foo').kind).toBe('unknown-reserved');
    expect(classifySheet('Agents.md').kind).toBe('agents');
    expect(classifySheet('DEL_App.tsx')).toEqual({ kind: 'delete', fileName: 'App.tsx' });
    expect(classifySheet('_app.tsx')).toEqual({ kind: 'code', fileName: '_app.tsx' });
  });
});

describe('状態判定', () => {
  const c = (hash: string) => ({ text: '', hash, lines: 0, chars: 0 });
  const p = { hash: 'P', lines: 0, chars: 0 };
  it.each([
    [p, c('P'), c('P'), 'clean'],
    [p, c('P'), c('X'), 'excel-changed'],
    [p, c('S'), c('P'), 'source-changed'],
    [p, c('S'), c('X'), 'conflict'],
    [p, c('S'), c('S'), 'clean'],
    [p, undefined, c('P'), 'source-deleted'],
    [p, undefined, c('X'), 'conflict'],
    [p, c('P'), undefined, 'sheet-missing'],
    [undefined, c('S'), undefined, 'source-new'],
    [undefined, undefined, c('X'), 'excel-new'],
    [undefined, c('S'), c('X'), 'conflict'],
    [p, undefined, undefined, 'gone'],
  ] as const)('%#', (prev, src, xl, expected) => {
    expect(decideStatus(prev, src, xl)).toBe(expected);
  });
});

describe('ブックの往復（Excel の自動変換対策）', () => {
  it('数式・先行ゼロ・日付・指数・先頭空白・空行を文字列のまま保持する', async () => {
    const lines = [
      '=SUM(A1:A3)',
      '007',
      '2026-01-01',
      '1e5',
      '    indented',
      '',
      '\ttab',
      'TRUE',
      '日本語 😀',
      "'quoted",
    ];
    const file = path.join(await mkdtemp(path.join(tmpdir(), 'xlb-')), 'x.xlsx');
    const b = Book.create();
    b.writeLines('a.ts', lines);
    await b.save(file);
    const data = (await Book.load(file)).readSheet('a.ts');
    expect(data.lines).toEqual(lines);
    expect(data.issues).toEqual([]);
  });
  it('数値・真偽値に変換されたセルを検出する', async () => {
    const file = path.join(await mkdtemp(path.join(tmpdir(), 'xlb-')), 'x.xlsx');
    const b = Book.create();
    const ws = b.wb.addWorksheet('a.ts');
    ws.getCell('A1').value = 'ok';
    ws.getCell('A2').value = 7;
    ws.getCell('A3').value = true;
    ws.getCell('A4').value = { formula: 'SUM(1,2)' } as never;
    await b.save(file);
    const data = (await Book.load(file)).readSheet('a.ts');
    expect(data.issues.map((i) => [i.row, i.type])).toEqual([
      [2, 'number'],
      [3, 'boolean'],
      [4, 'formula'],
    ]);
  });
  it('32,767 文字を超える行は書き込まない', () => {
    expect(() => Book.create().writeLines('a.js', ['x'.repeat(32768)])).toThrow(/32767/);
  });
});
