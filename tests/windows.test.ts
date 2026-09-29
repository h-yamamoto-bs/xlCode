import { mkdir, mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { bookStatus, build, createBook, initProject, sync } from '../src/core';
import { excelPathError, isBookCopy } from '../src/core/fsutil';
import { parseStatusZ, uncommittedChanges } from '../src/core/git';
import { validateFileName } from '../src/core/sheetName';
import { Book, escapeCell } from '../src/core/workbook';
import { fixture } from './helpers';

const BOOK = 'app/app.xlcode.xlsx';

describe('xlsx に保存すると化ける文字列', () => {
  const tricky = ['a_x000D_b', '_x0041_', 'c\u0001d', 'e\u000Bf', 'x\u001Fy', '_x005F_', "'_x0020_'"];

  it('_xHHHH_ と制御文字をエスケープして保存し、元に戻る', async () => {
    const file = path.join(await mkdtemp(path.join(tmpdir(), 'xlb-')), 'x.xlsx');
    const b = Book.create();
    b.writeLines('a.ts', tricky);
    await b.save(file);
    expect((await Book.load(file)).readSheet('a.ts').lines).toEqual(tricky);
  });

  it('エスケープの規則', () => {
    expect(escapeCell('_x0041_')).toBe('_x005F_x0041_');
    expect(escapeCell('a\u0001')).toBe('a_x0001_');
    expect(escapeCell('tab\tok')).toBe('tab\tok');
    expect(escapeCell('_x00_')).toBe('_x00_');
  });

  it('壊れた文字（孤立サロゲート）は保存しない', () => {
    expect(() => Book.create().writeLines('a.ts', ['ok', 'bad \uD800 x'])).toThrow(/2 行目に壊れた文字/);
  });

  it('ソース → Sync → Build の往復でも化けない', async () => {
    const src = `export const cr = "_x000D_";\nexport const ctrl = "\u0001";\n`;
    const f = await fixture({ 'app/esc.ts': src });
    await initProject(f.root);
    await createBook(f.root, f.file('app'));
    expect((await f.sheet(BOOK, 'esc.ts'))[0]).toBe('export const cr = "_x000D_";');
    // Excel 側で別の行を編集して Build しても、他の行は元のまま
    await f.editBook(BOOK, (b) => b.writeLines('esc.ts', [...b.readSheet('esc.ts').lines, 'export const x = 1;']));
    const r = await build(f.root, f.file(BOOK), { confirmed: true });
    expect(r.errors).toEqual([]);
    expect(await f.read('app/esc.ts')).toBe(src + 'export const x = 1;\n');
    expect((await sync(f.root, f.file(BOOK), { confirmed: true })).changes).toEqual([]);
  });
});

describe('Windows で使えないファイル名', () => {
  it.each(['CON.ts', 'con', 'nul.txt', 'COM1.js', 'lpt9.md', 'a<b.ts', 'a|b', 'a"b', 'end.', 'end ', 'a\u0001b'])(
    '%s',
    (name) => {
      expect(validateFileName(name)).not.toBeNull();
    },
  );
  it.each(['console.ts', 'con_fig.ts', 'nullable.ts', 'App.tsx', '_app.tsx', '.gitignore'])('%s は使える', (name) => {
    expect(validateFileName(name)).toBeNull();
  });
  it('Excel で作られた予約名のシートは Build でエラーにする', async () => {
    const f = await fixture({ 'app/a.ts': 'export {};\n' });
    await initProject(f.root);
    await createBook(f.root, f.file('app'));
    await f.editBook(BOOK, (b) => b.writeLines('con.ts', ['x']));
    const r = await build(f.root, f.file(BOOK), { confirmed: true });
    expect(r.status).toBe('error');
    expect(r.errors.join()).toMatch(/con\.ts.*予約名/);
  });
});

describe('Git の日本語ファイル名', () => {
  it('status -z を解析する（名前変更は新しいパス）', () => {
    expect(parseStatusZ(' M app/日本語.ts\0R  app/new.ts\0app/old.ts\0?? app/a b.ts\0')).toEqual([
      'app/日本語.ts',
      'app/new.ts',
      'app/a b.ts',
    ]);
  });
  it('未コミット変更の日本語ファイル名が化けない', async () => {
    const f = await fixture({ 'app/a.ts': 'x\n' }, { git: true });
    await f.write('app/日本語 ファイル.ts', 'y\n');
    expect(await uncommittedChanges(f.root, 'app')).toEqual(['app/日本語 ファイル.ts']);
  });
});

describe('OneDrive・Excel のパス', () => {
  it('同期がぶつかったときの複製を見分ける', () => {
    expect(isBookCopy('app.xlcode-DESKTOP-1234.xlsx')).toBe(true);
    expect(isBookCopy('app.xlcode (1).xlsx')).toBe(true);
    expect(isBookCopy('app.xlcode.xlsx')).toBe(false);
    expect(isBookCopy('~$app.xlcode.xlsx')).toBe(false);
    expect(isBookCopy('report.xlsx')).toBe(false);
  });
  it('複製があれば警告する', async () => {
    const f = await fixture({ 'app/a.ts': 'x\n' });
    await initProject(f.root);
    await createBook(f.root, f.file('app'));
    const { copyFile } = await import('node:fs/promises');
    await copyFile(f.file(BOOK), f.file('app/app.xlcode-DESKTOP-1234.xlsx'));
    const st = await bookStatus(f.root, f.file(BOOK));
    expect(st.warnings.join()).toMatch(/ブックの複製があります: app\.xlcode-DESKTOP-1234\.xlsx/);
  });
  it('Excel が開けない長さのパスはブックを作らない', async () => {
    expect(excelPathError('C:\\' + 'a'.repeat(215))).toBeNull();
    expect(excelPathError('C:\\' + 'a'.repeat(216))).toMatch(/219 文字/);
    const f = await fixture({});
    const deep = path.join(f.root, 'd'.repeat(120), 'e'.repeat(100));
    await mkdir(deep, { recursive: true });
    await expect(createBook(f.root, deep)).rejects.toThrow(/Excel が開ける 218 文字を超えています/);
  });
});
