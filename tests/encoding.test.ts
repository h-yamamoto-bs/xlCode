import { readFile, writeFile } from 'node:fs/promises';
import iconv from 'iconv-lite';
import { describe, expect, it } from 'vitest';
import { bookStatus, build, createBook, initProject, sync } from '../src/core';
import { decodeFile, defaultFormat, encodeFile, formatLabel } from '../src/core/encoding';
import { fixture } from './helpers';

const BOOK = 'app/app.xlcode.xlsx';
const sjis = (s: string) => iconv.encode(s, 'cp932');
const bom = (s: string) => Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(s, 'utf8')]);
const utf16 = (s: string) => Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(s, 'utf16le')]);

describe('文字コード・改行コードの判定', () => {
  it.each([
    ['a.ts', Buffer.from('const あ = 1;\n'), 'UTF-8 / LF'],
    ['a.ts', Buffer.from('a\r\nb\r\n'), 'UTF-8 / CRLF'],
    ['a.ps1', bom('Write-Host "日本語"\r\n'), 'UTF-8 BOM / CRLF'],
    ['a.reg', utf16('Windows Registry Editor Version 5.00\r\n'), 'UTF-16 LE / CRLF'],
    ['a.bat', sjis('@echo off\r\necho 日本語\r\n'), 'Shift_JIS / CRLF'],
    ['a.txt', sjis('日本語だけ'), 'Shift_JIS / LF'],
    // ASCII だけのファイルは拡張子の標準に合わせる
    ['a.bat', Buffer.from('@echo off\r\n'), 'Shift_JIS / CRLF'],
    ['a.ps1', Buffer.from('Get-Date\r\n'), 'UTF-8 BOM / CRLF'],
    ['a.reg', Buffer.from('REGEDIT4\r\n'), 'UTF-8 / CRLF'],
    ['a.ts', Buffer.from('x'), 'UTF-8 / LF'],
  ])('%s %s', (name, buf, label) => {
    const d = decodeFile(buf, name);
    expect(d.kind).toBe('text');
    if (d.kind === 'text') expect(formatLabel(d.format)).toBe(label);
  });

  it('日本語を正しく読む', () => {
    const d = decodeFile(sjis('echo こんにちは\r\n'), 'a.bat');
    expect(d.kind === 'text' && d.text).toBe('echo こんにちは\r\n');
  });

  it('バイナリ・UTF-16 BE・壊れたファイルはスキップ', () => {
    expect(decodeFile(Buffer.from([0x89, 0x50, 0, 1]), 'a.png').kind).toBe('binary');
    expect(decodeFile(Buffer.from([0xfe, 0xff, 0, 0x41]), 'a.txt').kind).toBe('binary');
    expect(decodeFile(Buffer.from([0x81, 0xff, 0xfd]), 'a.txt').kind).toBe('binary');
  });

  it('新しいファイルの形式（Windows 11 日本語環境）', () => {
    expect(formatLabel(defaultFormat('run.BAT'))).toBe('Shift_JIS / CRLF');
    expect(formatLabel(defaultFormat('run.cmd'))).toBe('Shift_JIS / CRLF');
    expect(formatLabel(defaultFormat('tool.ps1'))).toBe('UTF-8 BOM / CRLF');
    expect(formatLabel(defaultFormat('a.reg'))).toBe('UTF-16 LE / CRLF');
    expect(formatLabel(defaultFormat('App.tsx'))).toBe('UTF-8 / LF');
  });
});

describe('書き出し', () => {
  it('形式どおりのバイト列にする', () => {
    expect(encodeFile('a\nb\n', { encoding: 'sjis', eol: 'crlf' }, 'a.bat')).toEqual(sjis('a\r\nb\r\n'));
    expect(encodeFile('あ\n', { encoding: 'utf8bom', eol: 'crlf' }, 'a.ps1')).toEqual(bom('あ\r\n'));
    expect(encodeFile('x\n', { encoding: 'utf16le', eol: 'crlf' }, 'a.reg')).toEqual(utf16('x\r\n'));
  });
  it('Shift_JIS で表せない文字はエラー', () => {
    expect(() => encodeFile('echo 😀 ①\n', { encoding: 'sjis', eol: 'crlf' }, 'a.bat')).toThrow(
      /Shift_JIS で表せない文字があります: 😀/,
    );
  });
});

describe('Build / Sync で形式を保つ', () => {
  async function setup(files: Record<string, string | Buffer>) {
    const f = await fixture({});
    for (const [rel, data] of Object.entries(files)) {
      await f.write(rel, '');
      await writeFile(f.file(rel), data);
    }
    await initProject(f.root);
    await createBook(f.root, f.file('app'));
    return f;
  }

  it('Shift_JIS・CRLF の .bat を Excel で編集しても Shift_JIS・CRLF のまま', async () => {
    const f = await setup({ 'app/run.bat': sjis('@echo off\r\necho 開始\r\n') });
    expect(await f.sheet(BOOK, 'run.bat')).toEqual(['@echo off', 'echo 開始']);
    await f.editBook(BOOK, (b) => b.writeLines('run.bat', ['@echo off', 'echo 開始', 'echo 終了']));
    const r = await build(f.root, f.file(BOOK), { confirmed: true });
    expect(r.errors).toEqual([]);
    expect(await readFile(f.file('app/run.bat'))).toEqual(sjis('@echo off\r\necho 開始\r\necho 終了\r\n'));
    expect((await sync(f.root, f.file(BOOK), { confirmed: true })).changes).toEqual([]);
    const st = await bookStatus(f.root, f.file(BOOK));
    expect(st.files.find((x) => x.name === 'run.bat')?.format).toBe('Shift_JIS / CRLF');
  });

  it('Excel で新しく作ったファイルは拡張子の標準形式で書き出す', async () => {
    const f = await setup({ 'app/a.ts': 'export {};\n' });
    await f.editBook(BOOK, (b) => {
      b.writeLines('setup.cmd', ['@echo off', 'echo セットアップ']);
      b.writeLines('tool.ps1', ['Write-Host "ツール"']);
      b.writeLines('b.ts', ['export const b = "日本語";']);
    });
    const r = await build(f.root, f.file(BOOK), { confirmed: true });
    expect(r.errors).toEqual([]);
    expect(await readFile(f.file('app/setup.cmd'))).toEqual(sjis('@echo off\r\necho セットアップ\r\n'));
    expect(await readFile(f.file('app/tool.ps1'))).toEqual(bom('Write-Host "ツール"\r\n'));
    expect(await readFile(f.file('app/b.ts'), 'utf8')).toBe('export const b = "日本語";\n');
  });

  it('UTF-8 BOM・CRLF のファイルはそのまま保つ', async () => {
    const f = await setup({ 'app/notes.txt': bom('メモ\r\n') });
    await f.editBook(BOOK, (b) => b.writeLines('notes.txt', ['メモ', '追記']));
    await build(f.root, f.file(BOOK), { confirmed: true });
    expect(await readFile(f.file('app/notes.txt'))).toEqual(bom('メモ\r\n追記\r\n'));
  });

  it('Shift_JIS のファイルに表せない文字が入ったら、何も書かずに止める', async () => {
    const f = await setup({ 'app/run.bat': sjis('echo a\r\n'), 'app/a.ts': 'export {};\n' });
    await f.editBook(BOOK, (b) => {
      b.writeLines('run.bat', ['echo 😀']);
      b.writeLines('a.ts', ['export const x = 1;']);
    });
    const r = await build(f.root, f.file(BOOK), { confirmed: true });
    expect(r.status).toBe('error');
    expect(r.errors[0]).toMatch(/run\.bat.*😀/);
    expect(await readFile(f.file('app/run.bat'))).toEqual(sjis('echo a\r\n'));
    expect(await f.read('app/a.ts')).toBe('export {};\n');
  });
});
