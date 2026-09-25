import { chmod, mkdir, readdir, readFile, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { atomicWrite } from '../src/core/atomic';
import { build, createBook, initProject, loadConfig, sync } from '../src/core';
import { StateError } from '../src/core/state';
import { fixture } from './helpers';

const BOOK = 'app/app.xlcode.xlsx';

async function setup() {
  const f = await fixture({ 'app/add.ts': 'export const a = 1;\n' });
  await initProject(f.root);
  await createBook(f.root, f.file('app'));
  return f;
}

describe('一時ファイル経由の保存', () => {
  it('内容・実行権限を保ち、一時ファイルを残さない', async () => {
    const f = await fixture({ 'run.sh': 'echo old\n' });
    await chmod(f.file('run.sh'), 0o755);
    await atomicWrite(f.file('run.sh'), 'echo new\n');
    expect(await f.read('run.sh')).toBe('echo new\n');
    if (process.platform !== 'win32') expect((await stat(f.file('run.sh'))).mode & 0o777).toBe(0o755);
    expect((await readdir(f.root)).filter((n) => n.startsWith('~$'))).toEqual([]);
  });

  it('置き換えに失敗したら一時ファイルを消し、元を残す', async () => {
    const f = await fixture({});
    await mkdir(f.file('target'));
    await writeFile(f.file('target/keep.txt'), 'keep');
    await expect(atomicWrite(f.file('target'), 'x')).rejects.toThrow();
    expect(await f.read('target/keep.txt')).toBe('keep');
    expect((await readdir(f.root)).filter((n) => n.startsWith('~$'))).toEqual([]);
  });

  it('Build 後にブックのディレクトリへ一時ファイルが残らない', async () => {
    const f = await setup();
    await f.editBook(BOOK, (b) => b.writeLines('add.ts', ['export const a = 2;']));
    const r = await build(f.root, f.file(BOOK), { confirmed: true });
    expect(r.status).toBe('ok');
    expect((await readdir(f.file('app'))).filter((n) => n.startsWith('~$'))).toEqual([]);
  });
});

describe('state.json の破損', () => {
  it('壊れていたら黙って初期化せず、エラーで止める', async () => {
    const f = await setup();
    await writeFile(f.file('.xlcode/state.json'), '{ "version": 1, "books": ');
    await expect(sync(f.root, f.file(BOOK), { confirmed: true })).rejects.toThrow(StateError);
    await expect(sync(f.root, f.file(BOOK), { confirmed: true })).rejects.toThrow(/壊れています.*state\.json\.bak/);
  });

  it('形式が違っても止める', async () => {
    const f = await setup();
    await writeFile(f.file('.xlcode/state.json'), JSON.stringify({ version: 1, books: { x: { files: { a: { hash: 1 } } } } }));
    await expect(build(f.root, f.file(BOOK), { confirmed: true })).rejects.toThrow(/x の a の形式が違います/);
  });

  it('保存時に前回分を .bak に残し、.bak から復元できる', async () => {
    const f = await setup();
    const before = await f.read('.xlcode/state.json');
    await f.editBook(BOOK, (b) => b.writeLines('add.ts', ['export const a = 3;']));
    await build(f.root, f.file(BOOK), { confirmed: true });
    expect(await f.read('.xlcode/state.json.bak')).toBe(before);

    // 壊れた → .bak を戻す → 動く（.bak の時点より後の Build 分は Excel 側変更として扱われる）
    await writeFile(f.file('.xlcode/state.json'), 'broken');
    await writeFile(f.file('.xlcode/state.json'), await readFile(f.file('.xlcode/state.json.bak')));
    const r = await sync(f.root, f.file(BOOK), { confirmed: true });
    expect(r.status).not.toBe('error');
  });

  it('config.json が壊れていたら止める', async () => {
    const f = await setup();
    await writeFile(path.join(f.root, '.xlcode/config.json'), '{ shrink');
    await expect(loadConfig(f.root)).rejects.toThrow(/config\.json が JSON として読めません/);
  });
});
