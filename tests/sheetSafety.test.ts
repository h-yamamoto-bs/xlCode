import { readFile, writeFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { bookStatus, build, createBook, initProject, refreshTree, saveConfig, sync } from '../src/core';
import { DEFAULT_CONFIG } from '../src/core/config';
import { isCodeName } from '../src/core/sheetName';
import { Book } from '../src/core/workbook';
import { fixture, type Fixture } from './helpers';

const BOOK = 'app/app.xlcode.xlsx';

describe('コードのシートの見分け方', () => {
  it.each([
    ['App.tsx', true],
    ['run.bat', true],
    ['.gitignore', true],
    ['Makefile', true],
    ['dockerfile', true],
    ['UI', false],
    ['データ', false],
    ['Sheet1', false],
    ['集計表', false],
    ['end.', false],
  ])('%s → %s', (name, expected) => {
    expect(isCodeName(name)).toBe(expected);
  });
  it('設定で追加した名前はコード', () => {
    expect(isCodeName('run', ['run'])).toBe(true);
  });
});

/**
 * Excel で UI 用シートを作った状態を再現する。
 * ExcelJS では保持できない図形（drawing）・印刷範囲・VBA（vbaProject.bin）を持たせる。
 */
async function addUiSheet(f: Fixture, rel: string): Promise<void> {
  const zip = await JSZip.loadAsync(await readFile(f.file(rel)));
  const sheet =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheetData><row r="1"><c r="A1"><v>123</v></c><c r="B1"><f>A1*2</f><v>246</v></c></row></sheetData>' +
    '<drawing r:id="rId1"/></worksheet>';
  zip.file('xl/worksheets/sheet50.xml', sheet);
  zip.file(
    'xl/worksheets/_rels/sheet50.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>',
  );
  zip.file(
    'xl/drawings/drawing1.xml',
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">ボタン</xdr:wsDr>',
  );
  zip.file('xl/vbaProject.bin', Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3, 4]));
  let types = await zip.file('[Content_Types].xml')!.async('string');
  types = types.replace(
    '</Types>',
    '<Override PartName="/xl/worksheets/sheet50.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' +
      '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>',
  );
  zip.file('[Content_Types].xml', types);
  let rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string');
  rels = rels.replace(
    '</Relationships>',
    '<Relationship Id="rId50" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet50.xml"/></Relationships>',
  );
  zip.file('xl/_rels/workbook.xml.rels', rels);
  let wb = await zip.file('xl/workbook.xml')!.async('string');
  const count = (wb.match(/<sheet\b/g) ?? []).length;
  wb = wb.replace('</sheets>', '<sheet name="UI" sheetId="99" r:id="rId50"/></sheets>');
  // UI シートの印刷範囲（シートの位置で参照する）
  const dn = `<definedNames><definedName name="_xlnm.Print_Area" localSheetId="${count}">UI!$A$1:$B$1</definedName></definedNames>`;
  wb = wb.includes('<calcPr') ? wb.replace('<calcPr', `${dn}<calcPr`) : wb.replace('</workbook>', `${dn}</workbook>`);
  zip.file('xl/workbook.xml', wb);
  await writeFile(f.file(rel), await zip.generateAsync({ type: 'nodebuffer' }));
}

async function parts(f: Fixture, rel: string) {
  const zip = await JSZip.loadAsync(await readFile(f.file(rel)));
  const get = async (p: string) => (await zip.file(p)?.async('nodebuffer')) ?? null;
  return {
    ui: await get('xl/worksheets/sheet50.xml'),
    uiRels: await get('xl/worksheets/_rels/sheet50.xml.rels'),
    drawing: await get('xl/drawings/drawing1.xml'),
    vba: await get('xl/vbaProject.bin'),
    workbook: (await zip.file('xl/workbook.xml')!.async('string')) as string,
  };
}

async function setup() {
  const f = await fixture({ 'app/a.ts': 'export const a = 1;\n', 'app/b.ts': 'export const b = 1;\n' });
  await initProject(f.root);
  await createBook(f.root, f.file('app'));
  await addUiSheet(f, BOOK);
  return f;
}

describe('UI 用シートを壊さない', () => {
  it('UI シートは Build / Sync の対象外（数値や数式があってもエラーにしない・ファイルにしない）', async () => {
    const f = await setup();
    const st = await bookStatus(f.root, f.file(BOOK));
    expect(st.errors).toEqual([]);
    expect(st.files.map((x) => x.name)).not.toContain('UI');
    await f.editBook(BOOK, (b) => b.writeLines('a.ts', ['export const a = 2;']));
    const r = await build(f.root, f.file(BOOK), { confirmed: true });
    expect(r.errors).toEqual([]);
    const { readdir } = await import('node:fs/promises');
    expect(await readdir(f.file('app'))).not.toContain('UI');
  });

  it('コードのシートの編集・追加・削除・Sync・Refresh Tree をしても、UI シート・図形・VBA はそのまま', async () => {
    const f = await setup();
    const before = await parts(f, BOOK);

    // 編集・追加・削除（DEL_）→ Build
    await f.editBook(BOOK, (b) => {
      b.writeLines('a.ts', ['export const a = 2;']);
      b.writeLines('c.ts', ['export const c = 3;']);
      const lines = b.readSheet('b.ts').lines;
      b.deleteSheet('b.ts');
      b.writeLines('DEL_b.ts', lines);
    });
    const r = await build(f.root, f.file(BOOK), { confirmed: true });
    expect(r.errors).toEqual([]);
    // ソース側の変更 → Sync、Agents.md の変更 → Refresh Tree
    await f.write('app/a.ts', 'export const a = 20;\n');
    await sync(f.root, f.file(BOOK), { confirmed: true });
    await f.write('Agents.md', '# Agents.md\n\n- 追加\n');
    await refreshTree(f.root);

    const after = await parts(f, BOOK);
    expect(after.ui).toEqual(before.ui);
    expect(after.uiRels).toEqual(before.uiRels);
    expect(after.drawing).toEqual(before.drawing);
    expect(after.vba).toEqual(before.vba);
    // 印刷範囲は UI シートの新しい位置を指す
    const book = await Book.load(f.file(BOOK));
    const uiIndex = book.sheetNames().indexOf('UI');
    expect(after.workbook).toContain(`localSheetId="${uiIndex}">UI!$A$1:$B$1`);
    // コードは正しく反映されている
    expect(book.readSheet('a.ts').lines).toEqual(['export const a = 20;']);
    expect(book.hasSheet('b.ts')).toBe(false);
    expect(book.hasSheet('DEL_b.ts')).toBe(false);
    expect(await f.read('app/c.ts')).toBe('export const c = 3;\n');
  });

  it('拡張子のないソースファイルは対象外（設定で追加できる）', async () => {
    const f = await fixture({ 'app/a.ts': 'x\n', 'app/run': '#!/bin/sh\n', 'app/Makefile': 'all:\n' });
    await initProject(f.root);
    const cb = await createBook(f.root, f.file('app'));
    expect(cb.sheets).toContain('Makefile');
    expect(cb.sheets).not.toContain('run');
    const st = await bookStatus(f.root, f.file(BOOK));
    expect(st.warnings.join()).toMatch(/拡張子のないファイル「run」は対象外/);

    await saveConfig(f.root, { ...DEFAULT_CONFIG, extraCodeNames: ['run'] });
    const r = await sync(f.root, f.file(BOOK), { confirmed: true });
    expect(r.changes).toContainEqual({ action: 'write-sheet', target: 'run' });
  });
});
