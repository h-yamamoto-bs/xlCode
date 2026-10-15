import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import JSZip from 'jszip';
import { afterEach, describe, expect, it } from 'vitest';
import {
  applyTheme,
  bookStatus,
  createBook,
  getHostTheme,
  initProject,
  resolveBookTheme,
  saveConfig,
  setHostTheme,
  sync,
  THEMES,
} from '../src/core';
import { DEFAULT_CONFIG, loadConfig } from '../src/core/config';
import { restyleSheetXml } from '../src/core/xlsxPatch';
import { Book } from '../src/core/workbook';
import { fixture, type Fixture } from './helpers';

const require = createRequire(import.meta.url);
// sax（ExcelJS が使っている XML パーサー）には型定義が無いので、使う分だけ書く
const sax = require('sax') as { parser: (strict: boolean) => { write: (xml: string) => { close: () => void } } };

afterEach(() => setHostTheme(null));

/** XML として壊れていないか（Excel の「修復しますか」の原因になる）を、厳密モードの sax で確かめる */
function assertWellFormed(xml: string): void {
  const parser = sax.parser(true);
  parser.write(xml).close();
}

async function parts(file: string) {
  const zip = await JSZip.loadAsync(await readFile(file));
  const wb = (await zip.file('xl/workbook.xml')!.async('string'))!;
  const rels = (await zip.file('xl/_rels/workbook.xml.rels')!.async('string'))!;
  const sheet = async (name: string) => {
    const tag = [...wb.matchAll(/<sheet\b[^>]*\/>/g)].map((m) => m[0]).find((t) => t.includes(`name="${name}"`))!;
    const rid = /r:id="([^"]+)"/.exec(tag)![1];
    const rel = [...rels.matchAll(/<Relationship\b[^>]*\/>/g)].map((m) => m[0]).find((t) => t.includes(`Id="${rid}"`))!;
    return (await zip.file(`xl/${/Target="([^"]+)"/.exec(rel)![1]}`)!.async('string'))!;
  };
  return { zip, sheet, styles: (await zip.file('xl/styles.xml')!.async('string'))! };
}

const BOOK = 'app/app.xlcode.xlsx';
async function project(theme?: 'dark' | 'light' | 'auto'): Promise<Fixture> {
  const f = await fixture({ 'app/a.ts': 'export const a = 1;\n', 'app/b.ts': 'export const b = 2;\n' });
  await initProject(f.root);
  if (theme) await saveConfig(f.root, { ...DEFAULT_CONFIG, bookTheme: theme });
  await createBook(f.root, f.file('app'));
  return f;
}

describe('ブックのテーマの決め方', () => {
  it('auto は xlCode の表示モードに合わせ、分からなければダーク。固定ならそのまま', () => {
    expect(resolveBookTheme('auto', 'light')).toBe('light');
    expect(resolveBookTheme(undefined, 'dark')).toBe('dark');
    expect(resolveBookTheme('auto', null)).toBe('dark');
    expect(resolveBookTheme('light', 'dark')).toBe('light');
    expect(resolveBookTheme('dark', 'light')).toBe('dark');
    setHostTheme('light');
    expect(getHostTheme()).toBe('light');
  });

  it('設定に知らない値があれば auto として扱う', async () => {
    const f = await fixture({ 'a.ts': 'x\n' });
    await saveConfig(f.root, { ...DEFAULT_CONFIG, bookTheme: 'pink' as never });
    expect((await loadConfig(f.root)).bookTheme).toBe('auto');
  });
});

describe('ブックを作るときにテーマを当てる', () => {
  it('ダーク: 黒に近い背景・明るい文字・枠線なし・タブの色。データの無い範囲まで塗りつぶす', async () => {
    const f = await project('dark');
    const p = await parts(f.file(BOOK));
    const xml = await p.sheet('a.ts');
    assertWellFormed(xml);
    expect(xml).toContain('showGridLines="0"');
    expect(xml).toContain(`<tabColor rgb="FF${THEMES.dark.tab.code}"/>`);
    // A 列は文字列書式のスタイル、B 列から最後の列（XFD）までは塗りつぶしだけのスタイル
    expect(xml).toMatch(/<col min="1" max="1" width="120" customWidth="1" style="\d+"\/>/);
    expect(xml).toMatch(/<col min="2" max="16384" width="9\.140625" style="\d+"\/>/);
    // sheetPr は dimension より前、cols は sheetData より前（スキーマの順序）
    expect(xml.indexOf('<sheetPr>')).toBeLessThan(xml.indexOf('<dimension'));
    expect(xml.indexOf('<cols>')).toBeLessThan(xml.indexOf('<sheetData>'));
    assertWellFormed(p.styles);
    expect(p.styles).toContain(`<fgColor rgb="FF${THEMES.dark.fill}"/>`);
    expect(p.styles).toContain(`<color rgb="FF${THEMES.dark.text}"/>`);
    expect(p.styles).toContain('<name val="Consolas"/>');
    // #tree と Agents.md は meta の色
    expect(await p.sheet('#tree')).toContain(`<tabColor rgb="FF${THEMES.dark.tab.meta}"/>`);
    expect(await p.sheet('Agents.md')).toContain(`<tabColor rgb="FF${THEMES.dark.tab.meta}"/>`);
    expect(await p.sheet('LocalAgents.md')).toContain(`<tabColor rgb="FF${THEMES.dark.tab.meta}"/>`);
  });

  it('ライト: ふつうの白い背景（Excel がダークモードでも白）・枠線なし', async () => {
    const f = await project('light');
    const p = await parts(f.file(BOOK));
    const xml = await p.sheet('a.ts');
    expect(xml).toContain('showGridLines="0"');
    expect(xml).toContain(`<tabColor rgb="FF${THEMES.light.tab.code}"/>`);
    expect(p.styles).toContain(`<fgColor rgb="FF${THEMES.light.fill}"/>`);
    expect(p.styles).toContain(`<color rgb="FF${THEMES.light.text}"/>`);
    expect(p.styles).not.toContain(`<fgColor rgb="FF${THEMES.dark.fill}"/>`);
  });

  it('auto は xlCode の表示モードで決まる', async () => {
    setHostTheme('light');
    const f = await project('auto');
    expect(await (await parts(f.file(BOOK))).sheet('a.ts')).toContain(`FF${THEMES.light.tab.code}`);
    setHostTheme('dark');
    const g = await project('auto');
    expect(await (await parts(g.file(BOOK))).sheet('a.ts')).toContain(`FF${THEMES.dark.tab.code}`);
  });

  it('中身は変わらず、何度書いても styles.xml が増え続けない。ExcelJS でも読める', async () => {
    const f = await project('dark');
    const before = (await parts(f.file(BOOK))).styles;
    const count = (s: string, tag: string) => Number(new RegExp(`<${tag}\\b[^>]*count="(\\d+)"`).exec(s)![1]);
    for (let i = 0; i < 4; i++) {
      await f.write('app/a.ts', `export const a = ${i + 10};\n`);
      await sync(f.root, f.file(BOOK), { confirmed: true });
    }
    const after = (await parts(f.file(BOOK))).styles;
    for (const tag of ['fonts', 'fills', 'cellXfs']) expect(count(after, tag)).toBe(count(before, tag));
    expect(await f.sheet(BOOK, 'a.ts')).toEqual(['export const a = 13;']);
    const b = await Book.load(f.file(BOOK));
    expect(b.sheetTheme('a.ts')).toBe('dark');
    expect(b.sheetTheme('#tree')).toBe('dark');
  });

  it('Sync で書き直したシートにもテーマが当たる', async () => {
    const f = await project('dark');
    await f.write('app/c.ts', 'export const c = 3;\n');
    await sync(f.root, f.file(BOOK), { confirmed: true });
    const p = await parts(f.file(BOOK));
    const xml = await p.sheet('c.ts');
    assertWellFormed(xml);
    expect(xml).toContain('showGridLines="0"');
    expect(xml).toContain(`FF${THEMES.dark.tab.code}`);
  });
});

describe('テーマを当て直す', () => {
  it('設定を変えると状態に出て、当て直すと中身はそのままテーマだけが変わる', async () => {
    const f = await project('dark');
    const st0 = await bookStatus(f.root, f.file(BOOK));
    expect(st0.theme).toMatchObject({ desired: 'dark', outdated: [] });
    expect(st0.theme.managed).toBe(5);

    await saveConfig(f.root, { ...DEFAULT_CONFIG, bookTheme: 'light' });
    const st1 = await bookStatus(f.root, f.file(BOOK));
    expect(st1.theme.desired).toBe('light');
    expect(st1.theme.outdated.sort()).toEqual(['#tree', 'Agents.md', 'LocalAgents.md', 'a.ts', 'b.ts']);

    const r = await applyTheme(f.root, f.file(BOOK));
    expect(r.theme).toBe('light');
    expect(r.sheets).toHaveLength(5);
    const p = await parts(f.file(BOOK));
    const xml = await p.sheet('a.ts');
    assertWellFormed(xml);
    expect(xml).toContain(`<tabColor rgb="FF${THEMES.light.tab.code}"/>`);
    expect(xml).not.toContain(THEMES.dark.tab.code);
    expect(xml).toContain('showGridLines="0"');
    // セルのスタイルは 1 つ（s が重複していない）で、中身はそのまま
    expect(xml).not.toMatch(/ s="\d+"[^>]* s="\d+"/);
    expect(await f.sheet(BOOK, 'a.ts')).toEqual(['export const a = 1;']);
    expect((await bookStatus(f.root, f.file(BOOK))).theme.outdated).toEqual([]);
    // もう一度当てても何もしない（ブックを保存しない）
    expect((await applyTheme(f.root, f.file(BOOK))).sheets).toEqual([]);
  });

  it('Copilot が作ったテーマの無いシートも見つけて当てる。画面・データ用のシートは触らない', async () => {
    const f = await project('dark');
    // Excel で保存し直した状態: 新しいコードのシートと、UI 用のシートを足す
    await f.saveAsExcel(BOOK, (wb) => {
      wb.addWorksheet('new.ts').getCell('A1').value = 'export {};';
      wb.addWorksheet('UI').getCell('B2').value = 'ボタンの説明';
    });
    const uiBefore = await (await parts(f.file(BOOK))).sheet('UI');
    const st = await bookStatus(f.root, f.file(BOOK));
    expect(st.theme.outdated).toContain('new.ts');
    expect(st.theme.outdated).not.toContain('UI');
    await applyTheme(f.root, f.file(BOOK));
    const p = await parts(f.file(BOOK));
    const xml = await p.sheet('new.ts');
    assertWellFormed(xml);
    expect(xml).toContain(`FF${THEMES.dark.tab.code}`);
    expect(xml).toContain('showGridLines="0"');
    expect(xml).toMatch(/<c r="A1"[^>]* s="\d+"/);
    expect(await p.sheet('UI')).toBe(uiBefore);
    expect(await f.sheet(BOOK, 'new.ts')).toEqual(['export {};']);
  });

  it('開かれているブックには当てない', async () => {
    const f = await project('dark');
    await saveConfig(f.root, { ...DEFAULT_CONFIG, bookTheme: 'light' });
    const { writeFile } = await import('node:fs/promises');
    await writeFile(f.file('app/~$app.xlcode.xlsx'), '');
    await expect(applyTheme(f.root, f.file(BOOK))).rejects.toThrow(/開かれています/);
  });

  it('列の指定が無い・sheetPr がある・枠線の指定がある、よくある形の XML も壊さずに当てる', () => {
    const style = { text: 7, fill: 8, tab: '569CD6', hideGrid: true };
    const excel =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetPr codeName="Sheet1"/><dimension ref="A1"/><sheetViews><sheetView showGridLines="1" tabSelected="1" workbookViewId="0"/></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="A1" s="3" t="s"><v>0</v></c><c r="B1"><v>1</v></c><c r="C1" s="2"/></row></sheetData>' +
      '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';
    const out = restyleSheetXml(excel, style);
    assertWellFormed(out);
    expect(out).toContain('<sheetPr codeName="Sheet1"><tabColor rgb="FF569CD6"/></sheetPr>');
    expect(out).toContain('<sheetView showGridLines="0" tabSelected="1" workbookViewId="0">'.replace('>', '/>'));
    expect(out).toContain(
      '<cols><col min="1" max="1" width="120" customWidth="1" style="7"/><col min="2" max="16384" width="9.140625" style="8"/></cols>',
    );
    expect(out.match(/<c r="[A-C]1"[^>]* s="7"/g)).toHaveLength(3);
    // 元の並び（r が先頭）を保つ
    expect(out).toContain('<c r="A1" s="7" t="s">');
    expect(out.indexOf('<cols>')).toBeLessThan(out.indexOf('<sheetData>'));
    // 既にテーマが当たっていても、同じ結果になる（何度でも当てられる）
    expect(restyleSheetXml(out, style)).toBe(out);
  });
});
