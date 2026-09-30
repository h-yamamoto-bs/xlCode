import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, readdir, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { atomicWrite } from './atomic';
import { AGENTS_SHEET, BOOK_SUFFIX, LOCAL_AGENTS_SHEET, REFS_SHEET, XLCODE_DIR } from './constants';
import { decodeFile, EncodeError, encodeFile } from './encoding';
import { excelPathError } from './fsutil';
import { autoCommit, isGitRepo, uncommittedChanges } from './git';
import { checkBookOpen } from './lock';
import { linesToText, normalizeText, sha256, textToLines } from './normalize';
import { bookRef, bookRootOf, exists, openProject, type BookRef, type ProjectContext } from './project';
import { newResult, type OpResult } from './result';
import { classifySheet } from './sheetName';
import { bookState, saveState, type FileState } from './state';
import { isVbaIdentifier, parseFormSheet, type FormDef } from './vbaForm';
import { parseRefs, renderRefs, type VbaReference } from './vbaRefs';
import { Book } from './workbook';
import { patchXlsx, type PatchOp } from './xlsxPatch';

/**
 * VBA モード
 *
 *   <ブックの置き場所>/販売管理.xlcode.xlsx   編集用（UI・データのシートと、.bas / .cls / .frm のシート）
 *        │ Build
 *        ▼
 *   <プロジェクト>/販売管理.xlsm               ビルド結果（UI・データのシート + VBA。拡張子付きのシートは除く）
 *   <プロジェクト>/vba/Module1.bas など         Git 用の控え（xlCode は読み込まない）
 *
 * 正は編集用ブック。ビルド結果は Build のたびに作り直す。
 */

export type VbaKind = 'standard' | 'class' | 'form' | 'workbook' | 'sheet';

export const VBA_KIND_LABEL: Record<VbaKind, string> = {
  standard: '標準モジュール',
  class: 'クラスモジュール',
  form: 'ユーザーフォーム',
  workbook: 'ブックのコード',
  sheet: 'シートのコード',
};

export interface VbaModule {
  /** シート名（Module1.bas） */
  sheet: string;
  /** モジュール名（Module1）。シートのコードの場合はシート名 */
  name: string;
  kind: VbaKind;
  /** シートのコードの書き込み先（UI シートの名前） */
  targetSheet?: string;
  /** 正規化したシートの内容（LF・末尾改行1つ）。比較と控えに使う */
  text: string;
  hash: string;
  lines: number;
  chars: number;
  /** VBE に書き込むコード（フォームは配置を除いた部分） */
  code: string[];
  form?: FormDef;
}

export interface VbaCollect {
  modules: VbaModule[];
  /** ビルド結果に残すシート（UI・データ） */
  keep: string[];
  /** ビルド結果から除くシート */
  remove: string[];
  /** #refs シートの参照設定 */
  references: VbaReference[];
  errors: string[];
  warnings: string[];
}

/** VBA のシートの拡張子 */
export function vbaExt(sheet: string): 'bas' | 'cls' | 'frm' | null {
  const m = /\.(bas|cls|frm)$/i.exec(sheet);
  return m ? (m[1].toLowerCase() as 'bas' | 'cls' | 'frm') : null;
}

/** VBE の「エクスポート」で付く行（VERSION 1.0 CLASS 〜 END、Attribute VB_…）を先頭から除く */
export function stripExportHeader(lines: readonly string[]): { lines: string[]; stripped: boolean } {
  let i = 0;
  if (/^VERSION\s+[\d.]+\s+CLASS\s*$/i.test(lines[0] ?? '')) {
    const end = lines.findIndex((l, j) => j > 0 && /^END\s*$/i.test(l.trim()));
    i = end < 0 ? 1 : end + 1;
  }
  while (i < lines.length && /^Attribute\s+VB_\w+\s*=/i.test(lines[i])) i++;
  return { lines: lines.slice(i), stripped: i > 0 };
}

/** 編集用ブックのシートを、ビルド結果に残すもの・VBA として書き込むもの・除くものに分ける */
export function collectModules(
  book: Book,
  opts: { trimTrailingWhitespace: boolean; extraCodeNames?: string[] },
): VbaCollect {
  const out: VbaCollect = { modules: [], keep: [], remove: [], references: [], errors: [], warnings: [] };
  const names = book.sheetNames();
  for (const name of names) {
    if (classifySheet(name, opts.extraCodeNames).kind === 'other') out.keep.push(name);
    else out.remove.push(name);
  }
  if (book.hasSheet(REFS_SHEET)) {
    const refs = parseRefs(REFS_SHEET, book.readSheet(REFS_SHEET).lines);
    out.references = refs.refs;
    out.errors.push(...refs.errors);
  }
  const seen = new Map<string, string>();
  for (const sheet of out.remove) {
    const info = classifySheet(sheet, opts.extraCodeNames);
    if (info.kind === 'delete') {
      out.warnings.push(
        `「${sheet}」: VBA モードでは DEL_ は不要です。シートを削除すれば、次の Build でモジュールも消えます`,
      );
      continue;
    }
    if (info.kind !== 'code') continue;
    const ext = vbaExt(sheet);
    if (!ext) {
      if (sheet !== AGENTS_SHEET && sheet !== LOCAL_AGENTS_SHEET) {
        out.warnings.push(
          `「${sheet}」は .bas / .cls / .frm ではないため VBA には書き込みません（ビルド結果からは除きます）`,
        );
      }
      continue;
    }
    const data = book.readSheet(sheet);
    for (const is of data.issues) {
      out.errors.push(
        `「${sheet}」${is.row} 行目が文字列ではありません（${is.type}）。Excel の自動変換の可能性があります`,
      );
    }
    if (data.hasExtraColumns) out.warnings.push(`「${sheet}」の B 列以降の内容は無視します`);

    const text = normalizeText(data.lines.join('\n'), sheet, { trimTrailingWhitespace: opts.trimTrailingWhitespace });
    const base = sheet.slice(0, -4);
    let lines = textToLines(text);
    let kind: VbaKind;
    let targetSheet: string | undefined;
    let form: FormDef | undefined;
    if (ext === 'frm') {
      const parsed = parseFormSheet(sheet, lines);
      out.errors.push(...parsed.errors);
      kind = 'form';
      form = parsed.form ?? undefined;
      lines = parsed.code;
    } else if (ext === 'cls' && base.toLowerCase() === 'thisworkbook') {
      kind = 'workbook';
    } else if (ext === 'cls' && out.keep.some((k) => k.toLowerCase() === base.toLowerCase())) {
      kind = 'sheet';
      targetSheet = out.keep.find((k) => k.toLowerCase() === base.toLowerCase());
    } else {
      kind = ext === 'bas' ? 'standard' : 'class';
    }
    if ((kind === 'standard' || kind === 'class' || kind === 'form') && !isVbaIdentifier(base)) {
      out.errors.push(
        `「${sheet}」: モジュール名「${base}」は使えません（先頭は文字、以降は文字・数字・_ で 31 文字以内）` +
          (ext === 'cls' ? '。シートのコードを書く場合は、シート名を「UI シートの名前.cls」にしてください' : ''),
      );
    }
    const header = stripExportHeader(lines);
    if (header.stripped) {
      out.warnings.push(`「${sheet}」の先頭の Attribute 行などは不要なので除いて書き込みます`);
      lines = header.lines;
    }
    const key = kind === 'sheet' ? `sheet:${targetSheet!.toLowerCase()}` : base.toLowerCase();
    if (seen.has(key)) {
      out.errors.push(`「${seen.get(key)}」と「${sheet}」は同じモジュールになります。どちらかの名前を変えてください`);
      continue;
    }
    seen.set(key, sheet);
    out.modules.push({
      sheet,
      name: kind === 'sheet' ? targetSheet! : base,
      kind,
      targetSheet,
      text,
      hash: sha256(text),
      lines: textToLines(text).length,
      chars: text.length,
      code: lines,
      form,
    });
  }
  return out;
}

/** ビルド結果の .xlsm（<プロジェクト>/<ディレクトリ>/<名前>.xlsm） */
export function vbaOutputPath(ref: BookRef): string {
  const name = path.basename(ref.abs).slice(0, -BOOK_SUFFIX.length);
  return path.join(ref.dirAbs, `${name}.xlsm`);
}

export const VBA_COPY_DIR = 'vba';

/** Git 用の控えのファイル名 */
export function copyFileName(m: VbaModule): string {
  return m.kind === 'form' ? `${m.sheet}.txt` : m.sheet;
}

/** 参照設定の控え */
export const REFS_COPY = 'references.txt';

function isCopyFile(name: string): boolean {
  return /\.(bas|cls|frm\.txt)$/i.test(name) || name === REFS_COPY;
}

/** 控えの内容。VBE の「ファイルのインポート」で取り込めるよう、標準・クラスモジュールには見出しを付ける */
export function copyText(m: VbaModule): string {
  const body = m.code.join('\n');
  const tail = body === '' ? '' : `${body}\n`;
  if (m.kind === 'standard') return `Attribute VB_Name = "${m.name}"\n${tail}`;
  if (m.kind === 'class') {
    return (
      [
        'VERSION 1.0 CLASS',
        'BEGIN',
        "  MultiUse = -1  'True",
        'END',
        `Attribute VB_Name = "${m.name}"`,
        'Attribute VB_GlobalNameSpace = False',
        'Attribute VB_Creatable = False',
        'Attribute VB_PredeclaredId = False',
        'Attribute VB_Exposed = False',
        '',
      ].join('\n') + tail
    );
  }
  // フォームは配置も含めてシートの内容のまま。シート・ブックのコードはコードだけ
  return m.kind === 'form' ? m.text : tail;
}

/** 控え（vba/ の中）からシートの内容に戻す（ブックを作り直すとき用） */
export function sheetFromCopy(fileName: string, text: string): { sheet: string; lines: string[] } | null {
  const lines = textToLines(normalizeText(text, fileName, { trimTrailingWhitespace: false }));
  if (fileName === REFS_COPY) return { sheet: REFS_SHEET, lines };
  if (/\.frm\.txt$/i.test(fileName)) return { sheet: fileName.slice(0, -4), lines };
  if (/\.(bas|cls)$/i.test(fileName)) return { sheet: fileName, lines: stripExportHeader(lines).lines };
  return null;
}

// ---- Excel への書き込み（Windows の PowerShell + COM。src/main/excelCom.ts） ----

export interface VbaJobModule {
  sheet: string;
  name: string;
  kind: VbaKind;
  targetSheet?: string;
  /** CRLF 区切りのコード */
  code: string;
  form?: FormDef;
}

export interface VbaJob {
  /** 拡張子付きのシートを除いた .xlsx */
  input: string;
  /** 保存先の .xlsm */
  output: string;
  modules: VbaJobModule[];
  /** 追加する参照設定 */
  references: VbaReference[];
}

export interface VbaRunResult {
  ok: boolean;
  error?: string;
  /** no-excel: Excel が無い / vbom: VBA プロジェクトへのアクセスが信頼されていない */
  errorKind?: 'no-excel' | 'vbom' | 'timeout' | 'other';
}

export type VbaRunner = (job: VbaJob) => Promise<VbaRunResult>;

export const VBOM_HELP =
  'Excel の「ファイル → オプション → トラスト センター → トラスト センターの設定 → マクロの設定」で' +
  '「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」にチェックを付けてください';

function fileHash(buf: Uint8Array): string {
  return createHash('sha256').update(buf).digest('hex');
}

export async function outputHash(file: string): Promise<string | null> {
  try {
    return fileHash(await readFile(file));
  } catch {
    return null;
  }
}

export interface VbaBookInfo {
  /** ビルド結果の絶対パス */
  output: string;
  exists: boolean;
  /** 前回の Build の後に、ビルド結果が直接変更された */
  changed: boolean;
}

/** ビルド結果の状態 */
export async function vbaOutputInfo(ctx: ProjectContext, ref: BookRef): Promise<VbaBookInfo> {
  const output = vbaOutputPath(ref);
  const hash = await outputHash(output);
  const prev = bookState(ctx.state, ref.rel).outputHash;
  return { output, exists: hash !== null, changed: hash !== null && prev !== undefined && hash !== prev };
}

/** 前回の Build からシートの内容が減りすぎていないか（Copilot の省略対策） */
function shrinkMessage(ctx: ProjectContext, prev: FileState | undefined, m: VbaModule): string | null {
  const { shrinkThreshold: th, shrinkMinLines } = ctx.config;
  if (!prev || prev.lines < shrinkMinLines) return null;
  const lineDrop = 1 - m.lines / prev.lines;
  const charDrop = prev.chars === 0 ? 0 : 1 - m.chars / prev.chars;
  if (lineDrop < th && charDrop < th) return null;
  const pct = (v: number) => `${Math.round(v * 100)}%`;
  return `${m.sheet}: ${prev.lines}→${m.lines} 行（-${pct(lineDrop)}）、${prev.chars}→${m.chars} 文字（-${pct(charDrop)}）`;
}

const BACKUP_KEEP = 5;

/** 上書き前のビルド結果を .xlcode/backup/ に残す（新しいものから 5 世代） */
async function backupOutput(root: string, ref: BookRef, output: string): Promise<string> {
  const dir = path.join(root, XLCODE_DIR, 'backup', ref.dirRel === '' ? '_root' : ref.dirRel.split('/').join('__'));
  await mkdir(dir, { recursive: true });
  const base = path.basename(output, '.xlsm');
  const stamp = new Date().toISOString().replace(/[-:]/g, '').replace('T', '-').slice(0, 15);
  const dest = path.join(dir, `${base}-${stamp}.xlsm`);
  await writeFile(dest, await readFile(output));
  const olds = (await readdir(dir)).filter((n) => n.startsWith(`${base}-`) && n.endsWith('.xlsm')).sort();
  for (const n of olds.slice(0, Math.max(0, olds.length - BACKUP_KEEP))) await rm(path.join(dir, n), { force: true });
  return path.relative(root, dest).split(path.sep).join('/');
}

export interface VbaBuildOptions {
  confirmed?: boolean;
}

/**
 * VBA モードの Build。
 * 1. 編集用ブックから拡張子付きのシートを除いたコピーを作る（UI シートの図形・ボタンはそのまま）
 * 2. Excel で開き、.bas / .cls / .frm のシートを VBA に書き込んで .xlsm として保存する
 * 3. できた .xlsm でビルド結果を置き換え、Git 用の控えを書き出す
 * 失敗したときは、ビルド結果・控えのどちらも変更しない。
 */
export async function vbaBuild(
  root: string,
  bookAbs: string,
  opts: VbaBuildOptions,
  runner: VbaRunner | null,
): Promise<OpResult> {
  const ctx = await openProject(root);
  const ref = bookRef(root, bookAbs, bookRootOf(root, ctx.config));
  const r = newResult();
  const fail = (msg: string): OpResult => {
    r.errors.push(msg);
    return { ...r, status: 'error' };
  };

  const open = await checkBookOpen(ref.abs);
  if (open.open)
    return fail(`${ref.rel} が開かれています。保存して Excel を閉じてから実行してください（${open.reason}）`);
  const bytes = await readFile(ref.abs);
  const book = await Book.fromBuffer(bytes, ref.abs);
  const c = collectModules(book, ctx.config);
  r.warnings.push(...c.warnings);
  if (c.errors.length > 0) {
    r.errors.push(...c.errors);
    return { ...r, status: 'error' };
  }
  if (c.modules.length === 0 && c.keep.length === 0) return fail('ブックに UI のシートも VBA のシートもありません');

  const output = vbaOutputPath(ref);
  const longPath = excelPathError(output);
  if (longPath) return fail(longPath);
  const outOpen = await checkBookOpen(output);
  if (outOpen.open)
    return fail(
      `ビルド結果 ${path.basename(output)} が開かれています。閉じてから実行してください（${outOpen.reason}）`,
    );

  // 控えの内容を先に作る（Shift_JIS で表せない文字は VBE でも化けるため、ここで止める）
  const copies = new Map<string, Buffer>();
  for (const m of c.modules) {
    try {
      copies.set(copyFileName(m), encodeFile(copyText(m), { encoding: 'sjis', eol: 'crlf' }, m.sheet));
    } catch (e) {
      if (!(e instanceof EncodeError)) throw e;
      r.errors.push(`${e.message}（VBA は Shift_JIS のため、この文字は使えません）`);
    }
  }
  if (c.references.length > 0) {
    copies.set(
      REFS_COPY,
      encodeFile(linesToText(renderRefs(c.references)), { encoding: 'sjis', eol: 'crlf' }, REFS_COPY),
    );
  }
  if (r.errors.length > 0) return { ...r, status: 'error' };

  const bs = bookState(ctx.state, ref.rel);
  const shrinks = c.modules.map((m) => shrinkMessage(ctx, bs.files[m.sheet], m)).filter((s): s is string => s !== null);
  if (shrinks.length > 0) {
    r.confirmations.push({
      kind: 'shrink',
      message: 'コードが大きく減っています。Copilot が省略した可能性があります',
      files: shrinks,
    });
  }
  const removed = Object.keys(bs.files).filter((s) => !c.modules.some((m) => m.sheet === s));
  if (removed.length > 0) {
    r.confirmations.push({
      kind: 'delete',
      message: '次のシートが無くなっています。ビルド結果からこのモジュールを削除します',
      files: removed,
    });
  }
  const current = await outputHash(output);
  if (current !== null && current !== bs.outputHash) {
    r.confirmations.push({
      kind: 'overwrite',
      message:
        bs.outputHash === undefined
          ? `既にある ${path.basename(output)} を上書きします（元のファイルは .xlcode/backup に残します）`
          : `${path.basename(output)} が前回の Build の後に変更されています。上書きすると、直接入力したデータや VBE で直したコードは失われます（元のファイルは .xlcode/backup に残します）`,
      files: [path.relative(root, output).split(path.sep).join('/')],
    });
  }
  const useGit = await gitConfirmations(ctx, ref, r);
  if (r.confirmations.length > 0 && !opts.confirmed) return { ...r, status: 'confirm' };

  if (!runner) {
    return fail('VBA の書き込みには Windows とデスクトップ版 Excel が必要です（この環境では Build できません）');
  }
  if (useGit) {
    try {
      if (await autoCommit(ctx.root, ref.dirRel, `xlcode: Build 前の自動コミット (${ref.rel})`, [VBA_COPY_DIR])) {
        r.changes.push({ action: 'commit', target: ref.dirRel || '.' });
      }
    } catch (e) {
      return fail(`自動コミットに失敗しました: ${e instanceof Error ? e.message : String(e)}`);
    }
  }

  const work = await mkdtemp(path.join(tmpdir(), 'xlcode-build-'));
  try {
    // UI シートが 1 枚も無いと空のブックになれないため、白紙のシートを 1 枚置く
    const ops: PatchOp[] = [];
    if (c.keep.length === 0) ops.push({ kind: 'write', name: 'Sheet1', columns: [] });
    for (const s of c.remove) ops.push({ kind: 'delete', name: s });
    const input = path.join(work, 'input.xlsx');
    await writeFile(input, await patchXlsx(bytes, ops));
    const built = path.join(work, 'output.xlsm');
    const run = await runner({
      input,
      output: built,
      modules: c.modules.map((m) => ({
        sheet: m.sheet,
        name: m.name,
        kind: m.kind,
        targetSheet: m.targetSheet,
        code: m.code.join('\r\n'),
        form: m.form,
      })),
      references: c.references,
    });
    if (!run.ok) {
      r.errors.push(`Excel での書き込みに失敗しました: ${run.error ?? '原因不明'}`);
      if (run.errorKind === 'vbom') r.errors.push(VBOM_HELP);
      r.errors.push('ビルド結果と控えは変更していません');
      return { ...r, status: 'error' };
    }
    const result = await readFile(built);

    await mkdir(ref.dirAbs, { recursive: true });
    const again = await checkBookOpen(output);
    if (again.open)
      return fail(`ビルド結果 ${path.basename(output)} が途中で開かれたため中断しました（${again.reason}）`);
    if (current !== null) r.changes.push({ action: 'backup', target: await backupOutput(root, ref, output) });
    await atomicWrite(output, result);
    r.changes.push({ action: 'write-file', target: path.basename(output) });
    for (const m of c.modules) {
      const prev = bs.files[m.sheet];
      if (!prev || prev.hash !== m.hash)
        r.changes.push({ action: 'write-module', target: `${m.sheet}（${VBA_KIND_LABEL[m.kind]}）` });
    }
    for (const s of removed) r.changes.push({ action: 'delete-module', target: s });

    // 控え
    const copyDir = path.join(ref.dirAbs, VBA_COPY_DIR);
    await mkdir(copyDir, { recursive: true });
    for (const [name, buf] of copies) {
      const file = path.join(copyDir, name);
      const old = await readFile(file).catch(() => null);
      if (old && old.equals(buf)) continue;
      await atomicWrite(file, buf);
      r.changes.push({ action: 'write-file', target: `${VBA_COPY_DIR}/${name}` });
    }
    for (const name of await readdir(copyDir)) {
      if (isCopyFile(name) && !copies.has(name)) {
        await rm(path.join(copyDir, name), { force: true });
        r.changes.push({ action: 'delete-file', target: `${VBA_COPY_DIR}/${name}` });
      }
    }

    bs.files = Object.fromEntries(c.modules.map((m) => [m.sheet, { hash: m.hash, lines: m.lines, chars: m.chars }]));
    bs.outputHash = fileHash(result);
    bs.lastBuildAt = new Date().toISOString();
    await saveState(root, ctx.state);
    return r;
  } finally {
    await rm(work, { recursive: true, force: true });
  }
}

async function gitConfirmations(ctx: ProjectContext, ref: BookRef, r: OpResult): Promise<boolean> {
  if (!ctx.config.autoCommit) return false;
  if (!(await isGitRepo(ctx.root))) {
    r.confirmations.push({
      kind: 'no-git',
      message: 'Git リポジトリではないため、実行前のバックアップが作られません',
      files: [],
    });
    return false;
  }
  const changes = await uncommittedChanges(ctx.root, ref.dirRel, [VBA_COPY_DIR]);
  if (changes.length > 0) {
    r.confirmations.push({
      kind: 'uncommitted',
      message: '未コミットの変更があります。実行前に自動コミットします',
      files: changes,
    });
  }
  return true;
}

/** VBA モードの画面表示用の状態 */
export async function vbaFileStatus(
  ctx: ProjectContext,
  ref: BookRef,
  book: Book,
): Promise<{
  collect: VbaCollect;
  files: { name: string; status: 'clean' | 'excel-changed' | 'excel-new' | 'removed'; format: string }[];
}> {
  const collect = collectModules(book, ctx.config);
  const prev = bookState(ctx.state, ref.rel).files;
  const files: { name: string; status: 'clean' | 'excel-changed' | 'excel-new' | 'removed'; format: string }[] =
    collect.modules.map((m) => ({
      name: m.sheet,
      status: !prev[m.sheet]
        ? ('excel-new' as const)
        : prev[m.sheet].hash === m.hash
          ? ('clean' as const)
          : ('excel-changed' as const),
      format: VBA_KIND_LABEL[m.kind],
    }));
  for (const s of Object.keys(prev)) {
    if (!collect.modules.some((m) => m.sheet === s)) files.push({ name: s, status: 'removed', format: '' });
  }
  return { collect, files };
}

/** 控え（vba/）から、ブックに入れるシートを読み出す */
export async function readCopies(dirAbs: string): Promise<{ sheet: string; lines: string[] }[]> {
  const dir = path.join(dirAbs, VBA_COPY_DIR);
  if (!(await exists(dir))) return [];
  const out: { sheet: string; lines: string[] }[] = [];
  for (const name of (await readdir(dir)).sort()) {
    if (!isCopyFile(name)) continue;
    const d = decodeFile(await readFile(path.join(dir, name)), name);
    if (d.kind !== 'text') continue;
    const s = sheetFromCopy(name, d.text);
    if (s) out.push(s);
  }
  return out;
}
