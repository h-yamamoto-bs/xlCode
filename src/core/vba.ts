import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, readdir, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { atomicWrite } from './atomic';
import { BOOK_SUFFIX, XLCODE_DIR } from './constants';
import { decodeFile } from './encoding';
import { excelPathError } from './fsutil';
import { checkBookOpen } from './lock';
import { normalizeText, textToLines } from './normalize';
import { build } from './ops';
import { bookRef, bookRootOf, openProject, type BookRef, type ProjectContext } from './project';
import type { Confirmation, OpResult } from './result';
import { newResult } from './result';
import { classifySheet } from './sheetName';
import { bookState, loadState, saveState } from './state';
import { isVbaIdentifier, parseFormSheet, type FormDef } from './vbaForm';
import { parseRefs, type VbaReference } from './vbaRefs';
import { Book } from './workbook';
import { patchXlsx, type PatchOp } from './xlsxPatch';

/**
 * VBA モード
 *
 *   <ブックの置き場所>/販売管理.xlcode.xlsx   編集用。画面・データのシートと、Module1.bas などのコードのシート
 *        │ Build ①（ソースコードモードと同じ）
 *        ▼
 *   <プロジェクト>/Module1.bas など            ソースコード（Git 管理。正本）
 *        │ Build ②（Windows のデスクトップ版 Excel で書き込む）
 *        ▼
 *   <プロジェクト>/販売管理.xlsm               ビルド結果（画面・データのシート + VBA。Git 管理しない）
 *
 * ①はソースコードモードの Build そのもの（衝突・省略検知・DEL_ も同じ）。Sync もそのまま使える。
 * ②はソースコードのファイルから VBA を組み立てる。ビルド結果は Build のたびに作り直す。
 */

export type VbaKind = 'standard' | 'class' | 'form' | 'workbook' | 'sheet';

export const VBA_KIND_LABEL: Record<VbaKind, string> = {
  standard: '標準モジュール',
  class: 'クラスモジュール',
  form: 'ユーザーフォーム',
  workbook: 'ブックのコード',
  sheet: 'シートのコード',
};

/** 参照設定のファイル（1 ブックに 1 つ） */
export const REFS_FILE = 'References.refs';

/** ビルド結果の .xlsm を Git 管理から外すための .gitignore の行 */
export const VBA_GITIGNORE = '*.xlsm';

export interface VbaModule {
  /** ファイル名（= シート名。Module1.bas） */
  file: string;
  /** モジュール名（Module1）。シートのコードの場合はシート名 */
  name: string;
  kind: VbaKind;
  /** シートのコードの書き込み先（画面のシートの名前） */
  targetSheet?: string;
  /** VBE に書き込むコード（フォームは配置を除いた部分） */
  code: string[];
  form?: FormDef;
}

export interface VbaModules {
  modules: VbaModule[];
  references: VbaReference[];
  errors: string[];
  warnings: string[];
}

/** VBA のファイルの拡張子 */
export function vbaExt(name: string): 'bas' | 'cls' | 'frm' | null {
  const m = /\.(bas|cls|frm)$/i.exec(name);
  return m ? (m[1].toLowerCase() as 'bas' | 'cls' | 'frm') : null;
}

export function isRefsFile(name: string): boolean {
  return name.toLowerCase() === REFS_FILE.toLowerCase();
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

/**
 * ファイル（またはシート）の内容から、VBE に書き込むモジュールを組み立てる。
 * @param entries .bas / .cls / .frm と References.refs の名前と行
 * @param uiSheets 画面・データのシート名（「シート名.cls」をシートのコードとして扱うため）
 */
export function modulesFrom(
  entries: readonly { name: string; lines: readonly string[] }[],
  uiSheets: readonly string[],
): VbaModules {
  const out: VbaModules = { modules: [], references: [], errors: [], warnings: [] };
  const seen = new Map<string, string>();
  for (const e of entries) {
    if (isRefsFile(e.name)) {
      const refs = parseRefs(e.name, e.lines);
      out.references.push(...refs.refs);
      out.errors.push(...refs.errors);
      continue;
    }
    const ext = vbaExt(e.name);
    if (!ext) continue;
    const base = e.name.slice(0, -4);
    let lines = [...e.lines];
    let kind: VbaKind;
    let targetSheet: string | undefined;
    let form: FormDef | undefined;
    if (ext === 'frm') {
      const parsed = parseFormSheet(e.name, lines);
      out.errors.push(...parsed.errors);
      kind = 'form';
      form = parsed.form ?? undefined;
      lines = parsed.code;
    } else if (ext === 'cls' && base.toLowerCase() === 'thisworkbook') {
      kind = 'workbook';
    } else if (ext === 'cls' && uiSheets.some((k) => k.toLowerCase() === base.toLowerCase())) {
      kind = 'sheet';
      targetSheet = uiSheets.find((k) => k.toLowerCase() === base.toLowerCase());
    } else {
      kind = ext === 'bas' ? 'standard' : 'class';
    }
    if ((kind === 'standard' || kind === 'class' || kind === 'form') && !isVbaIdentifier(base)) {
      out.errors.push(
        `「${e.name}」: モジュール名「${base}」は使えません（先頭は文字、以降は文字・数字・_ で 31 文字以内）` +
          (ext === 'cls' ? '。シートのコードを書く場合は「画面のシート名.cls」にしてください' : ''),
      );
    }
    const header = stripExportHeader(lines);
    if (header.stripped) {
      out.warnings.push(`「${e.name}」の先頭の Attribute 行などは不要なので除いて書き込みます`);
      lines = header.lines;
    }
    const key = kind === 'sheet' ? `sheet:${targetSheet!.toLowerCase()}` : base.toLowerCase();
    if (seen.has(key)) {
      out.errors.push(`「${seen.get(key)}」と「${e.name}」は同じモジュールになります。どちらかの名前を変えてください`);
      continue;
    }
    seen.set(key, e.name);
    out.modules.push({
      file: e.name,
      name: kind === 'sheet' ? targetSheet! : base,
      kind,
      targetSheet,
      code: lines,
      form,
    });
  }
  return out;
}

/** 画面・データのシート（ビルド結果に残すもの）と、それ以外（ビルド結果から除くもの） */
export function splitSheets(book: Book, extraCodeNames?: string[]): { keep: string[]; remove: string[] } {
  const keep: string[] = [];
  const remove: string[] = [];
  for (const name of book.sheetNames()) {
    if (classifySheet(name, extraCodeNames).kind === 'other') keep.push(name);
    else remove.push(name);
  }
  return { keep, remove };
}

function isVbaEntry(name: string): boolean {
  return vbaExt(name) !== null || isRefsFile(name);
}

/** 編集用ブックのシートから組み立てる（画面の問題表示用） */
export function modulesFromBook(book: Book, extraCodeNames?: string[]): VbaModules {
  const { keep, remove } = splitSheets(book, extraCodeNames);
  const entries = remove
    .filter((s) => classifySheet(s, extraCodeNames).kind === 'code' && isVbaEntry(s))
    .map((s) => ({
      name: s,
      lines: textToLines(normalizeText(book.readSheet(s).lines.join('\n'), s, { trimTrailingWhitespace: false })),
    }));
  return modulesFrom(entries, keep);
}

/** ソースコードのファイルから組み立てる（Build 用） */
export async function modulesFromFiles(dirAbs: string, uiSheets: readonly string[]): Promise<VbaModules> {
  const entries: { name: string; lines: string[] }[] = [];
  const errors: string[] = [];
  for (const name of (await readdir(dirAbs)).sort()) {
    if (!isVbaEntry(name) || name.startsWith('~$')) continue;
    const d = decodeFile(await readFile(path.join(dirAbs, name)), name);
    if (d.kind !== 'text') {
      errors.push(`「${name}」を読めません（${d.reason}）`);
      continue;
    }
    entries.push({ name, lines: textToLines(normalizeText(d.text, name, { trimTrailingWhitespace: false })) });
  }
  const m = modulesFrom(entries, uiSheets);
  m.errors.unshift(...errors);
  return m;
}

/** ビルド結果の .xlsm（<プロジェクト>/<ディレクトリ>/<名前>.xlsm） */
export function vbaOutputPath(ref: BookRef): string {
  const name = path.basename(ref.abs).slice(0, -BOOK_SUFFIX.length);
  return path.join(ref.dirAbs, `${name}.xlsm`);
}

// ---- Excel への書き込み（Windows の PowerShell + COM。src/main/excelCom.ts） ----

export interface VbaJobModule {
  file: string;
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
  /** ビルド結果の更新日時（ISO 8601。無ければ null） */
  builtAt: string | null;
  /** 上書き前のバックアップ（.xlcode/backup/ の中。新しい順、ルートからの相対パス） */
  backups: string[];
}

/** 上書き前のバックアップを置くフォルダ */
export function vbaBackupDir(root: string, ref: BookRef): string {
  return path.join(root, XLCODE_DIR, 'backup', ref.dirRel === '' ? '_root' : ref.dirRel.split('/').join('__'));
}

async function listBackups(root: string, ref: BookRef, output: string): Promise<string[]> {
  const dir = vbaBackupDir(root, ref);
  const base = path.basename(output, '.xlsm');
  const names = await readdir(dir).catch(() => [] as string[]);
  return names
    .filter((n) => n.startsWith(`${base}-`) && n.endsWith('.xlsm'))
    .sort()
    .reverse()
    .map((n) => path.relative(root, path.join(dir, n)).split(path.sep).join('/'));
}

/** ビルド結果の状態 */
export async function vbaOutputInfo(ctx: ProjectContext, ref: BookRef): Promise<VbaBookInfo> {
  const output = vbaOutputPath(ref);
  const hash = await outputHash(output);
  const prev = bookState(ctx.state, ref.rel).outputHash;
  const builtAt =
    hash === null
      ? null
      : await stat(output).then(
          (s) => s.mtime.toISOString(),
          () => null,
        );
  return {
    output,
    exists: hash !== null,
    changed: hash !== null && prev !== undefined && hash !== prev,
    builtAt,
    backups: await listBackups(ctx.root, ref, output),
  };
}

const BACKUP_KEEP = 5;

/** 上書き前のビルド結果を .xlcode/backup/ に残す（新しいものから 5 世代） */
async function backupOutput(root: string, ref: BookRef, output: string): Promise<string> {
  const dir = vbaBackupDir(root, ref);
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
 * ① シート → ソースコード（ソースコードモードの Build。確認・衝突もそのまま）
 * ② ソースコード → .xlsm（編集用ブックから拡張子付きのシートを除いたコピーを Excel で開き、VBA を書き込んで保存）
 * ②で失敗したときは、ビルド結果を変更しない（①のソースコードは出力済み）。
 */
export async function vbaBuild(
  root: string,
  bookAbs: string,
  opts: VbaBuildOptions,
  runner: VbaRunner | null,
): Promise<OpResult> {
  const ctx = await openProject(root);
  const ref = bookRef(root, bookAbs, bookRootOf(root, ctx.config));
  const fail = (msg: string): OpResult => {
    const r = newResult();
    r.errors.push(msg);
    return { ...r, status: 'error' };
  };
  if (!runner)
    return fail('VBA の書き込みには Windows とデスクトップ版 Excel が必要です（この環境では Build できません）');

  const output = vbaOutputPath(ref);
  const longPath = excelPathError(output);
  if (longPath) return fail(longPath);
  const outOpen = await checkBookOpen(output);
  if (outOpen.open) {
    return fail(
      `ビルド結果 ${path.basename(output)} が開かれています。閉じてから実行してください（${outOpen.reason}）`,
    );
  }
  const extra: Confirmation[] = [];
  const current = await outputHash(output);
  const prevHash = bookState(ctx.state, ref.rel).outputHash;
  if (current !== null && current !== prevHash) {
    extra.push({
      kind: 'overwrite',
      message:
        prevHash === undefined
          ? `既にある ${path.basename(output)} を上書きします（元のファイルは .xlcode/backup に残します）`
          : `${path.basename(output)} が前回の Build の後に変更されています。上書きすると、直接入力したデータや VBE で直したコードは失われます（元のファイルは .xlcode/backup に残します）`,
      files: [path.relative(root, output).split(path.sep).join('/')],
    });
  }

  // ① シート → ソースコード
  // ビルド結果も「元に戻す」の控えに含める（②で作り直すため）
  const r = await build(root, bookAbs, {
    confirmed: opts.confirmed,
    extraConfirmations: extra,
    snapshotOutput: output,
  });
  if (r.status !== 'ok') return r;

  // ② ソースコード → .xlsm
  const stop = (msg: string, detail?: string): OpResult => {
    r.errors.push(msg);
    if (detail) r.errors.push(detail);
    r.errors.push('ソースコードへの出力は完了しています。ビルド結果（.xlsm）は変更していません');
    return { ...r, status: 'error' };
  };
  const bytes = await readFile(ref.abs);
  const book = await Book.fromBuffer(bytes, ref.abs);
  const { keep, remove } = splitSheets(book, ctx.config.extraCodeNames);
  const mods = await modulesFromFiles(ref.dirAbs, keep);
  r.warnings.push(...mods.warnings);
  if (mods.errors.length > 0) return stop(`VBA を組み立てられません:\n${mods.errors.join('\n')}`);

  const work = await mkdtemp(path.join(tmpdir(), 'xlcode-build-'));
  try {
    // 画面のシートが 1 枚も無いと空のブックになれないため、白紙のシートを 1 枚置く
    const ops: PatchOp[] = [];
    if (keep.length === 0) ops.push({ kind: 'write', name: 'Sheet1', columns: [] });
    for (const s of remove) ops.push({ kind: 'delete', name: s });
    const input = path.join(work, 'input.xlsx');
    await writeFile(input, await patchXlsx(bytes, ops));
    const built = path.join(work, 'output.xlsm');
    const run = await runner({
      input,
      output: built,
      modules: mods.modules.map((m) => ({
        file: m.file,
        name: m.name,
        kind: m.kind,
        targetSheet: m.targetSheet,
        code: m.code.join('\r\n'),
        form: m.form,
      })),
      references: mods.references,
    });
    if (!run.ok) {
      return stop(
        `Excel での書き込みに失敗しました: ${run.error ?? '原因不明'}`,
        run.errorKind === 'vbom' ? VBOM_HELP : undefined,
      );
    }
    const result = await readFile(built);
    const again = await checkBookOpen(output);
    if (again.open)
      return stop(`ビルド結果 ${path.basename(output)} が途中で開かれたため中断しました（${again.reason}）`);
    if (current !== null) r.changes.push({ action: 'backup', target: await backupOutput(root, ref, output) });
    await atomicWrite(output, result);
    r.changes.push({ action: 'write-file', target: path.basename(output) });
    for (const m of mods.modules)
      r.changes.push({ action: 'write-module', target: `${m.file}（${VBA_KIND_LABEL[m.kind]}）` });

    // ①が保存した状態に、ビルド結果のハッシュを足す
    const state = await loadState(root);
    bookState(state, ref.rel).outputHash = fileHash(result);
    await saveState(root, state);
    return r;
  } finally {
    await rm(work, { recursive: true, force: true });
  }
}
