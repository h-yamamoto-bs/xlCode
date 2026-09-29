import { rm } from 'node:fs/promises';
import { atomicWrite } from './atomic';
import { defaultFormat, EncodeError, encodeFile } from './encoding';
import type { Canon } from './canonical';
import { CONFLICT_PREFIX } from './constants';
import { autoCommit, isGitRepo, uncommittedChanges } from './git';
import { checkBookOpen } from './lock';
import { textToLines } from './normalize';
import { bookRef, openProject, type BookRef, type ProjectContext } from './project';
import { newResult, type OpResult } from './result';
import { scanBook, type FileEntry, type Scan } from './scan';
import { bookState, saveState, type BookState } from './state';

export interface BuildOptions {
  /** 確認ダイアログでユーザーが続行を選んだ */
  confirmed?: boolean;
}

export interface SyncOptions {
  confirmed?: boolean;
  /** 5.1 (A): 未 Build のシート変更を破棄して Sync する */
  discardExcelChanges?: boolean;
}

function toState(c: Canon) {
  return { hash: c.hash, lines: c.lines, chars: c.chars };
}

function sameLines(a: readonly string[] | undefined, b: readonly string[]): boolean {
  if (!a) return false;
  // 末尾の空セルは比較対象外（正規化で落ちるため）
  let n = a.length;
  while (n > 0 && a[n - 1] === '') n--;
  return n === b.length && b.every((l, i) => l === a[i]);
}

/** 省略検知（4.6）。前回値から閾値以上減っていればメッセージを返す */
function shrinkMessage(ctx: ProjectContext, e: FileEntry, next: Canon): string | null {
  const prev = e.prev;
  const { shrinkThreshold: th, shrinkMinLines } = ctx.config;
  if (!prev || prev.lines < shrinkMinLines) return null;
  const lineDrop = 1 - next.lines / prev.lines;
  const charDrop = prev.chars === 0 ? 0 : 1 - next.chars / prev.chars;
  if (lineDrop < th && charDrop < th) return null;
  const pct = (v: number) => `${Math.round(v * 100)}%`;
  return `${e.name}: ${prev.lines}→${next.lines} 行（-${pct(lineDrop)}）、${prev.chars}→${next.chars} 文字（-${pct(charDrop)}）`;
}

/** 書き込み直前の再確認（確認〜保存の間に Excel で開かれた場合に備える） */
async function assertClosed(ref: BookRef, r: OpResult, detail: string): Promise<boolean> {
  const open = await checkBookOpen(ref.abs);
  if (!open.open) return true;
  r.errors.push(`${ref.rel} が途中で開かれたため中断しました。${detail}（${open.reason}）`);
  r.status = 'error';
  return false;
}

async function preflight(ctx: ProjectContext, ref: BookRef, r: OpResult): Promise<Scan | null> {
  const open = await checkBookOpen(ref.abs);
  if (open.open) {
    r.errors.push(`${ref.rel} が開かれています。Excel を閉じてから実行してください（${open.reason}）`);
    return null;
  }
  const scan = await scanBook(ctx, ref);
  r.warnings.push(...scan.warnings);
  if (scan.errors.length > 0) {
    r.errors.push(...scan.errors);
    return null;
  }
  return scan;
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
  const changes = await uncommittedChanges(ctx.root, ref.dirRel);
  if (changes.length > 0) {
    r.confirmations.push({
      kind: 'uncommitted',
      message: '未コミットの変更があります。実行前に自動コミットします',
      files: changes,
    });
  }
  return true;
}

async function commitBefore(ctx: ProjectContext, ref: BookRef, label: string, r: OpResult): Promise<boolean> {
  try {
    if (await autoCommit(ctx.root, ref.dirRel, `xlcode: ${label} 前の自動コミット (${ref.rel})`)) {
      r.changes.push({ action: 'commit', target: ref.dirRel || '.' });
    }
    return true;
  } catch (e) {
    r.errors.push(`自動コミットに失敗しました: ${e instanceof Error ? e.message : String(e)}`);
    r.status = 'error';
    return false;
  }
}

/** 衝突シートを作成する（5.6）。同じファイルの衝突シートが既にあれば上書きする */
function writeConflictSheet(scan: Scan, e: FileEntry): string {
  const { book, ref } = scan;
  const fileRel = ref.dirRel ? `${ref.dirRel}/${e.name}` : e.name;
  const existing = book
    .sheetNames()
    .filter((n) => n.startsWith(CONFLICT_PREFIX))
    .find((n) => book.readColumn(n, 1, 1)[0] === fileRel);
  let sheet = existing;
  if (!sheet) {
    for (let i = 1; ; i++) {
      const cand = `${CONFLICT_PREFIX}${String(i).padStart(2, '0')}`;
      if (!book.hasSheet(cand)) {
        sheet = cand;
        break;
      }
    }
    book.writeLines(sheet, []);
  }
  const xl = e.xl ? textToLines(e.xl.text) : [];
  const src = e.src ? textToLines(e.src.text) : [];
  book.writeLines(
    sheet,
    [fileRel, 'Excel側', ...xl],
    [
      [
        `A列とB列を統合して「${e.name}」シートに書き、このシートを削除してください`,
        e.src ? 'ソース側' : 'ソース側（削除済み）',
        ...src,
      ],
    ],
  );
  return sheet;
}

/** 衝突シートを作り、前回値をソース側に合わせる（統合後の Build で Excel 側が採用されるように） */
function recordConflicts(scan: Scan, bs: BookState, entries: FileEntry[], r: OpResult): void {
  for (const e of entries) {
    const sheet = writeConflictSheet(scan, e);
    r.conflicts.push({ sheet, file: e.name });
    r.changes.push({ action: 'conflict-sheet', target: sheet });
    if (e.src) bs.files[e.name] = toState(e.src);
    else delete bs.files[e.name];
  }
}

/** ソース側の内容をシートへ反映する */
function pullToSheet(scan: Scan, bs: BookState, e: FileEntry, r: OpResult): void {
  if (e.src) {
    scan.book.writeLines(e.sheetName ?? e.name, textToLines(e.src.text));
    bs.files[e.name] = toState(e.src);
    r.changes.push({ action: 'write-sheet', target: e.name });
    if (e.status === 'sheet-missing')
      r.warnings.push(
        `「${e.name}」シートを復元しました。ファイルを削除するにはシート名を DEL_${e.name} にして Build してください`,
      );
  } else {
    if (e.sheetName) scan.book.deleteSheet(e.sheetName);
    delete bs.files[e.name];
    r.changes.push({ action: 'delete-sheet', target: e.name });
  }
}

/** 両側一致のファイル。シートが未整形なら整形結果で書き直す */
function settleClean(scan: Scan, bs: BookState, e: FileEntry, r: OpResult): boolean {
  bs.files[e.name] = toState(e.src!);
  const lines = textToLines(e.src!.text);
  if (sameLines(e.xlRaw, lines)) return false;
  scan.book.writeLines(e.sheetName!, lines);
  r.changes.push({ action: 'reformat-sheet', target: e.name });
  return true;
}

/**
 * Build（Excel → ソースコード）4章
 * 実行直前に強制 Sync を兼ねる: ソース側だけの変更はシートへ取り込み、両側変更は衝突として扱う。
 */
export async function build(root: string, bookAbs: string, opts: BuildOptions = {}): Promise<OpResult> {
  const ctx = await openProject(root);
  const ref = bookRef(root, bookAbs);
  const r = newResult();
  const scan = await preflight(ctx, ref, r);
  if (!scan) return { ...r, status: 'error' };
  if (scan.conflictSheets.length > 0) {
    r.errors.push(`未解決の衝突シートがあります: ${scan.conflictSheets.join(', ')}`);
    return { ...r, status: 'error' };
  }
  const bs = bookState(ctx.state, ref.rel);

  const conflicts = scan.entries.filter((e) => e.status === 'conflict');
  if (conflicts.length > 0) {
    recordConflicts(scan, bs, conflicts, r);
    if (!(await assertClosed(ref, r, '何も変更していません'))) return r;
    await scan.book.save(ref.abs);
    await saveState(root, ctx.state);
    r.errors.push(
      'ソース側と Excel 側の両方で変更されたファイルがあります。衝突シートを統合してから再度 Build してください',
    );
    return { ...r, status: 'conflict' };
  }

  const toBuild = scan.entries.filter((e) => e.status === 'excel-changed' || e.status === 'excel-new');
  // 書き出す内容を先に作る（Shift_JIS で表せない文字などは、何も書かずに中断する）
  const encoded = new Map<string, Buffer>();
  for (const e of toBuild) {
    try {
      encoded.set(e.name, encodeFile(e.xl!.text, e.format ?? defaultFormat(e.name), e.name));
    } catch (err) {
      if (!(err instanceof EncodeError)) throw err;
      r.errors.push(err.message);
    }
  }
  if (r.errors.length > 0) return { ...r, status: 'error' };
  const shrinks = toBuild.map((e) => shrinkMessage(ctx, e, e.xl!)).filter((m): m is string => m !== null);
  if (shrinks.length > 0) {
    r.confirmations.push({
      kind: 'shrink',
      message: 'コードが大きく減っています。Copilot が省略した可能性があります',
      files: shrinks,
    });
  }
  if (scan.deletes.length > 0) {
    r.confirmations.push({
      kind: 'delete',
      message: '次のファイルとシートを削除します',
      files: scan.deletes.map((d) => d.fileName),
    });
  }
  const useGit = await gitConfirmations(ctx, ref, r);
  if (r.confirmations.length > 0 && !opts.confirmed) return { ...r, status: 'confirm' };
  if (useGit && !(await commitBefore(ctx, ref, 'Build', r))) return r;
  if (!(await assertClosed(ref, r, 'ソースは変更していません'))) return r;

  let bookChanged = false;
  for (const e of scan.entries) {
    switch (e.status) {
      case 'excel-changed':
      case 'excel-new': {
        await atomicWrite(e.srcAbs, encoded.get(e.name)!);
        bs.files[e.name] = toState(e.xl!);
        r.changes.push({ action: 'write-file', target: e.name });
        const lines = textToLines(e.xl!.text);
        if (!sameLines(e.xlRaw, lines)) {
          scan.book.writeLines(e.sheetName!, lines);
          r.changes.push({ action: 'reformat-sheet', target: e.name });
          bookChanged = true;
        }
        break;
      }
      case 'source-changed':
      case 'source-new':
      case 'source-deleted':
      case 'sheet-missing':
        pullToSheet(scan, bs, e, r);
        bookChanged = true;
        break;
      case 'clean':
        bookChanged = settleClean(scan, bs, e, r) || bookChanged;
        break;
      case 'gone':
        delete bs.files[e.name];
        break;
    }
  }
  for (const d of scan.deletes) {
    await rm(`${ref.dirAbs}/${d.fileName}`, { force: true });
    scan.book.deleteSheet(d.sheet);
    delete bs.files[d.fileName];
    r.changes.push({ action: 'delete-file', target: d.fileName }, { action: 'delete-sheet', target: d.sheet });
    bookChanged = true;
  }
  // ソースは出力済み。ブックが保存できなくても、次回の Build / Sync で両側一致として扱われる
  if (
    bookChanged &&
    !(await assertClosed(ref, r, 'ソースへの出力は完了しています。Excel を閉じて再度 Build してください'))
  )
    return r;
  if (bookChanged) await scan.book.save(ref.abs);
  bs.lastBuildAt = new Date().toISOString();
  await saveState(root, ctx.state);
  return r;
}

/** Sync（ソースコード → Excel）5章 */
export async function sync(root: string, bookAbs: string, opts: SyncOptions = {}): Promise<OpResult> {
  const ctx = await openProject(root);
  const ref = bookRef(root, bookAbs);
  const r = newResult();
  const scan = await preflight(ctx, ref, r);
  if (!scan) return { ...r, status: 'error' };
  const bs = bookState(ctx.state, ref.rel);

  const unbuilt = scan.entries.filter((e) => e.status === 'excel-changed' || e.status === 'excel-new');
  if (unbuilt.length > 0 && !opts.discardExcelChanges) {
    r.unbuilt = unbuilt.map((e) => e.name);
    return { ...r, status: 'needs-decision' };
  }
  if (scan.deletes.length > 0) {
    r.warnings.push(`削除マーク付きのシートは Build で処理されます: ${scan.deletes.map((d) => d.sheet).join(', ')}`);
  }

  const toSync = scan.entries.filter((e) => (e.status === 'source-changed' || e.status === 'conflict') && e.src);
  const shrinks = toSync.map((e) => shrinkMessage(ctx, e, e.src!)).filter((m): m is string => m !== null);
  if (shrinks.length > 0) {
    r.confirmations.push({ kind: 'shrink', message: 'ソースコードが大きく減っています', files: shrinks });
  }
  const useGit = await gitConfirmations(ctx, ref, r);
  if (r.confirmations.length > 0 && !opts.confirmed) return { ...r, status: 'confirm' };
  if (useGit && !(await commitBefore(ctx, ref, 'Sync', r))) return r;

  let bookChanged = false;
  const conflicts: FileEntry[] = [];
  for (const e of scan.entries) {
    switch (e.status) {
      case 'source-changed':
      case 'source-new':
      case 'source-deleted':
      case 'sheet-missing':
        pullToSheet(scan, bs, e, r);
        bookChanged = true;
        break;
      case 'excel-changed':
      case 'excel-new':
        // ここに来るのは discardExcelChanges のときだけ
        pullToSheet(scan, bs, e, r);
        bookChanged = true;
        break;
      case 'conflict':
        if (opts.discardExcelChanges) pullToSheet(scan, bs, e, r);
        else conflicts.push(e);
        bookChanged = true;
        break;
      case 'clean':
        bookChanged = settleClean(scan, bs, e, r) || bookChanged;
        break;
      case 'gone':
        delete bs.files[e.name];
        break;
    }
  }
  if (conflicts.length > 0) recordConflicts(scan, bs, conflicts, r);
  if (bookChanged && !(await assertClosed(ref, r, 'ブックは変更していません'))) return r;
  if (bookChanged) await scan.book.save(ref.abs);
  bs.lastSyncAt = new Date().toISOString();
  await saveState(root, ctx.state);
  return conflicts.length > 0 ? { ...r, status: 'conflict' } : r;
}
