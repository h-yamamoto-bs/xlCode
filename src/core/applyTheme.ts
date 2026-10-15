import { checkBookOpen } from './lock';
import { bookRef, bookRootOf, openProject } from './project';
import { classifySheet } from './sheetName';
import type { BookTheme } from './theme';
import { Book } from './workbook';

/** xlCode が管理するシートか（コード・削除マーク・ルール・予約・衝突）。画面・データ用のシートは対象外 */
export function isManagedSheet(name: string, extraCodeNames: readonly string[] = []): boolean {
  const kind = classifySheet(name, extraCodeNames).kind;
  return kind !== 'other' && kind !== 'unknown-reserved';
}

export interface ThemeState {
  /** 今の設定で使うテーマ */
  desired: BookTheme;
  /** xlCode が管理するシートの数 */
  managed: number;
  /** そのうち、desired のテーマが当たっていないシート（Copilot が作ったシート、別のテーマのシートなど） */
  outdated: string[];
}

/** ブックのシートのうち、今のテーマが当たっていないものを調べる（読み取りのみ） */
export function themeState(book: Book, desired: BookTheme, extraCodeNames: readonly string[] = []): ThemeState {
  const managed = book.sheetNames().filter((n) => isManagedSheet(n, extraCodeNames));
  return { desired, managed: managed.length, outdated: managed.filter((n) => book.sheetTheme(n) !== desired) };
}

export interface ApplyThemeResult {
  theme: BookTheme;
  /** テーマを当て直したシート */
  sheets: string[];
}

/**
 * ブックのテーマを当て直す。xlCode が管理するシートの、背景・文字色・枠線・タブの色だけを変える
 * （セルの中身は読み直さず、XML のスタイルだけを書き換える）。画面・データ用のシートは触らない。
 * ブックが開かれているときは何もせずエラーにする
 */
export async function applyTheme(root: string, bookAbs: string): Promise<ApplyThemeResult> {
  const ctx = await openProject(root);
  const ref = bookRef(root, bookAbs, bookRootOf(root, ctx.config));
  const open = await checkBookOpen(ref.abs);
  if (open.open) throw new Error(`${ref.rel} が開かれています。Excel を閉じてから実行してください（${open.reason}）`);
  const book = await Book.load(ref.abs, ctx.bookTheme);
  const state = themeState(book, ctx.bookTheme, ctx.config.extraCodeNames);
  if (state.outdated.length === 0) return { theme: ctx.bookTheme, sheets: [] };
  book.restyle(state.outdated);
  // 書き込み直前にもう一度確認する（確認〜保存の間に Excel で開かれた場合に備える）
  const again = await checkBookOpen(ref.abs);
  if (again.open)
    throw new Error(`${ref.rel} が途中で開かれたため中断しました。何も変更していません（${again.reason}）`);
  await book.save(ref.abs);
  return { theme: ctx.bookTheme, sheets: state.outdated };
}
