import {
  AGENTS_SHEET,
  CONFLICT_PREFIX,
  DELETE_PREFIX,
  FORBIDDEN_CHARS,
  MAX_SHEET_NAME,
  RESERVED_PREFIX,
  RESERVED_SHEETS,
  TREE_SHEET,
} from './constants';

export type SheetKind = 'tree' | 'reserved' | 'conflict' | 'unknown-reserved' | 'agents' | 'delete' | 'code';

export interface SheetInfo {
  kind: SheetKind;
  /** code / delete の場合の対象ファイル名 */
  fileName?: string;
}

export function classifySheet(name: string): SheetInfo {
  if (name === TREE_SHEET) return { kind: 'tree' };
  if (RESERVED_SHEETS.includes(name)) return { kind: 'reserved' };
  if (/^#conflict_\d+$/.test(name)) return { kind: 'conflict' };
  if (name.startsWith(CONFLICT_PREFIX) || name.startsWith(RESERVED_PREFIX)) return { kind: 'unknown-reserved' };
  if (name === AGENTS_SHEET) return { kind: 'agents' };
  if (name.startsWith(DELETE_PREFIX)) return { kind: 'delete', fileName: name.slice(DELETE_PREFIX.length) };
  return { kind: 'code', fileName: name };
}

/** ファイル名がシート名として使えるか検証する（2.3）。問題があればエラーメッセージを返す */
export function validateFileName(name: string): string | null {
  if (name.length === 0) return 'ファイル名が空です';
  if (name.length > MAX_SHEET_NAME) {
    return `「${name}」は ${name.length} 文字です（拡張子込み ${MAX_SHEET_NAME} 文字以内）`;
  }
  const bad = FORBIDDEN_CHARS.filter((c) => name.includes(c));
  if (bad.length > 0) return `「${name}」に使用できない記号 ${bad.join(' ')} が含まれています`;
  if (name.startsWith("'") || name.endsWith("'")) return `「${name}」の先頭・末尾にアポストロフィは使えません`;
  if (name.startsWith(RESERVED_PREFIX)) return `「${name}」は予約接頭辞「#」で始まっています`;
  return null;
}

/** Excel のシート名は大文字小文字を区別しないため、重複判定は小文字で行う */
export function sheetKey(name: string): string {
  return name.toLowerCase();
}
