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

export type SheetKind =
  | 'tree'
  | 'reserved'
  | 'conflict'
  | 'unknown-reserved'
  | 'agents'
  | 'delete'
  | 'code'
  /** コードではないシート（UI・データ表示など）。xlCode は一切触らない */
  | 'other';

/** 拡張子のないファイル名のうち、コードとして扱うもの（大文字小文字は区別しない） */
export const EXTENSIONLESS_CODE_NAMES: readonly string[] = [
  'Makefile',
  'GNUmakefile',
  'Dockerfile',
  'Containerfile',
  'Procfile',
  'Gemfile',
  'Rakefile',
  'Brewfile',
  'Pipfile',
  'Jenkinsfile',
  'Vagrantfile',
  'Caddyfile',
  'Justfile',
  'LICENSE',
  'README',
  'CHANGELOG',
  'CODEOWNERS',
];

/**
 * コードのシート・ファイルとして扱う名前か。
 * 拡張子がある（例: App.tsx, run.bat, .gitignore）か、拡張子のない決まった名前（Makefile など）。
 * UI やデータ表示用のシート（例: 「UI」「データ」）はコードではない。
 */
export function isCodeName(name: string, extra: readonly string[] = []): boolean {
  if (/\.[^.]+$/.test(name)) return true;
  const n = name.toLowerCase();
  return [...EXTENSIONLESS_CODE_NAMES, ...extra].some((x) => x.toLowerCase() === n);
}

export interface SheetInfo {
  kind: SheetKind;
  /** code / delete の場合の対象ファイル名 */
  fileName?: string;
}

export function classifySheet(name: string, extraCodeNames: readonly string[] = []): SheetInfo {
  if (name === TREE_SHEET) return { kind: 'tree' };
  if (RESERVED_SHEETS.includes(name)) return { kind: 'reserved' };
  if (/^#conflict_\d+$/.test(name)) return { kind: 'conflict' };
  if (name.startsWith(CONFLICT_PREFIX) || name.startsWith(RESERVED_PREFIX)) return { kind: 'unknown-reserved' };
  if (name === AGENTS_SHEET) return { kind: 'agents' };
  if (name.startsWith(DELETE_PREFIX)) {
    const fileName = name.slice(DELETE_PREFIX.length);
    return isCodeName(fileName, extraCodeNames) ? { kind: 'delete', fileName } : { kind: 'other' };
  }
  return isCodeName(name, extraCodeNames) ? { kind: 'code', fileName: name } : { kind: 'other' };
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
  return windowsNameError(name);
}

/** Windows の予約デバイス名（拡張子が付いていても使えない） */
const WINDOWS_RESERVED = /^(con|prn|aux|nul|com[0-9¹²³]|lpt[0-9¹²³])(\.|$)/i;

/** Windows でファイル名に使えないかを検証する */
export function windowsNameError(name: string): string | null {
  const bad = ['<', '>', '"', '|'].filter((c) => name.includes(c));
  if (bad.length > 0) return `「${name}」に Windows で使えない記号 ${bad.join(' ')} が含まれています`;
  // eslint-disable-next-line no-control-regex
  if (/[\u0000-\u001f]/.test(name)) return `「${name}」に制御文字が含まれています`;
  if (/[. ]$/.test(name)) return `「${name}」は末尾が「.」または空白のため Windows で使えません`;
  if (WINDOWS_RESERVED.test(name)) return `「${name}」は Windows の予約名のため使えません`;
  return null;
}

/** Excel のシート名は大文字小文字を区別しないため、重複判定は小文字で行う */
export function sheetKey(name: string): string {
  return name.toLowerCase();
}
