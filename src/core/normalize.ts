import { createHash } from 'node:crypto';

export interface NormalizeOptions {
  trimTrailingWhitespace: boolean;
}

function isMarkdown(fileName: string): boolean {
  return /\.(md|markdown|mdx)$/i.test(fileName);
}

/**
 * 正規化ルール（3.7）
 * - BOM 除去、改行コードを LF に統一
 * - 行末空白の削除（設定で切り替え。Markdown は改行記法を壊すため対象外）
 * - 末尾の空行を除き、ファイル末尾に改行を1つだけ付与（空ファイルは空のまま）
 */
export function normalizeText(text: string, fileName: string, opts: NormalizeOptions): string {
  let lines = text
    .replace(/^\uFEFF/, '')
    .replace(/\r\n?/g, '\n')
    .split('\n');
  if (opts.trimTrailingWhitespace && !isMarkdown(fileName)) {
    lines = lines.map((l) => l.replace(/[ \t]+$/, ''));
  }
  while (lines.length > 0 && lines[lines.length - 1] === '') lines.pop();
  return linesToText(lines);
}

/** 行配列 → テキスト（末尾改行1つ） */
export function linesToText(lines: readonly string[]): string {
  return lines.length === 0 ? '' : lines.join('\n') + '\n';
}

/** 正規化済みテキスト → 行配列（1行 = 1セル） */
export function textToLines(text: string): string[] {
  if (text === '') return [];
  return (text.endsWith('\n') ? text.slice(0, -1) : text).split('\n');
}

export function sha256(text: string): string {
  return createHash('sha256').update(text, 'utf8').digest('hex');
}
