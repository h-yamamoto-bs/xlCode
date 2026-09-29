/**
 * xlsx（XML）に保存できる形へ変換する。読み込み時は ExcelJS が元に戻す。
 * - 文字列中の「_xHHHH_」は xlsx では文字コードとして解釈されるため、先頭の「_」を _x005F_ にする
 * - XML に書けない制御文字は _xHHHH_ で表す（そのまま書くと読み込めないブックになる）
 */
export function escapeCell(text: string): string {
  return text
    .replace(/_(x[0-9A-Fa-f]{4}_)/g, '_x005F_$1')
    .replace(XML_INVALID, (c) => `_x${c.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}_`);
}

/** XML 1.0 に書けない文字（タブ・改行・復帰以外の制御文字など） */
// eslint-disable-next-line no-control-regex
const XML_INVALID = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g;

/** 対になっていないサロゲート（壊れた UTF-16）を含むか。xlsx に保存できない */
export function hasLoneSurrogate(text: string): boolean {
  return /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(text);
}
