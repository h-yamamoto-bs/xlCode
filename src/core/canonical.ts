import { formatText } from './format';
import { normalizeText, sha256, textToLines } from './normalize';
import type { XlcodeConfig } from './config';

export interface Canon {
  text: string;
  hash: string;
  lines: number;
  chars: number;
  formatError?: string;
}

/**
 * 正規化 → 整形 → 正規化 を行い、比較・保存用の正準テキストを得る。
 * Build と Sync の両方で同じ処理を通すことで、往復で差分が出ないようにする（3.7 / No.7）。
 *
 * 正準テキストは不動点なので、既知の正準ハッシュ（state.json の値）と一致する入力は
 * 整形を省略する。未変更ファイルに毎回 Prettier をかけないための高速化。
 */
export class Canonicalizer {
  constructor(
    private readonly config: XlcodeConfig,
    private readonly known: Set<string> = new Set(),
  ) {}

  addKnown(hash: string): void {
    this.known.add(hash);
  }

  async canonical(fileName: string, absPath: string, raw: string): Promise<Canon> {
    const opts = { trimTrailingWhitespace: this.config.trimTrailingWhitespace };
    let text = normalizeText(raw, fileName, opts);
    let hash = sha256(text);
    let formatError: string | undefined;
    if (this.config.format && !this.known.has(hash)) {
      const r = await formatText(text, absPath);
      formatError = r.error;
      text = normalizeText(r.text, fileName, opts);
      hash = sha256(text);
    }
    return { text, hash, lines: textToLines(text).length, chars: text.length, formatError };
  }
}
