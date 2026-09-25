import * as prettier from 'prettier';

export interface FormatResult {
  text: string;
  /** 構文エラー等で整形できなかった場合のメッセージ（text は未整形のまま） */
  error?: string;
}

/**
 * Prettier で整形する。Prettier が対応しない拡張子はそのまま返す。
 * absPath は設定ファイル（.prettierrc）の解決とパーサ推定に使う。存在しなくてもよい。
 */
export async function formatText(text: string, absPath: string): Promise<FormatResult> {
  if (text === '') return { text };
  const info = await prettier.getFileInfo(absPath, { resolveConfig: false });
  if (info.ignored || !info.inferredParser) return { text };
  try {
    const options = (await prettier.resolveConfig(absPath, { editorconfig: true })) ?? {};
    const out = await prettier.format(text, { ...options, filepath: absPath, endOfLine: 'lf' });
    return { text: out };
  } catch (e) {
    const msg = e instanceof Error ? e.message.split('\n')[0] : String(e);
    return { text, error: msg };
  }
}
