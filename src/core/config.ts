import { mkdir, readFile } from 'node:fs/promises';
import path from 'node:path';
import { atomicWrite } from './atomic';
import { CONFIG_FILE, XLCODE_DIR } from './constants';

/**
 * プロジェクトの種類（最初のブックを作るときに決める）
 * - source: シートをソースコードのファイルとして書き出す
 * - vba: .bas / .cls / .frm のシートを VBA として書き込んだ .xlsm を生成する
 */
export type ProjectMode = 'source' | 'vba';

export interface XlcodeConfig {
  /** 未設定なら source */
  mode?: ProjectMode;
  /** 省略検知の閾値（0.3 = 30%以上の減少で警告） */
  shrinkThreshold: number;
  /** 省略検知の対象とする最小行数（前回の行数がこれ未満なら検知しない） */
  shrinkMinLines: number;
  /** 行末空白を削除するか（未決 No.5 の暫定値。Markdown は対象外） */
  trimTrailingWhitespace: boolean;
  /** Prettier で整形するか */
  format: boolean;
  /** Build / Sync 直前に Git へ自動コミットするか */
  autoCommit: boolean;
  /**
   * プロジェクトルートに対応する Web 上の URL（Web 版 Excel で開く用）。
   * 未設定なら Windows の OneDrive 設定から自動で求める。
   */
  webUrlBase?: string;
  /**
   * ブックの置き場所（例: OneDrive 内の専用フォルダ）。ソースと同じフォルダ構成でブックだけを置く。
   * 未設定ならソースの各フォルダの中に置く。
   */
  bookRoot?: string;
  /** 拡張子のないファイル名のうち、追加でコードとして扱うもの（Makefile などは既定で扱う） */
  extraCodeNames?: string[];
}

export const DEFAULT_CONFIG: XlcodeConfig = {
  shrinkThreshold: 0.3,
  shrinkMinLines: 10,
  trimTrailingWhitespace: true,
  format: true,
  autoCommit: true,
};

/** .xlcode/config.json を読み込み、既定値とマージする。壊れていれば例外（設定が黙って無視されないように） */
export async function loadConfig(root: string): Promise<XlcodeConfig> {
  let raw: string;
  try {
    raw = await readFile(path.join(root, XLCODE_DIR, CONFIG_FILE), 'utf8');
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code === 'ENOENT') return { ...DEFAULT_CONFIG };
    throw e;
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    throw new Error(`${XLCODE_DIR}/${CONFIG_FILE} が JSON として読めません。修正するか削除してください`);
  }
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new Error(`${XLCODE_DIR}/${CONFIG_FILE} の形式が違います。修正するか削除してください`);
  }
  return { ...DEFAULT_CONFIG, ...(parsed as Partial<XlcodeConfig>) };
}

/** .xlcode/config.json に保存する（既定値と同じ項目も含めて書く） */
export async function saveConfig(root: string, config: XlcodeConfig): Promise<void> {
  await mkdir(path.join(root, XLCODE_DIR), { recursive: true });
  await atomicWrite(path.join(root, XLCODE_DIR, CONFIG_FILE), JSON.stringify(config, null, 2) + '\n');
}
