import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { CONFIG_FILE, XLCODE_DIR } from './constants';

export interface XlcodeConfig {
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
}

export const DEFAULT_CONFIG: XlcodeConfig = {
  shrinkThreshold: 0.3,
  shrinkMinLines: 10,
  trimTrailingWhitespace: true,
  format: true,
  autoCommit: true,
};

/** .xlcode/config.json を読み込み、既定値とマージする */
export async function loadConfig(root: string): Promise<XlcodeConfig> {
  try {
    const raw = await readFile(path.join(root, XLCODE_DIR, CONFIG_FILE), 'utf8');
    return { ...DEFAULT_CONFIG, ...(JSON.parse(raw) as Partial<XlcodeConfig>) };
  } catch {
    return { ...DEFAULT_CONFIG };
  }
}
