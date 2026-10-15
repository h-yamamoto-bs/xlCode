import { readFileSync } from 'node:fs';
import { writeFile } from 'node:fs/promises';
import path from 'node:path';
import { app, nativeTheme, type BrowserWindow } from 'electron';
import { setHostTheme } from '../core';
import type { AppThemeSetting } from '../shared/api';

/** ウィンドウの色（本体の背景・タイトルバーのボタン部分）。画面側の配色（index.css）と合わせる */
export const WINDOW_COLORS = {
  dark: { background: '#1f1f1f', bar: '#181818', symbol: '#cccccc' },
  light: { background: '#ffffff', bar: '#f8f8f8', symbol: '#3b3b3b' },
} as const;

const file = () => path.join(app.getPath('userData'), 'appearance.json');

/**
 * 前回の表示モード。画面（localStorage）が決めたものを、起動直後（画面を読み込む前）に使うために控えておく
 * （最初に一瞬だけ別の色で光るのを避ける）
 */
export function readAppearance(): AppThemeSetting {
  try {
    const v = (JSON.parse(readFileSync(file(), 'utf8')) as { setting?: unknown }).setting;
    return v === 'light' || v === 'system' ? v : 'dark';
  } catch {
    return 'dark';
  }
}

/** 設定を OS（Electron）に反映し、実際に使うモードを返す */
export function resolveAppearance(setting: AppThemeSetting): 'dark' | 'light' {
  nativeTheme.themeSource = setting;
  const resolved = nativeTheme.shouldUseDarkColors ? 'dark' : 'light';
  setHostTheme(resolved);
  return resolved;
}

/** 画面が表示モードを決めたとき（起動時・設定を変えたとき・OS の設定が変わったとき） */
export async function applyAppearance(
  win: BrowserWindow | null,
  setting: AppThemeSetting,
  resolved: 'dark' | 'light',
): Promise<void> {
  nativeTheme.themeSource = setting;
  // ブックのテーマ（auto）の基準
  setHostTheme(resolved);
  const c = WINDOW_COLORS[resolved];
  try {
    win?.setBackgroundColor(c.background);
    if (process.platform !== 'darwin') win?.setTitleBarOverlay({ color: c.bar, symbolColor: c.symbol, height: 35 });
  } catch {
    // タイトルバーを変えられない環境でも、画面の配色だけは切り替わる
  }
  await writeFile(file(), JSON.stringify({ setting }) + '\n').catch(() => undefined);
}
