import type { AppThemeSetting } from '../../shared/api';
import { storageGet } from './api';

/** 画面の表示モード（xlCode の画面そのものの配色）。ブックのテーマの「auto」の基準にもなる */
export const APP_THEME_SETTINGS: { id: AppThemeSetting; label: string; desc: string }[] = [
  { id: 'dark', label: 'ダーク', desc: '暗い配色（VS Code の Dark Modern）。' },
  { id: 'light', label: 'ライト', desc: '明るい配色（VS Code の Light Modern）。' },
  {
    id: 'system',
    label: 'OS に合わせる',
    desc: 'Windows の「色」の設定（ライト / ダーク）に従い、変わったら自動で切り替えます。',
  },
];

export function isAppThemeSetting(v: unknown): v is AppThemeSetting {
  return v === 'dark' || v === 'light' || v === 'system';
}

/** 保存してある設定（無い・壊れていればダーク。これまでの見た目） */
export function readSavedAppTheme(): AppThemeSetting {
  const v = storageGet<unknown>('appTheme', 'dark');
  return isAppThemeSetting(v) ? v : 'dark';
}

/** 設定と OS の今の状態から、実際に使うモードを決める */
export function resolveAppTheme(setting: AppThemeSetting, prefersDark: boolean): 'dark' | 'light' {
  if (setting === 'system') return prefersDark ? 'dark' : 'light';
  return setting;
}

/** 画面全体の配色を切り替える（CSS の data-theme と、フォームやスクロールバーの color-scheme） */
export function applyDocumentTheme(resolved: 'dark' | 'light'): void {
  document.documentElement.dataset.theme = resolved;
  document.documentElement.style.colorScheme = resolved;
}

export const PREFERS_DARK = '(prefers-color-scheme: dark)';
