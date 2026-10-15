import { AGENTS_SHEET, LOCAL_AGENTS_SHEET, RESERVED_PREFIX } from './constants';

/**
 * 編集用ブックのテーマ（シートの見た目）。
 * xlCode が書くシートは、このテーマの塗りつぶし・文字色・枠線なし・タブの色で作る。
 * - dark: エディタ（VS Code）で編集しているような、黒に近い背景
 * - light: ふつうの白い背景
 * 画面（renderer）の色見本にも使うため、node のモジュールには依存しない
 */
export type BookTheme = 'dark' | 'light';

/** 設定の値。auto は xlCode の表示モードに合わせる */
export type BookThemeSetting = 'auto' | BookTheme;

export const BOOK_THEME_SETTINGS: readonly BookThemeSetting[] = ['auto', 'dark', 'light'];

/** シートの役割（タブの色を分ける） */
export type SheetRole = 'code' | 'meta' | 'conflict';

export interface ThemePalette {
  label: string;
  /** セルの背景（RRGGBB） */
  fill: string;
  /** 文字色（RRGGBB） */
  text: string;
  /** タブの色（RRGGBB）。テーマの目印も兼ねるので、テーマ間で重ならない値にする */
  tab: Record<SheetRole, string>;
}

export const THEMES: Record<BookTheme, ThemePalette> = {
  dark: {
    label: 'ダーク',
    fill: '1E1E1E',
    text: 'D4D4D4',
    tab: { code: '569CD6', meta: '858585', conflict: 'F14C4C' },
  },
  light: {
    label: 'ライト',
    fill: 'FFFFFF',
    text: '1F1F1F',
    tab: { code: '0078D4', meta: 'A0A0A0', conflict: 'D13438' },
  },
};

/** 文字のフォント（コードを書くので等幅） */
export const THEME_FONT = { name: 'Consolas', size: 10 } as const;

/** シートの役割。Agents.md・LocalAgents.md・# で始まる予約シートは meta、衝突シートは conflict */
export function sheetRole(name: string): SheetRole {
  if (/^#conflict_\d+$/.test(name)) return 'conflict';
  if (name.startsWith(RESERVED_PREFIX) || name === AGENTS_SHEET || name === LOCAL_AGENTS_SHEET) return 'meta';
  return 'code';
}

/** タブの色（ARGB）→ そのテーマ。テーマの色でなければ null */
export function themeOfTabColor(argb: string | undefined): BookTheme | null {
  if (!argb) return null;
  const rgb = argb.slice(-6).toUpperCase();
  for (const t of ['dark', 'light'] as const) {
    if (Object.values(THEMES[t].tab).includes(rgb)) return t;
  }
  return null;
}

/**
 * 設定から、実際に使うテーマを決める。
 * auto は xlCode の表示モード（host）に合わせる。表示モードが分からない（CLI など）ときはダーク
 */
export function resolveBookTheme(setting: BookThemeSetting | undefined, host: BookTheme | null): BookTheme {
  if (setting === 'dark' || setting === 'light') return setting;
  return host ?? 'dark';
}

/** xlCode の表示モード（画面が決めて、メインプロセスが覚えておく）。auto のときの基準 */
let hostTheme: BookTheme | null = null;

export function setHostTheme(mode: BookTheme | null): void {
  hostTheme = mode;
}

export function getHostTheme(): BookTheme | null {
  return hostTheme;
}
