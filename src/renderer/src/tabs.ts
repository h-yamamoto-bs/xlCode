/**
 * エディタ領域のタブ（VS Code と同じ振る舞い）。
 * - サイドバーで 1 回押すと「プレビュー」（斜体）として開く。別のものをプレビューで開くと置き換わる
 * - 2 回押す・編集する・タブを 2 回押すと固定される
 * - 閉じたタブは開き直せる
 * 状態は純粋な関数で更新する（テストしやすく、画面はこの結果を描くだけ）
 */

export type TabKind = 'book' | 'doc' | 'settings' | 'help';

export interface Tab {
  /** kind:rel（settings / help は kind だけ） */
  id: string;
  kind: TabKind;
  /** book: ブックの相対パス、doc: ドキュメントの相対パス */
  rel?: string;
  preview: boolean;
}

export interface TabsState {
  tabs: Tab[];
  active: string | null;
  /** 閉じたタブ（新しい順）。開き直す用 */
  closed: Tab[];
}

export const EMPTY_TABS: TabsState = { tabs: [], active: null, closed: [] };

export function tabId(kind: TabKind, rel?: string): string {
  return rel === undefined ? kind : `${kind}:${rel}`;
}

export function activeTab(s: TabsState): Tab | null {
  return s.tabs.find((t) => t.id === s.active) ?? null;
}

export interface OpenOptions {
  /** プレビューとして開く（既定: false = 固定） */
  preview?: boolean;
  /** 開いたタブを表示する（既定: true） */
  activate?: boolean;
}

/** タブを開く。既にあれば表示する（preview が false なら固定にする） */
export function openTab(s: TabsState, kind: TabKind, rel: string | undefined, opts: OpenOptions = {}): TabsState {
  const id = tabId(kind, rel);
  const preview = opts.preview ?? false;
  const activate = opts.activate ?? true;
  const existing = s.tabs.find((t) => t.id === id);
  if (existing) {
    const tabs = preview ? s.tabs : s.tabs.map((t) => (t.id === id ? { ...t, preview: false } : t));
    return { ...s, tabs, active: activate ? id : s.active };
  }
  const tab: Tab = { id, kind, rel, preview };
  const closed = s.closed.filter((t) => t.id !== id);
  const prev = s.tabs.findIndex((t) => t.preview);
  if (preview && prev >= 0) {
    // プレビューは 1 つだけ。置き換える
    const tabs = s.tabs.map((t, i) => (i === prev ? tab : t));
    return { tabs, active: activate ? id : s.active === s.tabs[prev].id ? id : s.active, closed };
  }
  // 表示中のタブの右に開く
  const at = s.tabs.findIndex((t) => t.id === s.active);
  const tabs = [...s.tabs];
  tabs.splice(at < 0 ? tabs.length : at + 1, 0, tab);
  return { tabs, active: activate ? id : s.active, closed };
}

/** プレビューを固定にする */
export function pinTab(s: TabsState, id: string): TabsState {
  if (!s.tabs.some((t) => t.id === id && t.preview)) return s;
  return { ...s, tabs: s.tabs.map((t) => (t.id === id ? { ...t, preview: false } : t)) };
}

/** 閉じる。表示中なら右隣（無ければ左隣）を表示する */
export function closeTab(s: TabsState, id: string): TabsState {
  const i = s.tabs.findIndex((t) => t.id === id);
  if (i < 0) return s;
  const tabs = s.tabs.filter((t) => t.id !== id);
  let active = s.active;
  if (s.active === id) active = (tabs[i] ?? tabs[i - 1])?.id ?? null;
  return { tabs, active, closed: [s.tabs[i], ...s.closed.filter((t) => t.id !== id)].slice(0, 20) };
}

export function closeOthers(s: TabsState, id: string): TabsState {
  return s.tabs.filter((t) => t.id !== id).reduce((acc, t) => closeTab(acc, t.id), s);
}

export function closeAll(s: TabsState): TabsState {
  return s.tabs.reduce((acc, t) => closeTab(acc, t.id), s);
}

/** 最後に閉じたタブを開き直す */
export function reopenClosed(s: TabsState): TabsState {
  const [t, ...rest] = s.closed;
  if (!t) return s;
  const next = openTab({ ...s, closed: rest }, t.kind, t.rel, { preview: false });
  return next;
}

/** 表示するタブを delta だけずらす（端で折り返す） */
export function cycleTab(s: TabsState, delta: number): TabsState {
  if (s.tabs.length === 0) return s;
  const i = s.tabs.findIndex((t) => t.id === s.active);
  const n = ((((i < 0 ? 0 : i) + delta) % s.tabs.length) + s.tabs.length) % s.tabs.length;
  return { ...s, active: s.tabs[n].id };
}

/** n 番目（0 始まり）のタブを表示する。範囲外なら何もしない */
export function activateIndex(s: TabsState, i: number): TabsState {
  const t = s.tabs[i];
  return t ? { ...s, active: t.id } : s;
}

export function activate(s: TabsState, id: string): TabsState {
  return s.tabs.some((t) => t.id === id) ? { ...s, active: id } : s;
}

/** タブを to の位置へ移動する（ドラッグで並べ替え） */
export function moveTab(s: TabsState, id: string, to: number): TabsState {
  const from = s.tabs.findIndex((t) => t.id === id);
  if (from < 0) return s;
  const tabs = [...s.tabs];
  const [t] = tabs.splice(from, 1);
  tabs.splice(Math.max(0, Math.min(to, tabs.length)), 0, t);
  return { ...s, tabs };
}

/** 無くなったブックなどのタブを閉じる */
export function pruneTabs(s: TabsState, keep: (t: Tab) => boolean): TabsState {
  return s.tabs.filter((t) => !keep(t)).reduce((acc, t) => closeTab(acc, t.id), s);
}

/** 保存する形（閉じたタブの履歴は保存しない） */
export function serializeTabs(s: TabsState): { tabs: Tab[]; active: string | null } {
  return { tabs: s.tabs, active: s.active };
}

export function deserializeTabs(v: unknown): TabsState {
  if (typeof v !== 'object' || v === null) return EMPTY_TABS;
  const o = v as { tabs?: unknown; active?: unknown };
  const tabs = Array.isArray(o.tabs)
    ? o.tabs.filter(
        (t): t is Tab =>
          typeof t === 'object' &&
          t !== null &&
          typeof (t as Tab).id === 'string' &&
          ['book', 'doc', 'settings', 'help'].includes((t as Tab).kind),
      )
    : [];
  const active = typeof o.active === 'string' && tabs.some((t) => t.id === o.active) ? o.active : (tabs[0]?.id ?? null);
  return { tabs: tabs.map((t) => ({ ...t, preview: !!t.preview })), active, closed: [] };
}
