/**
 * キーボードショートカット（VS Code と同じ書き方）。
 * 割り当ては "ctrl+shift+b" のような文字列。2 段階（コード）は "ctrl+k ctrl+s" のように空白で区切る。
 * macOS では ctrl を ⌘ として扱う（押すのは Cmd、表示は ⌘）。
 */

const KEY_NAMES: Record<string, string> = {
  ' ': 'space',
  arrowup: 'up',
  arrowdown: 'down',
  arrowleft: 'left',
  arrowright: 'right',
  escape: 'escape',
  enter: 'enter',
  tab: 'tab',
  backspace: 'backspace',
  delete: 'delete',
  pageup: 'pageup',
  pagedown: 'pagedown',
  home: 'home',
  end: 'end',
};

const MODIFIER_KEYS = new Set(['control', 'shift', 'alt', 'meta', 'os']);

export interface KeyLike {
  key: string;
  code?: string;
  ctrlKey: boolean;
  metaKey: boolean;
  shiftKey: boolean;
  altKey: boolean;
}

/**
 * キー入力を "ctrl+shift+b" の形にする。修飾キーだけなら null。
 * Shift+数字で記号になる配列（Shift+1 = ! など）に備えて、数字は物理キー（code）から取る
 */
export function eventToKey(e: KeyLike, platform = 'win32'): string | null {
  const raw = e.key.toLowerCase();
  if (MODIFIER_KEYS.has(raw)) return null;
  let key = KEY_NAMES[raw] ?? raw;
  const digit = e.code?.match(/^Digit(\d)$/);
  if (digit) key = digit[1];
  const mac = platform === 'darwin';
  const mod = mac ? e.metaKey : e.ctrlKey;
  const parts: string[] = [];
  if (mod) parts.push('ctrl');
  // macOS の Control キーは別物として扱わない（使わない）
  if (e.shiftKey) parts.push('shift');
  if (e.altKey) parts.push('alt');
  parts.push(key);
  return parts.join('+');
}

/** 割り当ての文字列を正規化する（大文字小文字・並び順の違いを吸収） */
export function normalizeBinding(binding: string): string {
  return binding
    .trim()
    .split(/\s+/)
    .map((chord) => {
      const parts = chord.toLowerCase().split('+');
      const key = parts.pop()!;
      const mods = ['ctrl', 'shift', 'alt'].filter((m) => parts.includes(m));
      return [...mods, key].join('+');
    })
    .join(' ');
}

const LABELS: Record<string, string> = {
  up: '↑',
  down: '↓',
  left: '←',
  right: '→',
  pageup: 'PageUp',
  pagedown: 'PageDown',
  escape: 'Esc',
  enter: 'Enter',
  tab: 'Tab',
  space: 'Space',
  backspace: 'Backspace',
  delete: 'Delete',
  ',': ',',
  '`': '`',
};

/** 画面に出す形（Ctrl+Shift+B、macOS では ⇧⌘B）にする */
export function formatKey(binding: string, platform = 'win32'): string {
  const mac = platform === 'darwin';
  return normalizeBinding(binding)
    .split(' ')
    .map((chord) => {
      const parts = chord.split('+');
      const key = parts.pop()!;
      const label = LABELS[key] ?? (key.length === 1 ? key.toUpperCase() : key[0].toUpperCase() + key.slice(1));
      if (mac) {
        return (
          (parts.includes('ctrl') ? '⌘' : '') +
          (parts.includes('alt') ? '⌥' : '') +
          (parts.includes('shift') ? '⇧' : '') +
          label
        );
      }
      const mods = parts.map((m) => ({ ctrl: 'Ctrl', shift: 'Shift', alt: 'Alt' })[m] ?? m);
      return [...mods, label].join('+');
    })
    .join(' ');
}

/** 文字入力中（textarea・input）か。文字入力に使うキーはショートカットにしない */
export function isTyping(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) return false;
  const tag = target.tagName;
  return tag === 'TEXTAREA' || tag === 'INPUT' || tag === 'SELECT' || target.isContentEditable;
}

/**
 * 文字入力中でも効かせてよいショートカットか。
 * 修飾キー（Ctrl / Alt）付きか、F キー・Escape は通す。それ以外（矢印・Enter・文字）は入力欄のもの
 */
export function allowedWhileTyping(key: string): boolean {
  if (key.includes('ctrl+') || key.includes('alt+')) return true;
  const last = key.split('+').pop()!;
  return /^f\d{1,2}$/.test(last) || last === 'escape';
}
