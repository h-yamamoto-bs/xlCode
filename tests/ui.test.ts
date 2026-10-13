import { describe, expect, it } from 'vitest';
import { chordPrefixes, findCommand, type Command } from '../src/renderer/src/commands';
import { fuzzyMatch, splitHighlights } from '../src/renderer/src/fuzzy';
import { allowedWhileTyping, eventToKey, formatKey, normalizeBinding } from '../src/renderer/src/keys';
import {
  activateIndex,
  closeAll,
  closeTab,
  cycleTab,
  deserializeTabs,
  EMPTY_TABS,
  moveTab,
  openTab,
  pinTab,
  pruneTabs,
  reopenClosed,
  serializeTabs,
} from '../src/renderer/src/tabs';

const ev = (
  key: string,
  mods: Partial<{ ctrl: boolean; shift: boolean; alt: boolean; meta: boolean }> = {},
  code?: string,
) => ({
  key,
  code,
  ctrlKey: !!mods.ctrl,
  shiftKey: !!mods.shift,
  altKey: !!mods.alt,
  metaKey: !!mods.meta,
});

describe('タブ', () => {
  it('1 回押し（プレビュー）は 1 つだけで、次のプレビューに置き換わる。固定すると残る', () => {
    let s = openTab(EMPTY_TABS, 'book', 'app/app.xlcode.xlsx', { preview: true });
    s = openTab(s, 'book', 'api/api.xlcode.xlsx', { preview: true });
    expect(s.tabs.map((t) => t.rel)).toEqual(['api/api.xlcode.xlsx']);
    expect(s.active).toBe('book:api/api.xlcode.xlsx');
    s = pinTab(s, s.active!);
    s = openTab(s, 'doc', 'Agents.md', { preview: true });
    expect(s.tabs.map((t) => t.id)).toEqual(['book:api/api.xlcode.xlsx', 'doc:Agents.md']);
    // 固定で開き直すとプレビューが解除される
    s = openTab(s, 'doc', 'Agents.md');
    expect(s.tabs[1].preview).toBe(false);
  });

  it('表示中のタブの右に開き、閉じると右隣（無ければ左隣）を表示する', () => {
    let s = openTab(EMPTY_TABS, 'book', 'a');
    s = openTab(s, 'book', 'c');
    s = openTab(s, 'book', 'a'); // a を表示
    s = openTab(s, 'book', 'b'); // a の右に b
    expect(s.tabs.map((t) => t.rel)).toEqual(['a', 'b', 'c']);
    s = closeTab(s, 'book:b');
    expect(s.active).toBe('book:c');
    s = closeTab(s, 'book:c');
    expect(s.active).toBe('book:a');
    s = closeTab(s, 'book:a');
    expect(s.active).toBeNull();
    expect(s.closed.map((t) => t.rel)).toEqual(['a', 'c', 'b']);
    s = reopenClosed(s);
    expect(s.tabs.map((t) => t.rel)).toEqual(['a']);
    expect(s.closed.map((t) => t.rel)).toEqual(['c', 'b']);
  });

  it('切り替え・番号・並べ替え・すべて閉じる', () => {
    let s = ['a', 'b', 'c'].reduce((acc, r) => openTab(acc, 'book', r), EMPTY_TABS);
    expect(s.active).toBe('book:c');
    s = cycleTab(s, 1);
    expect(s.active).toBe('book:a');
    s = cycleTab(s, -1);
    expect(s.active).toBe('book:c');
    s = activateIndex(s, 1);
    expect(s.active).toBe('book:b');
    s = moveTab(s, 'book:c', 0);
    expect(s.tabs.map((t) => t.rel)).toEqual(['c', 'a', 'b']);
    s = closeAll(s);
    expect(s.tabs).toEqual([]);
    expect(s.closed).toHaveLength(3);
  });

  it('無くなったブックのタブを閉じ、保存の形から復元できる', () => {
    let s = ['a', 'b'].reduce((acc, r) => openTab(acc, 'book', r), EMPTY_TABS);
    s = openTab(s, 'settings', undefined);
    s = pruneTabs(s, (t) => t.kind !== 'book' || t.rel === 'b');
    expect(s.tabs.map((t) => t.id)).toEqual(['book:b', 'settings']);
    const restored = deserializeTabs(JSON.parse(JSON.stringify(serializeTabs(s))));
    expect(restored.tabs.map((t) => t.id)).toEqual(['book:b', 'settings']);
    expect(restored.active).toBe('settings');
    expect(deserializeTabs(null)).toEqual(EMPTY_TABS);
    expect(deserializeTabs({ tabs: [{ id: 'x', kind: 'bogus' }], active: 'x' }).tabs).toEqual([]);
  });
});

describe('ショートカット', () => {
  it('キー入力を VS Code の書き方にする（Shift+数字は物理キーで判定）', () => {
    expect(eventToKey(ev('b', { ctrl: true, shift: true }))).toBe('ctrl+shift+b');
    expect(eventToKey(ev('!', { alt: true }, 'Digit1'))).toBe('alt+1');
    expect(eventToKey(ev('Control', { ctrl: true }))).toBeNull();
    expect(eventToKey(ev('ArrowDown'))).toBe('down');
    expect(eventToKey(ev(',', { ctrl: true }))).toBe('ctrl+,');
    // macOS では Cmd を ctrl として扱う
    expect(eventToKey(ev('p', { meta: true }), 'darwin')).toBe('ctrl+p');
    expect(eventToKey(ev('p', { ctrl: true }), 'darwin')).toBe('p');
  });

  it('割り当ての正規化と表示', () => {
    expect(normalizeBinding('Shift+Ctrl+B')).toBe('ctrl+shift+b');
    expect(normalizeBinding('ctrl+k  ctrl+s')).toBe('ctrl+k ctrl+s');
    expect(formatKey('ctrl+shift+b')).toBe('Ctrl+Shift+B');
    expect(formatKey('ctrl+k ctrl+s')).toBe('Ctrl+K Ctrl+S');
    expect(formatKey('ctrl+pagedown')).toBe('Ctrl+PageDown');
    expect(formatKey('ctrl+shift+b', 'darwin')).toBe('⌘⇧B');
    expect(formatKey('f1')).toBe('F1');
  });

  it('文字入力中は Ctrl / Alt 付きと F キーだけ効く', () => {
    expect(allowedWhileTyping('ctrl+s')).toBe(true);
    expect(allowedWhileTyping('f1')).toBe(true);
    expect(allowedWhileTyping('escape')).toBe(true);
    expect(allowedWhileTyping('down')).toBe(false);
    expect(allowedWhileTyping('enter')).toBe(false);
  });

  it('コマンドの検索と 2 段階のキー', () => {
    const run = () => {};
    const commands: Command[] = [
      { id: 'a', title: 'A', keys: ['ctrl+shift+b'], run },
      { id: 'b', title: 'B', keys: ['ctrl+k ctrl+s'], run },
      { id: 'c', title: 'C', keys: ['ctrl+w'], when: false, run },
    ];
    expect(findCommand(commands, 'Ctrl+Shift+B')?.id).toBe('a');
    expect(findCommand(commands, 'ctrl+k ctrl+s')?.id).toBe('b');
    expect(findCommand(commands, 'ctrl+w')).toBeUndefined();
    expect([...chordPrefixes(commands)]).toEqual(['ctrl+k']);
  });
});

describe('あいまい検索', () => {
  it('部分一致を優先し、順番どおりの一致も拾う', () => {
    expect(fuzzyMatch('app/app.xlcode.xlsx', 'app')!.positions).toEqual([0, 1, 2]);
    expect(fuzzyMatch('Agents.md', 'agm')).not.toBeNull();
    expect(fuzzyMatch('Agents.md', 'zzz')).toBeNull();
    expect(fuzzyMatch('ブック: Build（Excel → ソース）', 'build excel')).not.toBeNull();
    expect(fuzzyMatch('anything', '')!.score).toBe(0);
    const a = fuzzyMatch('LocalAgents.md', 'local')!.score;
    const b = fuzzyMatch('xlocalx', 'local')!.score;
    expect(a).toBeGreaterThan(b);
  });

  it('強調表示の分割は通常から始まり、交互に並ぶ', () => {
    expect(splitHighlights('abcde', [1, 2])).toEqual(['a', 'bc', 'de']);
    expect(splitHighlights('abc', [0])).toEqual(['', 'a', 'bc']);
    expect(splitHighlights('abc', [])).toEqual(['abc']);
  });
});
