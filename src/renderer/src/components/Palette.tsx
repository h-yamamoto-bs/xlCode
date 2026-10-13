import clsx from 'clsx';
import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import { fuzzyMatch, splitHighlights } from '../fuzzy';
import { Icon } from './Icons';
import { Kbd } from './ui';

export interface PaletteItem {
  id: string;
  label: string;
  /** 薄く添える説明（フォルダなど） */
  detail?: string;
  /** 右端に出すキー */
  keys?: string;
  icon?: ReactNode;
  run: () => void;
}

export interface PaletteSpec {
  /** 入力欄の説明 */
  placeholder: string;
  /** 最初に入れておく文字（コマンドパレットは ">"） */
  initial?: string;
  /** 通常の一覧（クイックオープンならブックとドキュメント） */
  items: PaletteItem[];
  /** ">" で始めたときの一覧（コマンド）。無ければ切り替えない */
  commands?: PaletteItem[];
  /** 何も入力していないときに出す件数（最近使ったもの順などにしたい場合は items の並び順で決める） */
  maxEmpty?: number;
}

function Highlight({ text, positions }: { text: string; positions: number[] }) {
  const parts = splitHighlights(text, positions);
  return (
    <>
      {parts.map((p, i) =>
        i % 2 === 1 ? (
          <span key={i} className="font-semibold text-link">
            {p}
          </span>
        ) : (
          <span key={i}>{p}</span>
        ),
      )}
    </>
  );
}

/**
 * クイックオープン / コマンドパレット（Ctrl+P / Ctrl+Shift+P）。
 * 入力の先頭が ">" ならコマンドの一覧、それ以外はブック・ドキュメントの一覧を、あいまい検索で絞り込む
 */
export function Palette({ spec, onClose }: { spec: PaletteSpec; onClose: () => void }) {
  const [query, setQuery] = useState(spec.initial ?? '');
  const [cursor, setCursor] = useState(0);
  const input = useRef<HTMLInputElement>(null);
  const list = useRef<HTMLUListElement>(null);

  const commandMode = spec.commands !== undefined && query.startsWith('>');
  const q = commandMode ? query.slice(1) : query;
  const shown = useMemo(() => {
    const items = commandMode ? spec.commands! : spec.items;
    const trimmed = q.trim();
    if (!trimmed) return items.slice(0, spec.maxEmpty ?? 50).map((it) => ({ it, positions: [] as number[] }));
    return items
      .map((it) => ({ it, m: fuzzyMatch(it.detail ? `${it.label} ${it.detail}` : it.label, trimmed) }))
      .filter((x): x is { it: PaletteItem; m: NonNullable<typeof x.m> } => x.m !== null)
      .sort((a, b) => b.m.score - a.m.score)
      .slice(0, 50)
      .map(({ it, m }) => ({ it, positions: m.positions.filter((p) => p < it.label.length) }));
  }, [commandMode, q, spec]);

  useEffect(() => {
    input.current?.focus();
    input.current?.setSelectionRange(input.current.value.length, input.current.value.length);
  }, []);
  useEffect(() => setCursor(0), [query]);
  useEffect(() => {
    list.current?.children[cursor]?.scrollIntoView({ block: 'nearest' });
  }, [cursor]);

  const pick = (i: number) => {
    const row = shown[i];
    if (!row) return;
    onClose();
    row.it.run();
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-start justify-center pt-[8vh]"
      onMouseDown={onClose}
      role="presentation"
    >
      <div
        role="dialog"
        aria-label={commandMode ? 'コマンドパレット' : 'クイックオープン'}
        className="w-[600px] max-w-[92vw] overflow-hidden rounded-md border border-[#454545] bg-widget shadow-[0_0_16px_rgba(0,0,0,0.6)]"
        onMouseDown={(e) => e.stopPropagation()}
      >
        <div className="p-1.5">
          <input
            ref={input}
            aria-label={spec.placeholder}
            className="h-[26px] w-full rounded-[2px] border border-[#3c3c3c] bg-input px-2 text-[13px] text-fg outline-none focus:border-accent"
            placeholder={spec.placeholder}
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'ArrowDown') {
                e.preventDefault();
                setCursor((c) => Math.min(c + 1, shown.length - 1));
              } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                setCursor((c) => Math.max(c - 1, 0));
              } else if (e.key === 'PageDown') {
                e.preventDefault();
                setCursor((c) => Math.min(c + 10, shown.length - 1));
              } else if (e.key === 'PageUp') {
                e.preventDefault();
                setCursor((c) => Math.max(c - 10, 0));
              } else if (e.key === 'Enter') {
                e.preventDefault();
                pick(cursor);
              } else if (e.key === 'Escape') {
                e.preventDefault();
                onClose();
              }
            }}
          />
        </div>
        <ul ref={list} role="listbox" aria-label="候補" className="max-h-[50vh] overflow-auto pb-1">
          {shown.length === 0 && <li className="px-3 py-2 text-[12px] text-muted">一致するものがありません</li>}
          {shown.map(({ it, positions }, i) => (
            <li
              key={it.id}
              role="option"
              aria-selected={i === cursor}
              className={clsx(
                'flex h-[24px] cursor-pointer items-center gap-2 px-3 text-[13px]',
                i === cursor ? 'bg-focus text-fg-strong' : 'hover:bg-hover',
              )}
              onMouseEnter={() => setCursor(i)}
              onClick={() => pick(i)}
            >
              {it.icon && <span className="flex w-4 shrink-0 justify-center">{it.icon}</span>}
              <span className="truncate">
                <Highlight text={it.label} positions={positions} />
              </span>
              {it.detail && <span className="truncate text-[11px] text-faint">{it.detail}</span>}
              {it.keys && (
                <span className="ml-auto shrink-0">
                  <Kbd>{it.keys}</Kbd>
                </span>
              )}
            </li>
          ))}
        </ul>
        {spec.commands !== undefined && !commandMode && (
          <div className="flex items-center gap-1 border-t border-line px-3 py-1 text-[11px] text-faint">
            <Icon.Info size={11} /> 先頭に <Kbd>&gt;</Kbd> を付けるとコマンドを検索します
          </div>
        )}
      </div>
    </div>
  );
}
