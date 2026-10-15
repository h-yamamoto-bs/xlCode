import clsx from 'clsx';
import { useEffect, useMemo, useRef, useState } from 'react';
import type { FileContents } from '../../../core';
import type { BookSummary, OpenWith } from '../../../shared/api';
import { api } from '../api';
import { findMatches, highlight } from '../find';
import { STATUS } from '../status';
import { Icon } from './Icons';
import { Button, IconButton } from './ui';

/**
 * コードのファイル（シート）の閲覧。編集はしない（Markdown 以外は既定のアプリか VS Code で開く）。
 * Excel 側（シートの A 列）とソース側（ファイル）を切り替えて、整形せずにそのまま表示する
 */
export function FileView({
  root,
  book,
  name,
  onOpenWith,
  onReveal,
  findNonce = 0,
  onSelectForPrompt,
}: {
  root: string;
  book: BookSummary;
  name: string;
  onOpenWith: (app: OpenWith) => void;
  onReveal: () => void;
  /** 増えるたびに検索欄を開いて入力に移る（Ctrl+F） */
  findNonce?: number;
  /** 右のパネルの「指示」にこのファイル名を使う */
  onSelectForPrompt?: () => void;
}) {
  const [side, setSide] = useState<'source' | 'sheet'>('source');
  const [data, setData] = useState<FileContents | null>(null);
  const [error, setError] = useState<string | null>(null);
  const status = book.files?.find((f) => f.name === name)?.status;
  const statusKey = JSON.stringify((book.files ?? []).map((f) => [f.name, f.status]));
  useEffect(() => {
    let stale = false;
    setError(null);
    api.fileContents(root, book.rel, name).then((r) => {
      if (stale) return;
      if (r.ok) setData(r.value);
      else setError(r.error);
    });
    return () => {
      stale = true;
    };
    // statusKey: Build / Sync で中身が変わったら読み直す
  }, [root, book.rel, name, statusKey]);
  // ソース側が無ければ Excel 側を出す
  useEffect(() => {
    if (data && data.source === null && data.sheet !== null) setSide('sheet');
  }, [data]);
  // 開いたファイルを右のパネルの「指示」にも使う（開いたときに 1 回だけ）
  const selectRef = useRef(onSelectForPrompt);
  selectRef.current = onSelectForPrompt;
  useEffect(() => {
    selectRef.current?.();
  }, [name, book.rel]);

  const lines = side === 'sheet' ? data?.sheet : data?.source;
  const rel = book.dirRel ? `${book.dirRel}/${name}` : name;

  // ---- ファイル内の検索（Ctrl+F） ----
  const [findOpen, setFindOpen] = useState(false);
  const [query, setQuery] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [cursor, setCursor] = useState(0);
  const findInput = useRef<HTMLInputElement>(null);
  const matches = useMemo(() => findMatches(lines ?? [], query, matchCase), [lines, query, matchCase]);
  const current = matches.length === 0 ? -1 : ((cursor % matches.length) + matches.length) % matches.length;
  // 検索欄を開いて入力に移る。入力欄は開いた描画で現れるので、その描画のあとで focus する
  const [wantFocus, setWantFocus] = useState(0);
  const openFind = () => {
    setFindOpen(true);
    setWantFocus((n) => n + 1);
  };
  useEffect(() => {
    if (findNonce > 0) openFind();
    // findNonce が増えたときだけ
  }, [findNonce]);
  useEffect(() => {
    if (wantFocus === 0) return;
    findInput.current?.focus();
    findInput.current?.select();
  }, [wantFocus]);
  // 検索語・対象が変わったら先頭から
  useEffect(() => setCursor(0), [query, matchCase, side]);
  // 今の該当箇所を見える位置に
  const currentRef = useRef<HTMLElement>(null);
  useEffect(() => {
    currentRef.current?.scrollIntoView({ block: 'center' });
  }, [current, matches]);
  const closeFind = () => {
    setFindOpen(false);
    findInput.current?.blur();
  };
  const step = (d: number) => setCursor(matches.length === 0 ? 0 : (current + d + matches.length) % matches.length);
  // 行ごとの該当箇所
  const byLine = useMemo(() => {
    const m = new Map<number, { start: number; end: number; index: number }[]>();
    matches.forEach((x, index) => {
      const arr = m.get(x.line) ?? [];
      arr.push({ start: x.start, end: x.end, index });
      m.set(x.line, arr);
    });
    return m;
  }, [matches]);
  return (
    <div className="flex h-full min-h-0 flex-col">
      <div className="flex h-[35px] shrink-0 items-center gap-3 border-b border-line px-4">
        <Icon.File size={14} className="shrink-0 text-muted" />
        <span className="min-w-0 truncate font-mono text-[12.5px]" title={rel}>
          {rel}
        </span>
        {status && <span className={clsx('shrink-0 text-[12px]', STATUS[status].color)}>{STATUS[status].label}</span>}
        <span className="ml-auto flex shrink-0 gap-1.5">
          <Button onClick={openFind} title="ファイル内の文字を検索 (Ctrl+F)" aria-label="検索">
            <Icon.Search size={14} /> 検索
          </Button>
          <Button
            onClick={() => onOpenWith('default')}
            disabled={data?.source === null}
            title="Windows に関連付けられたアプリ（メモ帳など）で開く"
          >
            <Icon.Reveal size={14} /> 既定のアプリで開く
          </Button>
          <Button
            onClick={() => onOpenWith('vscode')}
            disabled={data?.source === null}
            title="PATH の code コマンドで開く"
          >
            VS Code で開く
          </Button>
          <Button onClick={onReveal} title="エクスプローラーでファイルを表示">
            <Icon.Folder size={14} /> フォルダ
          </Button>
        </span>
      </div>
      <div className="flex shrink-0 flex-wrap items-center gap-x-3 gap-y-1 border-b border-line bg-side px-4 py-1.5 text-[12px] text-muted">
        <span className="flex shrink-0 rounded-[3px] border border-line text-[11px]">
          {(
            [
              ['source', 'ソース側'],
              ['sheet', 'Excel 側'],
            ] as const
          ).map(([k, l]) => (
            <button
              key={k}
              role="tab"
              aria-selected={side === k}
              className={clsx('px-2 py-0.5', side === k ? 'bg-focus text-fg-strong' : 'text-muted hover:text-fg')}
              onClick={() => setSide(k)}
            >
              {l}
            </button>
          ))}
        </span>
        <span className="shrink-0 text-[11px] text-faint">
          {data &&
            (side === 'sheet'
              ? data.sheet
                ? `${data.sheet.length} 行（シートの A 列）`
                : 'シートがありません'
              : data.source
                ? `${data.source.length} 行（${data.format}）`
                : 'ファイルがありません')}
        </span>
        <span className="min-w-0 flex-1 basis-[280px]">
          閲覧だけです。直すときは「既定のアプリで開く」か「VS Code で開く」でエディタを開き、保存したら Sync
          します。Excel 側は Copilot が書いた内容（Build 前のシート）です。
        </span>
      </div>
      {findOpen && (
        <div
          className="flex shrink-0 items-center gap-1.5 border-b border-line bg-side px-3 py-1 text-[12px]"
          role="search"
          aria-label="ファイル内を検索"
        >
          <Icon.Search size={13} className="shrink-0 text-muted" />
          <input
            ref={findInput}
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            onKeyDown={(e) => {
              if (e.nativeEvent.isComposing) return;
              if (e.key === 'Enter') {
                e.preventDefault();
                step(e.shiftKey ? -1 : 1);
              } else if (e.key === 'Escape') {
                e.preventDefault();
                closeFind();
              }
            }}
            placeholder="検索（Enter で次、Shift+Enter で前）"
            aria-label="検索する文字"
            spellCheck={false}
            className="h-6 w-[260px] max-w-[50vw] rounded-[3px] border border-field-border bg-input px-2 font-mono text-[12px] text-fg outline-none focus:border-accent"
          />
          <button
            className={clsx(
              'h-6 rounded-[3px] border px-1.5 font-mono text-[11px]',
              matchCase ? 'border-accent bg-focus text-fg-strong' : 'border-transparent text-muted hover:bg-hover',
            )}
            aria-pressed={matchCase}
            aria-label="大文字と小文字を区別する"
            title="大文字と小文字を区別する"
            onClick={() => setMatchCase((v) => !v)}
          >
            Aa
          </button>
          <span className="min-w-[72px] text-[11px] text-muted" aria-live="polite" data-testid="find-count">
            {query === '' ? '' : matches.length === 0 ? '該当なし' : `${current + 1} / ${matches.length}`}
          </span>
          <IconButton title="前へ (Shift+Enter)" onClick={() => step(-1)} disabled={matches.length === 0}>
            <Icon.ChevronDown size={14} className="rotate-180" />
          </IconButton>
          <IconButton title="次へ (Enter)" onClick={() => step(1)} disabled={matches.length === 0}>
            <Icon.ChevronDown size={14} />
          </IconButton>
          <IconButton title="閉じる (Escape)" onClick={closeFind}>
            <Icon.Close size={14} />
          </IconButton>
        </div>
      )}
      {error && <div className="px-4 py-2 text-[12px] text-deleted">{error}</div>}
      <pre
        className="selectable m-0 min-h-0 flex-1 overflow-auto px-2 py-2 font-mono text-[13px] leading-[19px] text-fg"
        aria-label={`${name} の${side === 'sheet' ? 'Excel 側' : 'ソース側'}の内容`}
      >
        {lines?.map((l, i) => (
          <div key={i} className="flex hover:bg-hover">
            <span className="w-10 shrink-0 pr-3 text-right text-faint select-none">{i + 1}</span>
            <span className="whitespace-pre">
              {findOpen && byLine.has(i)
                ? highlight(l, byLine.get(i)!).map((seg, k) =>
                    seg.match === null ? (
                      seg.text
                    ) : (
                      <mark
                        key={k}
                        ref={seg.match === current ? currentRef : undefined}
                        className={clsx(
                          'rounded-[2px]',
                          seg.match === current ? 'bg-accent text-white' : 'bg-mark text-fg-strong',
                        )}
                        data-current={seg.match === current || undefined}
                      >
                        {seg.text}
                      </mark>
                    ),
                  )
                : l}
            </span>
          </div>
        ))}
        {data && lines === null && <div className="px-3 text-muted">（ありません）</div>}
      </pre>
    </div>
  );
}
