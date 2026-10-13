import clsx from 'clsx';
import { useEffect, useRef, useState } from 'react';
import type { FileContents } from '../../../core';
import type { BookSummary, OpenWith } from '../../../shared/api';
import { api } from '../api';
import { STATUS } from '../status';
import { Icon } from './Icons';
import { Button } from './ui';

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
  onSelectForPrompt,
}: {
  root: string;
  book: BookSummary;
  name: string;
  onOpenWith: (app: OpenWith) => void;
  onReveal: () => void;
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
  return (
    <div className="flex h-full min-h-0 flex-col">
      <div className="flex h-[35px] shrink-0 items-center gap-3 border-b border-line px-4">
        <Icon.File size={14} className="shrink-0 text-muted" />
        <span className="min-w-0 truncate font-mono text-[12.5px]" title={rel}>
          {rel}
        </span>
        {status && <span className={clsx('shrink-0 text-[12px]', STATUS[status].color)}>{STATUS[status].label}</span>}
        <span className="ml-auto flex shrink-0 gap-1.5">
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
      {error && <div className="px-4 py-2 text-[12px] text-deleted">{error}</div>}
      <pre
        className="selectable m-0 min-h-0 flex-1 overflow-auto px-2 py-2 font-mono text-[13px] leading-[19px] text-fg"
        aria-label={`${name} の${side === 'sheet' ? 'Excel 側' : 'ソース側'}の内容`}
      >
        {lines?.map((l, i) => (
          <div key={i} className="flex hover:bg-hover">
            <span className="w-10 shrink-0 pr-3 text-right text-faint select-none">{i + 1}</span>
            <span className="whitespace-pre">{l}</span>
          </div>
        ))}
        {data && lines === null && <div className="px-3 text-muted">（ありません）</div>}
      </pre>
    </div>
  );
}
