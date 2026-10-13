import clsx from 'clsx';
import { useEffect, useState } from 'react';
import type { CommitInfo, FileContents, ProjectMode } from '../../../core';
import type { BookSummary } from '../../../shared/api';
import { api } from '../api';
import { fillPrompt, PROMPT_GROUPS } from '../prompts';
import { STATUS } from '../status';
import { Icon } from './Icons';
import { IconButton } from './ui';

export type SecondaryTab = 'prompts' | 'contents' | 'history';

export const SECONDARY_TABS: { id: SecondaryTab; label: string; title: string }[] = [
  { id: 'prompts', label: '指示', title: 'Excel の Copilot に貼り付ける指示のテンプレート' },
  { id: 'contents', label: '内容', title: '選んでいるファイルの Excel 側・ソース側の内容' },
  { id: 'history', label: '履歴', title: 'このブックのフォルダに関わる Git のコミット' },
];

/**
 * 右のサイドバー（VS Code のセカンダリサイドバー）。
 * ブック画面の作業中に、横に置いておくと便利なものを出す
 */
export function SecondarySidebar({
  tab,
  onTab,
  onClose,
  root,
  book,
  projectMode,
  selectedFile,
  onLog,
}: {
  tab: SecondaryTab;
  onTab: (t: SecondaryTab) => void;
  onClose: () => void;
  root: string | null;
  /** 表示中のブック（無ければ null） */
  book: BookSummary | null;
  projectMode: ProjectMode;
  /** ブック画面で選んでいるファイル */
  selectedFile: string | null;
  onLog: (level: 'info' | 'error', text: string) => void;
}) {
  return (
    <div className="flex h-full min-h-0 flex-col bg-side">
      <div className="flex h-[35px] shrink-0 items-center gap-1 pr-1 pl-3">
        <div role="tablist" aria-label="右のパネル" className="flex min-w-0 flex-1 gap-3">
          {SECONDARY_TABS.map((t) => (
            <button
              key={t.id}
              role="tab"
              aria-selected={tab === t.id}
              title={t.title}
              className={clsx(
                'h-[35px] border-b text-[11px] tracking-wide uppercase',
                tab === t.id ? 'border-fg-strong text-fg-strong' : 'border-transparent text-muted hover:text-fg',
              )}
              onClick={() => onTab(t.id)}
            >
              {t.label}
            </button>
          ))}
        </div>
        <IconButton title="右のパネルを閉じる (Ctrl+Alt+B)" onClick={onClose}>
          <Icon.Close size={14} />
        </IconButton>
      </div>
      <div className="min-h-0 flex-1 overflow-auto">
        {tab === 'prompts' && (
          <PromptPanel book={book} projectMode={projectMode} selectedFile={selectedFile} onLog={onLog} />
        )}
        {tab === 'contents' && <ContentsPanel root={root} book={book} selectedFile={selectedFile} />}
        {tab === 'history' && <HistoryPanel root={root} book={book} />}
      </div>
    </div>
  );
}

async function copyText(text: string): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    // クリップボードが使えない環境では、選択して Ctrl+C してもらう
    return false;
  }
}

function PromptPanel({
  book,
  projectMode,
  selectedFile,
  onLog,
}: {
  book: BookSummary | null;
  projectMode: ProjectMode;
  selectedFile: string | null;
  onLog: (level: 'info' | 'error', text: string) => void;
}) {
  const [copied, setCopied] = useState<string | null>(null);
  const [openId, setOpenId] = useState<string | null>(null);
  const ctx = { file: selectedFile ?? undefined, conflict: book?.conflictSheets?.[0] };
  useEffect(() => {
    if (!copied) return;
    const t = setTimeout(() => setCopied(null), 1500);
    return () => clearTimeout(t);
  }, [copied]);
  return (
    <div className="px-3 py-2 text-[12px]">
      <p className="mb-2 leading-relaxed text-muted">
        Excel の Copilot に貼り付ける指示の例です。「コピー」で写し、必要なところを書き足して使います。
        {selectedFile ? (
          <>
            {' '}
            ファイル名は <span className="font-mono text-fg">{selectedFile}</span> で埋めています。
          </>
        ) : (
          ' ブック画面でファイルの行を選ぶと、ファイル名を埋めます。'
        )}
      </p>
      {PROMPT_GROUPS.map((g) => {
        const items = g.items.filter((it) => !it.mode || it.mode === projectMode);
        if (items.length === 0) return null;
        return (
          <section key={g.title} className="mb-3">
            <div className="mb-1 text-[11px] tracking-wide text-muted uppercase">{g.title}</div>
            <ul className="flex flex-col gap-1">
              {items.map((it) => {
                const text = fillPrompt(it.text, ctx);
                const open = openId === it.id;
                return (
                  <li key={it.id} className="rounded-[3px] border border-line bg-editor">
                    <div className="flex items-center gap-1 px-2 py-1">
                      <button
                        className="min-w-0 flex-1 text-left"
                        onClick={() => setOpenId(open ? null : it.id)}
                        title={it.when}
                        aria-expanded={open}
                      >
                        <span className="text-fg-strong">{it.title}</span>
                        <span className="block truncate text-[11px] text-faint">{it.when}</span>
                      </button>
                      <button
                        className={clsx(
                          'shrink-0 rounded-[3px] border border-line px-2 py-0.5 text-[11px] hover:bg-hover',
                          copied === it.id ? 'text-added' : 'text-fg',
                        )}
                        onClick={async () => {
                          if (await copyText(text)) {
                            setCopied(it.id);
                            onLog('info', `指示をコピーしました: ${it.title}`);
                          } else {
                            setOpenId(it.id);
                            onLog('error', 'クリップボードに書き込めませんでした。本文を選択してコピーしてください');
                          }
                        }}
                        aria-label={`${it.title} をコピー`}
                      >
                        {copied === it.id ? 'コピーしました' : 'コピー'}
                      </button>
                    </div>
                    {open && (
                      <div className="selectable border-t border-line px-2 py-1.5 leading-relaxed whitespace-pre-wrap text-fg">
                        {text}
                      </div>
                    )}
                  </li>
                );
              })}
            </ul>
          </section>
        );
      })}
    </div>
  );
}

function ContentsPanel({
  root,
  book,
  selectedFile,
}: {
  root: string | null;
  book: BookSummary | null;
  selectedFile: string | null;
}) {
  const [side, setSide] = useState<'sheet' | 'source'>('sheet');
  const [data, setData] = useState<FileContents | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const statusKey = JSON.stringify((book?.files ?? []).map((f) => [f.name, f.status]));
  useEffect(() => {
    setData(null);
    setError(null);
    if (!root || !book || !selectedFile) return;
    let stale = false;
    setLoading(true);
    api.fileContents(root, book.rel, selectedFile).then((r) => {
      if (stale) return;
      setLoading(false);
      if (r.ok) setData(r.value);
      else setError(r.error);
    });
    return () => {
      stale = true;
    };
    // statusKey: Build / Sync で中身が変わったら読み直す
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [root, book?.rel, selectedFile, statusKey]);

  if (!book) return <Hint>ブックのタブを表示すると、選んだファイルの内容が出ます。</Hint>;
  if (!selectedFile)
    return <Hint>ブック画面のファイル一覧で行を選ぶと、Excel 側とソース側の内容をそのまま表示します。</Hint>;
  const status = book.files?.find((f) => f.name === selectedFile)?.status;
  const lines = side === 'sheet' ? data?.sheet : data?.source;
  return (
    <div className="flex h-full min-h-0 flex-col">
      <div className="flex shrink-0 flex-wrap items-center gap-2 border-b border-line px-3 py-1.5">
        <span className="min-w-0 flex-1 truncate font-mono text-[12px] text-fg-strong" title={selectedFile}>
          {selectedFile}
        </span>
        {status && <span className={clsx('text-[11px]', STATUS[status].color)}>{STATUS[status].label}</span>}
        <span className="flex rounded-[3px] border border-line text-[11px]">
          {(
            [
              ['sheet', 'Excel 側'],
              ['source', 'ソース側'],
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
      </div>
      <div className="shrink-0 px-3 py-1 text-[11px] text-faint">
        {data &&
          (side === 'sheet'
            ? data.sheet
              ? `${data.sheet.length} 行（シートの A 列）`
              : 'シートがありません'
            : data.source
              ? `${data.source.length} 行（${data.format}）`
              : 'ファイルがありません')}
        {loading && '読み込んでいます...'}
        {error && <span className="text-deleted">{error}</span>}
      </div>
      <pre
        className="selectable m-0 min-h-0 flex-1 overflow-auto px-3 pb-3 font-mono text-[12px] leading-[18px] text-fg"
        aria-label={`${selectedFile} の${side === 'sheet' ? 'Excel 側' : 'ソース側'}の内容`}
      >
        {lines?.map((l, i) => (
          <div key={i} className="flex">
            <span className="w-8 shrink-0 pr-2 text-right text-faint select-none">{i + 1}</span>
            <span className="whitespace-pre">{l}</span>
          </div>
        ))}
      </pre>
    </div>
  );
}

function HistoryPanel({ root, book }: { root: string | null; book: BookSummary | null }) {
  const [commits, setCommits] = useState<CommitInfo[] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [copied, setCopied] = useState<string | null>(null);
  const undoAt = book?.undo?.at;
  useEffect(() => {
    setCommits(null);
    setError(null);
    if (!root || !book) return;
    let stale = false;
    api.recentCommits(root, book.dirRel).then((r) => {
      if (stale) return;
      if (r.ok) setCommits(r.value);
      else setError(r.error);
    });
    return () => {
      stale = true;
    };
    // undoAt: Build / Sync のたびに読み直す
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [root, book?.rel, book?.dirRel, undoAt]);
  if (!book) return <Hint>ブックのタブを表示すると、そのフォルダのコミットの履歴が出ます。</Hint>;
  if (error) return <Hint>{error}</Hint>;
  if (commits === null) return <Hint>読み込んでいます...</Hint>;
  if (commits.length === 0)
    return (
      <Hint>コミットがありません（Git で管理していないか、このフォルダのファイルをまだコミットしていません）。</Hint>
    );
  return (
    <div className="px-3 py-2 text-[12px]">
      <p className="mb-2 leading-relaxed text-muted">
        <span className="font-mono">{book.dirRel || '.'}</span> のファイルに関わるコミット。「xlcode:」は Build / Sync
        前の自動コミットです。前の状態に戻すには、ターミナルで{' '}
        <span className="font-mono">git checkout ハッシュ -- ファイル</span>
        のようにします。
      </p>
      <ul className="flex flex-col gap-1">
        {commits.map((c) => (
          <li key={c.hash} className="flex items-start gap-2 rounded-[3px] border border-line bg-editor px-2 py-1">
            <button
              className="shrink-0 font-mono text-[11px] text-link hover:underline"
              title="ハッシュをコピー"
              onClick={async () => {
                if (await copyText(c.hash)) {
                  setCopied(c.hash);
                  setTimeout(() => setCopied(null), 1500);
                }
              }}
            >
              {copied === c.hash ? 'コピー済' : c.hash}
            </button>
            <span className="min-w-0 flex-1">
              <span className={clsx('block truncate', c.auto ? 'text-muted' : 'text-fg')} title={c.subject}>
                {c.subject}
              </span>
              <span className="block text-[11px] text-faint">
                {new Date(c.date).toLocaleString('ja-JP', { hour12: false })}
              </span>
            </span>
          </li>
        ))}
      </ul>
    </div>
  );
}

function Hint({ children }: { children: React.ReactNode }) {
  return <div className="px-3 py-3 text-[12px] leading-relaxed text-muted">{children}</div>;
}
