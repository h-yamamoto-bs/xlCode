import clsx from 'clsx';
import { useState } from 'react';
import type { BookSummary, BookSync, ProjectInfo } from '../../../shared/api';
import { EXCEL_SIDE, SOURCE_SIDE, STATUS, SYNC_META } from '../status';
import { Icon } from './Icons';
import { IconButton, Section } from './ui';

function bookBadges(b: BookSummary, treeVersion: string) {
  const files = b.files ?? [];
  return {
    excel: files.filter((f) => EXCEL_SIDE.includes(f.status)).length,
    source: files.filter((f) => SOURCE_SIDE.includes(f.status)).length,
    conflict: files.filter((f) => f.status === 'conflict').length + (b.conflictSheets?.length ?? 0),
    problems: (b.errors?.length ?? 0) + (b.loadError ? 1 : 0),
    treeStale: b.treeVersion !== treeVersion,
  };
}

export function Sidebar({
  project,
  selected,
  busy,
  onSelect,
  sync,
  onRefreshTree,
  onReload,
  onCreateBook,
  onCreateAll,
  vba = false,
}: {
  /** ブックの無いディレクトリすべてにブックを作る（ソースコードモード） */
  onCreateAll?: () => void;
  /** VBA モード（ブックは既存の Excel ツールから取り込むか、空で作る） */
  vba?: boolean;
  project: ProjectInfo;
  selected: string | null;
  busy: boolean;
  onSelect: (rel: string) => void;
  sync: Record<string, BookSync>;
  onRefreshTree: () => void;
  onReload: () => void;
  onCreateBook: (dirRel: string) => void;
}) {
  const [openBooks, setOpenBooks] = useState(true);
  // ブックが無いときは、作成先を選べるよう最初から開いておく
  const [openDirs, setOpenDirs] = useState(project.books.length === 0);
  const [expanded, setExpanded] = useState<Record<string, boolean>>({});

  return (
    <div className="flex h-full flex-col bg-side">
      <div className="flex h-[35px] shrink-0 items-center px-5 text-[11px] tracking-wide text-muted uppercase">
        エクスプローラー
      </div>
      <Section
        title={`ブック — ${project.name}`}
        open={openBooks}
        onToggle={() => setOpenBooks(!openBooks)}
        grow
        actions={
          <>
            <IconButton title="Refresh Tree（全ブックの #tree を更新）" onClick={onRefreshTree} disabled={busy}>
              <Icon.Tree />
            </IconButton>
            <IconButton title="再読み込み" onClick={onReload} disabled={busy}>
              <Icon.Refresh />
            </IconButton>
          </>
        }
      >
        {project.books.length === 0 && (
          <div className="px-5 py-2 text-[12px] leading-relaxed text-muted">
            ブックがありません。下の「ブック未作成のディレクトリ」の ＋ から作成してください。
            {vba && '既存の Excel ツール（.xlsm など）から取り込むこともできます。'}
          </div>
        )}
        <ul role="tree" aria-label="ブック">
          {project.books.map((b) => {
            const badge = bookBadges(b, project.treeVersion);
            const isOpen = expanded[b.rel] ?? b.rel === selected;
            const fileName = b.rel.split('/').pop()!;
            const changed = (b.files ?? []).filter((f) => f.status !== 'clean');
            return (
              <li key={b.rel} role="treeitem" aria-expanded={isOpen} aria-selected={b.rel === selected}>
                <div
                  className={clsx(
                    'group flex h-[22px] cursor-pointer items-center gap-1 pr-2 pl-2',
                    b.rel === selected
                      ? 'bg-focus outline outline-1 -outline-offset-1 outline-accent'
                      : 'hover:bg-hover',
                  )}
                  onClick={() => onSelect(b.rel)}
                >
                  <span
                    className="flex w-4 justify-center text-muted"
                    onClick={(e) => {
                      e.stopPropagation();
                      setExpanded({ ...expanded, [b.rel]: !isOpen });
                    }}
                  >
                    {isOpen ? <Icon.ChevronDown size={14} /> : <Icon.ChevronRight size={14} />}
                  </span>
                  <Icon.Book size={15} className="shrink-0 text-excel" />
                  <span className="truncate text-fg">{fileName}</span>
                  <span className="truncate text-[11px] text-faint">{b.dirRel || '.'}</span>
                  <span className="ml-auto flex shrink-0 items-center gap-1.5 text-[11px]">
                    {b.open && <Icon.Lock size={12} className="text-warn" aria-label="開かれています" />}
                    {sync[b.rel] && SYNC_META[sync[b.rel].state].label && sync[b.rel].state !== 'outside' && (
                      <span title={`OneDrive: ${SYNC_META[sync[b.rel].state].label}`} className="flex">
                        <Icon.Cloud
                          size={13}
                          className={SYNC_META[sync[b.rel].state].color}
                          aria-label={`OneDrive: ${SYNC_META[sync[b.rel].state].label}`}
                        />
                      </span>
                    )}
                    {badge.problems > 0 && <Icon.Error size={12} className="text-deleted" aria-label="エラー" />}
                    {badge.conflict > 0 && (
                      <span className="text-conflict" title={`両側で変更: ${badge.conflict}（Build で衝突シート作成）`}>
                        C{badge.conflict}
                      </span>
                    )}
                    {badge.excel > 0 && (
                      <span className="text-modified" title={`Excel側で編集中: ${badge.excel}（Build で反映）`}>
                        E{badge.excel}
                      </span>
                    )}
                    {badge.source > 0 && (
                      <span className="text-info" title={`エディタ側で変更: ${badge.source}（Sync で反映）`}>
                        S{badge.source}
                      </span>
                    )}
                  </span>
                </div>
                {isOpen && (
                  <ul role="group">
                    {(b.files ?? []).map((f) => {
                      const m = STATUS[f.status];
                      return (
                        <li
                          key={f.name}
                          role="treeitem"
                          className="flex h-[22px] items-center gap-1.5 pr-3 pl-[38px] hover:bg-hover"
                          title={m.hint ? `${m.label} — ${m.hint}` : m.label}
                        >
                          <Icon.File size={14} className="shrink-0 text-muted" />
                          <span className={clsx('truncate', f.status === 'clean' ? 'text-fg' : m.color)}>{f.name}</span>
                          <span className={clsx('ml-auto w-4 text-right font-mono text-[11px]', m.color)}>
                            {m.letter}
                          </span>
                        </li>
                      );
                    })}
                    {b.files && changed.length === 0 && b.files.length === 0 && (
                      <li className="pl-[38px] text-[12px] text-faint">シートなし</li>
                    )}
                  </ul>
                )}
              </li>
            );
          })}
        </ul>
      </Section>
      <Section
        title="ブック未作成のディレクトリ"
        open={openDirs}
        onToggle={() => setOpenDirs(!openDirs)}
        actions={
          onCreateAll &&
          project.dirsWithoutBook.length > 1 && (
            <IconButton title="すべてのディレクトリにブックを作成" onClick={onCreateAll} disabled={busy}>
              <Icon.Plus size={14} />
            </IconButton>
          )
        }
      >
        <ul className="max-h-[30vh] overflow-auto pb-1">
          {project.dirsWithoutBook.map((d) => (
            <li key={d} className="group flex h-[22px] items-center gap-1.5 pr-2 pl-5 hover:bg-hover">
              <Icon.Folder size={14} className="shrink-0 text-muted" />
              <span className="truncate">{d || `${project.name}（ルート）`}</span>
              <IconButton
                title={`${d || 'ルート'} にブックを作成${vba ? '（既存の Excel ツールから取り込み / 空のブック）' : ''}`}
                className="ml-auto opacity-0 group-hover:opacity-100"
                onClick={() => onCreateBook(d)}
                disabled={busy}
              >
                <Icon.Plus size={14} />
              </IconButton>
            </li>
          ))}
        </ul>
      </Section>
    </div>
  );
}
