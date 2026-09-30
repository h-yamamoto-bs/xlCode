import clsx from 'clsx';
import { useMemo, useState, type ReactNode } from 'react';
import type { BookSummary, BookSync, ProjectInfo } from '../../../shared/api';
import { buildDirTree, countItems, type DirNode } from '../dirTree';
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
  // 開閉状態（キー: dir:パス / book:ブック / empty:パス）。未操作なら既定値
  const [expanded, setExpanded] = useState<Record<string, boolean>>({});
  const toggle = (key: string, def: boolean) => setExpanded((e) => ({ ...e, [key]: !(e[key] ?? def) }));
  const bookTree = useMemo(
    () =>
      buildDirTree(
        project.books,
        (b) => b.dirRel,
        (b) => b.rel.split('/').pop()!,
      ),
    [project.books],
  );
  const dirTree = useMemo(
    () =>
      buildDirTree(
        project.dirsWithoutBook,
        (d) => d,
        (d) => d,
      ),
    [project.dirsWithoutBook],
  );
  const createProps = (d: string) => ({
    title: `${d || 'ルート'} にブックを作成${vba ? '（既存の Excel ツールから取り込み / 空のブック）' : ''}`,
    busy,
  });

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
          <BookDir
            node={bookTree}
            depth={0}
            expanded={expanded}
            toggle={toggle}
            renderBook={(b, depth) => (
              <BookItem
                key={b.rel}
                book={b}
                depth={depth}
                treeVersion={project.treeVersion}
                selected={b.rel === selected}
                open={expanded[`book:${b.rel}`] ?? b.rel === selected}
                onToggle={() => toggle(`book:${b.rel}`, b.rel === selected)}
                onSelect={() => onSelect(b.rel)}
                sync={sync[b.rel]}
              />
            )}
          />
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
        <ul role="tree" aria-label="ブック未作成のディレクトリ" className="max-h-[30vh] overflow-auto pb-1">
          {dirTree.items.length > 0 && (
            <li role="treeitem" aria-label={`${project.name}（ルート）`}>
              <DirRow
                label={`${project.name}（ルート）`}
                depth={0}
                onCreate={() => onCreateBook('')}
                {...createProps('')}
              />
            </li>
          )}
          <EmptyDirs
            node={dirTree}
            depth={0}
            expanded={expanded}
            toggle={toggle}
            render={(n, depth, isOpen, onToggle) => (
              <DirRow
                key={n.path}
                label={n.label}
                depth={depth}
                isOpen={n.dirs.length > 0 ? isOpen : undefined}
                onToggle={onToggle}
                onCreate={n.items.includes(n.path) ? () => onCreateBook(n.path) : undefined}
                {...createProps(n.path)}
              />
            )}
          />
        </ul>
      </Section>
    </div>
  );
}

const INDENT = 12;
const pad = (depth: number) => ({ paddingLeft: 8 + depth * INDENT });

function Chevron({ open }: { open: boolean }) {
  return (
    <span className="flex w-4 shrink-0 justify-center text-muted">
      {open ? <Icon.ChevronDown size={14} /> : <Icon.ChevronRight size={14} />}
    </span>
  );
}

/** ブックの階層。ディレクトリ → その下のディレクトリとブック */
function BookDir({
  node,
  depth,
  expanded,
  toggle,
  renderBook,
}: {
  node: DirNode<BookSummary>;
  depth: number;
  expanded: Record<string, boolean>;
  toggle: (key: string, def: boolean) => void;
  renderBook: (b: BookSummary, depth: number) => ReactNode;
}) {
  return (
    <>
      {node.dirs.map((d) => {
        const key = `dir:${d.path}`;
        const isOpen = expanded[key] ?? true;
        return (
          <li key={key} role="treeitem" aria-label={d.label} aria-expanded={isOpen}>
            <div
              className="flex h-[22px] cursor-pointer items-center gap-1 pr-2 hover:bg-hover"
              style={pad(depth)}
              title={d.path}
              onClick={() => toggle(key, true)}
            >
              <Chevron open={isOpen} />
              <Icon.Folder size={14} className="shrink-0 text-muted" />
              <span className="truncate text-fg">{d.label}</span>
              {!isOpen && <span className="ml-auto shrink-0 text-[11px] text-faint">{countItems(d)}</span>}
            </div>
            {isOpen && (
              <ul role="group">
                <BookDir node={d} depth={depth + 1} expanded={expanded} toggle={toggle} renderBook={renderBook} />
              </ul>
            )}
          </li>
        );
      })}
      {node.items.map((b) => renderBook(b, depth))}
    </>
  );
}

function BookItem({
  book: b,
  depth,
  treeVersion,
  selected,
  open: isOpen,
  onToggle,
  onSelect,
  sync,
}: {
  book: BookSummary;
  depth: number;
  treeVersion: string;
  selected: boolean;
  open: boolean;
  onToggle: () => void;
  onSelect: () => void;
  sync: BookSync | undefined;
}) {
  const badge = bookBadges(b, treeVersion);
  const fileName = b.rel.split('/').pop()!;
  const syncMeta = sync && sync.state !== 'outside' ? SYNC_META[sync.state] : null;
  return (
    <li role="treeitem" aria-label={fileName} aria-expanded={isOpen} aria-selected={selected}>
      <div
        className={clsx(
          'group flex h-[22px] cursor-pointer items-center gap-1 pr-2',
          selected ? 'bg-focus outline outline-1 -outline-offset-1 outline-accent' : 'hover:bg-hover',
        )}
        style={pad(depth)}
        title={b.rel}
        onClick={onSelect}
      >
        <span
          onClick={(e) => {
            e.stopPropagation();
            onToggle();
          }}
        >
          <Chevron open={isOpen} />
        </span>
        <Icon.Book size={15} className="shrink-0 text-excel" />
        <span className="truncate text-fg">{fileName}</span>
        <span className="ml-auto flex shrink-0 items-center gap-1.5 text-[11px]">
          {b.open && <Icon.Lock size={12} className="text-warn" aria-label="開かれています" />}
          {syncMeta?.label && (
            <span title={`OneDrive: ${syncMeta.label}`} className="flex">
              <Icon.Cloud size={13} className={syncMeta.color} aria-label={`OneDrive: ${syncMeta.label}`} />
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
                className="flex h-[22px] items-center gap-1.5 pr-3 hover:bg-hover"
                style={{ paddingLeft: 8 + (depth + 1) * INDENT + 18 }}
                title={m.hint ? `${m.label} — ${m.hint}` : m.label}
              >
                <Icon.File size={14} className="shrink-0 text-muted" />
                <span className={clsx('truncate', f.status === 'clean' ? 'text-fg' : m.color)}>{f.name}</span>
                <span className={clsx('ml-auto w-4 text-right font-mono text-[11px]', m.color)}>{m.letter}</span>
              </li>
            );
          })}
          {b.files && b.files.length === 0 && (
            <li className="text-[12px] text-faint" style={{ paddingLeft: 8 + (depth + 1) * INDENT + 18 }}>
              シートなし
            </li>
          )}
        </ul>
      )}
    </li>
  );
}

/** ブック未作成のディレクトリの階層（既定では閉じておく） */
function EmptyDirs({
  node,
  depth,
  expanded,
  toggle,
  render,
}: {
  node: DirNode<string>;
  depth: number;
  expanded: Record<string, boolean>;
  toggle: (key: string, def: boolean) => void;
  render: (n: DirNode<string>, depth: number, isOpen: boolean, onToggle: () => void) => ReactNode;
}) {
  return (
    <>
      {node.dirs.map((d) => {
        const key = `empty:${d.path}`;
        const isOpen = expanded[key] ?? false;
        return (
          <li key={key} role="treeitem" aria-label={d.label} aria-expanded={d.dirs.length > 0 ? isOpen : undefined}>
            {render(d, depth, isOpen, () => toggle(key, false))}
            {isOpen && d.dirs.length > 0 && (
              <ul role="group">
                <EmptyDirs node={d} depth={depth + 1} expanded={expanded} toggle={toggle} render={render} />
              </ul>
            )}
          </li>
        );
      })}
    </>
  );
}

function DirRow({
  label,
  depth,
  isOpen,
  onToggle,
  onCreate,
  title,
  busy,
}: {
  label: string;
  depth: number;
  /** undefined なら子が無い（開閉しない） */
  isOpen?: boolean;
  onToggle?: () => void;
  onCreate?: () => void;
  title: string;
  busy: boolean;
}) {
  return (
    <div
      className={clsx('group flex h-[22px] items-center gap-1 pr-2 hover:bg-hover', onToggle && 'cursor-pointer')}
      style={pad(depth)}
      onClick={isOpen === undefined ? undefined : onToggle}
    >
      {isOpen === undefined ? <span className="w-4 shrink-0" /> : <Chevron open={isOpen} />}
      <Icon.Folder size={14} className={clsx('shrink-0', onCreate ? 'text-muted' : 'text-faint')} />
      <span className={clsx('truncate', !onCreate && 'text-faint')}>{label}</span>
      {onCreate && (
        <IconButton
          title={title}
          className="ml-auto opacity-0 group-hover:opacity-100"
          onClick={(e) => {
            e.stopPropagation();
            onCreate();
          }}
          disabled={busy}
        >
          <Icon.Plus size={14} />
        </IconButton>
      )}
    </div>
  );
}
