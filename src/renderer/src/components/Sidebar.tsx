import clsx from 'clsx';
import { useEffect, useMemo, useRef, useState, type KeyboardEvent, type ReactNode } from 'react';
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

/** 階層に並べるもの: ブック、またはドキュメント（Agents.md / LocalAgents.md / .xlcodeignore） */
type Item = { kind: 'book'; book: BookSummary } | { kind: 'doc'; rel: string; label: string; exists: boolean };

/** 画面に描く 1 行（階層を平らにしたもの。キーボードで上下に動くため） */
interface Row {
  key: string;
  depth: number;
  parent: string | null;
  /** 開閉できる（フォルダ・ブック） */
  collapsible: boolean;
  open: boolean;
  node:
    | { type: 'dir'; label: string; path: string; count: number }
    | { type: 'book'; book: BookSummary }
    | { type: 'sheet'; book: BookSummary; file: NonNullable<BookSummary['files']>[number] }
    | { type: 'doc'; rel: string; label: string; exists: boolean }
    | { type: 'emptydir'; label: string; path: string; canCreate: boolean };
}

export type OpenTarget = { kind: 'book' | 'doc'; rel: string };

export function Sidebar({
  project,
  selected,
  busy,
  onOpen,
  sync,
  onRefreshTree,
  onReload,
  onCreateBook,
  onCreateAll,
  onSettings,
  focusNonce = 0,
  vba = false,
}: {
  /** ブックの無いディレクトリすべてにブックを作る（ソースコードモード） */
  onCreateAll?: () => void;
  /** VBA モード（ブックは既存の Excel ツールから取り込むか、空で作る） */
  vba?: boolean;
  project: ProjectInfo;
  /** 表示中のタブ（book:rel / doc:rel） */
  selected: string | null;
  busy: boolean;
  /** ブックやドキュメントを開く。preview は 1 回押し（プレビュー）、固定は 2 回押しか Enter */
  onOpen: (target: OpenTarget, opts: { preview: boolean }) => void;
  sync: Record<string, BookSync>;
  onRefreshTree: () => void;
  onReload: () => void;
  onCreateBook: (dirRel: string) => void;
  /** ブックの置き場所の設定を開く */
  onSettings: () => void;
  /** 値が変わるたびに一覧へフォーカスを移す（Ctrl+Shift+E） */
  focusNonce?: number;
}) {
  const [openBooks, setOpenBooks] = useState(true);
  // ブックが無いときは、作成先を選べるよう最初から開いておく
  const [openDirs, setOpenDirs] = useState(project.books.length === 0);
  // 開閉状態（キー: dir:パス / book:ブック / empty:パス）。未操作なら既定値
  const [expanded, setExpanded] = useState<Record<string, boolean>>({});
  const toggle = (key: string, def: boolean) => setExpanded((e) => ({ ...e, [key]: !(e[key] ?? def) }));
  const setOpen = (key: string, open: boolean) => setExpanded((e) => ({ ...e, [key]: open }));

  const items = useMemo<Item[]>(() => {
    const docs = new Set(project.docs);
    const out: Item[] = [
      { kind: 'doc', rel: 'Agents.md', label: 'Agents.md', exists: docs.has('Agents.md') },
      { kind: 'doc', rel: '.xlcodeignore', label: '.xlcodeignore', exists: docs.has('.xlcodeignore') },
    ];
    for (const b of project.books) {
      out.push({ kind: 'book', book: b });
      const rel = b.dirRel ? `${b.dirRel}/LocalAgents.md` : 'LocalAgents.md';
      out.push({ kind: 'doc', rel, label: 'LocalAgents.md', exists: docs.has(rel) });
    }
    return out;
  }, [project.books, project.docs]);

  const bookTree = useMemo(
    () =>
      buildDirTree(
        items,
        (it) =>
          it.kind === 'book' ? it.book.dirRel : it.rel.includes('/') ? it.rel.slice(0, it.rel.lastIndexOf('/')) : '',
        // ブックを先に、ドキュメントを後に
        (it) => (it.kind === 'book' ? `0:${it.book.rel.split('/').pop()}` : `1:${it.label}`),
      ),
    [items],
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

  // 階層を平らにする
  const bookRows = useMemo(() => {
    const rows: Row[] = [];
    const walk = (node: DirNode<Item>, depth: number, parent: string | null) => {
      for (const d of node.dirs) {
        const key = `dir:${d.path}`;
        const open = expanded[key] ?? true;
        rows.push({
          key,
          depth,
          parent,
          collapsible: true,
          open,
          node: { type: 'dir', label: d.label, path: d.path, count: countItems(d) },
        });
        if (open) walk(d, depth + 1, key);
      }
      for (const it of node.items) {
        if (it.kind === 'book') {
          const key = `book:${it.book.rel}`;
          const open = expanded[key] ?? key === selected;
          rows.push({ key, depth, parent, collapsible: true, open, node: { type: 'book', book: it.book } });
          if (open) {
            for (const f of it.book.files ?? []) {
              rows.push({
                key: `sheet:${it.book.rel}:${f.name}`,
                depth: depth + 1,
                parent: key,
                collapsible: false,
                open: false,
                node: { type: 'sheet', book: it.book, file: f },
              });
            }
          }
        } else {
          rows.push({
            key: `doc:${it.rel}`,
            depth,
            parent,
            collapsible: false,
            open: false,
            node: { type: 'doc', rel: it.rel, label: it.label, exists: it.exists },
          });
        }
      }
    };
    walk(bookTree, 0, null);
    return rows;
  }, [bookTree, expanded, selected]);

  const dirRows = useMemo(() => {
    const rows: Row[] = [];
    if (dirTree.items.length > 0) {
      rows.push({
        key: 'empty:',
        depth: 0,
        parent: null,
        collapsible: false,
        open: false,
        node: { type: 'emptydir', label: `${project.name}（ルート）`, path: '', canCreate: true },
      });
    }
    const walk = (node: DirNode<string>, depth: number, parent: string | null) => {
      for (const d of node.dirs) {
        const key = `empty:${d.path}`;
        const open = expanded[key] ?? false;
        rows.push({
          key,
          depth,
          parent,
          collapsible: d.dirs.length > 0,
          open,
          node: { type: 'emptydir', label: d.label, path: d.path, canCreate: d.items.includes(d.path) },
        });
        if (open && d.dirs.length > 0) walk(d, depth + 1, key);
      }
    };
    walk(dirTree, 0, null);
    return rows;
  }, [dirTree, expanded, project.name]);

  const createProps = (d: string) => ({
    title: `${d || 'ルート'} にブックを作成${vba ? '（既存の Excel ツールから取り込み / 空のブック）' : ''}`,
    busy,
  });

  /** 行の既定の操作（Enter / クリック） */
  const activate = (row: Row, preview: boolean) => {
    switch (row.node.type) {
      case 'dir':
        toggle(row.key, true);
        break;
      case 'book':
        onOpen({ kind: 'book', rel: row.node.book.rel }, { preview });
        break;
      case 'sheet':
        onOpen({ kind: 'book', rel: row.node.book.rel }, { preview });
        break;
      case 'doc':
        onOpen({ kind: 'doc', rel: row.node.rel }, { preview });
        break;
      case 'emptydir':
        // 1 回押し（preview）は開閉だけ。Enter（固定）で作成
        if (row.collapsible) toggle(row.key, false);
        else if (!preview && row.node.canCreate) onCreateBook(row.node.path);
        break;
    }
  };

  const booksTree = useRef<HTMLUListElement>(null);
  useEffect(() => {
    if (focusNonce > 0) booksTree.current?.focus();
  }, [focusNonce]);

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
            <IconButton title="再読み込み (F5)" onClick={onReload} disabled={busy}>
              <Icon.Refresh />
            </IconButton>
          </>
        }
      >
        <button
          className="flex h-[22px] w-full items-center gap-1.5 px-5 text-left text-[11px] text-faint hover:bg-hover hover:text-fg"
          title={`ブックの置き場所: ${project.bookRoot ?? `${project.root}（ソースの中）`}\nクリックで設定を開き、置き場所を変えられます`}
          onClick={onSettings}
        >
          <Icon.Cloud size={12} className={clsx('shrink-0', project.bookRoot ? 'text-excel' : 'text-faint')} />
          <span className="truncate">
            {project.bookRoot ? `置き場所: ${project.bookRoot}` : '置き場所: ソースの中'}
          </span>
          <Icon.Gear size={11} className="ml-auto shrink-0" />
        </button>
        {project.books.length === 0 && (
          <div className="px-5 py-2 text-[12px] leading-relaxed text-muted">
            ブックがありません。下の「ブック未作成のディレクトリ」の ＋ から作成してください。
            {vba && '既存の Excel ツール（.xlsm など）から取り込むこともできます。'}
          </div>
        )}
        <Tree
          ref={booksTree}
          label="ブック"
          rows={bookRows}
          selected={selected}
          setOpen={setOpen}
          activate={activate}
          render={(row) => (
            <TreeRow
              row={row}
              treeVersion={project.treeVersion}
              selected={row.key === selected}
              sync={row.node.type === 'book' ? sync[row.node.book.rel] : undefined}
              onToggle={() => toggle(row.key, row.node.type === 'dir')}
            />
          )}
        />
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
        <Tree
          label="ブック未作成のディレクトリ"
          rows={dirRows}
          selected={null}
          setOpen={setOpen}
          activate={activate}
          className="max-h-[30vh] pb-1"
          render={(row) =>
            row.node.type === 'emptydir' && (
              <DirRow
                label={row.node.label}
                depth={row.depth}
                isOpen={row.collapsible ? row.open : undefined}
                onToggle={() => toggle(row.key, false)}
                onCreate={row.node.canCreate ? () => onCreateBook((row.node as { path: string }).path) : undefined}
                {...createProps(row.node.path)}
              />
            )
          }
        />
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

/**
 * 階層の一覧。キーボードで動ける（VS Code のエクスプローラーと同じ）:
 * ↑↓ で移動、→ で開く・中へ、← で閉じる・親へ、Enter で開く（固定）、Space でプレビュー、Home / End
 */
function Tree({
  ref,
  label,
  rows,
  selected,
  setOpen,
  activate,
  render,
  className,
}: {
  ref?: React.Ref<HTMLUListElement>;
  label: string;
  rows: Row[];
  selected: string | null;
  setOpen: (key: string, open: boolean) => void;
  activate: (row: Row, preview: boolean) => void;
  render: (row: Row) => ReactNode;
  className?: string;
}) {
  const [focusKey, setFocusKey] = useState<string | null>(null);
  const [focused, setFocused] = useState(false);
  const list = useRef<HTMLUListElement | null>(null);
  const cur = rows.findIndex((r) => r.key === (focusKey ?? selected));
  const move = (i: number) => {
    const row = rows[Math.max(0, Math.min(i, rows.length - 1))];
    if (!row) return;
    setFocusKey(row.key);
    list.current
      ?.querySelector<HTMLElement>(`[data-key="${CSS.escape(row.key)}"]`)
      ?.scrollIntoView({ block: 'nearest' });
  };
  const onKeyDown = (e: KeyboardEvent<HTMLUListElement>) => {
    if (e.ctrlKey || e.altKey || e.metaKey || e.nativeEvent.isComposing) return;
    const row = rows[cur];
    switch (e.key) {
      case 'ArrowDown':
        move(cur + 1);
        break;
      case 'ArrowUp':
        move(cur < 0 ? 0 : cur - 1);
        break;
      case 'Home':
        move(0);
        break;
      case 'End':
        move(rows.length - 1);
        break;
      case 'ArrowRight':
        if (!row) return;
        if (row.collapsible && !row.open) setOpen(row.key, true);
        else if (row.collapsible) move(cur + 1);
        break;
      case 'ArrowLeft':
        if (!row) return;
        if (row.collapsible && row.open) setOpen(row.key, false);
        else if (row.parent) move(rows.findIndex((r) => r.key === row.parent));
        break;
      case 'Enter':
        if (row) activate(row, false);
        break;
      case ' ':
        if (row) activate(row, true);
        break;
      default:
        return;
    }
    e.preventDefault();
  };
  return (
    <ul
      ref={(el) => {
        list.current = el;
        if (typeof ref === 'function') ref(el);
        else if (ref) (ref as React.RefObject<HTMLUListElement | null>).current = el;
      }}
      role="tree"
      aria-label={label}
      tabIndex={0}
      className={clsx('outline-none', className)}
      onKeyDown={onKeyDown}
      onFocus={() => {
        setFocused(true);
        if (cur < 0 && rows.length > 0) setFocusKey(rows[0].key);
      }}
      onBlur={() => setFocused(false)}
    >
      {rows.map((row) => (
        <li
          key={row.key}
          data-key={row.key}
          role="treeitem"
          aria-label={rowLabel(row)}
          aria-expanded={row.collapsible ? row.open : undefined}
          aria-selected={row.key === selected ? true : undefined}
          className={clsx(
            focused && row.key === (focusKey ?? selected) && 'outline outline-1 -outline-offset-1 outline-accent',
          )}
          onMouseDown={() => setFocusKey(row.key)}
          onClick={(e) => {
            if (e.detail === 2 && row.node.type !== 'dir' && row.node.type !== 'emptydir') return;
            activate(row, true);
          }}
          onDoubleClick={() => {
            if (row.node.type !== 'dir' && row.node.type !== 'emptydir') activate(row, false);
          }}
        >
          {render(row)}
        </li>
      ))}
    </ul>
  );
}

function rowLabel(row: Row): string {
  switch (row.node.type) {
    case 'dir':
    case 'emptydir':
      return row.node.label;
    case 'book':
      return row.node.book.rel.split('/').pop()!;
    case 'sheet':
      return row.node.file.name;
    case 'doc':
      return row.node.rel;
  }
}

function TreeRow({
  row,
  treeVersion,
  selected,
  sync,
  onToggle,
}: {
  row: Row;
  treeVersion: string;
  selected: boolean;
  sync: BookSync | undefined;
  onToggle: () => void;
}) {
  const n = row.node;
  if (n.type === 'dir') {
    return (
      <div
        className="flex h-[22px] cursor-pointer items-center gap-1 pr-2 hover:bg-hover"
        style={pad(row.depth)}
        title={n.path}
      >
        <Chevron open={row.open} />
        <Icon.Folder size={14} className="shrink-0 text-muted" />
        <span className="truncate text-fg">{n.label}</span>
        {!row.open && <span className="ml-auto shrink-0 text-[11px] text-faint">{n.count}</span>}
      </div>
    );
  }
  if (n.type === 'book') {
    const b = n.book;
    const badge = bookBadges(b, treeVersion);
    const fileName = b.rel.split('/').pop()!;
    const syncMeta = sync && sync.state !== 'outside' ? SYNC_META[sync.state] : null;
    return (
      <div
        className={clsx(
          'group flex h-[22px] cursor-pointer items-center gap-1 pr-2',
          selected ? 'bg-focus' : 'hover:bg-hover',
        )}
        style={pad(row.depth)}
        title={b.rel}
      >
        <span
          onClick={(e) => {
            e.stopPropagation();
            onToggle();
          }}
          onDoubleClick={(e) => e.stopPropagation()}
        >
          <Chevron open={row.open} />
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
    );
  }
  if (n.type === 'sheet') {
    const m = STATUS[n.file.status];
    return (
      <div
        className="flex h-[22px] cursor-pointer items-center gap-1.5 pr-3 hover:bg-hover"
        style={{ paddingLeft: 8 + row.depth * INDENT + 18 }}
        title={m.hint ? `${m.label} — ${m.hint}` : m.label}
      >
        <Icon.File size={14} className="shrink-0 text-muted" />
        <span className={clsx('truncate', n.file.status === 'clean' ? 'text-fg' : m.color)}>{n.file.name}</span>
        <span className={clsx('ml-auto w-4 text-right font-mono text-[11px]', m.color)}>{m.letter}</span>
      </div>
    );
  }
  if (n.type === 'doc') {
    const ignore = n.label === '.xlcodeignore';
    return (
      <div
        className={clsx(
          'flex h-[22px] cursor-pointer items-center gap-1.5 pr-2',
          selected ? 'bg-focus' : 'hover:bg-hover',
        )}
        style={{ paddingLeft: 8 + row.depth * INDENT + 18 }}
        title={
          ignore
            ? `${n.rel} — ブックにしないファイル・フォルダ${n.exists ? '' : '（まだありません。開いて保存すると作成します）'}`
            : `${n.rel} — Copilot に守らせるルール${n.exists ? '' : '（まだありません。開いて保存すると作成します）'}`
        }
      >
        {ignore ? (
          <Icon.EyeOff size={14} className={clsx('shrink-0', n.exists ? 'text-warn' : 'text-faint')} />
        ) : (
          <Icon.Rules size={14} className={clsx('shrink-0', n.exists ? 'text-info' : 'text-faint')} />
        )}
        <span className={clsx('truncate', n.exists ? 'text-fg' : 'text-faint italic')}>{n.label}</span>
      </div>
    );
  }
  return null;
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
