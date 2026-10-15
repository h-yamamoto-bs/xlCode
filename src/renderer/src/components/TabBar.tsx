import clsx from 'clsx';
import { useEffect, useRef, useState, type ReactNode } from 'react';
import type { Tab } from '../tabs';
import { Icon } from './Icons';

export interface TabView {
  tab: Tab;
  label: string;
  /** 同名のタブがあるときなどに添える説明（フォルダ） */
  detail?: string;
  icon: ReactNode;
  /** 保存していない変更がある（× の代わりに ● を出す） */
  dirty?: boolean;
  /** ツールチップ */
  title?: string;
}

interface Menu {
  x: number;
  y: number;
  id: string;
}

/**
 * エディタ領域のタブバー（VS Code と同じ操作）。
 * クリックで表示、ダブルクリックで固定、中クリック・× で閉じる、ドラッグで並べ替え、右クリックでメニュー
 */
export function TabBar({
  tabs,
  active,
  onActivate,
  onClose,
  onPin,
  onMove,
  onCloseOthers,
  onCloseAll,
  actions,
}: {
  /** 右端に出すボタン（サイドバーの切り替えなど） */
  actions?: ReactNode;
  tabs: TabView[];
  active: string | null;
  onActivate: (id: string) => void;
  onClose: (id: string) => void;
  onPin: (id: string) => void;
  onMove: (id: string, to: number) => void;
  onCloseOthers: (id: string) => void;
  onCloseAll: () => void;
}) {
  const [menu, setMenu] = useState<Menu | null>(null);
  const [dragging, setDragging] = useState<string | null>(null);
  const bar = useRef<HTMLDivElement>(null);

  // 表示中のタブが見えるようにスクロールする
  useEffect(() => {
    bar.current
      ?.querySelector<HTMLElement>('[aria-selected="true"]')
      ?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }, [active, tabs.length]);

  useEffect(() => {
    if (!menu) return;
    const close = () => setMenu(null);
    window.addEventListener('mousedown', close);
    window.addEventListener('keydown', close);
    return () => {
      window.removeEventListener('mousedown', close);
      window.removeEventListener('keydown', close);
    };
  }, [menu]);

  return (
    <div className="flex h-[35px] shrink-0 bg-side">
      <div
        ref={bar}
        role="tablist"
        aria-label="開いているタブ"
        className="flex min-w-0 flex-1 overflow-x-auto overflow-y-hidden [scrollbar-width:thin]"
        onWheel={(e) => {
          if (e.deltaY !== 0 && bar.current) bar.current.scrollLeft += e.deltaY;
        }}
      >
        {tabs.map((t, i) => {
          const selected = t.tab.id === active;
          return (
            <div
              key={t.tab.id}
              role="tab"
              aria-selected={selected}
              aria-label={t.detail ? `${t.label} ${t.detail}` : t.label}
              title={t.title ?? (t.detail ? `${t.detail}/${t.label}` : t.label)}
              draggable
              className={clsx(
                'group flex h-full shrink-0 cursor-pointer items-center gap-1.5 border-r border-r-line pr-1 pl-3 text-[13px] select-none',
                selected ? 'border-t border-t-accent bg-editor text-fg-strong' : 'text-muted hover:text-fg',
                dragging === t.tab.id && 'opacity-50',
              )}
              onMouseDown={(e) => {
                if (e.button === 1) {
                  e.preventDefault();
                  onClose(t.tab.id);
                }
              }}
              onClick={() => onActivate(t.tab.id)}
              onDoubleClick={() => onPin(t.tab.id)}
              onContextMenu={(e) => {
                e.preventDefault();
                setMenu({ x: e.clientX, y: e.clientY, id: t.tab.id });
              }}
              onDragStart={(e) => {
                e.dataTransfer.effectAllowed = 'move';
                e.dataTransfer.setData('text/xlcode-tab', t.tab.id);
                setDragging(t.tab.id);
              }}
              onDragEnd={() => setDragging(null)}
              onDragOver={(e) => {
                if (dragging && dragging !== t.tab.id) e.preventDefault();
              }}
              onDrop={(e) => {
                e.preventDefault();
                const id = e.dataTransfer.getData('text/xlcode-tab') || dragging;
                if (id && id !== t.tab.id) onMove(id, i);
                setDragging(null);
              }}
            >
              <span className="shrink-0">{t.icon}</span>
              <span className={clsx('max-w-[220px] truncate', t.tab.preview && 'italic')}>{t.label}</span>
              {t.detail && <span className="max-w-[160px] truncate text-[11px] text-faint">{t.detail}</span>}
              <button
                aria-label={t.dirty ? `${t.label} を閉じる（保存していない変更があります）` : `${t.label} を閉じる`}
                title={t.dirty ? '保存していない変更があります' : '閉じる'}
                className={clsx(
                  'ml-1 flex h-5 w-5 shrink-0 items-center justify-center rounded-[3px] hover:bg-hover-strong',
                  !t.dirty && !selected && 'opacity-0 group-hover:opacity-100',
                )}
                onClick={(e) => {
                  e.stopPropagation();
                  onClose(t.tab.id);
                }}
              >
                {t.dirty ? <span className="text-[12px] leading-none text-fg group-hover:hidden">●</span> : null}
                <Icon.Close size={13} className={clsx(t.dirty && 'hidden group-hover:block')} />
              </button>
            </div>
          );
        })}
        {/* 右の余白へのドロップは末尾へ移動 */}
        <div
          className="min-w-[40px] flex-1 border-b border-line"
          onDragOver={(e) => {
            if (dragging) e.preventDefault();
          }}
          onDrop={(e) => {
            const id = e.dataTransfer.getData('text/xlcode-tab') || dragging;
            if (id) onMove(id, tabs.length);
            setDragging(null);
          }}
        />
      </div>
      {actions && <div className="flex shrink-0 items-center gap-0.5 border-b border-line px-1">{actions}</div>}
      {menu && (
        <ul
          role="menu"
          className="fixed z-50 min-w-[180px] rounded-md border border-widget-border bg-widget py-1 text-[13px] shadow-[0_0_12px_rgba(0,0,0,0.6)]"
          style={{ left: menu.x, top: menu.y }}
          onMouseDown={(e) => e.stopPropagation()}
        >
          {(
            [
              ['閉じる', () => onClose(menu.id), 'Ctrl+W'],
              ['ほかのタブを閉じる', () => onCloseOthers(menu.id), ''],
              ['すべて閉じる', () => onCloseAll(), 'Ctrl+K Ctrl+W'],
              ['固定する', () => onPin(menu.id), ''],
            ] as const
          ).map(([label, fn, keys]) => (
            <li
              key={label}
              role="menuitem"
              className="flex cursor-pointer items-center gap-6 px-4 py-1 hover:bg-focus hover:text-fg-strong"
              onClick={() => {
                setMenu(null);
                fn();
              }}
            >
              <span className="flex-1">{label}</span>
              {keys && <span className="text-[11px] text-faint">{keys}</span>}
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
