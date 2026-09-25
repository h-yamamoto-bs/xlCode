import clsx from 'clsx';
import { useEffect, useRef } from 'react';
import { Icon } from './Icons';
import { IconButton } from './ui';

export interface LogEntry {
  time: string;
  level: 'info' | 'success' | 'warning' | 'error';
  text: string;
}

export interface Problem {
  level: 'error' | 'warning';
  source: string;
  text: string;
}

export function Panel({
  tab,
  onTab,
  logs,
  problems,
  onClose,
  onClear,
}: {
  tab: 'problems' | 'output';
  onTab: (t: 'problems' | 'output') => void;
  logs: LogEntry[];
  problems: Problem[];
  onClose: () => void;
  onClear: () => void;
}) {
  const end = useRef<HTMLDivElement>(null);
  useEffect(() => {
    // 新しい Chromium では scrollIntoView が Promise を返すため、戻り値を返さない
    end.current?.scrollIntoView({ block: 'end' });
  }, [logs, tab]);
  const tabs = [
    { id: 'problems' as const, label: '問題', count: problems.length },
    { id: 'output' as const, label: '出力' },
  ];
  return (
    <div className="flex h-full flex-col border-t border-line bg-side">
      <div className="flex h-[35px] shrink-0 items-center gap-4 px-4">
        {tabs.map((t) => (
          <button
            key={t.id}
            className={clsx(
              'flex h-full items-center gap-1.5 border-b text-[11px] tracking-wide uppercase',
              tab === t.id ? 'border-fg-strong text-fg-strong' : 'border-transparent text-muted hover:text-fg',
            )}
            onClick={() => onTab(t.id)}
          >
            {t.label}
            {t.count !== undefined && t.count > 0 && (
              <span className="rounded-full bg-[#616161] px-1.5 text-[10px] leading-4 text-white">{t.count}</span>
            )}
          </button>
        ))}
        <div className="ml-auto flex gap-0.5">
          {tab === 'output' && (
            <IconButton title="出力をクリア" onClick={onClear}>
              <Icon.Close size={14} />
            </IconButton>
          )}
          <IconButton title="パネルを閉じる" onClick={onClose}>
            <Icon.ChevronDown size={14} />
          </IconButton>
        </div>
      </div>
      <div className="selectable min-h-0 flex-1 overflow-auto px-4 pb-2 font-mono text-[12px] leading-[18px]">
        {tab === 'output' &&
          logs.map((l, i) => (
            <div
              key={i}
              className={clsx(
                'whitespace-pre-wrap',
                l.level === 'error' && 'text-deleted',
                l.level === 'warning' && 'text-warn',
                l.level === 'success' && 'text-added',
              )}
            >
              <span className="text-faint">[{l.time}]</span> {l.text}
            </div>
          ))}
        {tab === 'problems' &&
          (problems.length === 0 ? (
            <div className="font-sans text-muted">問題は検出されていません。</div>
          ) : (
            problems.map((p, i) => (
              <div key={i} className="flex items-start gap-2 py-0.5 font-sans text-[13px]">
                {p.level === 'error' ? (
                  <Icon.Error size={14} className="mt-0.5 shrink-0 text-deleted" />
                ) : (
                  <Icon.Warning size={14} className="mt-0.5 shrink-0 text-warn" />
                )}
                <span className="text-fg">{p.text}</span>
                <span className="shrink-0 text-faint">{p.source}</span>
              </div>
            ))
          ))}
        <div ref={end} />
      </div>
    </div>
  );
}
