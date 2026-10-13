import clsx from 'clsx';
import { useCallback, useRef } from 'react';

/**
 * サイドバーの幅を変えるつまみ（VS Code のサッシュ）。
 * ドラッグ中はウィンドウ全体で mousemove を拾い、幅を min〜max に収める
 */
export function Resizer({
  side,
  width,
  min = 180,
  max = 640,
  onResize,
  onReset,
  label,
}: {
  /** つまみがどちら側のサイドバーに付いているか */
  side: 'left' | 'right';
  width: number;
  min?: number;
  max?: number;
  onResize: (width: number) => void;
  /** ダブルクリックで既定の幅に戻す */
  onReset?: () => void;
  label: string;
}) {
  const start = useRef<{ x: number; width: number } | null>(null);
  const onMouseDown = useCallback(
    (e: React.MouseEvent) => {
      e.preventDefault();
      start.current = { x: e.clientX, width };
      const move = (ev: MouseEvent) => {
        if (!start.current) return;
        const delta = ev.clientX - start.current.x;
        const next = side === 'left' ? start.current.width + delta : start.current.width - delta;
        onResize(Math.max(min, Math.min(max, next)));
      };
      const up = () => {
        start.current = null;
        window.removeEventListener('mousemove', move);
        window.removeEventListener('mouseup', up);
        document.body.style.cursor = '';
        document.body.style.userSelect = '';
      };
      document.body.style.cursor = 'col-resize';
      document.body.style.userSelect = 'none';
      window.addEventListener('mousemove', move);
      window.addEventListener('mouseup', up);
    },
    [side, width, min, max, onResize],
  );
  return (
    <div
      role="separator"
      aria-orientation="vertical"
      aria-label={label}
      aria-valuenow={width}
      aria-valuemin={min}
      aria-valuemax={max}
      tabIndex={0}
      title="ドラッグで幅を変える（ダブルクリックで元に戻す）"
      className={clsx(
        'relative z-10 w-[4px] shrink-0 cursor-col-resize bg-transparent transition-colors hover:bg-accent/70 focus-visible:bg-accent/70',
        side === 'left' ? '-ml-[2px]' : '-mr-[2px]',
      )}
      onMouseDown={onMouseDown}
      onDoubleClick={onReset}
      onKeyDown={(e) => {
        if (e.key === 'ArrowLeft' || e.key === 'ArrowRight') {
          e.preventDefault();
          const d = (e.key === 'ArrowRight' ? 16 : -16) * (side === 'left' ? 1 : -1);
          onResize(Math.max(min, Math.min(max, width + d)));
        }
      }}
    />
  );
}
