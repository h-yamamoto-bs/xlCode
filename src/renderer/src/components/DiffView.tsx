import clsx from 'clsx';
import type { FileDiff } from '../../../core';

const SIDE = { excel: 'Excel 側', source: 'ソース側' } as const;

/** 差分の見出し（現在 → 反映後） */
export function diffCaption(d: FileDiff): string {
  if (d.status === 'conflict') return `食い違い: ${SIDE[d.from]}（現在のファイル）と ${SIDE[d.to]}（シート）`;
  const op = d.to === 'excel' ? 'Build' : 'Sync';
  const target = d.to === 'excel' ? 'ソースファイル' : 'シート';
  return `${op} 後の${target}: ${SIDE[d.from]}の現在 → ${SIDE[d.to]}の内容（${d.oldLines} → ${d.newLines} 行）`;
}

export function DiffView({ diff }: { diff: FileDiff }) {
  if (diff.hunks.length === 0) {
    return <div className="px-3 py-2 text-[12px] text-muted">内容は同じです（整形の違いのみ）</div>;
  }
  return (
    <div className="selectable overflow-auto font-mono text-[12px] leading-[18px]" aria-label={`${diff.name} の差分`}>
      {diff.hunks.map((h, i) => {
        let oldNo = h.oldStart;
        let newNo = h.newStart;
        return (
          <div key={i} className="border-t border-line first:border-t-0">
            <div className="bg-side px-3 text-[11px] text-info">
              @@ -{h.oldStart} +{h.newStart} @@
            </div>
            {h.lines.map((l, j) => {
              const o = l.kind === 'add' ? '' : String(oldNo++);
              const n = l.kind === 'del' ? '' : String(newNo++);
              return (
                <div
                  key={j}
                  className={clsx(
                    'flex whitespace-pre',
                    l.kind === 'add' && 'bg-[#2ea04326] text-added',
                    l.kind === 'del' && 'bg-[#f8514926] text-deleted',
                  )}
                >
                  <span className="w-10 shrink-0 pr-2 text-right text-faint tabular-nums">{o}</span>
                  <span className="w-10 shrink-0 pr-2 text-right text-faint tabular-nums">{n}</span>
                  <span className="w-4 shrink-0 text-center">
                    {l.kind === 'add' ? '+' : l.kind === 'del' ? '−' : ''}
                  </span>
                  <span className="pr-3">{l.text}</span>
                </div>
              );
            })}
          </div>
        );
      })}
    </div>
  );
}
