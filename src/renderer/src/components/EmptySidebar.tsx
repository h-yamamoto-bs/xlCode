import { formatKey } from '../keys';
import { Icon } from './Icons';
import { Button, Kbd } from './ui';

/** プロジェクトを開いていないときのサイドバー（VS Code の「フォルダーを開く」と同じ） */
export function EmptySidebar({
  recent,
  onOpen,
  onOpenRecent,
  platform = 'win32',
}: {
  recent: string[];
  onOpen: () => void;
  onOpenRecent: (p: string) => void;
  platform?: string;
}) {
  return (
    <div className="flex h-full flex-col bg-side">
      <div className="flex h-[35px] shrink-0 items-center px-5 text-[11px] tracking-wide text-muted uppercase">
        エクスプローラー
      </div>
      <div className="flex flex-col gap-3 px-5 py-2 text-[12px] leading-relaxed text-muted">
        <p>プロジェクト（ソースコードのフォルダ）をまだ開いていません。</p>
        <Button variant="primary" className="w-full justify-center" onClick={onOpen}>
          <Icon.Folder size={14} /> プロジェクトを開く
        </Button>
        <p className="text-faint">
          <Kbd>{formatKey('ctrl+o', platform)}</Kbd> でも開けます
        </p>
        {recent.length > 0 && (
          <div>
            <div className="mb-1 text-[11px] tracking-wide text-muted uppercase">最近使用したプロジェクト</div>
            <ul>
              {recent.map((p) => (
                <li key={p}>
                  <button
                    className="flex w-full min-w-0 items-baseline gap-2 rounded-[3px] px-1 py-0.5 text-left hover:bg-hover"
                    title={p}
                    onClick={() => onOpenRecent(p)}
                  >
                    <span className="shrink-0 text-fg">{p.split(/[\\/]/).pop()}</span>
                    <span className="truncate text-[11px] text-faint">{p}</span>
                  </button>
                </li>
              ))}
            </ul>
          </div>
        )}
      </div>
    </div>
  );
}
