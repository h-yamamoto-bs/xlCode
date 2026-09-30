import { Icon } from './Icons';
import { Kbd } from './ui';

export function Welcome({
  recent,
  onOpen,
  onOpenRecent,
}: {
  recent: string[];
  onOpen: () => void;
  onOpenRecent: (p: string) => void;
}) {
  return (
    <div className="flex h-full items-center justify-center overflow-auto bg-editor">
      <div className="w-[680px] max-w-[90%] py-10">
        <div className="mb-1 flex items-center gap-3 text-[34px] font-light text-fg-strong">
          <Icon.Logo size={56} crop={1.3} className="rounded-[10px]" />
          xlCode
        </div>
        <div className="mb-10 text-[15px] text-muted">
          Excel in Copilot を使ったコーディングを、Build / Sync でソースコードとつなぐ
        </div>
        <div className="grid grid-cols-1 gap-10 sm:grid-cols-2">
          <div>
            <div className="mb-2 text-[15px] text-fg-strong">開始</div>
            <button className="flex items-center gap-2 py-1 text-link hover:underline" onClick={onOpen}>
              <Icon.Folder size={16} /> プロジェクトを開く...
            </button>
            <div className="mt-6 mb-2 text-[15px] text-fg-strong">最近使用したプロジェクト</div>
            {recent.length === 0 && <div className="text-muted">なし</div>}
            {recent.map((p) => (
              <button
                key={p}
                className="flex w-full min-w-0 items-baseline gap-2 py-0.5 text-left"
                onClick={() => onOpenRecent(p)}
                title={p}
              >
                <span className="shrink-0 text-link hover:underline">{p.split(/[\\/]/).pop()}</span>
                <span className="truncate text-[12px] text-faint">{p}</span>
              </button>
            ))}
          </div>
          <div>
            <div className="mb-2 text-[15px] text-fg-strong">基本の流れ</div>
            <ol className="flex flex-col gap-2.5 text-[13px] leading-relaxed text-fg">
              <li>
                <span className="text-muted">1.</span> <b>Excelで開く</b>（Sync してから開く）
              </li>
              <li>
                <span className="text-muted">2.</span> Copilot にシートを編集させる
              </li>
              <li>
                <span className="text-muted">3.</span> Excel を閉じて <b>Build</b>
                <span className="ml-1 text-muted">— 変更をソースコードへ</span>
              </li>
              <li>
                <span className="text-muted">4.</span> エディタで触ったら <b>Sync</b>
                <span className="ml-1 text-muted">— 変更をシートへ</span>
              </li>
            </ol>
            <div className="mt-4 rounded-[3px] border border-line bg-side px-3 py-2 text-[12px] leading-relaxed">
              <div className="flex items-center gap-1.5">
                <Icon.Build size={13} className="text-modified" />
                <b>Build</b>
                <span className="text-muted">Excel</span>
                <Icon.ArrowRight size={12} className="text-muted" />
                <span className="text-muted">ソースコード</span>
              </div>
              <div className="mt-1 flex items-center gap-1.5">
                <Icon.Sync size={13} className="text-info" />
                <b>Sync</b>
                <span className="text-muted">ソースコード</span>
                <Icon.ArrowRight size={12} className="text-muted" />
                <span className="text-muted">Excel</span>
              </div>
              <div className="mt-1.5 text-faint">
                どちらも実行前に Git へ自動コミットし、直前の 1
                回は「元に戻す」で戻せます。画面の「次の操作」に従えば迷いません。
              </div>
            </div>
          </div>
        </div>
        <div className="mt-10 flex flex-wrap gap-x-5 gap-y-2 text-[12px] text-muted">
          <span>
            <Kbd>Ctrl+O</Kbd> 開く
          </span>
          <span>
            <Kbd>Ctrl+B</Kbd> Build
          </span>
          <span>
            <Kbd>Ctrl+Shift+S</Kbd> Sync
          </span>
          <span>
            <Kbd>Ctrl+J</Kbd> パネル
          </span>
        </div>
      </div>
    </div>
  );
}
