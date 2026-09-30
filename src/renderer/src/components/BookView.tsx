import clsx from 'clsx';
import type { ReactNode } from 'react';
import type { ProjectMode } from '../../../core';
import type { BookSummary, BookSync, ExcelMode, OpenVia } from '../../../shared/api';
import { EXCEL_SIDE, SOURCE_SIDE, STATUS, SYNC_META, VBA_STATUS } from '../status';
import { Icon } from './Icons';
import { Button } from './ui';

function Banner({
  kind,
  children,
  action,
}: {
  kind: 'error' | 'warning' | 'info';
  children: ReactNode;
  action?: ReactNode;
}) {
  const style = {
    error: 'border-deleted/60 bg-[#5a1d1d]/60',
    warning: 'border-warn/50 bg-[#4d3b00]/50',
    info: 'border-accent/60 bg-[#063b49]/50',
  }[kind];
  const icon = {
    error: <Icon.Error className="text-deleted" />,
    warning: <Icon.Warning className="text-warn" />,
    info: <Icon.Info className="text-info" />,
  }[kind];
  return (
    <div className={clsx('flex items-start gap-2.5 rounded-[3px] border px-3 py-2 leading-relaxed', style)}>
      <span className="mt-0.5 shrink-0">{icon}</span>
      <div className="selectable min-w-0 flex-1">{children}</div>
      {action}
    </div>
  );
}

function Stat({ label, value, color, hint }: { label: string; value: number; color: string; hint: string }) {
  return (
    <div className="rounded-[3px] border border-line bg-side px-4 py-3">
      <div className={clsx('text-[22px] font-light tabular-nums', value > 0 ? color : 'text-faint')}>{value}</div>
      <div className="text-[12px] text-fg">{label}</div>
      <div className="text-[11px] text-faint">{hint}</div>
    </div>
  );
}

export function BookView({
  book,
  treeVersion,
  busy,
  mode,
  sync,
  oneDriveRunning,
  onOpenExcel,
  onSync,
  onBuild,
  onTerminal,
  onReveal,
  onRefreshTree,
  projectMode = 'source',
  onOpenOutput,
}: {
  projectMode?: ProjectMode;
  /** VBA モード: ビルド結果を開く（reveal ならフォルダで表示） */
  onOpenOutput?: (reveal: boolean) => void;
  book: BookSummary;
  treeVersion: string;
  busy: boolean;
  mode: ExcelMode;
  sync: BookSync | undefined;
  oneDriveRunning: boolean | null;
  onOpenExcel: (via: OpenVia) => void;
  onSync: () => void;
  onBuild: () => void;
  onTerminal: () => void;
  onReveal: () => void;
  onRefreshTree: () => void;
}) {
  const files = book.files ?? [];
  const excel = files.filter((f) => EXCEL_SIDE.includes(f.status));
  const source = files.filter((f) => SOURCE_SIDE.includes(f.status));
  const conflict = files.filter((f) => f.status === 'conflict');
  const clean = files.filter((f) => f.status === 'clean');
  const fileName = book.rel.split('/').pop()!;
  const disabled = busy || book.open;
  const vba = projectMode === 'vba';
  const removed = files.filter((f) => f.status === 'removed');
  const out = book.vba;
  const outName = out?.output.split(/[\\/]/).pop();

  return (
    <div className="flex h-full min-h-0 flex-col">
      {/* パンくず */}
      <div className="flex h-[22px] shrink-0 items-center gap-1 px-4 text-[12px] text-muted">
        {(book.dirRel ? book.dirRel.split('/') : ['.']).map((seg, i) => (
          <span key={i} className="flex items-center gap-1">
            {seg}
            <Icon.ChevronRight size={12} />
          </span>
        ))}
        <Icon.Book size={13} className="text-excel" />
        <span className="text-fg">{fileName}</span>
      </div>

      {/* ツールバー */}
      <div className="flex shrink-0 flex-wrap items-center gap-2 border-b border-line px-4 pt-1 pb-3">
        {mode !== 'web' && (
          <Button
            variant="primary"
            onClick={() => onOpenExcel('desktop')}
            disabled={disabled}
            title={vba ? 'デスクトップ版 Excel で開く' : 'Sync してからデスクトップ版 Excel で開く'}
          >
            <Icon.Desktop size={15} />
            {mode === 'both' ? 'デスクトップで開く' : 'Excelで開く'}
          </Button>
        )}
        {mode !== 'desktop' && (
          <Button
            variant={mode === 'web' ? 'primary' : 'secondary'}
            onClick={() => onOpenExcel('web')}
            disabled={disabled}
            title={vba ? 'Web 版 Excel（ブラウザ）で開く' : 'Sync してから Web 版 Excel（ブラウザ）で開く'}
          >
            <Icon.Cloud size={15} />
            {mode === 'both' ? 'Webで開く' : 'Excelで開く'}
          </Button>
        )}
        {!vba && (
          <Button onClick={onSync} disabled={disabled} title="ソースコード → Excel">
            <Icon.Sync size={15} />
            Sync
          </Button>
        )}
        <Button
          onClick={onBuild}
          disabled={disabled}
          title={vba ? 'シートの VBA を書き込んだ .xlsm を生成する' : 'Excel → ソースコード'}
        >
          <Icon.Build size={15} />
          Build
        </Button>
        {vba && (
          <Button
            onClick={() => onOpenOutput?.(false)}
            disabled={busy || !out?.exists}
            title="ビルド結果の .xlsm をデスクトップ版 Excel で開く"
          >
            <Icon.Excel size={15} />
            ビルド結果を開く
          </Button>
        )}
        <div className="mx-1 h-4 w-px bg-line" />
        <Button onClick={onTerminal} disabled={busy} title="OS のターミナルをこのディレクトリで開く">
          <Icon.Terminal size={15} />
          ターミナル
        </Button>
        <Button onClick={onReveal} disabled={busy} title="フォルダを表示">
          <Icon.Reveal size={15} />
          フォルダ
        </Button>
      </div>

      <div className="min-h-0 flex-1 overflow-auto px-4 py-3">
        <div className="mx-auto flex max-w-[980px] flex-col gap-3">
          {book.loadError && <Banner kind="error">ブックを読み込めません: {book.loadError}</Banner>}
          {sync && oneDriveRunning === false && sync.state !== 'outside' && (
            <Banner kind="warning">OneDrive が起動していません。ブックの変更がクラウドと同期されません。</Banner>
          )}
          {sync && (SYNC_META[sync.state].busy || SYNC_META[sync.state].problem) && (
            <Banner kind={SYNC_META[sync.state].problem ? 'warning' : 'info'}>
              OneDrive: {SYNC_META[sync.state].label}
              <div className="text-[12px] text-muted">
                {SYNC_META[sync.state].busy
                  ? 'Build / Sync / 開く は、同期が終わるまで自動で待ちます。'
                  : 'OneDrive の画面で状態を確認してください。このまま Build / Sync すると、変更が失われたり複製ができたりする可能性があります。'}
              </div>
            </Banner>
          )}
          {book.open && (
            <Banner kind="warning">
              このブックはデスクトップ版 Excel で開かれています。Build / Sync は Excel
              を閉じてから実行してください（閉じると自動で解除されます）。
              <div className="text-[12px] text-muted">{book.openReason}</div>
            </Banner>
          )}
          {(book.conflictSheets?.length ?? 0) > 0 && (
            <Banner kind="error">
              未解決の衝突シートがあります: <span className="font-mono">{book.conflictSheets!.join(', ')}</span>
              <div className="text-[12px] text-muted">
                Copilot に「A列とB列を統合して元のシートに書き、衝突シートを削除して」と指示してください。解決するまで
                Build できません。
              </div>
            </Banner>
          )}
          {book.errors?.map((e) => (
            <Banner key={e} kind="error">
              {e}
            </Banner>
          ))}
          {book.treeVersion !== treeVersion && (
            <Banner
              kind="info"
              action={
                <Button onClick={onRefreshTree} disabled={busy} className="shrink-0">
                  Refresh Tree
                </Button>
              }
            >
              #tree が最新ではありません。Refresh Tree を実行してください。
            </Banner>
          )}
          {vba && out && (
            <Banner
              kind={out.changed ? 'warning' : 'info'}
              action={
                out.exists && (
                  <Button onClick={() => onOpenOutput?.(true)} disabled={busy} className="shrink-0">
                    <Icon.Reveal size={14} />
                    フォルダ
                  </Button>
                )
              }
            >
              ビルド結果: <span className="font-mono break-all">{out.output}</span>
              <div className="text-[12px] text-muted">
                {!out.exists
                  ? 'まだ Build していません。Build すると、拡張子付きのシートを除き、VBA を書き込んだ .xlsm を作ります。'
                  : out.changed
                    ? `${outName} が前回の Build の後に変更されています。次の Build で作り直すと、直接入力したデータや VBE で直したコードは失われます（元のファイルは .xlcode/backup に残します）。`
                    : 'ビルド結果は Build のたびに作り直す「ひな形」です。実際に使うときはコピーして使ってください（コードの修正は編集用ブックで行います）。'}
              </div>
            </Banner>
          )}
          {(book.deletes?.length ?? 0) > 0 && (
            <Banner kind="warning">
              削除マーク付きのシート: <span className="font-mono">{book.deletes!.join(', ')}</span>（Build
              時に確認のうえ削除します）
            </Banner>
          )}

          {vba ? (
            <div className="grid grid-cols-3 gap-3">
              <Stat label="未ビルドの変更" value={excel.length} color="text-modified" hint="Build で書き込み" />
              <Stat label="削除したシート" value={removed.length} color="text-deleted" hint="Build でモジュール削除" />
              <Stat
                label="ビルド済み"
                value={clean.length}
                color="text-fg"
                hint={`全 ${files.length - removed.length} モジュール`}
              />
            </div>
          ) : (
            <div className="grid grid-cols-2 gap-3 md:grid-cols-4">
              <Stat label="Excel側で編集中" value={excel.length} color="text-modified" hint="Build で反映" />
              <Stat label="エディタ側で変更" value={source.length} color="text-info" hint="Sync で取り込み" />
              <Stat label="両側で変更" value={conflict.length} color="text-conflict" hint="衝突シートで統合" />
              <Stat label="同期済み" value={clean.length} color="text-fg" hint={`全 ${files.length} ファイル`} />
            </div>
          )}

          <div className="overflow-hidden rounded-[3px] border border-line">
            <table className="w-full border-collapse text-left">
              <thead>
                <tr className="h-[26px] bg-side text-[11px] text-muted uppercase">
                  <th className="w-10 px-3 font-normal" />
                  <th className="px-2 font-normal">{vba ? 'シート' : 'ファイル（シート）'}</th>
                  <th className="px-2 font-normal">状態</th>
                  <th className="px-2 font-normal">{vba ? '種類' : '形式'}</th>
                  <th className="px-3 font-normal">次の操作</th>
                </tr>
              </thead>
              <tbody>
                {files.map((f) => {
                  const m = vba ? (VBA_STATUS[f.status] ?? STATUS[f.status]) : STATUS[f.status];
                  return (
                    <tr key={f.name} className="h-[24px] border-t border-line hover:bg-hover">
                      <td className={clsx('px-3 text-center font-mono text-[12px]', m.color)}>
                        {m.letter || <Icon.Check size={12} className="inline text-faint" />}
                      </td>
                      <td className="selectable px-2 font-mono text-[12.5px] text-fg">{f.name}</td>
                      <td className={clsx('px-2 text-[12px]', m.color)}>{m.label}</td>
                      <td className="px-2 text-[12px] whitespace-nowrap text-faint">{f.format}</td>
                      <td className="px-3 text-[12px] text-muted">{m.hint}</td>
                    </tr>
                  );
                })}
                {files.length === 0 && (
                  <tr>
                    <td colSpan={5} className="px-3 py-3 text-[12px] text-faint">
                      {vba ? 'VBA のシート（.bas / .cls / .frm）がありません' : 'コードシートがありません'}
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}
