import clsx from 'clsx';
import { useEffect, useState, type ReactNode } from 'react';
import type { FileDiff, OpResult, ProjectMode } from '../../../core';
import type { BookSummary, BookSync, ExcelMode, OpenVia } from '../../../shared/api';
import { api, unwrap } from '../api';
import { nextAction, type Next } from '../next';
import { EXCEL_SIDE, SOURCE_SIDE, STATUS, SYNC_META } from '../status';
import { DiffView, diffCaption } from './DiffView';
import { Icon } from './Icons';
import { Button, IconButton } from './ui';

function Banner({
  kind,
  children,
  action,
  onClose,
  role,
}: {
  kind: 'error' | 'warning' | 'info' | 'success' | 'next';
  children: ReactNode;
  action?: ReactNode;
  onClose?: () => void;
  role?: string;
}) {
  const style = {
    error: 'border-deleted/60 bg-[#5a1d1d]/60',
    warning: 'border-warn/50 bg-[#4d3b00]/50',
    info: 'border-accent/60 bg-[#063b49]/50',
    success: 'border-added/50 bg-[#1b3a26]/60',
    next: 'border-line bg-side',
  }[kind];
  const icon = {
    error: <Icon.Error className="text-deleted" />,
    warning: <Icon.Warning className="text-warn" />,
    info: <Icon.Info className="text-info" />,
    success: <Icon.Check className="text-added" />,
    next: <Icon.Lightbulb className="text-warn" />,
  }[kind];
  return (
    <div role={role} className={clsx('flex items-start gap-2.5 rounded-[3px] border px-3 py-2 leading-relaxed', style)}>
      <span className="mt-0.5 shrink-0">{icon}</span>
      <div className="selectable min-w-0 flex-1">{children}</div>
      {action}
      {onClose && (
        <IconButton title="閉じる" onClick={onClose} className="-mr-1 shrink-0">
          <Icon.Close size={14} />
        </IconButton>
      )}
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

/** ツールバーのボタンに添える件数 */
function Count({ n, color }: { n: number; color: string }) {
  if (n === 0) return null;
  return (
    <span className={clsx('rounded-full bg-[#ffffff1a] px-1.5 text-[11px] leading-4 tabular-nums', color)}>{n}</span>
  );
}

export const ACTION_LABEL: Record<string, string> = {
  'write-file': 'ファイル出力',
  'delete-file': 'ファイル削除',
  'write-sheet': 'シート更新',
  'delete-sheet': 'シート削除',
  'reformat-sheet': 'シート整形',
  'conflict-sheet': '衝突シート作成',
  commit: '自動コミット',
  'write-module': 'VBA 書き込み',
  'delete-module': 'VBA 削除',
  backup: '上書き前のバックアップ',
};

/** 直前の操作の結果（画面に出して、何が起きたかをその場で示す） */
export interface LastResult {
  label: string;
  time: string;
  result: OpResult;
}

function summarizeChanges(r: OpResult): string {
  const counts = new Map<string, number>();
  for (const c of r.changes) counts.set(c.action, (counts.get(c.action) ?? 0) + 1);
  return [...counts].map(([a, n]) => (a === 'commit' ? ACTION_LABEL[a] : `${ACTION_LABEL[a] ?? a} ${n}`)).join('、');
}

function ResultBanner({ last, onClose }: { last: LastResult; onClose: () => void }) {
  const r = last.result;
  const kind = r.status === 'ok' ? 'success' : r.status === 'error' ? 'error' : 'warning';
  const title = {
    ok: r.changes.length === 0 ? `${last.label}: 変更はありませんでした` : `${last.label} 完了`,
    error: `${last.label} を中断しました（何も書き換えていません）`,
    confirm: `${last.label} をキャンセルしました`,
    conflict: `${last.label}: 衝突シートを作成しました`,
    'needs-decision': `${last.label} を中止しました`,
  }[r.status];
  const summary = summarizeChanges(r);
  return (
    <Banner kind={kind} onClose={onClose} role="status">
      <div className="flex items-baseline gap-2">
        <span className="text-fg-strong">{title}</span>
        <span className="text-[11px] text-faint">{last.time}</span>
      </div>
      {summary && <div className="text-[12px] text-muted">{summary}</div>}
      {r.errors.map((e) => (
        <div key={e} className="mt-1 text-[12px] text-fg">
          {e}
        </div>
      ))}
      {r.warnings.map((w) => (
        <div key={w} className="mt-1 text-[12px] text-warn">
          {w}
        </div>
      ))}
      {r.changes.length > 0 && (
        <details className="mt-1 text-[12px]">
          <summary className="cursor-pointer text-muted">変更したもの（{r.changes.length}）</summary>
          <ul className="mt-1 rounded-[3px] border border-line bg-editor px-3 py-1.5 font-mono">
            {r.changes.map((c, i) => (
              <li key={i}>
                <span className="text-muted">{ACTION_LABEL[c.action] ?? c.action}:</span> {c.target}
              </li>
            ))}
          </ul>
        </details>
      )}
    </Banner>
  );
}

function fmtTime(iso: string): string {
  const d = new Date(iso);
  return isNaN(d.getTime()) ? iso : d.toLocaleString('ja-JP', { hour12: false });
}

export function BookView({
  root,
  book,
  treeVersion,
  busy,
  mode,
  sync,
  oneDriveRunning,
  lastResult,
  onDismissResult,
  onOpenExcel,
  onSync,
  onBuild,
  onUndo,
  onTerminal,
  onReveal,
  onRefreshTree,
  projectMode = 'source',
  onOpenOutput,
}: {
  root: string;
  projectMode?: ProjectMode;
  /** VBA モード: ビルド結果を開く（reveal ならフォルダで表示） */
  onOpenOutput?: (reveal: boolean) => void;
  book: BookSummary;
  treeVersion: string;
  busy: boolean;
  mode: ExcelMode;
  sync: BookSync | undefined;
  oneDriveRunning: boolean | null;
  lastResult: LastResult | undefined;
  onDismissResult: () => void;
  onOpenExcel: (via: OpenVia) => void;
  onSync: () => void;
  onBuild: () => void;
  onUndo: () => void;
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
  const next: Next = nextAction(book, treeVersion, busy);

  // 押せない理由をツールチップに出す（迷わないため）
  const blocked = busy
    ? '処理中です'
    : book.open
      ? 'デスクトップ版 Excel で開かれているため実行できません。Excel を閉じると自動で押せるようになります'
      : null;
  const reason = (normal: string) => blocked ?? normal;

  // 変更を確認（差分）。展開している行があるあいだは、ブックの状態が変わるたびに取り直す
  const [diffs, setDiffs] = useState<Record<string, FileDiff> | null>(null);
  const [diffError, setDiffError] = useState<string | null>(null);
  const [expanded, setExpanded] = useState<Record<string, boolean>>({});
  const statusKey = JSON.stringify(files.map((f) => [f.name, f.status]));
  const wantDiffs = Object.values(expanded).some(Boolean);
  useEffect(() => {
    setDiffs(null);
    setDiffError(null);
    if (!wantDiffs) return;
    let stale = false;
    unwrap(api.bookDiff(root, book.rel))
      .then((list) => {
        if (!stale) setDiffs(Object.fromEntries(list.map((d) => [d.name, d])));
      })
      .catch((e: Error) => {
        if (!stale) setDiffError(e.message);
      });
    return () => {
      stale = true;
    };
  }, [statusKey, book.rel, root, wantDiffs]);
  const toggle = (name: string) => setExpanded({ ...expanded, [name]: !expanded[name] });
  const changedCount = files.length - clean.length;
  const allOpen = changedCount > 0 && files.every((f) => f.status === 'clean' || expanded[f.name]);
  const toggleAll = () => {
    if (allOpen) setExpanded({});
    else setExpanded(Object.fromEntries(files.filter((f) => f.status !== 'clean').map((f) => [f.name, true])));
  };

  const openPrimary = next.kind === 'open' || next.kind === 'refresh';
  const vba = projectMode === 'vba';
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
            variant={openPrimary ? 'primary' : 'secondary'}
            onClick={() => onOpenExcel('desktop')}
            disabled={disabled}
            title={reason('Sync してからデスクトップ版 Excel で開く')}
          >
            <Icon.Desktop size={15} />
            {mode === 'both' ? 'デスクトップで開く' : 'Excelで開く'}
          </Button>
        )}
        {mode !== 'desktop' && (
          <Button
            variant={openPrimary && mode === 'web' ? 'primary' : 'secondary'}
            onClick={() => onOpenExcel('web')}
            disabled={disabled}
            title={reason('Sync してから Web 版 Excel（ブラウザ）で開く')}
          >
            <Icon.Cloud size={15} />
            {mode === 'both' ? 'Webで開く' : 'Excelで開く'}
          </Button>
        )}
        <div className="mx-1 h-4 w-px bg-line" />
        <Button
          variant={next.kind === 'build' ? 'primary' : 'secondary'}
          onClick={onBuild}
          disabled={disabled}
          aria-label="Build"
          title={reason(
            vba
              ? 'Build: Excel 側の変更をソースコードへ出力し、VBA を書き込んだ .xlsm を作る'
              : 'Build: Excel 側の変更をソースコードへ出力する（エディタ側だけの変更はシートへ取り込む）',
          )}
        >
          <Icon.Build size={15} />
          Build
          <span className="flex items-center gap-0.5 text-[11px] opacity-70">
            Excel
            <Icon.ArrowRight size={11} />
            ソース
          </span>
          <Count n={excel.length + conflict.length} color="text-modified" />
        </Button>
        <Button
          variant={next.kind === 'sync' ? 'primary' : 'secondary'}
          onClick={onSync}
          disabled={disabled}
          aria-label="Sync"
          title={reason('Sync: エディタ側の変更をシートへ反映する')}
        >
          <Icon.Sync size={15} />
          Sync
          <span className="flex items-center gap-0.5 text-[11px] opacity-70">
            ソース
            <Icon.ArrowRight size={11} />
            Excel
          </span>
          <Count n={source.length} color="text-info" />
        </Button>
        {book.undo && (
          <Button
            onClick={onUndo}
            disabled={disabled}
            aria-label={`${book.undo.label} を元に戻す`}
            title={reason(`${fmtTime(book.undo.at)} の ${book.undo.label} を元に戻す`)}
          >
            <Icon.Undo size={15} />
            元に戻す
          </Button>
        )}
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
          {book.loadError && (
            <Banner kind="error">
              ブックを読み込めません: {book.loadError}
              <div className="text-[12px] text-muted">
                ブックが壊れているか、Excel 以外の形式かもしれません。「フォルダ」から実物を確認してください。
              </div>
            </Banner>
          )}
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
                  ? 'まだ Build していません。Build すると、シートをソースコードに書き出してから、画面のシートと VBA を入れた .xlsm を作ります。'
                  : out.changed
                    ? `${outName} が前回の Build の後に変更されています。次の Build で作り直すと、直接入力したデータや VBE で直したコードは失われます（元のファイルは .xlcode/backup に残します）。`
                    : 'ビルド結果は Build のたびに作り直す「ひな形」です（Git 管理はしません）。実際に使うときはコピーして使ってください。'}
              </div>
            </Banner>
          )}
          {(book.deletes?.length ?? 0) > 0 && (
            <Banner kind="warning">
              削除マーク付きのシート: <span className="font-mono">{book.deletes!.join(', ')}</span>（Build
              時に確認のうえ削除します）
            </Banner>
          )}
          {lastResult && <ResultBanner last={lastResult} onClose={onDismissResult} />}

          {/* 次の操作 */}
          <Banner
            kind="next"
            role="note"
            action={
              next.kind === 'build' ? (
                <Button
                  variant="primary"
                  onClick={onBuild}
                  disabled={disabled}
                  className="shrink-0"
                  aria-label="次の操作: Build"
                >
                  <Icon.Build size={14} /> Build
                </Button>
              ) : next.kind === 'sync' ? (
                <Button
                  variant="primary"
                  onClick={onSync}
                  disabled={disabled}
                  className="shrink-0"
                  aria-label="次の操作: Sync"
                >
                  <Icon.Sync size={14} /> Sync
                </Button>
              ) : next.kind === 'refresh' ? (
                <Button
                  variant="primary"
                  onClick={onRefreshTree}
                  disabled={busy}
                  className="shrink-0"
                  aria-label="次の操作: Refresh Tree"
                >
                  Refresh Tree
                </Button>
              ) : undefined
            }
          >
            <span className="mr-2 text-[11px] tracking-wide text-muted uppercase">次の操作</span>
            <span className="text-fg-strong">{next.text}</span>
            {next.detail && <div className="text-[12px] text-muted">{next.detail}</div>}
          </Banner>

          <div className="grid grid-cols-2 gap-3 md:grid-cols-4">
            <Stat label="Excel側で編集中" value={excel.length} color="text-modified" hint="Build で反映" />
            <Stat label="エディタ側で変更" value={source.length} color="text-info" hint="Sync で取り込み" />
            <Stat label="両側で変更" value={conflict.length} color="text-conflict" hint="衝突シートで統合" />
            <Stat label="同期済み" value={clean.length} color="text-fg" hint={`全 ${files.length} ファイル`} />
          </div>

          <div className="overflow-hidden rounded-[3px] border border-line">
            <table className="w-full border-collapse text-left">
              <thead>
                <tr className="h-[26px] bg-side text-[11px] text-muted uppercase">
                  <th className="w-10 px-3 font-normal" />
                  <th className="px-2 font-normal">ファイル（シート）</th>
                  <th className="px-2 font-normal">状態</th>
                  <th className="px-2 font-normal">形式</th>
                  <th className="px-3 font-normal">次の操作</th>
                  <th className="w-[120px] px-3 text-right font-normal normal-case">
                    {changedCount > 0 && (
                      <button
                        className="inline-flex items-center gap-1 text-[11px] text-link hover:underline"
                        onClick={toggleAll}
                      >
                        <Icon.Diff size={12} />
                        {allOpen ? '差分を閉じる' : '変更を確認'}
                      </button>
                    )}
                  </th>
                </tr>
              </thead>
              <tbody>
                {files.map((f) => {
                  const m = STATUS[f.status];
                  const changed = f.status !== 'clean';
                  const open = changed && !!expanded[f.name];
                  const d = diffs?.[f.name];
                  return [
                    <tr
                      key={f.name}
                      className={clsx('h-[24px] border-t border-line', changed && 'cursor-pointer', 'hover:bg-hover')}
                      onClick={changed ? () => toggle(f.name) : undefined}
                      aria-expanded={changed ? open : undefined}
                    >
                      <td className={clsx('px-3 text-center font-mono text-[12px]', m.color)} title={m.label}>
                        {m.letter || <Icon.Check size={12} className="inline text-faint" />}
                      </td>
                      <td className="selectable px-2 font-mono text-[12.5px] text-fg">{f.name}</td>
                      <td className={clsx('px-2 text-[12px]', m.color)}>{m.label}</td>
                      <td className="px-2 text-[12px] whitespace-nowrap text-faint">{f.format}</td>
                      <td className="px-3 text-[12px] text-muted">{m.hint}</td>
                      <td className="px-3 text-right text-[11px] whitespace-nowrap">
                        {changed && (
                          <span className="inline-flex items-center gap-1 text-muted">
                            {d && (
                              <span className="font-mono tabular-nums">
                                <span className="text-added">+{d.added}</span>{' '}
                                <span className="text-deleted">−{d.removed}</span>
                              </span>
                            )}
                            {open ? <Icon.ChevronDown size={12} /> : <Icon.ChevronRight size={12} />}
                          </span>
                        )}
                      </td>
                    </tr>,
                    open && (
                      <tr key={`${f.name}:diff`} className="border-t border-line bg-editor">
                        <td colSpan={6} className="p-0">
                          {d ? (
                            <>
                              <div className="px-3 py-1 text-[11px] text-muted">{diffCaption(d)}</div>
                              <DiffView diff={d} />
                            </>
                          ) : diffError ? (
                            <div className="px-3 py-2 text-[12px] text-deleted">差分を取得できません: {diffError}</div>
                          ) : (
                            <div className="flex items-center gap-2 px-3 py-2 text-[12px] text-muted">
                              <Icon.Spinner size={12} /> 差分を読み込んでいます...
                            </div>
                          )}
                        </td>
                      </tr>
                    ),
                  ];
                })}
                {files.length === 0 && (
                  <tr>
                    <td colSpan={6} className="px-3 py-3 text-[12px] text-faint">
                      コードシートがありません
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
