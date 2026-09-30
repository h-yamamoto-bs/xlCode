import clsx from 'clsx';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { Confirmation, OpResult, ProjectMode } from '../../core';
import type { ExcelMode, OpenVia, ProjectInfo, SyncReport } from '../../shared/api';
import { api, storageGet, storageSet, unwrap } from './api';
import { BookView } from './components/BookView';
import { useDialog } from './components/Dialogs';
import { Icon } from './components/Icons';
import { Panel, type LogEntry, type Problem } from './components/Panel';
import { RulesEditor, RulesList, type Draft } from './components/RulesView';
import { EXCEL_MODES, SettingsView } from './components/SettingsView';
import { Sidebar } from './components/Sidebar';
import { StatusBar } from './components/StatusBar';
import { Welcome } from './components/Welcome';
import { SYNC_META } from './status';

type View = 'books' | 'rules' | 'settings';

/** xlCode からブックを開いた方法と、Web 版で開いた時点のブックの状態 */
interface Opened {
  via: OpenVia;
  stamp?: string;
}

const ACTION_LABEL: Record<string, string> = {
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

function now(): string {
  return new Date().toLocaleTimeString('ja-JP', { hour12: false });
}

export function App() {
  const [root, setRoot] = useState<string | null>(null);
  const [project, setProject] = useState<ProjectInfo | null>(null);
  const [selected, setSelected] = useState<string | null>(null);
  const [view, setView] = useState<View>('books');
  const [ruleRel, setRuleRel] = useState('Agents.md');
  // 保存していない Markdown の編集内容（ファイルや画面を切り替えても保持する）
  const [drafts, setDrafts] = useState<Record<string, Draft>>({});
  const setDraft = useCallback((rel: string, d: Draft | null) => {
    setDrafts((prev) => {
      const next = { ...prev };
      if (d) next[rel] = d;
      else delete next[rel];
      return next;
    });
  }, []);
  const hasDrafts = Object.keys(drafts).length > 0;
  const [logs, setLogs] = useState<LogEntry[]>([]);
  const [panelOpen, setPanelOpen] = useState(true);
  const [panelTab, setPanelTab] = useState<'problems' | 'output'>('output');
  const [busy, setBusy] = useState<string | null>(null);
  const [recent, setRecent] = useState<string[]>(() => storageGet('recent', []));
  // 使う Excel（このパソコンの設定）。未設定なら初回に確認する
  const [excelMode, setExcelModeState] = useState<ExcelMode | null>(() =>
    storageGet<ExcelMode | null>('excelMode', null),
  );
  const mode: ExcelMode = excelMode ?? 'both';
  const setExcelMode = useCallback((m: ExcelMode) => {
    setExcelModeState(m);
    storageSet('excelMode', m);
  }, []);
  // ブックを最後に xlCode からどちらで開いたか（キー: ルート|ブック）。
  // Web 版で開いたときは、その時点のブックの更新日時・サイズ（stamp）も覚えておく
  const lastOpened = useRef<Record<string, Opened>>(
    Object.fromEntries(
      Object.entries(storageGet<Record<string, Opened | OpenVia>>('lastOpened', {})).map(([k, v]) => [
        k,
        typeof v === 'string' ? { via: v } : v,
      ]),
    ),
  );
  const setOpened = (key: string, o: Opened) => {
    lastOpened.current = { ...lastOpened.current, [key]: o };
    storageSet('lastOpened', lastOpened.current);
  };
  // OneDrive の同期状態（Windows のみ）
  const [syncReport, setSyncReport] = useState<SyncReport | null>(null);
  const mergeSync = useCallback((r: SyncReport) => {
    setSyncReport((prev) => ({ oneDriveRunning: r.oneDriveRunning, books: { ...(prev?.books ?? {}), ...r.books } }));
  }, []);
  const webSkip = useRef(new Set<string>());
  const busyRef = useRef(false);
  const { ask, node: dialog } = useDialog();

  const log = useCallback((level: LogEntry['level'], text: string) => {
    setLogs((l) => [...l.slice(-500), { time: now(), level, text }]);
  }, []);

  const reload = useCallback(
    async (r: string | null = root) => {
      if (!r) return;
      try {
        const p = await unwrap(api.loadProject(r));
        setProject(p);
        setSelected((s) => (s && p.books.some((b) => b.rel === s) ? s : (p.books[0]?.rel ?? null)));
      } catch (e) {
        log('error', `読み込みに失敗しました: ${(e as Error).message}`);
      }
    },
    [root, log],
  );

  /** 操作の排他制御と後処理（再読み込み） */
  const withBusy = useCallback(
    async (label: string, fn: () => Promise<void>, r: string | null = root) => {
      if (busyRef.current) return;
      busyRef.current = true;
      setBusy(label);
      try {
        await fn();
      } catch (e) {
        log('error', (e as Error).message);
        setPanelOpen(true);
        setPanelTab('output');
      } finally {
        await reload(r);
        busyRef.current = false;
        setBusy(null);
      }
    },
    [root, reload, log],
  );

  const logResult = useCallback(
    (label: string, r: OpResult) => {
      for (const c of r.changes) log('info', `  ${ACTION_LABEL[c.action] ?? c.action}: ${c.target}`);
      for (const w of r.warnings) log('warning', `  警告: ${w}`);
      for (const e of r.errors) log('error', `  エラー: ${e}`);
      const summary: Record<OpResult['status'], [LogEntry['level'], string]> = {
        ok: ['success', r.changes.length === 0 ? '変更なし' : '完了'],
        error: ['error', '中断しました'],
        confirm: ['warning', 'キャンセルしました'],
        conflict: ['warning', '衝突シートを作成しました'],
        'needs-decision': ['warning', 'キャンセルしました'],
      };
      const [level, text] = summary[r.status];
      log(level, `${label}: ${text}`);
      if (r.status === 'error') setPanelOpen(true);
    },
    [log],
  );

  // ---- 確認ダイアログ ----

  const confirmWebClosed = useCallback(
    async (book: string): Promise<boolean> => {
      if (mode === 'desktop' || webSkip.current.has(book)) return true;
      // 両方モード: xlCode からデスクトップ版で開いたブックは、ロックファイルで閉じたことを検知できる
      if (mode === 'both' && lastOpened.current[`${root}|${book}`]?.via === 'desktop') return true;
      const { value, checked } = await ask({
        title: 'Web 版 Excel でこのブックを閉じましたか？',
        icon: 'warning',
        body: (
          <>
            <p>
              Web 版 Excel で開いている状態は xlCode
              から検知できません（デスクトップ版で開いている場合は自動で検知します）。
            </p>
            <p className="mt-2">
              開いたまま実行すると、自動保存とローカルの書き換えがぶつかり、片方の変更が丸ごと失われます。ブラウザのタブを閉じ、OneDrive
              の同期完了を確認してから続行してください。
            </p>
          </>
        ),
        checkbox: 'このブックについては、アプリを閉じるまで確認しない',
        buttons: [
          { label: '閉じたので続行', value: true, variant: 'primary' },
          { label: 'キャンセル', value: false },
        ],
        cancelValue: false,
      });
      if (value && checked) webSkip.current.add(book);
      return value;
    },
    [ask, mode, root],
  );

  const askConfirmations = useCallback(
    async (label: string, cs: Confirmation[]): Promise<boolean> => {
      const danger = cs.some((c) => c.kind === 'delete' || c.kind === 'shrink' || c.kind === 'overwrite');
      const { value } = await ask({
        title: `${label} の前に確認してください`,
        icon: 'warning',
        body: (
          <div className="flex flex-col gap-3">
            {cs.map((c) => (
              <div key={c.kind}>
                <div
                  className={clsx(
                    c.kind === 'shrink' || c.kind === 'delete' || c.kind === 'overwrite' ? 'text-warn' : 'text-fg',
                  )}
                >
                  {c.message}
                </div>
                {c.files.length > 0 && (
                  <ul className="mt-1 rounded-[3px] border border-line bg-editor px-3 py-1.5 font-mono text-[12px]">
                    {c.files.map((f) => (
                      <li key={f}>{f}</li>
                    ))}
                  </ul>
                )}
              </div>
            ))}
          </div>
        ),
        buttons: [
          { label: '続行', value: true, variant: danger ? 'danger' : 'primary' },
          { label: 'キャンセル', value: false },
        ],
        cancelValue: false,
      });
      return value;
    },
    [ask],
  );

  /** confirm が返ったら確認して confirmed: true で再実行する */
  const runConfirmed = useCallback(
    async (
      label: string,
      run: (confirmed: boolean) => Promise<OpResult>,
      autoAccept: Confirmation['kind'][] = [],
    ): Promise<OpResult> => {
      const r = await run(false);
      if (r.status !== 'confirm') return r;
      const rest = r.confirmations.filter((c) => !autoAccept.includes(c.kind));
      if (rest.length > 0 && !(await askConfirmations(label, rest))) return r;
      return run(true);
    },
    [askConfirmations],
  );

  const showConflicts = useCallback(
    async (r: OpResult) => {
      await ask({
        title: '両側で変更されたファイルがあります',
        icon: 'warning',
        body: (
          <>
            <p>次の衝突シートを作成しました。</p>
            <ul className="my-2 rounded-[3px] border border-line bg-editor px-3 py-1.5 font-mono text-[12px]">
              {r.conflicts.map((c) => (
                <li key={c.sheet}>
                  {c.sheet} ← {c.file}
                </li>
              ))}
            </ul>
            <p>
              Excel で Copilot に「A列（Excel側）と
              B列（ソース側）を統合して元のシートに書き、衝突シートを削除して」と指示し、その後 Build してください。
            </p>
          </>
        ),
        buttons: [{ label: 'OK', value: true, variant: 'primary' }],
        cancelValue: true,
      });
    },
    [ask],
  );

  // ---- 操作 ----

  const buildFlow = useCallback(
    async (book: string): Promise<boolean> => {
      log('info', `Build 開始: ${book}`);
      const r = await runConfirmed('Build', (confirmed) => unwrap(api.build(root!, book, { confirmed })));
      logResult('Build', r);
      if (r.status === 'conflict') await showConflicts(r);
      return r.status === 'ok';
    },
    [root, runConfirmed, logResult, showConflicts, log],
  );

  const syncFlow = useCallback(
    async (book: string): Promise<OpResult | null> => {
      log('info', `Sync 開始: ${book}`);
      let discard = false;
      let r = await runConfirmed('Sync', (confirmed) => unwrap(api.sync(root!, book, { confirmed })));
      if (r.status === 'needs-decision') {
        const { value } = await ask<'discard' | 'build' | 'cancel'>({
          title: 'Excel 側に Build していない変更があります',
          icon: 'warning',
          body: (
            <>
              <ul className="mb-2 rounded-[3px] border border-line bg-editor px-3 py-1.5 font-mono text-[12px]">
                {r.unbuilt.map((f) => (
                  <li key={f}>{f}</li>
                ))}
              </ul>
              <p>このまま Sync するとこれらの変更は失われます。どうしますか？</p>
            </>
          ),
          buttons: [
            { label: '先に Build してから Sync', value: 'build', variant: 'primary' },
            { label: 'シートの変更を破棄して Sync', value: 'discard', variant: 'danger' },
            { label: '中止', value: 'cancel' },
          ],
          cancelValue: 'cancel',
        });
        if (value === 'cancel') {
          log('warning', 'Sync: 中止しました');
          return null;
        }
        if (value === 'build') {
          if (!(await buildFlow(book))) return null;
          log('info', `Sync 開始: ${book}`);
        } else discard = true;
        // 直前の Build が出力したファイルは未コミットになるため、その確認は省く
        r = await runConfirmed(
          'Sync',
          (confirmed) => unwrap(api.sync(root!, book, { confirmed, discardExcelChanges: discard })),
          discard ? [] : ['uncommitted'],
        );
      }
      logResult('Sync', r);
      if (r.status === 'conflict') await showConflicts(r);
      return r.status === 'ok' || r.status === 'conflict' ? r : null;
    },
    [root, runConfirmed, logResult, showConflicts, ask, buildFlow, log],
  );

  /**
   * OneDrive の同期が終わるまで待つ（Windows のみ）。転送中なら最大2分待ち、
   * 一時停止・エラー・OneDrive 未起動なら続けるかを聞く。false なら中止
   */
  const waitForOneDrive = useCallback(
    async (book: string, label: string): Promise<boolean> => {
      if (api.platform !== 'win32' || !root) return true;
      const deadline = Date.now() + 120_000;
      for (;;) {
        const r = await api.syncStatus(root, [book]);
        if (!r.ok) return true;
        mergeSync(r.value);
        const state = r.value.books[book]?.state ?? 'unknown';
        if (state === 'unsupported' || state === 'outside') return true;
        if (r.value.oneDriveRunning === false) {
          return await askContinue(
            'OneDrive が起動していません',
            'このままではブックの変更がクラウドと同期されません。OneDrive を起動してから実行することをおすすめします。',
            label,
          );
        }
        const meta = SYNC_META[state];
        if (meta.busy) {
          if (Date.now() > deadline) {
            return askContinue(
              'OneDrive の同期が終わりません',
              `2分待ちましたが「${meta.label}」のままです。エクスプローラーや OneDrive の画面で状態を確認してください。`,
              label,
            );
          }
          setBusy(`OneDrive の同期を待っています（${meta.label}）`);
          await new Promise((res) => setTimeout(res, 2000));
          continue;
        }
        if (meta.problem) {
          return askContinue(
            `OneDrive: ${meta.label}`,
            '同期されていない変更が失われたり、OneDrive が複製（ブック名-PC名.xlsx）を作ったりする可能性があります。OneDrive の状態を確認してください。',
            label,
          );
        }
        return true;
      }
    },
    // askContinue は下で定義（ask のみに依存）
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [root, mergeSync, ask],
  );

  const askContinue = async (title: string, body: string, label: string): Promise<boolean> => {
    const { value } = await ask({
      title,
      icon: 'warning',
      body: <p>{body}</p>,
      buttons: [
        { label: 'キャンセル', value: false, variant: 'primary' },
        { label: `このまま${label}`, value: true },
      ],
      cancelValue: false,
    });
    if (value) log('warning', `${title}（このまま${label}しました）`);
    return value;
  };

  /**
   * Web 版で開いたブックに、Web 版での編集が届いているか確認する。
   * 開いた時点からブックが変わっていなければ、まだ OneDrive から届いていない可能性がある
   */
  const checkWebEdits = useCallback(
    async (book: string, label: string): Promise<boolean> => {
      const key = `${root}|${book}`;
      const opened = lastOpened.current[key];
      if (opened?.via !== 'web' || !opened.stamp) return true;
      const now = await api.bookStamp(root!, book);
      if (!now.ok || now.value !== opened.stamp) return true;
      const { value } = await ask<'wait' | 'go' | 'cancel'>({
        title: 'Web 版での編集が、まだこの PC に届いていない可能性があります',
        icon: 'warning',
        body: (
          <>
            <p>Web 版で開いてから、このブックはこの PC 上で一度も更新されていません。</p>
            <p className="mt-2">
              Web 版で編集した場合は、OneDrive
              がダウンロードするまで待ってください。何も編集していなければ、このまま続けてかまいません。
            </p>
          </>
        ),
        buttons: [
          { label: '届くまで待つ（最大2分）', value: 'wait', variant: 'primary' },
          { label: `編集していないので${label}`, value: 'go' },
          { label: 'キャンセル', value: 'cancel' },
        ],
        cancelValue: 'cancel',
      });
      if (value === 'cancel') return false;
      if (value === 'go') return true;
      const deadline = Date.now() + 120_000;
      setBusy('Web 版での編集が届くのを待っています');
      while (Date.now() < deadline) {
        await new Promise((res) => setTimeout(res, 2000));
        const s = await api.bookStamp(root!, book);
        if (s.ok && s.value !== opened.stamp) {
          log('info', `Web 版での編集が届きました: ${book}`);
          // ダウンロードの途中かもしれないので、同期の完了も待つ
          return waitForOneDrive(book, label);
        }
      }
      return askContinue('Web 版での編集が届きませんでした', '2分待ちましたが、ブックは更新されていません。', label);
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [root, ask, log, waitForOneDrive],
  );

  /** Build / Sync の後、Web 版で開いていたブックは、比較の基準を今の状態に更新する */
  const refreshStamp = useCallback(
    async (book: string) => {
      const key = `${root}|${book}`;
      const opened = lastOpened.current[key];
      if (opened?.via !== 'web') return;
      const s = await api.bookStamp(root!, book);
      if (s.ok) setOpened(key, { ...opened, stamp: s.value });
    },

    [root],
  );

  const vbaMode = project?.mode === 'vba';

  /**
   * Excel 側・OneDrive 側の準備ができているか（Build / Sync / 開く の前）。
   * VBA モードはブックに書き込まないため、Web 版で閉じたかの確認は省く（編集が届いているかは確認する）
   */
  const ready = useCallback(
    async (book: string, label: string) =>
      (vbaMode || (await confirmWebClosed(book))) &&
      (await waitForOneDrive(book, label)) &&
      (await checkWebEdits(book, label)),
    [confirmWebClosed, waitForOneDrive, checkWebEdits, vbaMode],
  );

  const onBuild = useCallback(
    (book: string) =>
      withBusy('Build 中', async () => {
        if (!(await ready(book, 'Build'))) return;
        setBusy('Build 中');
        await buildFlow(book);
        await refreshStamp(book);
      }),
    [withBusy, ready, buildFlow, refreshStamp],
  );

  const onSync = useCallback(
    (book: string) =>
      withBusy('Sync 中', async () => {
        if (!(await ready(book, 'Sync'))) return;
        setBusy('Sync 中');
        await syncFlow(book);
        await refreshStamp(book);
      }),
    [withBusy, ready, syncFlow, refreshStamp],
  );

  /** 5.5-2: Sync → 起動を1操作にまとめる */
  const onOpenExcel = useCallback(
    (book: string, via: OpenVia) =>
      withBusy(vbaMode ? 'Excel を起動中' : 'Sync して Excel を起動中', async () => {
        if (!vbaMode) {
          if (!(await ready(book, 'Sync'))) return;
          setBusy('Sync して Excel を起動中');
          if (!(await syncFlow(book))) return;
        }
        // Web 版は、Sync で書き換えたブックが OneDrive にアップロードされてから開く（古い内容が表示されないように）
        if (via === 'web' && !(await waitForOneDrive(book, '開く'))) return;
        const opened = await api.openInExcel(root!, book, via);
        if (!opened.ok) {
          if (via !== 'web') throw new Error(opened.error);
          const { value } = await ask({
            title: 'Web 版で開けませんでした',
            icon: 'warning',
            body: <p>{opened.error}</p>,
            buttons: [
              { label: '設定を開く', value: true, variant: 'primary' },
              { label: '閉じる', value: false },
            ],
            cancelValue: false,
          });
          if (value) setView('settings');
          log('error', opened.error);
          return;
        }
        const url = opened.value;
        const stamp = via === 'web' ? await api.bookStamp(root!, book) : null;
        setOpened(`${root}|${book}`, { via, stamp: stamp?.ok ? stamp.value : undefined });
        // Web 版で開いたら、閉じたかの確認を再び出す
        if (via === 'web') webSkip.current.delete(book);
        log('info', via === 'web' ? `Web 版 Excel で開きました: ${url}` : `デスクトップ版 Excel で開きました: ${book}`);
      }),

    [withBusy, ready, syncFlow, root, log, ask, waitForOneDrive, vbaMode],
  );

  const onRefreshTree = useCallback(
    (r: string | null = root) =>
      withBusy(
        'Refresh Tree 中',
        async () => {
          const res = await unwrap(api.refreshTree(r!));
          for (const b of res.books)
            log(
              b.ok ? 'info' : 'error',
              `  ${b.ok ? (b.unchanged ? '最新' : '更新') : '失敗'}: ${b.book}${b.error ? `（${b.error}）` : ''}`,
            );
          const failed = res.books.filter((b) => !b.ok).length;
          if (res.partial)
            log(
              'warning',
              'Refresh Tree: 一部のブックだけ更新されました。#tree が不整合です。Excel を閉じて再実行してください',
            );
          else if (failed > 0) log('error', 'Refresh Tree: 更新できませんでした');
          else {
            const updated = res.books.filter((b) => !b.unchanged).length;
            log(
              'success',
              updated === 0
                ? `Refresh Tree: すべてのブックが最新です（${res.version}）`
                : `Refresh Tree: ${updated} ブックを更新しました（${res.version}）`,
            );
          }
          if (failed > 0) setPanelOpen(true);
        },
        r,
      ),
    [withBusy, root, log],
  );

  // 設定画面を読み直すためのキー（ブックの置き場所を変えた後など）
  const [settingsKey, setSettingsKey] = useState(0);

  /** ブックの置き場所を変える（既存のブックも移動する）。newRoot が null ならソースの中へ戻す */
  const onRelocate = useCallback(
    (newRoot: string | null) =>
      withBusy('ブックを移動中', async () => {
        const count = project?.books.length ?? 0;
        const { value } = await ask({
          title: 'ブックの置き場所を変えますか？',
          icon: 'info',
          body: (
            <>
              <p className="font-mono text-[12px] break-all">
                {project?.bookRoot ?? `${project?.root}（ソースの中）`}
                <br />→ {newRoot ?? `${project?.root}（ソースの中）`}
              </p>
              {count > 0 && (
                <p className="mt-2">
                  既存のブック {count} 冊を、同じフォルダ構成のまま移動します。Excel で開いている場合は閉じてください。
                </p>
              )}
            </>
          ),
          buttons: [
            { label: count > 0 ? '移動して変更' : '変更', value: true, variant: 'primary' },
            { label: 'キャンセル', value: false },
          ],
          cancelValue: false,
        });
        if (!value) return;
        const r = await unwrap(api.relocateBooks(root!, newRoot));
        for (const m of r.moved) log('info', `  移動: ${m}`);
        log('success', `ブックの置き場所を変更しました: ${r.bookRoot ?? 'ソースの中'}`);
        webSkip.current.clear();
        setSettingsKey((k) => k + 1);
      }),
    [withBusy, ask, project, root, log],
  );

  /** プロジェクトの種類を選ぶ（最初のブックを作るとき）。キャンセルなら null */
  const chooseMode = useCallback(async (): Promise<ProjectMode | null> => {
    const { value } = await ask<ProjectMode | null>({
      title: 'プロジェクトの種類を選んでください',
      icon: 'info',
      body: (
        <div className="flex flex-col gap-3">
          <div>
            <div className="text-fg">ソースコード</div>
            <div className="text-[12px] text-muted">
              シートをソースコードのファイルとして書き出します（Build / Sync）。1 シート = 1 ファイル。
            </div>
          </div>
          <div>
            <div className="text-fg">VBA</div>
            <div className="text-[12px] text-muted">
              UI・データのシートと、.bas / .cls / .frm のシートを 1 冊で編集します。Build
              すると、拡張子付きのシートを除き VBA を書き込んだ .xlsm をプロジェクトのフォルダに作ります。Windows
              のデスクトップ版 Excel と「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」の設定が必要です。
            </div>
          </div>
          <div className="text-[12px] text-faint">あとから変えることはできません。</div>
        </div>
      ),
      buttons: [
        { label: 'ソースコード', value: 'source', variant: 'primary' },
        { label: 'VBA', value: 'vba', variant: 'primary' },
        { label: 'キャンセル', value: null },
      ],
      cancelValue: null,
    });
    return value;
  }, [ask]);

  const onCreateBook = useCallback(
    (dirRel: string) =>
      withBusy('ブック作成中', async () => {
        const name = (dirRel.split('/').pop() || project?.name) ?? '';
        let pmode = project?.mode ?? 'source';
        if (project && !project.modeSet) {
          const chosen = await chooseMode();
          if (!chosen) return;
          await unwrap(api.setProjectMode(root!, chosen));
          log('info', `プロジェクトの種類: ${chosen === 'vba' ? 'VBA' : 'ソースコード'}`);
          pmode = chosen;
        }
        const { value } = await ask({
          title: 'ブックを作成しますか？',
          icon: 'info',
          body: (
            <>
              <p>
                <span className="font-mono break-all">
                  {project?.bookRoot ? `${project.bookRoot}/` : ''}
                  {dirRel ? `${dirRel}/` : ''}
                  {name}.xlcode.xlsx
                </span>{' '}
                を作成し、
                {pmode === 'vba'
                  ? `Agents.md などのルールのシートを入れます（${dirRel ? `${dirRel}/` : ''}vba/ に控えがあれば、そのコードもシートにします）。Build 結果は ${dirRel ? `${dirRel}/` : ''}${name}.xlsm です。`
                  : `${dirRel || 'ルート'} のファイルをシートとして取り込みます。`}
              </p>
              <p className="mt-2 text-muted">
                あわせて .gitignore に xlCode 用の除外（*.xlcode.xlsx, ~$*,
                .xlcode/）を追記し、Agents.md・LocalAgents.md が無ければ雛形を作成します。
              </p>
            </>
          ),
          buttons: [
            { label: '作成', value: true, variant: 'primary' },
            { label: 'キャンセル', value: false },
          ],
          cancelValue: false,
        });
        if (!value) return;
        for (const l of await unwrap(api.initProject(root!))) log('info', `  ${l}`);
        const r = await unwrap(api.createBook(root!, dirRel));
        log('success', `ブックを作成しました: ${r.book}（${r.sheets.length} シート）`);
        for (const s of r.skipped) log('warning', `  スキップ: ${s}`);
        setSelected(r.book);
        setView('books');
      }),
    [withBusy, ask, root, project, log, chooseMode],
  );

  const openProject = useCallback(
    async (p: string) => {
      if (hasDrafts) {
        const { value } = await ask({
          title: '保存していない Markdown があります',
          icon: 'warning',
          body: (
            <ul className="rounded-[3px] border border-line bg-editor px-3 py-1.5 font-mono text-[12px]">
              {Object.keys(drafts).map((r) => (
                <li key={r}>{r}</li>
              ))}
            </ul>
          ),
          buttons: [
            { label: '保存せずにプロジェクトを開く', value: true, variant: 'danger' },
            { label: 'キャンセル', value: false },
          ],
          cancelValue: false,
        });
        if (!value) return;
        setDrafts({});
      }
      setRoot(p);
      setProject(null);
      setSelected(null);
      webSkip.current.clear();
      const next = [p, ...recent.filter((x) => x !== p)].slice(0, 8);
      setRecent(next);
      storageSet('recent', next);
      log('info', `プロジェクトを開きました: ${p}`);
      if (excelMode === null) {
        const { value } = await ask<ExcelMode>({
          title: 'どの Excel でブックを編集しますか？',
          icon: 'info',
          body: (
            <div className="flex flex-col gap-2">
              {EXCEL_MODES.map((m) => (
                <div key={m.id}>
                  <div className="text-fg">{m.label}</div>
                  <div className="text-[12px] text-muted">{m.desc}</div>
                </div>
              ))}
              <div className="text-[12px] text-faint">あとから設定やステータスバーで変更できます。</div>
            </div>
          ),
          buttons: [
            { label: '両方', value: 'both', variant: 'primary' },
            { label: 'デスクトップ版のみ', value: 'desktop' },
            { label: 'Web 版のみ', value: 'web' },
          ],
          cancelValue: 'both',
        });
        setExcelMode(value);
      }
      // 6.4: プロジェクトを開いたときに Refresh Tree を自動実行
      await onRefreshTree(p);
    },
    [recent, log, onRefreshTree, excelMode, ask, setExcelMode, hasDrafts, drafts],
  );

  const pickProject = useCallback(async () => {
    const p = await api.pickProject();
    if (p) await openProject(p);
  }, [openProject]);

  // ウィンドウに戻ったとき（Excel での作業後など）に状態を更新する
  useEffect(() => {
    let t: ReturnType<typeof setTimeout> | undefined;
    const onFocus = () => {
      clearTimeout(t);
      t = setTimeout(() => {
        if (!busyRef.current) void reload();
      }, 300);
    };
    window.addEventListener('focus', onFocus);
    return () => {
      window.removeEventListener('focus', onFocus);
      clearTimeout(t);
    };
  }, [reload]);

  // 保存していない Markdown があるときにウィンドウを閉じようとしたら、メインプロセスで確認する
  useEffect(() => {
    if (!hasDrafts) return;
    const onBeforeUnload = (e: BeforeUnloadEvent) => {
      e.preventDefault();
      e.returnValue = '';
    };
    window.addEventListener('beforeunload', onBeforeUnload);
    return () => window.removeEventListener('beforeunload', onBeforeUnload);
  }, [hasDrafts]);

  // OneDrive の同期状態を定期的に取得する（Windows のみ。ウィンドウが前面のときだけ）
  useEffect(() => {
    if (!project || api.platform !== 'win32') return;
    const rels = project.books.map((b) => b.rel);
    if (rels.length === 0) return;
    let stopped = false;
    const tick = async () => {
      if (stopped || busyRef.current || !document.hasFocus()) return;
      const r = await api.syncStatus(project.root, rels);
      if (!stopped && r.ok) setSyncReport(r.value);
    };
    void tick();
    const timer = setInterval(tick, 10_000);
    return () => {
      stopped = true;
      clearInterval(timer);
    };
  }, [project]);

  const logSyncDiagnostics = useCallback(() => {
    if (!syncReport) {
      log('info', 'OneDrive の同期状態はまだ取得していません（Windows のみ）');
      return;
    }
    log(
      'info',
      `OneDrive: ${syncReport.oneDriveRunning === null ? '確認していません' : syncReport.oneDriveRunning ? '起動中' : '起動していません'}`,
    );
    for (const [rel, s] of Object.entries(syncReport.books)) {
      log('info', `  ${rel}: ${SYNC_META[s.state].label || s.state}${s.detail ? ` ${s.detail}` : ''}`);
    }
    setPanelOpen(true);
    setPanelTab('output');
  }, [syncReport, log]);

  // デスクトップ版 Excel の開閉を監視し、変わったら状態を更新する
  useEffect(() => {
    if (!project || mode === 'web') return;
    const rels = project.books.map((b) => b.rel);
    if (rels.length === 0) return;
    const timer = setInterval(async () => {
      if (busyRef.current) return;
      const r = await api.bookLocks(project.root, rels);
      if (!r.ok) return;
      const changed = project.books.filter((b) => r.value[b.rel] !== undefined && r.value[b.rel] !== b.open);
      if (changed.length === 0) return;
      for (const b of changed)
        log('info', r.value[b.rel] ? `Excel で開かれました: ${b.rel}` : `Excel が閉じられました: ${b.rel}`);
      void reload();
    }, 2500);
    return () => clearInterval(timer);
  }, [project, mode, reload, log]);

  // キーボードショートカット
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const mod = e.ctrlKey || e.metaKey;
      if (!mod) return;
      const k = e.key.toLowerCase();
      if (k === 'o') {
        e.preventDefault();
        void pickProject();
      } else if (k === 'j') {
        e.preventDefault();
        setPanelOpen((v) => !v);
      } else if (k === 'b' && selected) {
        e.preventDefault();
        void onBuild(selected);
      } else if (k === 's' && e.shiftKey && selected && !vbaMode) {
        e.preventDefault();
        void onSync(selected);
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [pickProject, onBuild, onSync, selected, vbaMode]);

  const problems = useMemo<Problem[]>(() => {
    if (!project) return [];
    return project.books.flatMap((b) => [
      ...(b.loadError ? [{ level: 'error' as const, source: b.rel, text: b.loadError }] : []),
      ...(b.errors ?? []).map((text) => ({ level: 'error' as const, source: b.rel, text })),
      ...(b.warnings ?? []).map((text) => ({ level: 'warning' as const, source: b.rel, text })),
    ]);
  }, [project]);

  const book = project?.books.find((b) => b.rel === selected) ?? null;
  const isMac = api.platform === 'darwin';

  return (
    <div className="flex h-full flex-col">
      {/* タイトルバー */}
      <div
        className={clsx(
          'drag flex h-[35px] shrink-0 items-center border-b border-line bg-side',
          isMac ? 'pl-20' : 'pl-3',
          'pr-[140px]',
        )}
      >
        <Icon.Excel size={16} className="text-excel" />
        <div className="flex-1 text-center text-[12px] text-muted">
          {project ? `${project.name} — xlCode` : 'xlCode'}
        </div>
      </div>

      <div className="flex min-h-0 flex-1">
        {/* アクティビティバー */}
        <nav className="flex w-12 shrink-0 flex-col border-r border-line bg-side" aria-label="ビュー">
          {(
            [
              { id: 'books', icon: <Icon.Files size={24} />, title: 'エクスプローラー' },
              { id: 'rules', icon: <Icon.Rules size={24} />, title: 'ルール（Agents.md）' },
            ] as const
          ).map((it) => (
            <button
              key={it.id}
              title={it.title}
              aria-label={it.title}
              disabled={!project}
              onClick={() => setView(it.id)}
              className={clsx(
                'flex h-12 items-center justify-center border-l-2 disabled:opacity-40',
                view === it.id && project
                  ? 'border-fg-strong text-fg-strong'
                  : 'border-transparent text-faint hover:text-fg',
              )}
            >
              {it.icon}
            </button>
          ))}
          <div className="flex-1" />
          <button
            title="設定"
            aria-label="設定"
            disabled={!project}
            onClick={() => setView('settings')}
            className={clsx(
              'flex h-12 items-center justify-center border-l-2 disabled:opacity-40',
              view === 'settings' && project
                ? 'border-fg-strong text-fg-strong'
                : 'border-transparent text-faint hover:text-fg',
            )}
          >
            <Icon.Gear size={22} />
          </button>
          <button
            title="プロジェクトを開く (Ctrl+O)"
            aria-label="プロジェクトを開く"
            onClick={pickProject}
            className="flex h-12 items-center justify-center text-faint hover:text-fg"
          >
            <Icon.Folder size={22} />
          </button>
        </nav>

        {/* サイドバー */}
        {project && view !== 'settings' && (
          <aside className="w-[280px] shrink-0 border-r border-line">
            {view === 'books' ? (
              <Sidebar
                project={project}
                selected={selected}
                busy={busy !== null}
                onSelect={setSelected}
                sync={syncReport?.books ?? {}}
                onRefreshTree={() => onRefreshTree()}
                onReload={() => withBusy('再読み込み中', async () => {})}
                onCreateBook={onCreateBook}
              />
            ) : (
              <RulesList project={project} selected={ruleRel} drafts={drafts} onSelect={setRuleRel} />
            )}
          </aside>
        )}

        {/* エディタ領域 */}
        <main className="flex min-w-0 flex-1 flex-col bg-editor">
          {project && (
            <div className="flex h-[35px] shrink-0 bg-side">
              <div className="flex items-center gap-1.5 border-t border-t-accent border-r border-r-line bg-editor px-3 text-[13px] text-fg-strong">
                {view === 'books' ? (
                  <>
                    <Icon.Book size={14} className="text-excel" />
                    {book ? book.rel.split('/').pop() : 'ブック'}
                  </>
                ) : view === 'rules' ? (
                  <>
                    <Icon.Rules size={14} className="text-info" />
                    {ruleRel}
                  </>
                ) : (
                  <>
                    <Icon.Gear size={14} />
                    設定
                  </>
                )}
              </div>
              <div className="flex-1 border-b border-line" />
            </div>
          )}
          <div className="min-h-0 flex-1">
            {!root || (!project && !busy) ? (
              <Welcome recent={recent} onOpen={pickProject} onOpenRecent={(p) => void openProject(p)} />
            ) : !project ? (
              <div className="flex h-full items-center justify-center gap-2 text-muted">
                <Icon.Spinner /> 読み込み中...
              </div>
            ) : view === 'settings' ? (
              <SettingsView
                key={settingsKey}
                root={project.root}
                bookRoot={project.bookRoot}
                bookCount={project.books.length}
                onRelocate={onRelocate}
                mode={mode}
                onMode={setExcelMode}
                onError={(m) => log('error', m)}
                onSaved={() => log('success', '設定を保存しました')}
              />
            ) : view === 'rules' ? (
              <RulesEditor
                key={ruleRel}
                root={project.root}
                rel={ruleRel}
                vba={project.mode === 'vba'}
                draft={drafts[ruleRel]}
                onDraft={setDraft}
                onError={(m) => log('error', m)}
                onSaved={(rel) => {
                  log('success', `保存しました: ${rel}`);
                  void reload();
                }}
              />
            ) : book ? (
              <BookView
                book={book}
                treeVersion={project.treeVersion}
                busy={busy !== null}
                mode={mode}
                sync={syncReport?.books[book.rel]}
                oneDriveRunning={syncReport?.oneDriveRunning ?? null}
                onOpenExcel={(via) => void onOpenExcel(book.rel, via)}
                onSync={() => void onSync(book.rel)}
                onBuild={() => void onBuild(book.rel)}
                onTerminal={() =>
                  void unwrap(api.openTerminal(project.root, book.dirRel)).catch((e: Error) => log('error', e.message))
                }
                onReveal={() =>
                  void unwrap(api.revealInFolder(project.root, book.rel)).catch((e: Error) => log('error', e.message))
                }
                onRefreshTree={() => void onRefreshTree()}
                projectMode={project.mode}
                onOpenOutput={(reveal) =>
                  void unwrap(api.openOutput(project.root, book.rel, reveal)).catch((e: Error) =>
                    log('error', e.message),
                  )
                }
              />
            ) : (
              <div className="flex h-full flex-col items-center justify-center gap-2 text-muted">
                <Icon.Book size={40} className="text-faint" />
                <div>ブックがありません</div>
                <div className="text-[12px]">サイドバーの「ブック未作成のディレクトリ」から作成できます</div>
              </div>
            )}
          </div>
          {project && panelOpen && (
            <div className="h-[200px] shrink-0">
              <Panel
                tab={panelTab}
                onTab={setPanelTab}
                logs={logs}
                problems={problems}
                onClose={() => setPanelOpen(false)}
                onClear={() => setLogs([])}
              />
            </div>
          )}
        </main>
      </div>

      <StatusBar
        project={project}
        busy={busy}
        errors={problems.filter((p) => p.level === 'error').length}
        warnings={problems.filter((p) => p.level === 'warning').length}
        mode={mode}
        onMode={() => setView('settings')}
        sync={syncReport}
        onSync={logSyncDiagnostics}
        onProblems={() => {
          setPanelOpen(true);
          setPanelTab('problems');
        }}
      />
      {dialog}
    </div>
  );
}
