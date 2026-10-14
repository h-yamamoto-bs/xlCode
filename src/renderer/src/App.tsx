import clsx from 'clsx';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { CommitInfo, Confirmation, OpResult, ProjectMode } from '../../core';
import type { BookSummary, ExcelMode, OpenVia, OpenWith, ProjectInfo, SyncReport } from '../../shared/api';
import { api, storageGet, storageSet, unwrap } from './api';
import { chordPrefixes, commandLabel, findCommand, type Command } from './commands';
import { ACTION_LABEL, BookView, type LastResult } from './components/BookView';
import { useDialog } from './components/Dialogs';
import { DocEditor, type Draft } from './components/DocEditor';
import { docKind } from './docs';
import { HelpView } from './components/HelpView';
import { Icon } from './components/Icons';
import { EmptySidebar } from './components/EmptySidebar';
import { FileView } from './components/FileView';
import { Palette, type PaletteItem, type PaletteSpec } from './components/Palette';
import { Resizer } from './components/Resizer';
import { NOTES_REL, SECONDARY_TABS, SecondarySidebar, type SecondaryTab } from './components/SecondarySidebar';
import { Panel, type LogEntry, type Problem } from './components/Panel';
import { EXCEL_MODES, SettingsView, type RelocateMode } from './components/SettingsView';
import { Sidebar, type OpenTarget } from './components/Sidebar';
import { StatusBar } from './components/StatusBar';
import { TabBar, type TabView } from './components/TabBar';
import { IconButton, Kbd } from './components/ui';
import { Welcome } from './components/Welcome';
import { allowedWhileTyping, eventToKey, formatKey, isTyping } from './keys';
import { SOURCE_SIDE, SYNC_META } from './status';
import {
  activate,
  activateIndex,
  activeTab,
  closeAll,
  closeOthers,
  closeTab,
  cycleTab,
  deserializeTabs,
  EMPTY_TABS,
  fileRef,
  moveTab,
  parseFileRef,
  openTab,
  pinTab,
  pruneTabs,
  reopenClosed,
  serializeTabs,
  tabId,
  type Tab,
  type TabsState,
} from './tabs';

/** 開く前に Sync が要るか（ソース側の変更・衝突がブックにまだ反映されていない） */
function needsSyncBeforeOpen(b: BookSummary): boolean {
  if (b.loadError) return false;
  return (b.files ?? []).some((f) => SOURCE_SIDE.includes(f.status) || f.status === 'conflict');
}

/** xlCode からブックを開いた方法と、Web 版で開いた時点のブックの状態 */
interface Opened {
  via: OpenVia;
  stamp?: string;
}

function now(): string {
  return new Date().toLocaleTimeString('ja-JP', { hour12: false });
}

/** LocalAgents.md の相対パス */
function localAgentsOf(dirRel: string): string {
  return dirRel ? `${dirRel}/LocalAgents.md` : 'LocalAgents.md';
}

/** 左右のサイドバーの開閉と幅 */
interface Layout {
  sidebarOpen: boolean;
  sidebarWidth: number;
  rightOpen: boolean;
  rightWidth: number;
  rightTab: SecondaryTab;
}
const DEFAULT_LAYOUT: Layout = {
  sidebarOpen: true,
  sidebarWidth: 280,
  rightOpen: false,
  rightWidth: 340,
  rightTab: 'prompts',
};

export function App() {
  const [root, setRoot] = useState<string | null>(null);
  const [project, setProject] = useState<ProjectInfo | null>(null);
  // ---- タブ（エディタ領域に開いているもの） ----
  const [tabs, setTabsState] = useState<TabsState>(EMPTY_TABS);
  const tabsRef = useRef(tabs);
  tabsRef.current = tabs;
  const setTabs = useCallback((fn: (s: TabsState) => TabsState) => setTabsState((s) => fn(s)), []);
  const active = activeTab(tabs);
  /** 表示中のブック（Build / Sync などのショートカットの対象） */
  const selected =
    active?.kind === 'book' ? (active.rel ?? null) : active?.kind === 'file' ? parseFileRef(active.rel!).book : null;
  const open = useCallback(
    (kind: Tab['kind'], rel?: string, opts: { preview?: boolean } = {}) =>
      setTabs((s) => openTab(s, kind, rel, { preview: opts.preview ?? false })),
    [setTabs],
  );
  // ヘルプで表示する見出し（nonce は同じ見出しへもう一度移動するため）
  const [help, setHelp] = useState<{ anchor?: string; nonce: number }>({ nonce: 0 });
  const openHelp = useCallback(
    (anchor?: string) => {
      setHelp((h) => ({ anchor, nonce: h.nonce + 1 }));
      open('help');
    },
    [open],
  );
  const [settingsFocus, setSettingsFocus] = useState<'bookRoot' | undefined>(undefined);
  const openSettings = useCallback(
    (focus?: 'bookRoot') => {
      setSettingsFocus(focus);
      open('settings');
    },
    [open],
  );
  // ---- 画面の配置（左右のサイドバーの開閉と幅。このパソコンで覚えておく） ----
  const [layout, setLayout] = useState<Layout>(() => ({
    ...DEFAULT_LAYOUT,
    ...storageGet<Partial<Layout>>('layout', {}),
  }));
  useEffect(() => storageSet('layout', layout), [layout]);
  const sidebarOpen = layout.sidebarOpen;
  const setSidebarOpen = useCallback(
    (v: boolean | ((prev: boolean) => boolean)) =>
      setLayout((l) => ({ ...l, sidebarOpen: typeof v === 'function' ? v(l.sidebarOpen) : v })),
    [],
  );
  const setRightOpen = useCallback(
    (v: boolean | ((prev: boolean) => boolean)) =>
      setLayout((l) => ({ ...l, rightOpen: typeof v === 'function' ? v(l.rightOpen) : v })),
    [],
  );
  const showRightTab = useCallback((t: SecondaryTab) => setLayout((l) => ({ ...l, rightOpen: true, rightTab: t })), []);
  // ブック画面の一覧で選んでいるファイル（キー: ブック）
  const [selectedFiles, setSelectedFiles] = useState<Record<string, string>>({});
  const [explorerFocus, setExplorerFocus] = useState(0);
  // 保存していない Markdown などの編集内容（ファイルや画面を切り替えても保持する）
  const [drafts, setDrafts] = useState<Record<string, Draft>>({});
  const draftsRef = useRef(drafts);
  draftsRef.current = drafts;
  const setDraft = useCallback(
    (rel: string, d: Draft | null) => {
      setDrafts((prev) => {
        const next = { ...prev };
        if (d) next[rel] = d;
        else delete next[rel];
        return next;
      });
      // 編集を始めたタブはプレビューから固定に変える（VS Code と同じ）
      if (d) setTabs((s) => pinTab(s, tabId('doc', rel)));
    },
    [setTabs],
  );
  const hasDrafts = Object.keys(drafts).length > 0;
  const [logs, setLogs] = useState<LogEntry[]>([]);
  // ブックごとの直前の操作の結果（画面に出して、何が起きたかをその場で示す）
  const [lastResults, setLastResults] = useState<Record<string, LastResult>>({});
  const setLastResult = useCallback((book: string, r: LastResult | null) => {
    setLastResults((prev) => {
      const next = { ...prev };
      if (r) next[book] = r;
      else delete next[book];
      return next;
    });
  }, []);
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
  // コマンドパレット / クイックオープン
  const [palette, setPalette] = useState<PaletteSpec | null>(null);
  // 2 段階のショートカット（Ctrl+K …）の 1 段目
  const [chord, setChord] = useState<string | null>(null);

  const log = useCallback((level: LogEntry['level'], text: string) => {
    setLogs((l) => [...l.slice(-500), { time: now(), level, text }]);
  }, []);

  /** 無くなったブックなどのタブを閉じる */
  const pruneForProject = useCallback(
    (p: ProjectInfo) =>
      setTabs((s) =>
        pruneTabs(s, (t) => {
          if (t.kind === 'book') return p.books.some((b) => b.rel === t.rel);
          if (t.kind === 'file') {
            const { book, name } = parseFileRef(t.rel ?? '');
            return p.books.some((b) => b.rel === book && (b.files ?? []).some((f) => f.name === name));
          }
          if (t.kind === 'doc') {
            const rel = t.rel ?? '';
            if (!rel.includes('/')) return true;
            return p.books.some((b) => localAgentsOf(b.dirRel) === rel);
          }
          return true;
        }),
      ),
    [setTabs],
  );

  const reload = useCallback(
    async (r: string | null = root) => {
      if (!r) return;
      try {
        const p = await unwrap(api.loadProject(r));
        setProject(p);
        pruneForProject(p);
      } catch (e) {
        log('error', `読み込みに失敗しました: ${(e as Error).message}`);
      }
    },
    [root, log, pruneForProject],
  );

  // 開いているタブをプロジェクトごとに覚えておく（次に開いたときに復元する）
  useEffect(() => {
    if (root && project) storageSet(`tabs:${root}`, serializeTabs(tabs));
  }, [tabs, root, project]);

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
    (label: string, book: string, r: OpResult) => {
      setLastResult(book, { label, time: now(), result: r });
      for (const c of r.changes) log('info', `  ${ACTION_LABEL[c.action] ?? c.action}: ${c.target}`);
      for (const w of r.warnings) log('warning', `  警告: ${w}`);
      for (const e of r.errors) log('error', `  エラー: ${e}`);
      const written = r.changes.some((c) => c.action !== 'commit');
      const summary: Record<OpResult['status'], [LogEntry['level'], string]> = {
        ok: ['success', r.changes.length === 0 ? '変更なし' : '完了'],
        error: ['error', written ? '途中で中断しました（上の変更までは書き換えています）' : '中断しました'],
        confirm: ['warning', 'キャンセルしました'],
        conflict: ['warning', '衝突シートを作成しました'],
        'needs-decision': ['warning', 'キャンセルしました'],
      };
      const [level, text] = summary[r.status];
      log(level, `${label}: ${text}`);
      if (r.status === 'error') setPanelOpen(true);
    },
    [log, setLastResult],
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
      logResult('Build', book, r);
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
      logResult('Sync', book, r);
      if (r.status === 'conflict') await showConflicts(r);
      return r.status === 'ok' || r.status === 'conflict' ? r : null;
    },
    [root, runConfirmed, logResult, showConflicts, ask, buildFlow, log],
  );

  const askContinue = useCallback(
    async (title: string, body: string, label: string): Promise<boolean> => {
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
    },
    [ask, log],
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
    [root, mergeSync, askContinue],
  );

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
    [root, ask, log, waitForOneDrive, askContinue],
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

  /** Excel 側・OneDrive 側の準備ができているか（Build / Sync / 開く の前） */
  const ready = useCallback(
    async (book: string, label: string) =>
      (await confirmWebClosed(book)) && (await waitForOneDrive(book, label)) && (await checkWebEdits(book, label)),
    [confirmWebClosed, waitForOneDrive, checkWebEdits],
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

  /**
   * 5.5-2: Sync → 起動を1操作にまとめる。
   * ブックに反映していないソースの変更が無ければ、ブックを書き換えないので確認も Sync もせずにそのまま開く
   */
  const onOpenExcel = useCallback(
    (book: string, via: OpenVia) =>
      withBusy('Excel を起動中', async () => {
        const latest = await api.loadProject(root!);
        const summary = (latest.ok ? latest.value : project)?.books.find((b) => b.rel === book);
        const needsSync = !summary || needsSyncBeforeOpen(summary);
        if (needsSync) {
          setBusy('Sync して Excel を起動中');
          if (!(await ready(book, 'Sync'))) return;
          setBusy('Sync して Excel を起動中');
          if (!(await syncFlow(book))) return;
          // Web 版は、Sync で書き換えたブックが OneDrive にアップロードされてから開く（古い内容が表示されないように）
          if (via === 'web' && !(await waitForOneDrive(book, '開く'))) return;
          setBusy('Excel を起動中');
        }
        const opened = await api.openInExcel(root!, book, via);
        if (!opened.ok) {
          if (via !== 'web') throw new Error(opened.error);
          log('error', opened.error);
          const { value } = await ask({
            title: 'Web 版で開けませんでした',
            icon: 'warning',
            body: <p className="break-all whitespace-pre-wrap">{opened.error}</p>,
            buttons: [
              { label: '設定を開く', value: true, variant: 'primary' },
              { label: '閉じる', value: false },
            ],
            cancelValue: false,
          });
          if (value) openSettings();
          return;
        }
        const url = opened.value;
        const stamp = via === 'web' ? await api.bookStamp(root!, book) : null;
        setOpened(`${root}|${book}`, { via, stamp: stamp?.ok ? stamp.value : undefined });
        // Web 版で開いたら、閉じたかの確認を再び出す
        if (via === 'web') webSkip.current.delete(book);
        log(
          'info',
          via === 'web'
            ? `Web 版 Excel で開きました: ${url ?? `${book}（OneDrive の「オンラインで表示」）`}`
            : `デスクトップ版 Excel で開きました: ${book}`,
        );
      }),
    [withBusy, ready, syncFlow, root, project, log, ask, waitForOneDrive, openSettings],
  );

  /** 直前の Build / Sync を元に戻す（ゴール: 戻せない操作だけ確認する。これは戻せないので確認する） */
  const onUndo = useCallback(
    (book: string) =>
      withBusy('元に戻しています', async () => {
        const info = project?.books.find((b) => b.rel === book)?.undo;
        if (!info) return;
        const when = new Date(info.at).toLocaleString('ja-JP', { hour12: false });
        const { value } = await ask({
          title: `${when} の ${info.label} を元に戻しますか？`,
          icon: 'warning',
          body: (
            <>
              <p>
                ブック・state.json
                {info.output && (
                  <>
                    ・ビルド結果 <span className="font-mono">{info.output}</span>
                  </>
                )}
                {info.files.length > 0 && '・次のソースファイル'}を、{info.label} の直前の状態に戻します。
              </p>
              {info.files.length > 0 && (
                <ul className="my-2 rounded-[3px] border border-line bg-editor px-3 py-1.5 font-mono text-[12px]">
                  {info.files.map((f) => (
                    <li key={f}>{f}</li>
                  ))}
                </ul>
              )}
              <p className="mt-2 text-muted">
                {info.label} の後にこれらのファイルやブックへ加えた編集は失われます。戻せるのは直前の 1 回だけです（Git
                の自動コミットはそのまま残ります）。
              </p>
            </>
          ),
          buttons: [
            { label: '元に戻す', value: true, variant: 'danger' },
            { label: 'キャンセル', value: false },
          ],
          cancelValue: false,
        });
        if (!value) return;
        const r = await unwrap(api.undoLast(root!, book));
        for (const f of r.restored) log('info', `  書き戻し: ${f}`);
        for (const f of r.removed) log('info', `  削除: ${f}`);
        log('success', `${r.label} を元に戻しました: ${book}`);
        setLastResult(book, null);
        webSkip.current.delete(book);
      }),
    [withBusy, project, ask, root, log, setLastResult],
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

  /**
   * ブックの置き場所を変える。move なら既存のブックも移動し、keep なら設定だけを変える。
   * newRoot が null ならソースの中へ戻す
   */
  const onRelocate = useCallback(
    (newRoot: string | null, relocateMode: RelocateMode = 'move') =>
      withBusy('ブックを移動中', async () => {
        const count = project?.books.length ?? 0;
        const move = relocateMode === 'move' && count > 0;
        const { value } = await ask({
          title: 'ブックの置き場所を変えますか？',
          icon: 'info',
          body: (
            <>
              <p className="font-mono text-[12px] break-all">
                {project?.bookRoot ?? `${project?.root}（ソースの中）`}
                <br />→ {newRoot ?? `${project?.root}（ソースの中）`}
              </p>
              {move && (
                <p className="mt-2">
                  既存のブック {count} 冊を、同じフォルダ構成のまま移動します。Excel で開いている場合は閉じてください。
                </p>
              )}
              {!move && count > 0 && (
                <p className="mt-2">
                  ブックは移動しません。新しい置き場所にあるブックを使います（元の場所のブック {count}
                  冊はそのまま残ります。不要なら自分で削除してください）。
                </p>
              )}
            </>
          ),
          buttons: [
            { label: move ? '移動して変更' : '変更', value: true, variant: 'primary' },
            { label: 'キャンセル', value: false },
          ],
          cancelValue: false,
        });
        if (!value) return;
        const r = await unwrap(api.relocateBooks(root!, newRoot, { move }));
        for (const m of r.moved) log('info', `  移動: ${m}`);
        log('success', `ブックの置き場所を変更しました: ${r.bookRoot ?? 'ソースの中'}`);
        webSkip.current.clear();
        setSettingsKey((k) => k + 1);
      }),
    [withBusy, ask, project, root, log],
  );

  /**
   * 最初のブックを作るときに、編集用ブックの置き場所を決める（Web 版・Copilot で開けるよう OneDrive を勧める）。
   * 置き場所（未設定ならプロジェクトの中は null）を返す。キャンセルなら undefined
   */
  const chooseBookRoot = useCallback(async (): Promise<string | null | undefined> => {
    if (!project || project.books.length > 0 || project.bookRoot) return project?.bookRoot ?? null;
    const { value } = await ask<'onedrive' | 'inside' | null>({
      title: '編集用ブックをどこに置きますか？',
      icon: 'info',
      body: (
        <div className="flex flex-col gap-2">
          <p>
            Web 版の Excel と Copilot で開けるよう、
            <span className="text-fg">OneDrive の中の、このプロジェクト専用のフォルダ</span>
            をおすすめします（ソースコード・Git はこのままプロジェクトのフォルダに置きます）。
          </p>
          <p className="text-[12px] text-muted">
            例: OneDrive\xlCode\{project.name}
            （フォルダの選択画面で新しく作れます）。あとから設定画面で変えることもできます。
          </p>
        </div>
      ),
      buttons: [
        { label: 'OneDrive のフォルダを選ぶ…', value: 'onedrive', variant: 'primary' },
        { label: 'プロジェクトの中に置く', value: 'inside' },
        { label: 'キャンセル', value: null },
      ],
      cancelValue: null,
    });
    if (!value) return undefined;
    if (value === 'inside') return null;
    const dir = await api.pickFolder(`編集用ブックの置き場所（OneDrive の中の「${project.name}」専用のフォルダ）`);
    if (!dir) return undefined;
    const r = await unwrap(api.relocateBooks(root!, dir));
    log('info', `編集用ブックの置き場所: ${r.bookRoot}`);
    setSettingsKey((k) => k + 1);
    return r.bookRoot;
  }, [ask, project, root, log]);

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

  /**
   * ブックを作る前に、ブックにしないファイルの指定（.xlcodeignore）が無ければ確認する。
   * .gitignore の除外だけに従ってよいか、.gitignore も無ければ全ファイルがシートになることを知らせる。
   * 続けてよければ true。「先に編集」なら .xlcodeignore を開いて false
   */
  const confirmIgnore = useCallback(
    async (dirs: string[]): Promise<boolean> => {
      if (!project || !root) return true;
      // 保存した直後でも最新の状態を見る（画面の再読み込みを待たない）
      const latest = await api.loadProject(root);
      const ignoreInfo = latest.ok ? latest.value.ignore : project.ignore;
      if ((ignoreInfo.xlcode ?? 0) > 0) return true;
      if (storageGet<Record<string, boolean>>('ignoreConfirmed', {})[root]) return true;
      // 作る対象のファイルを数える（多いときは最初の数フォルダだけ）
      let sheets = 0;
      let ignored = 0;
      let nonCode = 0;
      const samples: string[] = [];
      for (const d of dirs.slice(0, 50)) {
        const r = await api.previewBook(root, d);
        if (!r.ok) continue;
        sheets += r.value.sheets.length;
        ignored += r.value.ignored.length;
        nonCode += r.value.nonCode.length;
        for (const f of r.value.sheets) if (samples.length < 8) samples.push(d ? `${d}/${f}` : f);
      }
      const git = ignoreInfo.git;
      const { value, checked } = await ask<'go' | 'edit' | null>({
        title: 'ブックにしないファイルの指定がありません',
        icon: git === null ? 'warning' : 'info',
        body: (
          <div className="flex flex-col gap-2">
            <p>
              <span className="font-mono">.xlcodeignore</span> に除外の指定が
              {ignoreInfo.xlcode === null ? 'ありません（ファイルがありません）' : '書かれていません'}。
            </p>
            {git === null ? (
              <p className="text-warn">
                .gitignore
                もありません。フォルダの中のすべてのファイル（生成ファイルや依存パッケージがあればそれも）がシートになります。
              </p>
            ) : (
              <p>
                .gitignore の除外（{git} 件
                {ignoreInfo.gitSample.length > 0 && (
                  <>
                    : <span className="font-mono">{ignoreInfo.gitSample.join(' ')}</span>
                    {git > ignoreInfo.gitSample.length && ' …'}
                  </>
                )}
                ）だけに従ってブックを作ります。ドキュメントや生成ファイルなど、Git
                では管理するがブックにはしたくないものがあれば、先に .xlcodeignore に書いてください。
              </p>
            )}
            <p className="text-[12px] text-muted">
              {dirs.length > 1 ? `${dirs.length} 個のフォルダで` : ''}シートになるファイル: {sheets} 件
              {ignored > 0 && `（除外: ${ignored} 件）`}
              {nonCode > 0 && `（拡張子が無いため対象外: ${nonCode} 件）`}
              {samples.length > 0 && (
                <span className="mt-1 block font-mono text-faint">
                  {samples.join('  ')}
                  {sheets > samples.length && ' …'}
                </span>
              )}
            </p>
          </div>
        ),
        checkbox: 'このプロジェクトでは今後確認しない',
        buttons: [
          { label: '先に .xlcodeignore を編集', value: 'edit', variant: git === null ? 'primary' : 'secondary' },
          { label: 'このまま作成', value: 'go', variant: git === null ? 'secondary' : 'primary' },
          { label: 'キャンセル', value: null },
        ],
        cancelValue: null,
      });
      if (value === 'edit') {
        open('doc', '.xlcodeignore');
        log('info', '.xlcodeignore を編集してから、もう一度ブックを作成してください');
        return false;
      }
      if (value === 'go' && checked) {
        storageSet('ignoreConfirmed', { ...storageGet<Record<string, boolean>>('ignoreConfirmed', {}), [root]: true });
      }
      return value === 'go';
    },
    [project, root, ask, open, log],
  );

  const onCreateBook = useCallback(
    (dirRel: string) =>
      withBusy('ブック作成中', async () => {
        if (!(await confirmIgnore([dirRel]))) return;
        const name = (dirRel.split('/').pop() || project?.name) ?? '';
        let pmode = project?.mode ?? 'source';
        if (project && !project.modeSet) {
          const chosen = await chooseMode();
          if (!chosen) return;
          await unwrap(api.setProjectMode(root!, chosen));
          log('info', `プロジェクトの種類: ${chosen === 'vba' ? 'VBA' : 'ソースコード'}`);
          pmode = chosen;
        }
        const bookRoot = await chooseBookRoot();
        if (bookRoot === undefined) return;
        const bookPath = `${bookRoot ? `${bookRoot}/` : ''}${dirRel ? `${dirRel}/` : ''}${name}.xlcode.xlsx`;
        const { value } = await ask<'empty' | 'tool' | null>({
          title: 'ブックを作成しますか？',
          icon: 'info',
          body: (
            <>
              <p>
                <span className="font-mono break-all">{bookPath}</span> を作成します。
              </p>
              {pmode === 'vba' ? (
                <ul className="mt-2 list-disc pl-5">
                  <li>
                    <span className="text-fg">既存の Excel ツールから作成</span>
                    ：.xlsm などを選ぶと、シート（画面・データ）はそのまま、VBA
                    のモジュール・フォーム・参照設定をシートにして取り込みます。元のファイルは変更しません（Windows
                    のデスクトップ版 Excel が必要）。
                  </li>
                  <li>
                    <span className="text-fg">空のブックを作成</span>：{dirRel || 'ルート'} に .bas / .cls / .frm
                    があればそれをシートにし、無ければ Module1.bas と References.refs の雛形を置きます。
                  </li>
                  <li>
                    Build 結果は {dirRel ? `${dirRel}/` : ''}
                    {name}.xlsm です。作ったあと、画面の「次の操作」に従って Build してください。
                  </li>
                </ul>
              ) : (
                <p className="mt-2">
                  {dirRel || 'ルート'} のファイルをシートとして取り込みます（.gitignore / .xlcodeignore
                  で除外したものは除きます）。
                </p>
              )}
              <p className="mt-2 text-muted">
                あわせて .gitignore に xlCode 用の除外（*.xlcode.xlsx, ~$*,
                .xlcode/）を追記し、Agents.md・LocalAgents.md・.xlcodeignore が無ければ雛形を作成します。
              </p>
            </>
          ),
          buttons:
            pmode === 'vba'
              ? [
                  { label: '既存の Excel ツールから作成…', value: 'tool', variant: 'primary' },
                  { label: '空のブックを作成', value: 'empty' },
                  { label: 'キャンセル', value: null },
                ]
              : [
                  { label: '作成', value: 'empty', variant: 'primary' },
                  { label: 'キャンセル', value: null },
                ],
          cancelValue: null,
        });
        if (!value) return;
        let tool: string | null = null;
        if (value === 'tool') {
          tool = await api.pickToolFile();
          if (!tool) return;
        }
        for (const l of await unwrap(api.initProject(root!))) log('info', `  ${l}`);
        const label = tool ? '取り込み' : 'ブック作成';
        if (tool) {
          setBusy('Excel ツールを取り込み中（画面に出ない Excel を使います）');
          log('info', `取り込み開始: ${tool}`);
        }
        const res = await (tool ? api.importTool(root!, dirRel, tool) : api.createBook(root!, dirRel));
        if (!res.ok) {
          // 失敗の理由と対処を、出力パネルだけでなくその場で示す
          log('error', `${label}: ${res.error}`);
          await ask({
            title: tool ? '取り込めませんでした' : 'ブックを作成できませんでした',
            icon: 'error',
            body: (
              <>
                <p className="whitespace-pre-wrap">{res.error}</p>
                <p className="mt-2 text-[12px] text-muted">
                  {tool
                    ? '元のファイルは変更していません。原因を直してから、もう一度 ＋ →「既存の Excel ツールから作成…」を選んでください。'
                    : 'ブックは作成していません。原因を直してから、もう一度 ＋ を押してください。'}
                </p>
              </>
            ),
            buttons: [{ label: '閉じる', value: true, variant: 'primary' }],
            cancelValue: true,
          });
          return;
        }
        const r = res.value;
        log(
          'success',
          `ブックを作成しました: ${r.book}（${r.sheets.length} シート${tool ? '：' + r.sheets.join(', ') : ''}）`,
        );
        for (const s of r.skipped) log('warning', `  ${tool ? '注意' : 'スキップ'}: ${s}`);
        if (tool)
          log(
            'info',
            '  取り込んだ内容を確認してから Build してください（最初の Build では、既存の .xlsm を上書きする確認が出ます）',
          );
        // 何をシートにし、何を取り込めなかったかを、ブック画面にそのまま出す
        setLastResult(r.book, {
          label,
          time: now(),
          title: tool
            ? `${tool.split(/[\\/]/).pop()} から取り込みました（${r.sheets.length} シート）`
            : `ブックを作成しました（${r.sheets.length} シート）`,
          result: {
            status: 'ok',
            errors: [],
            warnings: r.skipped,
            confirmations: [],
            changes: r.sheets.map((s) => ({ action: 'create-sheet', target: s })),
            conflicts: [],
            unbuilt: [],
          },
        });
        open('book', r.book);
      }),
    [withBusy, ask, root, project, log, chooseMode, chooseBookRoot, setLastResult, open, confirmIgnore],
  );

  /** ブックの無いディレクトリすべてにブックを作る（ソースコードモード） */
  const onCreateAll = useCallback(
    () =>
      withBusy('ブック作成中', async () => {
        const dirs = project?.dirsWithoutBook ?? [];
        if (!project || dirs.length === 0) return;
        if (!(await confirmIgnore(dirs))) return;
        if (!project.modeSet) {
          const chosen = await chooseMode();
          if (!chosen) return;
          await unwrap(api.setProjectMode(root!, chosen));
          if (chosen === 'vba') {
            log('info', 'プロジェクトの種類: VBA（ブックは 1 つずつ作成してください）');
            return;
          }
        }
        if ((await chooseBookRoot()) === undefined) return;
        const { value } = await ask({
          title: `${dirs.length} 個のディレクトリにブックを作成しますか？`,
          icon: 'info',
          body: (
            <>
              <p>それぞれのディレクトリのファイルを、シートとして取り込みます。</p>
              <ul className="mt-2 max-h-[40vh] overflow-auto rounded-[3px] border border-line bg-editor px-3 py-1.5 font-mono text-[12px]">
                {dirs.map((d) => (
                  <li key={d}>{d || `${project.name}（ルート）`}</li>
                ))}
              </ul>
            </>
          ),
          buttons: [
            { label: 'すべて作成', value: true, variant: 'primary' },
            { label: 'キャンセル', value: false },
          ],
          cancelValue: false,
        });
        if (!value) return;
        for (const l of await unwrap(api.initProject(root!))) log('info', `  ${l}`);
        let made = 0;
        for (const d of dirs) {
          const r = await api.createBook(root!, d);
          if (!r.ok) {
            log('error', `  ${d || 'ルート'}: ${r.error}`);
            continue;
          }
          made++;
          log('info', `  作成: ${r.value.book}（${r.value.sheets.length} シート）`);
          for (const s of r.value.skipped) log('warning', `    スキップ: ${s}`);
        }
        log(made === dirs.length ? 'success' : 'warning', `ブックを ${made} / ${dirs.length} 冊作成しました`);
      }),
    [withBusy, ask, root, project, log, chooseMode, chooseBookRoot, confirmIgnore],
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
      // 前回開いていたタブを復元する
      setTabsState(deserializeTabs(storageGet(`tabs:${p}`, null)));
      setLastResults({});
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

  // プロジェクトを開いた直後、タブが何も無ければ最初のブックをプレビューで開いておく（迷わないように）。
  // 利用者がすべてのタブを閉じたあとの再読み込みでは開き直さない
  const autoOpenedFor = useRef<string | null>(null);
  useEffect(() => {
    if (!project || autoOpenedFor.current === project.root) return;
    autoOpenedFor.current = project.root;
    setTabs((s) =>
      s.tabs.length === 0 && project.books[0] ? openTab(s, 'book', project.books[0].rel, { preview: true }) : s,
    );
  }, [project, setTabs]);

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

  // ---- タブの操作 ----

  /** タブを閉じる。保存していない編集があれば確認する */
  const requestClose = useCallback(
    async (id: string) => {
      const t = tabsRef.current.tabs.find((x) => x.id === id);
      if (!t) return;
      const draft = t.kind === 'doc' && t.rel ? draftsRef.current[t.rel] : undefined;
      if (draft && t.rel) {
        const rel = t.rel;
        const { value } = await ask<'save' | 'discard' | 'cancel'>({
          title: `${rel} を保存しますか？`,
          icon: 'warning',
          body: <p>保存しないと、編集内容は失われます。</p>,
          buttons: [
            { label: '保存', value: 'save', variant: 'primary' },
            { label: '保存しない', value: 'discard', variant: 'danger' },
            { label: 'キャンセル', value: 'cancel' },
          ],
          cancelValue: 'cancel',
        });
        if (value === 'cancel') return;
        if (value === 'save') {
          try {
            await unwrap(api.writeDoc(root!, rel, draft.text));
            log('success', `保存しました: ${rel}`);
          } catch (e) {
            log('error', (e as Error).message);
            return;
          }
        }
        setDraft(rel, null);
      }
      setTabs((s) => closeTab(s, id));
    },
    [ask, root, log, setDraft, setTabs],
  );

  const onOpenTarget = useCallback(
    (target: OpenTarget, opts: { preview: boolean }) => {
      if (target.kind === 'file') {
        setSelectedFiles((m) => ({ ...m, [target.book]: target.name }));
        open('file', fileRef(target.book, target.name), { preview: opts.preview });
      } else open(target.kind, target.rel, { preview: opts.preview });
    },
    [open],
  );

  /** ブック画面の一覧からファイルを開く（Markdown は編集、それ以外は閲覧） */
  const openFileOf = useCallback(
    (b: BookSummary, name: string) => {
      if (name.toLowerCase().endsWith('.md')) open('doc', b.dirRel ? `${b.dirRel}/${name}` : name);
      else open('file', fileRef(b.rel, name));
    },
    [open],
  );

  /** ソースファイルを既定のアプリ / VS Code で開く、またはフォルダで表示する */
  const openWithApp = useCallback(
    (b: BookSummary, name: string, app: OpenWith | 'reveal') => {
      const rel = b.dirRel ? `${b.dirRel}/${name}` : name;
      const p = app === 'reveal' ? api.revealFile(root!, rel) : api.openFile(root!, rel, app);
      void unwrap(p)
        .then(() =>
          log(
            'info',
            `${app === 'reveal' ? 'フォルダで表示' : app === 'vscode' ? 'VS Code で開きました' : '既定のアプリで開きました'}: ${rel}`,
          ),
        )
        .catch((e: Error) => log('error', e.message));
    },
    [root, log],
  );

  // プロジェクト全体の最近のコミット（サイドバーの「Git」）。状態を読み直すたびに取り直す
  const [commits, setCommits] = useState<CommitInfo[] | null>(null);
  useEffect(() => {
    if (!project) {
      setCommits(null);
      return;
    }
    let stale = false;
    api.recentCommits(project.root, null).then((r) => {
      if (!stale) setCommits(r.ok ? r.value : []);
    });
    return () => {
      stale = true;
    };
  }, [project]);

  // ---- コマンド（ショートカット・コマンドパレット） ----

  const book = project?.books.find((b) => b.rel === selected) ?? null;
  const isBusy = busy !== null;
  const canOperate = !!book && !isBusy && !book.open && !book.loadError;

  const openQuickOpen = useCallback(
    (initial = '') => {
      if (!project) return;
      const items: PaletteItem[] = [
        ...project.books.map((b) => ({
          id: `book:${b.rel}`,
          label: b.rel.split('/').pop()!,
          detail: b.dirRel || undefined,
          icon: <Icon.Book size={14} className="text-excel" />,
          run: () => open('book', b.rel),
        })),
        {
          id: 'doc:Agents.md',
          label: 'Agents.md',
          detail: 'プロジェクト共通のルール',
          icon: <Icon.Rules size={14} className="text-info" />,
          run: () => open('doc', 'Agents.md'),
        },
        ...project.books.map((b) => ({
          id: `doc:${localAgentsOf(b.dirRel)}`,
          label: 'LocalAgents.md',
          detail: b.dirRel || '.',
          icon: <Icon.Rules size={14} className="text-info" />,
          run: () => open('doc', localAgentsOf(b.dirRel)),
        })),
        {
          id: 'doc:.xlcodeignore',
          label: '.xlcodeignore',
          detail: 'ブックにしないファイル',
          icon: <Icon.EyeOff size={14} className="text-warn" />,
          run: () => open('doc', '.xlcodeignore'),
        },
        { id: 'settings', label: '設定', icon: <Icon.Gear size={14} />, run: () => openSettings() },
        { id: 'help', label: 'ヘルプ', icon: <Icon.Help size={14} />, run: () => openHelp() },
      ];
      setPalette({
        placeholder: 'ブック・ドキュメントの名前で検索（先頭に > でコマンド）',
        initial,
        items,
        // コマンドの一覧は開くときに作る（下で定義）
        commands: [],
      });
    },
    [project, open, openSettings, openHelp],
  );

  const openRecent = useCallback(() => {
    setPalette({
      placeholder: '最近使用したプロジェクト',
      items: [
        ...recent.map((p) => ({
          id: p,
          label: p.split(/[\\/]/).pop()!,
          detail: p,
          icon: <Icon.Folder size={14} className="text-muted" />,
          run: () => void openProject(p),
        })),
        {
          id: 'pick',
          label: 'プロジェクトを開く…',
          keys: formatKey('ctrl+o', api.platform),
          icon: <Icon.Folder size={14} className="text-link" />,
          run: () => void pickProject(),
        },
      ],
    });
  }, [recent, openProject, pickProject]);

  const openCreatePicker = useCallback(() => {
    if (!project) return;
    setPalette({
      placeholder: 'ブックを作成するディレクトリ',
      items: project.dirsWithoutBook.map((d) => ({
        id: d,
        label: d || `${project.name}（ルート）`,
        icon: <Icon.Folder size={14} className="text-muted" />,
        run: () => void onCreateBook(d),
      })),
    });
  }, [project, onCreateBook]);

  const commands = useMemo<Command[]>(() => {
    const hasProject = !!project;
    const excelOpen = (via: OpenVia): Command => ({
      id: `excel-${via}`,
      category: 'ブック',
      title: via === 'desktop' ? 'デスクトップ版 Excel で開く' : 'Web 版 Excel で開く',
      when: canOperate && (via === 'desktop' ? mode !== 'web' : mode !== 'desktop'),
      run: () => void onOpenExcel(book!.rel, via),
    });
    const list: Command[] = [
      {
        id: 'open-project',
        category: 'プロジェクト',
        title: 'プロジェクトを開く…',
        keys: ['ctrl+o', 'ctrl+k ctrl+o'],
        run: () => void pickProject(),
      },
      {
        id: 'recent',
        category: 'プロジェクト',
        title: '最近使用したプロジェクトを開く…',
        keys: ['ctrl+r'],
        run: openRecent,
      },
      {
        id: 'quick-open',
        category: '移動',
        title: 'ブック・ドキュメントを開く…（クイックオープン）',
        keys: ['ctrl+p', 'ctrl+e'],
        when: hasProject,
        run: () => openQuickOpen(''),
      },
      {
        id: 'palette',
        category: '移動',
        title: 'コマンドパレット',
        keys: ['ctrl+shift+p'],
        when: hasProject,
        run: () => openQuickOpen('>'),
      },
      {
        id: 'build',
        category: 'ブック',
        title: 'Build（Excel → ソース）',
        keys: ['ctrl+shift+b'],
        when: canOperate,
        run: () => void onBuild(book!.rel),
      },
      {
        id: 'sync',
        category: 'ブック',
        title: 'Sync（ソース → Excel）',
        keys: ['ctrl+shift+s'],
        when: canOperate,
        run: () => void onSync(book!.rel),
      },
      excelOpen('desktop'),
      excelOpen('web'),
      {
        id: 'undo',
        category: 'ブック',
        title: `直前の ${book?.undo?.label ?? 'Build / Sync'} を元に戻す`,
        when: canOperate && !!book?.undo,
        run: () => void onUndo(book!.rel),
      },
      {
        id: 'refresh-tree',
        category: 'プロジェクト',
        title: 'Refresh Tree（全ブックの #tree を更新）',
        when: hasProject && !isBusy,
        run: () => void onRefreshTree(),
      },
      {
        id: 'reload',
        category: 'プロジェクト',
        title: '再読み込み',
        keys: ['f5'],
        when: hasProject && !isBusy,
        run: () => void withBusy('再読み込み中', async () => {}),
      },
      {
        id: 'create-book',
        category: 'プロジェクト',
        title: 'ブックを作成…',
        when: hasProject && !isBusy && (project?.dirsWithoutBook.length ?? 0) > 0,
        run: openCreatePicker,
      },
      {
        id: 'create-all',
        category: 'プロジェクト',
        title: 'すべてのディレクトリにブックを作成',
        when: hasProject && !isBusy && project?.mode !== 'vba' && (project?.dirsWithoutBook.length ?? 0) > 1,
        run: () => void onCreateAll(),
      },
      {
        id: 'terminal',
        category: 'ブック',
        title: 'ターミナルをこのフォルダで開く',
        keys: ['ctrl+`', 'ctrl+@'],
        when: !!book && !isBusy,
        run: () =>
          void unwrap(api.openTerminal(project!.root, book!.dirRel)).catch((e: Error) => log('error', e.message)),
      },
      {
        id: 'reveal',
        category: 'ブック',
        title: 'ブックのフォルダを表示',
        when: !!book && !isBusy,
        run: () =>
          void unwrap(api.revealInFolder(project!.root, book!.rel)).catch((e: Error) => log('error', e.message)),
      },
      {
        id: 'agents',
        category: 'ドキュメント',
        title: 'Agents.md を開く（プロジェクト共通のルール）',
        when: hasProject,
        run: () => open('doc', 'Agents.md'),
      },
      {
        id: 'local-agents',
        category: 'ドキュメント',
        title: 'LocalAgents.md を開く（このブックのルール）',
        when: !!book,
        run: () => open('doc', localAgentsOf(book!.dirRel)),
      },
      {
        id: 'notes',
        category: 'ドキュメント',
        title: '作業メモを開く（.xlcode/notes.md）',
        when: hasProject,
        run: () => showRightTab('notes'),
      },
      {
        id: 'open-default',
        category: 'ファイル',
        title: '既定のアプリで開く（閲覧中のファイル）',
        when: !!book && active?.kind === 'file',
        run: () => openWithApp(book!, parseFileRef(active!.rel!).name, 'default'),
      },
      {
        id: 'open-vscode',
        category: 'ファイル',
        title: 'VS Code で開く（閲覧中のファイル）',
        when: !!book && active?.kind === 'file',
        run: () => openWithApp(book!, parseFileRef(active!.rel!).name, 'vscode'),
      },
      {
        id: 'ignore',
        category: 'ドキュメント',
        title: '.xlcodeignore を開く（ブックにしないファイル）',
        when: hasProject,
        run: () => open('doc', '.xlcodeignore'),
      },
      {
        id: 'close-tab',
        category: 'タブ',
        title: 'タブを閉じる',
        keys: ['ctrl+w', 'ctrl+f4'],
        when: !!active,
        run: () => void requestClose(active!.id),
      },
      {
        id: 'close-others',
        category: 'タブ',
        title: 'ほかのタブを閉じる',
        when: tabs.tabs.length > 1,
        run: () => setTabs((s) => closeOthers(s, s.active!)),
      },
      {
        id: 'close-all',
        category: 'タブ',
        title: 'すべてのタブを閉じる',
        keys: ['ctrl+k ctrl+w'],
        when: tabs.tabs.length > 0,
        run: () => setTabs((s) => closeAll(s)),
      },
      {
        id: 'reopen',
        category: 'タブ',
        title: '閉じたタブを開き直す',
        keys: ['ctrl+shift+t'],
        when: tabs.closed.length > 0,
        run: () => setTabs((s) => reopenClosed(s)),
      },
      {
        id: 'next-tab',
        category: 'タブ',
        title: '次のタブ',
        keys: ['ctrl+tab', 'ctrl+pagedown'],
        when: tabs.tabs.length > 1,
        run: () => setTabs((s) => cycleTab(s, 1)),
      },
      {
        id: 'prev-tab',
        category: 'タブ',
        title: '前のタブ',
        keys: ['ctrl+shift+tab', 'ctrl+pageup'],
        when: tabs.tabs.length > 1,
        run: () => setTabs((s) => cycleTab(s, -1)),
      },
      {
        id: 'pin-tab',
        category: 'タブ',
        title: 'タブを固定する（プレビューをやめる）',
        when: !!active?.preview,
        run: () => setTabs((s) => pinTab(s, s.active!)),
      },
      ...Array.from({ length: 9 }, (_, i): Command => ({
        id: `tab-${i + 1}`,
        title: `${i + 1} 番目のタブ`,
        keys: [`alt+${i + 1}`, `ctrl+${i + 1}`],
        hidden: true,
        when: tabs.tabs.length > i,
        run: () => setTabs((s) => activateIndex(s, i)),
      })),
      {
        id: 'toggle-sidebar',
        category: '表示',
        title: '左のサイドバーの表示 / 非表示',
        keys: ['ctrl+b'],
        run: () => setSidebarOpen((v) => !v),
      },
      {
        id: 'toggle-secondary',
        category: '表示',
        title: '右のパネルの表示 / 非表示（指示・内容・履歴）',
        keys: ['ctrl+alt+b'],
        run: () => setRightOpen((v) => !v),
      },
      ...SECONDARY_TABS.map((t): Command => ({
        id: `secondary-${t.id}`,
        category: '表示',
        title: `右のパネル: ${t.label}を表示`,
        run: () => showRightTab(t.id),
      })),
      {
        id: 'focus-explorer',
        category: '表示',
        title: 'エクスプローラーにフォーカス',
        keys: ['ctrl+shift+e'],
        when: hasProject,
        run: () => {
          setSidebarOpen(true);
          setExplorerFocus((n) => n + 1);
        },
      },
      {
        id: 'toggle-panel',
        category: '表示',
        title: 'パネルの表示 / 非表示',
        keys: ['ctrl+j'],
        when: hasProject,
        run: () => setPanelOpen((v) => !v),
      },
      {
        id: 'problems',
        category: '表示',
        title: '問題を表示',
        keys: ['ctrl+shift+m'],
        when: hasProject,
        run: () => {
          setPanelOpen(true);
          setPanelTab('problems');
        },
      },
      {
        id: 'output',
        category: '表示',
        title: '出力を表示',
        keys: ['ctrl+shift+u'],
        when: hasProject,
        run: () => {
          setPanelOpen(true);
          setPanelTab('output');
        },
      },
      { id: 'clear-output', category: '表示', title: '出力をクリア', when: hasProject, run: () => setLogs([]) },
      {
        id: 'settings',
        category: '表示',
        title: '設定を開く',
        keys: ['ctrl+,'],
        when: hasProject,
        run: () => openSettings(),
      },
      {
        id: 'book-root',
        category: '設定',
        title: 'ブックの置き場所を変える…',
        when: hasProject,
        run: () => openSettings('bookRoot'),
      },
      { id: 'help', category: 'ヘルプ', title: 'ヘルプを開く', keys: ['f1'], run: () => openHelp() },
      {
        id: 'shortcuts',
        category: 'ヘルプ',
        title: 'キーボードショートカットの一覧',
        keys: ['ctrl+k ctrl+s'],
        run: () => openHelp('ショートカット'),
      },
      {
        id: 'onedrive',
        category: 'OneDrive',
        title: '同期状態を出力に表示',
        when: hasProject,
        run: logSyncDiagnostics,
      },
      ...EXCEL_MODES.map((m): Command => ({
        id: `excel-mode-${m.id}`,
        category: '設定',
        title: `使う Excel: ${m.label}`,
        when: mode !== m.id,
        run: () => setExcelMode(m.id),
      })),
    ];
    return list;
  }, [
    project,
    book,
    canOperate,
    isBusy,
    mode,
    active,
    tabs,
    pickProject,
    openRecent,
    openQuickOpen,
    onBuild,
    onSync,
    onOpenExcel,
    onUndo,
    onRefreshTree,
    withBusy,
    openCreatePicker,
    onCreateAll,
    log,
    open,
    requestClose,
    setTabs,
    openSettings,
    openHelp,
    logSyncDiagnostics,
    setExcelMode,
    setSidebarOpen,
    setRightOpen,
    showRightTab,
    openWithApp,
  ]);
  const commandsRef = useRef(commands);
  commandsRef.current = commands;

  // クイックオープンのコマンド一覧は、開いた時点の最新のコマンドから作る
  const paletteSpec = useMemo<PaletteSpec | null>(() => {
    if (!palette || palette.commands === undefined) return palette;
    const cmds: PaletteItem[] = commands
      .filter((c) => c.when !== false && !c.hidden)
      .map((c) => ({
        id: c.id,
        label: commandLabel(c),
        keys: c.keys?.[0] ? formatKey(c.keys[0], api.platform) : undefined,
        icon: <Icon.Command size={14} className="text-muted" />,
        run: c.run,
      }));
    return { ...palette, commands: cmds };
  }, [palette, commands]);

  // キーボードショートカット
  const prefixes = useMemo(() => chordPrefixes(commands), [commands]);
  const dialogOpen = dialog !== null;
  useEffect(() => {
    if (!chord) return;
    const t = setTimeout(() => setChord(null), 3000);
    return () => clearTimeout(t);
  }, [chord]);
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (palette || dialogOpen || e.isComposing) return;
      const key = eventToKey(e, api.platform);
      if (!key) return;
      if (isTyping(e.target) && !allowedWhileTyping(key)) return;
      if (chord) {
        setChord(null);
        const c = findCommand(commandsRef.current, `${chord} ${key}`);
        e.preventDefault();
        if (c) c.run();
        return;
      }
      if (prefixes.has(key)) {
        e.preventDefault();
        setChord(key);
        return;
      }
      const c = findCommand(commandsRef.current, key);
      if (c) {
        e.preventDefault();
        c.run();
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [palette, dialogOpen, chord, prefixes]);

  const problems = useMemo<Problem[]>(() => {
    if (!project) return [];
    return project.books.flatMap((b) => [
      ...(b.loadError ? [{ level: 'error' as const, source: b.rel, text: b.loadError }] : []),
      ...(b.errors ?? []).map((text) => ({ level: 'error' as const, source: b.rel, text })),
      ...(b.warnings ?? []).map((text) => ({ level: 'warning' as const, source: b.rel, text })),
    ]);
  }, [project]);

  // ---- タブバーの表示 ----
  const tabViews = useMemo<TabView[]>(() => {
    const labelOf = (t: Tab) =>
      t.kind === 'book'
        ? t.rel!.split('/').pop()!
        : t.kind === 'doc'
          ? t.rel!.split('/').pop()!
          : t.kind === 'file'
            ? parseFileRef(t.rel!).name
            : t.kind === 'settings'
              ? '設定'
              : 'ヘルプ';
    const counts = new Map<string, number>();
    for (const t of tabs.tabs) counts.set(labelOf(t), (counts.get(labelOf(t)) ?? 0) + 1);
    return tabs.tabs.map((t) => {
      const label = labelOf(t);
      const dir =
        t.kind === 'file'
          ? parseFileRef(t.rel!).book.split('/').slice(0, -1).join('/')
          : t.rel?.includes('/')
            ? t.rel.slice(0, t.rel.lastIndexOf('/'))
            : '';
      const dup = (counts.get(label) ?? 0) > 1;
      const icon =
        t.kind === 'book' ? (
          <Icon.Book size={14} className="text-excel" />
        ) : t.kind === 'doc' ? (
          docKind(t.rel!) === 'ignore' ? (
            <Icon.EyeOff size={14} className="text-warn" />
          ) : (
            <Icon.Rules size={14} className="text-info" />
          )
        ) : t.kind === 'file' ? (
          <Icon.File size={14} className="text-muted" />
        ) : t.kind === 'settings' ? (
          <Icon.Gear size={14} />
        ) : (
          <Icon.Help size={14} className="text-info" />
        );
      return {
        tab: t,
        label,
        detail: (t.kind === 'doc' && label === 'LocalAgents.md') || t.kind === 'file' || dup ? dir || '.' : undefined,
        icon,
        dirty: t.kind === 'doc' && !!drafts[t.rel!],
        title: t.kind === 'file' ? `${dir ? `${dir}/` : ''}${label}（閲覧）` : t.rel,
      };
    });
  }, [tabs.tabs, drafts]);

  const isMac = api.platform === 'darwin';
  const key = (b: string) => formatKey(b, api.platform);

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
        <Icon.Logo size={22} />
        <div className="flex-1 text-center text-[12px] text-muted">
          {project ? `${project.name} — xlCode` : 'xlCode'}
        </div>
      </div>

      <div className="flex min-h-0 flex-1">
        {/* アクティビティバー */}
        <nav className="flex w-12 shrink-0 flex-col border-r border-line bg-side" aria-label="ビュー">
          <button
            title={`エクスプローラー (${key('ctrl+shift+e')})`}
            aria-label="エクスプローラー"
            // 表示中にもう一度押すとサイドバーを隠す（VS Code と同じ）
            onClick={() => setSidebarOpen((v) => !v)}
            className={clsx(
              'flex h-12 items-center justify-center border-l-2',
              sidebarOpen ? 'border-fg-strong text-fg-strong' : 'border-transparent text-faint hover:text-fg',
            )}
          >
            <Icon.Files size={24} />
          </button>
          <button
            title={`検索・コマンド (${key('ctrl+p')} / ${key('ctrl+shift+p')})`}
            aria-label="検索・コマンド"
            disabled={!project}
            onClick={() => openQuickOpen('')}
            className="flex h-12 items-center justify-center border-l-2 border-transparent text-faint hover:text-fg disabled:opacity-40"
          >
            <Icon.Search size={22} />
          </button>
          <div className="flex-1" />
          <button
            title={`ヘルプ (${key('f1')})`}
            aria-label="ヘルプ"
            onClick={() => openHelp()}
            className={clsx(
              'flex h-12 items-center justify-center border-l-2',
              active?.kind === 'help'
                ? 'border-fg-strong text-fg-strong'
                : 'border-transparent text-faint hover:text-fg',
            )}
          >
            <Icon.Help size={22} />
          </button>
          <button
            title={`設定 (${key('ctrl+,')})`}
            aria-label="設定"
            disabled={!project}
            onClick={() => openSettings()}
            className={clsx(
              'flex h-12 items-center justify-center border-l-2 disabled:opacity-40',
              active?.kind === 'settings'
                ? 'border-fg-strong text-fg-strong'
                : 'border-transparent text-faint hover:text-fg',
            )}
          >
            <Icon.Gear size={22} />
          </button>
          <button
            title={`プロジェクトを開く (${key('ctrl+o')})`}
            aria-label="プロジェクトを開く"
            onClick={pickProject}
            className="flex h-12 items-center justify-center text-faint hover:text-fg"
          >
            <Icon.Folder size={22} />
          </button>
        </nav>

        {/* サイドバー（プロジェクトを開いていなければ「プロジェクトを開く」） */}
        {sidebarOpen && (
          <>
            <aside className="shrink-0 border-r border-line" style={{ width: layout.sidebarWidth }}>
              {project ? (
                <Sidebar
                  project={project}
                  selected={
                    active && (active.kind === 'book' || active.kind === 'doc' || active.kind === 'file')
                      ? active.id
                      : null
                  }
                  busy={busy !== null}
                  onOpen={onOpenTarget}
                  sync={syncReport?.books ?? {}}
                  onRefreshTree={() => onRefreshTree()}
                  onReload={() => withBusy('再読み込み中', async () => {})}
                  onCreateBook={onCreateBook}
                  onCreateAll={project.mode === 'vba' && project.modeSet ? undefined : () => void onCreateAll()}
                  onSettings={() => openSettings('bookRoot')}
                  focusNonce={explorerFocus}
                  vba={project.mode === 'vba' && project.modeSet}
                  oneDriveRunning={syncReport?.oneDriveRunning ?? null}
                  onSyncDetails={logSyncDiagnostics}
                  platform={api.platform}
                  commits={commits}
                  onTerminal={() =>
                    void unwrap(api.openTerminal(project.root, '')).catch((e: Error) => log('error', e.message))
                  }
                />
              ) : (
                <EmptySidebar
                  recent={recent}
                  onOpen={() => void pickProject()}
                  onOpenRecent={(p) => void openProject(p)}
                  platform={api.platform}
                />
              )}
            </aside>
            <Resizer
              side="left"
              width={layout.sidebarWidth}
              label="サイドバーの幅"
              onResize={(w) => setLayout((l) => ({ ...l, sidebarWidth: w }))}
              onReset={() => setLayout((l) => ({ ...l, sidebarWidth: DEFAULT_LAYOUT.sidebarWidth }))}
            />
          </>
        )}

        {/* エディタ領域 */}
        <main className="flex min-w-0 flex-1 flex-col bg-editor">
          {(project || tabs.tabs.length > 0) && (
            <TabBar
              tabs={tabViews}
              active={tabs.active}
              onActivate={(id) => setTabs((s) => activate(s, id))}
              onClose={(id) => void requestClose(id)}
              onPin={(id) => setTabs((s) => pinTab(s, id))}
              onMove={(id, to) => setTabs((s) => moveTab(s, id, to))}
              onCloseOthers={(id) => setTabs((s) => closeOthers(s, id))}
              onCloseAll={() => setTabs((s) => closeAll(s))}
              actions={
                <>
                  <IconButton
                    title={`左のサイドバー${sidebarOpen ? 'を隠す' : 'を表示'} (${key('ctrl+b')})`}
                    onClick={() => setSidebarOpen((v) => !v)}
                    className={sidebarOpen ? 'text-fg' : 'text-faint'}
                  >
                    <Icon.PanelLeft size={15} />
                  </IconButton>
                  <IconButton
                    title={`右のパネル${layout.rightOpen ? 'を隠す' : 'を表示'}（指示・内容・履歴） (${key('ctrl+alt+b')})`}
                    onClick={() => setRightOpen((v) => !v)}
                    className={layout.rightOpen ? 'text-fg' : 'text-faint'}
                  >
                    <Icon.PanelRight size={15} />
                  </IconButton>
                </>
              }
            />
          )}
          <div className="min-h-0 flex-1">
            {active?.kind === 'help' ? (
              <HelpView anchor={help.anchor} nonce={help.nonce} />
            ) : !root || (!project && !busy) ? (
              <Welcome
                recent={recent}
                onOpen={pickProject}
                onOpenRecent={(p) => void openProject(p)}
                onHelp={openHelp}
                platform={api.platform}
              />
            ) : !project ? (
              <div className="flex h-full items-center justify-center gap-2 text-muted">
                <Icon.Spinner /> 読み込み中...
              </div>
            ) : active?.kind === 'settings' ? (
              <SettingsView
                key={settingsKey}
                root={project.root}
                bookRoot={project.bookRoot}
                bookCount={project.books.length}
                onRelocate={onRelocate}
                mode={mode}
                onMode={setExcelMode}
                onError={(m) => log('error', m)}
                onSaved={() => {
                  log('success', '設定を保存しました');
                  void reload();
                }}
                onOpenDoc={(rel) => open('doc', rel)}
                onHelp={openHelp}
                focus={settingsFocus}
              />
            ) : active?.kind === 'file' ? (
              book ? (
                <FileView
                  key={active.rel}
                  root={project.root}
                  book={book}
                  name={parseFileRef(active.rel!).name}
                  onOpenWith={(app) => openWithApp(book, parseFileRef(active.rel!).name, app)}
                  onReveal={() => openWithApp(book, parseFileRef(active.rel!).name, 'reveal')}
                  onSelectForPrompt={() =>
                    setSelectedFiles((m) => ({ ...m, [book.rel]: parseFileRef(active.rel!).name }))
                  }
                />
              ) : (
                <div className="flex h-full items-center justify-center text-muted">ブックがありません</div>
              )
            ) : active?.kind === 'doc' ? (
              <DocEditor
                key={active.rel}
                root={project.root}
                rel={active.rel!}
                draft={drafts[active.rel!]}
                onDraft={setDraft}
                onError={(m) => log('error', m)}
                onSaved={(rel) => {
                  log('success', `保存しました: ${rel}`);
                  void reload();
                }}
                onHelp={openHelp}
              />
            ) : book ? (
              <BookView
                key={book.rel}
                root={project.root}
                book={book}
                treeVersion={project.treeVersion}
                busy={busy !== null}
                mode={mode}
                sync={syncReport?.books[book.rel]}
                oneDriveRunning={syncReport?.oneDriveRunning ?? null}
                onOpenExcel={(via) => void onOpenExcel(book.rel, via)}
                onSync={() => void onSync(book.rel)}
                onBuild={() => void onBuild(book.rel)}
                onUndo={() => void onUndo(book.rel)}
                lastResult={lastResults[book.rel]}
                onDismissResult={() => setLastResult(book.rel, null)}
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
                onHelp={openHelp}
                onOpenBackups={() =>
                  void unwrap(api.openBackups(project.root, book.rel)).catch((e: Error) => log('error', e.message))
                }
                onOpenDoc={(rel) => open('doc', rel)}
                platform={api.platform}
                selectedFile={selectedFiles[book.rel] ?? null}
                onOpenFile={(name) => openFileOf(book, name)}
                onSelectFile={(name) =>
                  setSelectedFiles((m) => {
                    const next = { ...m };
                    if (name) next[book.rel] = name;
                    else delete next[book.rel];
                    return next;
                  })
                }
              />
            ) : (
              <div className="flex h-full flex-col items-center justify-center gap-3 text-muted">
                <Icon.Logo size={56} className="opacity-40" />
                {project.books.length === 0 ? (
                  <>
                    <div>ブックがありません</div>
                    <div className="text-[12px]">サイドバーの「ブック未作成のディレクトリ」から作成できます</div>
                  </>
                ) : (
                  <div className="text-[12px]">サイドバーでブックを選ぶか、次のキーで開きます</div>
                )}
                <div className="grid grid-cols-[auto_auto] gap-x-4 gap-y-1.5 text-[12px]">
                  <span className="text-right">ブック・ドキュメントを開く</span>
                  <Kbd>{key('ctrl+p')}</Kbd>
                  <span className="text-right">コマンドパレット</span>
                  <Kbd>{key('ctrl+shift+p')}</Kbd>
                  <span className="text-right">エクスプローラー</span>
                  <Kbd>{key('ctrl+shift+e')}</Kbd>
                  <span className="text-right">ヘルプ</span>
                  <Kbd>{key('f1')}</Kbd>
                </div>
              </div>
            )}
          </div>
          {project && panelOpen && active?.kind !== 'help' && (
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

        {/* 右のサイドバー（指示・内容・履歴） */}
        {layout.rightOpen && (
          <>
            <Resizer
              side="right"
              width={layout.rightWidth}
              label="右のパネルの幅"
              onResize={(w) => setLayout((l) => ({ ...l, rightWidth: w }))}
              onReset={() => setLayout((l) => ({ ...l, rightWidth: DEFAULT_LAYOUT.rightWidth }))}
            />
            <aside
              className="shrink-0 border-l border-line"
              style={{ width: layout.rightWidth }}
              aria-label="右のパネル"
            >
              <SecondarySidebar
                tab={layout.rightTab}
                onTab={showRightTab}
                onClose={() => setRightOpen(false)}
                root={project?.root ?? null}
                book={book}
                projectMode={project?.mode ?? 'source'}
                selectedFile={book ? (selectedFiles[book.rel] ?? null) : null}
                onLog={log}
                onOpenNotesTab={() => open('doc', NOTES_REL)}
              />
            </aside>
          </>
        )}
      </div>

      <StatusBar
        project={project}
        busy={busy}
        errors={problems.filter((p) => p.level === 'error').length}
        warnings={problems.filter((p) => p.level === 'warning').length}
        mode={mode}
        onMode={() => openSettings()}
        sync={syncReport}
        onSync={logSyncDiagnostics}
        onProblems={() => {
          setPanelOpen(true);
          setPanelTab('problems');
        }}
        chord={chord ? formatKey(chord, api.platform) : null}
        onPalette={() => openQuickOpen('>')}
      />
      {paletteSpec && <Palette spec={paletteSpec} onClose={() => setPalette(null)} />}
      {dialog}
    </div>
  );
}
