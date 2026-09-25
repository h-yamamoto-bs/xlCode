import clsx from 'clsx';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { Confirmation, OpResult } from '../../core';
import type { ExcelMode, OpenVia, ProjectInfo } from '../../shared/api';
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

type View = 'books' | 'rules' | 'settings';

const ACTION_LABEL: Record<string, string> = {
  'write-file': 'ファイル出力',
  'delete-file': 'ファイル削除',
  'write-sheet': 'シート更新',
  'delete-sheet': 'シート削除',
  'reformat-sheet': 'シート整形',
  'conflict-sheet': '衝突シート作成',
  commit: '自動コミット',
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
  // ブックを最後に xlCode からどちらで開いたか（キー: ルート|ブック）
  const lastOpened = useRef<Record<string, OpenVia>>(storageGet('lastOpened', {}));
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
      if (mode === 'both' && lastOpened.current[`${root}|${book}`] === 'desktop') return true;
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
      const danger = cs.some((c) => c.kind === 'delete' || c.kind === 'shrink');
      const { value } = await ask({
        title: `${label} の前に確認してください`,
        icon: 'warning',
        body: (
          <div className="flex flex-col gap-3">
            {cs.map((c) => (
              <div key={c.kind}>
                <div className={clsx(c.kind === 'shrink' || c.kind === 'delete' ? 'text-warn' : 'text-fg')}>
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

  const onBuild = useCallback(
    (book: string) =>
      withBusy('Build 中', async () => {
        if (await confirmWebClosed(book)) await buildFlow(book);
      }),
    [withBusy, confirmWebClosed, buildFlow],
  );

  const onSync = useCallback(
    (book: string) =>
      withBusy('Sync 中', async () => {
        if (await confirmWebClosed(book)) await syncFlow(book);
      }),
    [withBusy, confirmWebClosed, syncFlow],
  );

  /** 5.5-2: Sync → 起動を1操作にまとめる */
  const onOpenExcel = useCallback(
    (book: string, via: OpenVia) =>
      withBusy('Sync して Excel を起動中', async () => {
        if (!(await confirmWebClosed(book))) return;
        if (!(await syncFlow(book))) return;
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
        lastOpened.current = { ...lastOpened.current, [`${root}|${book}`]: via };
        storageSet('lastOpened', lastOpened.current);
        // Web 版で開いたら、閉じたかの確認を再び出す
        if (via === 'web') webSkip.current.delete(book);
        log('info', via === 'web' ? `Web 版 Excel で開きました: ${url}` : `デスクトップ版 Excel で開きました: ${book}`);
        if (via === 'web') log('info', '  OneDrive への同期が終わる前に開くと、古い内容が表示されることがあります');
      }),
    [withBusy, confirmWebClosed, syncFlow, root, log, ask],
  );

  const onRefreshTree = useCallback(
    (r: string | null = root) =>
      withBusy(
        'Refresh Tree 中',
        async () => {
          const res = await unwrap(api.refreshTree(r!));
          for (const b of res.books)
            log(b.ok ? 'info' : 'error', `  ${b.ok ? '更新' : '失敗'}: ${b.book}${b.error ? `（${b.error}）` : ''}`);
          const failed = res.books.filter((b) => !b.ok).length;
          if (res.partial)
            log(
              'warning',
              'Refresh Tree: 一部のブックだけ更新されました。#tree が不整合です。Excel を閉じて再実行してください',
            );
          else if (failed > 0) log('error', 'Refresh Tree: 更新できませんでした');
          else log('success', `Refresh Tree: ${res.books.length} ブックを更新しました（${res.version}）`);
          if (failed > 0) setPanelOpen(true);
        },
        r,
      ),
    [withBusy, root, log],
  );

  const onCreateBook = useCallback(
    (dirRel: string) =>
      withBusy('ブック作成中', async () => {
        const name = (dirRel.split('/').pop() || project?.name) ?? '';
        const { value } = await ask({
          title: 'ブックを作成しますか？',
          icon: 'info',
          body: (
            <>
              <p>
                <span className="font-mono">
                  {dirRel ? `${dirRel}/` : ''}
                  {name}.xlcode.xlsx
                </span>{' '}
                を作成し、ディレクトリ内のファイルをシートとして取り込みます。
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
    [withBusy, ask, root, project, log],
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
      } else if (k === 's' && e.shiftKey && selected) {
        e.preventDefault();
        void onSync(selected);
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [pickProject, onBuild, onSync, selected]);

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
                root={project.root}
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
        onProblems={() => {
          setPanelOpen(true);
          setPanelTab('problems');
        }}
      />
      {dialog}
    </div>
  );
}
