import clsx from 'clsx';
import { useEffect, useRef, useState } from 'react';
import type { ProjectMode } from '../../../core';
import type { BookSummary } from '../../../shared/api';
import { api, storageGet, storageSet } from '../api';
import { fillPrompt, PROMPT_GROUPS, type PromptTemplate } from '../prompts';
import { Icon } from './Icons';
import { Button, IconButton } from './ui';

export type SecondaryTab = 'prompts' | 'notes';

export const SECONDARY_TABS: { id: SecondaryTab; label: string; title: string }[] = [
  { id: 'prompts', label: '指示', title: 'Excel の Copilot に貼り付ける指示を組み立てる' },
  { id: 'notes', label: 'メモ', title: 'このプロジェクトの作業メモ（.xlcode/notes.md）' },
];

/** メモの置き場所（ツール用フォルダ。PC ごと） */
export const NOTES_REL = '.xlcode/notes.md';

/**
 * 右のサイドバー（VS Code のセカンダリサイドバー）。
 * 指示: 作成欄にテンプレートを足して書き足し、コピーする。自分のテンプレートも保存できる
 * メモ: プロジェクトごとの作業メモ
 */
export function SecondarySidebar({
  tab,
  onTab,
  onClose,
  root,
  book,
  projectMode,
  selectedFile,
  onLog,
  onOpenNotesTab,
}: {
  tab: SecondaryTab;
  onTab: (t: SecondaryTab) => void;
  onClose: () => void;
  root: string | null;
  /** 表示中のブック（無ければ null） */
  book: BookSummary | null;
  projectMode: ProjectMode;
  /** ブック画面や閲覧で選んでいるファイル */
  selectedFile: string | null;
  onLog: (level: 'info' | 'error' | 'success', text: string) => void;
  /** メモを大きなタブで開く */
  onOpenNotesTab: () => void;
}) {
  return (
    <div className="flex h-full min-h-0 flex-col bg-side">
      <div className="flex h-[35px] shrink-0 items-center gap-1 pr-1 pl-3">
        <div role="tablist" aria-label="右のパネル" className="flex min-w-0 flex-1 gap-3">
          {SECONDARY_TABS.map((t) => (
            <button
              key={t.id}
              role="tab"
              aria-selected={tab === t.id}
              title={t.title}
              className={clsx(
                'h-[35px] border-b text-[11px] tracking-wide uppercase',
                tab === t.id ? 'border-fg-strong text-fg-strong' : 'border-transparent text-muted hover:text-fg',
              )}
              onClick={() => onTab(t.id)}
            >
              {t.label}
            </button>
          ))}
        </div>
        <IconButton title="右のパネルを閉じる (Ctrl+Alt+B)" onClick={onClose}>
          <Icon.Close size={14} />
        </IconButton>
      </div>
      <div className="min-h-0 flex-1 overflow-auto">
        {tab === 'prompts' && (
          <PromptPanel root={root} book={book} projectMode={projectMode} selectedFile={selectedFile} onLog={onLog} />
        )}
        {tab === 'notes' && <NotesPanel root={root} onLog={onLog} onOpenTab={onOpenNotesTab} />}
      </div>
    </div>
  );
}

async function copyText(text: string): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    return false;
  }
}

/** 自分のテンプレート（プロジェクトごと。この PC に保存） */
interface CustomPrompt {
  id: string;
  title: string;
  text: string;
}

interface PromptState {
  draft: string;
  custom: CustomPrompt[];
}

function promptKey(root: string | null): string {
  return `prompts:${root ?? ''}`;
}

function PromptPanel({
  root,
  book,
  projectMode,
  selectedFile,
  onLog,
}: {
  root: string | null;
  book: BookSummary | null;
  projectMode: ProjectMode;
  selectedFile: string | null;
  onLog: (level: 'info' | 'error' | 'success', text: string) => void;
}) {
  const [state, setState] = useState<PromptState>(() => storageGet(promptKey(root), { draft: '', custom: [] }));
  useEffect(() => setState(storageGet(promptKey(root), { draft: '', custom: [] })), [root]);
  useEffect(() => storageSet(promptKey(root), state), [state, root]);
  const [copied, setCopied] = useState(false);
  const [openGroup, setOpenGroup] = useState<string | null>(PROMPT_GROUPS[0].title);
  const area = useRef<HTMLTextAreaElement>(null);
  const ctx = { file: selectedFile ?? undefined, conflict: book?.conflictSheets?.[0] };
  useEffect(() => {
    if (!copied) return;
    const t = setTimeout(() => setCopied(false), 1500);
    return () => clearTimeout(t);
  }, [copied]);

  const setDraft = (draft: string) => setState((s) => ({ ...s, draft }));
  /** テンプレートを作成欄の末尾に足す（空なら置く） */
  const append = (text: string) => {
    setState((s) => ({ ...s, draft: s.draft.trim() ? `${s.draft.replace(/\s+$/, '')}\n\n${text}` : text }));
    requestAnimationFrame(() => {
      area.current?.focus();
      area.current?.setSelectionRange(area.current.value.length, area.current.value.length);
    });
  };
  const copyDraft = async () => {
    if (!state.draft.trim()) return;
    if (await copyText(state.draft)) {
      setCopied(true);
      onLog('info', '指示をコピーしました。Excel の Copilot に貼り付けてください');
    } else {
      area.current?.select();
      onLog('error', 'クリップボードに書き込めませんでした。作成欄を選択してコピーしてください');
    }
  };
  const saveCustom = () => {
    const text = state.draft.trim();
    if (!text) return;
    const title = window.prompt('テンプレートの名前', text.split('\n')[0].slice(0, 30));
    if (!title) return;
    setState((s) => ({ ...s, custom: [...s.custom, { id: `${Date.now()}`, title, text }] }));
    onLog('success', `テンプレートに保存しました: ${title}`);
  };
  const removeCustom = (id: string) => setState((s) => ({ ...s, custom: s.custom.filter((c) => c.id !== id) }));

  const Template = ({ it, onRemove }: { it: PromptTemplate | CustomPrompt; onRemove?: () => void }) => (
    <li className="flex items-center gap-1 rounded-[3px] border border-line bg-editor px-2 py-1">
      <button
        className="min-w-0 flex-1 text-left"
        title={`${'when' in it ? it.when + '\n\n' : ''}${fillPrompt(it.text, ctx)}`}
        onClick={() => append(fillPrompt(it.text, ctx))}
        aria-label={`${it.title} を作成欄に足す`}
      >
        <span className="text-fg-strong">{it.title}</span>
        {'when' in it && <span className="block truncate text-[11px] text-faint">{it.when}</span>}
      </button>
      <IconButton title="作成欄に足す" onClick={() => append(fillPrompt(it.text, ctx))}>
        <Icon.Plus size={13} />
      </IconButton>
      {onRemove && (
        <IconButton title="このテンプレートを削除" onClick={onRemove}>
          <Icon.Close size={13} />
        </IconButton>
      )}
    </li>
  );

  return (
    <div className="flex h-full min-h-0 flex-col">
      <div className="shrink-0 border-b border-line px-3 py-2">
        <div className="mb-1 flex items-center gap-2 text-[11px] tracking-wide text-muted uppercase">
          作成欄
          <span className="normal-case">
            {selectedFile ? (
              <>
                （ファイル: <span className="font-mono text-fg">{selectedFile}</span>）
              </>
            ) : (
              '（ファイルを選ぶと名前を埋めます）'
            )}
          </span>
        </div>
        <textarea
          ref={area}
          aria-label="指示の作成欄"
          className="h-[140px] w-full resize-y rounded-[3px] border border-line bg-editor px-2 py-1.5 text-[12.5px] leading-relaxed text-fg outline-none focus:border-accent"
          placeholder="下のテンプレートを足して、具体的な内容を書き足してから「コピー」。Excel の Copilot に貼り付けます"
          value={state.draft}
          onChange={(e) => setDraft(e.target.value)}
          spellCheck={false}
        />
        <div className="mt-1.5 flex flex-wrap items-center gap-1.5">
          <Button variant="primary" onClick={copyDraft} disabled={!state.draft.trim()} aria-label="作成欄をコピー">
            {copied ? 'コピーしました' : 'コピー'}
          </Button>
          <Button
            onClick={saveCustom}
            disabled={!state.draft.trim()}
            title="作成欄の内容を、自分のテンプレートとして保存"
          >
            テンプレートに保存
          </Button>
          <Button onClick={() => setDraft('')} disabled={!state.draft} title="作成欄を空にする">
            クリア
          </Button>
        </div>
      </div>
      <div className="min-h-0 flex-1 overflow-auto px-3 py-2 text-[12px]">
        {state.custom.length > 0 && (
          <section className="mb-3">
            <div className="mb-1 text-[11px] tracking-wide text-muted uppercase">自分のテンプレート</div>
            <ul className="flex flex-col gap-1">
              {state.custom.map((c) => (
                <Template key={c.id} it={c} onRemove={() => removeCustom(c.id)} />
              ))}
            </ul>
          </section>
        )}
        {PROMPT_GROUPS.map((g) => {
          const items = g.items.filter((it) => !it.mode || it.mode === projectMode);
          if (items.length === 0) return null;
          const open = openGroup === g.title;
          return (
            <section key={g.title} className="mb-2">
              <button
                className="flex w-full items-center gap-1 py-1 text-left text-[11px] tracking-wide text-muted uppercase hover:text-fg"
                onClick={() => setOpenGroup(open ? null : g.title)}
                aria-expanded={open}
              >
                {open ? <Icon.ChevronDown size={12} /> : <Icon.ChevronRight size={12} />}
                {g.title}
                <span className="text-faint">{items.length}</span>
              </button>
              {open && (
                <ul className="flex flex-col gap-1">
                  {items.map((it) => (
                    <Template key={it.id} it={it} />
                  ))}
                </ul>
              )}
            </section>
          );
        })}
      </div>
    </div>
  );
}

/** 作業メモ（.xlcode/notes.md）。少し待ってから自動で保存する */
function NotesPanel({
  root,
  onLog,
  onOpenTab,
}: {
  root: string | null;
  onLog: (level: 'info' | 'error' | 'success', text: string) => void;
  onOpenTab: () => void;
}) {
  const [text, setText] = useState<string | null>(null);
  const [savedText, setSavedText] = useState<string>('');
  const [savedAt, setSavedAt] = useState<string | null>(null);
  const timer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  useEffect(() => {
    setText(null);
    if (!root) return;
    api.readDoc(root, NOTES_REL).then((r) => {
      if (!r.ok) {
        onLog('error', r.error);
        return;
      }
      setText(r.value ?? '');
      setSavedText(r.value ?? '');
    });
  }, [root, onLog]);
  const save = async (value: string) => {
    if (!root || value === savedText) return;
    const r = await api.writeDoc(root, NOTES_REL, value);
    if (r.ok) {
      setSavedText(value);
      setSavedAt(new Date().toLocaleTimeString('ja-JP', { hour12: false }));
    } else onLog('error', r.error);
  };
  const edit = (value: string) => {
    setText(value);
    clearTimeout(timer.current);
    timer.current = setTimeout(() => void save(value), 800);
  };
  if (!root) return <div className="px-3 py-3 text-[12px] text-muted">プロジェクトを開くと、メモを書けます。</div>;
  return (
    <div className="flex h-full min-h-0 flex-col">
      <div className="flex shrink-0 items-center gap-2 px-3 py-1.5 text-[11px] text-muted">
        <span className="min-w-0 flex-1 truncate">
          <span className="font-mono">.xlcode/notes.md</span>（この PC だけ。Git には入りません）
        </span>
        <span className="text-faint">
          {text !== null && text !== savedText ? '保存待ち' : savedAt ? `保存 ${savedAt}` : ''}
        </span>
        <button className="text-link hover:underline" onClick={onOpenTab} title="メモを大きなタブで開く">
          タブで開く
        </button>
      </div>
      <textarea
        aria-label="作業メモ"
        className="min-h-0 flex-1 resize-none bg-editor px-3 py-2 font-mono text-[12.5px] leading-[19px] text-fg outline-none"
        placeholder="気づいたこと、Copilot に頼んだこと、次にやることなど。入力をやめると自動で保存します"
        value={text ?? ''}
        readOnly={text === null}
        onChange={(e) => edit(e.target.value)}
        onBlur={() => text !== null && void save(text)}
        spellCheck={false}
      />
    </div>
  );
}
