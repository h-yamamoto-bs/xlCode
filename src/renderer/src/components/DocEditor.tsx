import clsx from 'clsx';
import { useEffect, useRef, useState } from 'react';
import type { DocKind } from '../../../shared/api';
import { api, unwrap } from '../api';
import { Icon } from './Icons';
import { Button, Kbd } from './ui';

/** Agents.md が厚すぎると Copilot が生成を停止する（1.4）ための目安 */
const LINE_LIMIT = 100;

/** 保存していない編集内容。base は編集を始めたときのディスク上の内容 */
export interface Draft {
  text: string;
  base: string;
}

/** ドキュメントの相対パスから種類を決める */
export function docKind(rel: string): DocKind {
  const base = rel.split('/').pop()!;
  if (base === '.xlcodeignore') return 'ignore';
  return base === 'Agents.md' ? 'agents' : 'local-agents';
}

export const DOC_META: Record<DocKind, { title: string; hint: string; icon: 'rules' | 'ignore' }> = {
  agents: {
    title: 'プロジェクト共通のルール',
    hint: 'Copilot に守らせたいルール。保存後に Refresh Tree を実行すると、全ブックの Agents.md シートへ配布されます。',
    icon: 'rules',
  },
  'local-agents': {
    title: 'このフォルダ固有のルール',
    hint: '保存後に Sync すると、ブックの LocalAgents.md シートへ反映されます（Copilot に編集させて Build もできます）。',
    icon: 'rules',
  },
  ignore: {
    title: 'ブックにしないファイル・フォルダ',
    hint: '.gitignore と同じ書き方（1 行に 1 つ。docs/ や *.min.js など）。Git では管理したまま、シートにも #tree にも載せません。保存後に Sync と Refresh Tree で反映されます。',
    icon: 'ignore',
  },
};

/**
 * ドキュメント（Agents.md / LocalAgents.md / .xlcodeignore）の編集。
 * 保存していない編集は draft として親が持つ（タブやプロジェクトを切り替えても残す）
 */
export function DocEditor({
  root,
  rel,
  draft,
  onDraft,
  onSaved,
  onError,
  onHelp,
}: {
  root: string;
  rel: string;
  draft: Draft | undefined;
  onDraft: (rel: string, draft: Draft | null) => void;
  onSaved: (rel: string) => void;
  onError: (msg: string) => void;
  onHelp?: (anchor: string) => void;
}) {
  // null = 読み込み中
  const [saved, setSaved] = useState<string | null>(null);
  const [exists, setExists] = useState(true);
  const gutter = useRef<HTMLDivElement>(null);
  const area = useRef<HTMLTextAreaElement>(null);
  const kind = docKind(rel);
  const meta = DOC_META[kind];

  useEffect(() => {
    unwrap(api.readDoc(root, rel))
      .then((t) => {
        setSaved(t ?? '');
        setExists(t !== null);
      })
      .catch((e: Error) => onError(e.message));
  }, [root, rel, onError]);

  // ウィンドウに戻ったとき（エディタで編集した後など）にディスクの内容を読み直す
  useEffect(() => {
    const onFocus = () => {
      void api.readDoc(root, rel).then((r) => {
        if (r.ok) {
          setSaved(r.value ?? '');
          setExists(r.value !== null);
        }
      });
    };
    window.addEventListener('focus', onFocus);
    return () => window.removeEventListener('focus', onFocus);
  }, [root, rel]);

  const text = draft?.text ?? saved ?? '';
  const dirty = draft !== undefined;
  // 下書きを始めた後に、エディタなど別の場所でファイルが変更された
  const changedOnDisk = draft !== undefined && saved !== null && draft.base !== saved;

  const edit = (next: string) => {
    if (saved === null) return;
    if (next === saved) onDraft(rel, null);
    else onDraft(rel, { text: next, base: draft?.base ?? saved });
  };
  const save = async () => {
    if (!dirty) return;
    try {
      await unwrap(api.writeDoc(root, rel, text));
      setSaved(text);
      setExists(true);
      onDraft(rel, null);
      onSaved(rel);
    } catch (e) {
      onError((e as Error).message);
    }
  };
  const lines = text === '' ? 1 : text.replace(/\n$/, '').split('\n').length;
  const isAgents = kind === 'agents';

  return (
    <div className="flex h-full min-h-0 flex-col">
      <div className="flex h-[35px] shrink-0 items-center gap-3 border-b border-line px-4">
        {meta.icon === 'ignore' ? (
          <Icon.EyeOff size={14} className="text-warn" />
        ) : (
          <Icon.Rules size={14} className="text-info" />
        )}
        <span className="font-mono text-[12.5px]">
          {rel}
          {dirty && <span className="ml-1 text-fg">●</span>}
        </span>
        <span className="text-[12px] text-muted">{meta.title}</span>
        <span className={clsx('text-[12px]', isAgents && lines > LINE_LIMIT ? 'text-warn' : 'text-faint')}>
          {lines} 行
          {isAgents &&
            lines > LINE_LIMIT &&
            `（目安 ${LINE_LIMIT} 行を超えています。厚すぎると Copilot が生成を停止することがあります）`}
        </span>
        <Button variant="primary" className="ml-auto" onClick={save} disabled={!dirty} title="保存 (Ctrl+S)">
          <Icon.Save size={14} />
          保存
        </Button>
      </div>
      <div className="flex shrink-0 items-center gap-2 border-b border-line bg-side px-4 py-1.5 text-[12px] text-muted">
        <span className="min-w-0 flex-1">
          {meta.hint}
          {!exists && ' ファイルはまだありません。保存すると作成します。'}
        </span>
        {onHelp && (
          <button
            className="shrink-0 text-link hover:underline"
            onClick={() => onHelp(kind === 'ignore' ? 'ブックにしないファイル' : 'ルールagentsmd--localagentsmd')}
          >
            詳しく
          </button>
        )}
      </div>
      {changedOnDisk && (
        <div className="flex shrink-0 items-center gap-3 border-b border-warn/50 bg-[#4d3b00]/50 px-4 py-1.5 text-[12px]">
          <Icon.Warning size={14} className="shrink-0 text-warn" />
          <span className="flex-1">
            編集中に、このファイルがほかの場所で変更されました。保存するとほかの場所での変更は上書きされます。
          </span>
          <Button onClick={() => onDraft(rel, null)}>編集を破棄して読み込み直す</Button>
        </div>
      )}
      <div className="flex min-h-0 flex-1 bg-editor">
        <div
          ref={gutter}
          className="w-12 shrink-0 overflow-hidden py-2 pr-3 text-right font-mono text-[13px] leading-[19px] text-faint"
        >
          {Array.from({ length: Math.max(lines, 1) }, (_, i) => (
            <div key={i}>{i + 1}</div>
          ))}
        </div>
        <textarea
          ref={area}
          aria-label={rel}
          className="min-h-0 flex-1 resize-none bg-transparent py-2 pr-4 font-mono text-[13px] leading-[19px] text-fg outline-none"
          spellCheck={false}
          value={text}
          readOnly={saved === null}
          onChange={(e) => edit(e.target.value)}
          onScroll={(e) => {
            if (gutter.current) gutter.current.scrollTop = e.currentTarget.scrollTop;
          }}
          onKeyDown={(e) => {
            if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's' && !e.shiftKey) {
              e.preventDefault();
              void save();
            } else if (e.key === 'Tab' && !e.ctrlKey && !e.altKey) {
              // Tab で字下げ（VS Code と同じく 2 つの空白）。Shift+Tab で戻す
              e.preventDefault();
              const el = e.currentTarget;
              const { selectionStart: s, selectionEnd: en, value } = el;
              if (e.shiftKey) {
                const lineStart = value.lastIndexOf('\n', s - 1) + 1;
                const removed = value.slice(lineStart).match(/^( {1,2}|\t)/)?.[0].length ?? 0;
                if (removed === 0) return;
                const next = value.slice(0, lineStart) + value.slice(lineStart + removed);
                edit(next);
                requestAnimationFrame(() =>
                  el.setSelectionRange(Math.max(lineStart, s - removed), Math.max(lineStart, en - removed)),
                );
              } else {
                const next = value.slice(0, s) + '  ' + value.slice(en);
                edit(next);
                requestAnimationFrame(() => el.setSelectionRange(s + 2, s + 2));
              }
            }
          }}
        />
      </div>
      <div className="flex h-[22px] shrink-0 items-center gap-3 border-t border-line bg-side px-4 text-[11px] text-faint">
        <span>
          <Kbd>Ctrl+S</Kbd> 保存
        </span>
        <span>
          <Kbd>Tab</Kbd> 字下げ
        </span>
        {kind === 'ignore' && <span>例: docs/ ・ *.min.js ・ package-lock.json ・ !keep.ts（除外から外す）</span>}
      </div>
    </div>
  );
}
