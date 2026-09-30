import clsx from 'clsx';
import { useEffect, useRef, useState } from 'react';
import type { ProjectInfo } from '../../../shared/api';
import { api, unwrap } from '../api';
import { Icon } from './Icons';
import { Button } from './ui';

/** Agents.md が厚すぎると Copilot が生成を停止する（1.4）ための目安 */
const LINE_LIMIT = 100;

/** 保存していない編集内容。base は編集を始めたときのディスク上の内容 */
export interface Draft {
  text: string;
  base: string;
}

export function RulesList({
  project,
  selected,
  drafts,
  onSelect,
}: {
  project: ProjectInfo;
  selected: string;
  drafts: Record<string, Draft>;
  onSelect: (rel: string) => void;
}) {
  const items = [
    { rel: 'Agents.md', label: 'Agents.md', sub: 'プロジェクト共通' },
    ...project.books.map((b) => ({
      rel: b.dirRel ? `${b.dirRel}/LocalAgents.md` : 'LocalAgents.md',
      label: 'LocalAgents.md',
      sub: b.dirRel || '.',
    })),
  ];
  return (
    <div className="flex h-full flex-col bg-side">
      <div className="flex h-[35px] shrink-0 items-center px-5 text-[11px] tracking-wide text-muted uppercase">
        ルール
      </div>
      <ul>
        {items.map((it) => (
          <li
            key={it.rel}
            className={clsx(
              'flex h-[22px] cursor-pointer items-center gap-1.5 pr-2 pl-5',
              it.rel === selected ? 'bg-focus outline outline-1 -outline-offset-1 outline-accent' : 'hover:bg-hover',
            )}
            onClick={() => onSelect(it.rel)}
          >
            <Icon.Rules size={14} className="shrink-0 text-info" />
            <span className="truncate">{it.label}</span>
            <span className="truncate text-[11px] text-faint">{it.sub}</span>
            {drafts[it.rel] && (
              <span className="ml-auto text-fg" title="保存していない変更があります">
                ●
              </span>
            )}
          </li>
        ))}
      </ul>
    </div>
  );
}

export function RulesEditor({
  root,
  rel,
  vba,
  draft,
  onDraft,
  onSaved,
  onError,
}: {
  root: string;
  rel: string;
  /** VBA モード（LocalAgents.md も Refresh Tree で配布する） */
  vba: boolean;
  draft: Draft | undefined;
  onDraft: (rel: string, draft: Draft | null) => void;
  onSaved: (rel: string) => void;
  onError: (msg: string) => void;
}) {
  // null = 読み込み中
  const [saved, setSaved] = useState<string | null>(null);
  const [exists, setExists] = useState(true);
  const gutter = useRef<HTMLDivElement>(null);

  useEffect(() => {
    unwrap(api.readRuleFile(root, rel))
      .then((t) => {
        setSaved(t ?? '');
        setExists(t !== null);
      })
      .catch((e: Error) => onError(e.message));
  }, [root, rel, onError]);

  // ウィンドウに戻ったとき（エディタで編集した後など）にディスクの内容を読み直す
  useEffect(() => {
    const onFocus = () => {
      void api.readRuleFile(root, rel).then((r) => {
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
      await unwrap(api.writeRuleFile(root, rel, text));
      setSaved(text);
      setExists(true);
      onDraft(rel, null);
      onSaved(rel);
    } catch (e) {
      onError((e as Error).message);
    }
  };
  const lines = text === '' ? 1 : text.replace(/\n$/, '').split('\n').length;
  const isAgents = rel === 'Agents.md';

  return (
    <div className="flex h-full min-h-0 flex-col">
      <div className="flex h-[35px] shrink-0 items-center gap-3 border-b border-line px-4">
        <Icon.Rules size={14} className="text-info" />
        <span className="font-mono text-[12.5px]">
          {rel}
          {dirty && <span className="ml-1 text-fg">●</span>}
        </span>
        <span className={clsx('text-[12px]', lines > LINE_LIMIT ? 'text-warn' : 'text-faint')}>
          {lines} 行
          {lines > LINE_LIMIT &&
            `（目安 ${LINE_LIMIT} 行を超えています。厚すぎると Copilot が生成を停止することがあります）`}
        </span>
        <Button variant="primary" className="ml-auto" onClick={save} disabled={!dirty}>
          <Icon.Save size={14} />
          保存
        </Button>
      </div>
      <div className="shrink-0 border-b border-line bg-side px-4 py-1.5 text-[12px] text-muted">
        {isAgents
          ? 'プロジェクト共通のルール。保存後に Refresh Tree を実行すると、全ブックの Agents.md シートへ配布されます。'
          : vba
            ? 'このブック固有のルール。保存後に Refresh Tree を実行すると、ブックの LocalAgents.md シートへ反映されます。'
            : 'このディレクトリ固有のルール。保存後に Sync すると、ブックの LocalAgents.md シートへ反映されます。'}
        {!exists && ' ファイルはまだありません。保存すると作成します。'}
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
            if ((e.ctrlKey || e.metaKey) && e.key === 's') {
              e.preventDefault();
              void save();
            }
          }}
        />
      </div>
    </div>
  );
}
