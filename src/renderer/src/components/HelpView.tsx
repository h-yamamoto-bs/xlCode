import clsx from 'clsx';
import { Fragment, useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import guide from '../../../../docs/guide.md?raw';
import { parseInline, parseMarkdown, toSections, type Block, type Inline } from '../markdown';
import { Icon } from './Icons';

const SECTIONS = toSections(parseMarkdown(guide));

/** 検索語を <mark> で強調する */
function highlight(text: string, query: string): ReactNode {
  if (!query) return text;
  const lower = text.toLowerCase();
  const parts: ReactNode[] = [];
  let from = 0;
  for (let at = lower.indexOf(query); at >= 0; at = lower.indexOf(query, from)) {
    parts.push(text.slice(from, at));
    parts.push(
      <mark key={at} className="rounded-[2px] bg-[#613214] text-fg-strong">
        {text.slice(at, at + query.length)}
      </mark>,
    );
    from = at + query.length;
  }
  parts.push(text.slice(from));
  return parts;
}

function Inlines({ nodes, ctx }: { nodes: Inline[]; ctx: Ctx }): ReactNode {
  return nodes.map((n, i) => {
    switch (n.type) {
      case 'text':
        return <Fragment key={i}>{highlight(n.text, ctx.query)}</Fragment>;
      case 'code':
        return (
          <code
            key={i}
            className="rounded-[3px] bg-[#ffffff14] px-1 py-px font-mono text-[0.92em] break-words text-fg-strong"
          >
            {highlight(n.text, ctx.query)}
          </code>
        );
      case 'strong':
        return (
          <strong key={i} className="font-semibold text-fg-strong">
            <Inlines nodes={n.children} ctx={ctx} />
          </strong>
        );
      case 'link':
        return (
          <a
            key={i}
            href={n.href}
            className="text-link hover:underline"
            onClick={(e) => {
              e.preventDefault();
              if (n.href.startsWith('#')) ctx.jump(decodeURIComponent(n.href.slice(1)));
              else window.open(n.href);
            }}
          >
            <Inlines nodes={n.children} ctx={ctx} />
          </a>
        );
    }
  });
}

interface Ctx {
  query: string;
  jump: (id: string) => void;
}

const text = (s: string, ctx: Ctx) => <Inlines nodes={parseInline(s)} ctx={ctx} />;

function Blocks({ blocks, ctx }: { blocks: Block[]; ctx: Ctx }): ReactNode {
  return blocks.map((b, i) => <BlockView key={i} b={b} ctx={ctx} />);
}

function BlockView({ b, ctx }: { b: Block; ctx: Ctx }): ReactNode {
  switch (b.type) {
    case 'heading': {
      const cls = {
        1: 'mb-3 text-[26px] font-light text-fg-strong',
        2: 'mt-10 mb-3 border-b border-line pb-1.5 text-[20px] font-normal text-fg-strong',
        3: 'mt-7 mb-2 text-[16px] font-semibold text-fg-strong',
        4: 'mt-5 mb-1.5 text-[14px] font-semibold text-fg-strong',
      }[Math.min(b.level, 4)];
      const Tag = `h${Math.min(b.level, 4)}` as 'h1' | 'h2' | 'h3' | 'h4';
      return (
        <Tag id={b.id} data-heading={b.level} className={clsx('scroll-mt-4 first:mt-0', cls)}>
          {text(b.text, ctx)}
        </Tag>
      );
    }
    case 'paragraph':
      return (
        <p className="my-2.5">
          {b.lines.map((l, i) => (
            <Fragment key={i}>
              {i > 0 && <br />}
              {text(l, ctx)}
            </Fragment>
          ))}
        </p>
      );
    case 'list': {
      const Tag = b.ordered ? 'ol' : 'ul';
      return (
        <Tag
          start={b.ordered ? b.start : undefined}
          className={clsx('my-2.5 pl-6', b.ordered ? 'list-decimal' : 'list-disc', '[&_ol]:my-1 [&_ul]:my-1')}
        >
          {b.items.map((item, i) => (
            <li key={i} className="my-1 pl-0.5 marker:text-muted [&>p]:my-0">
              <Blocks blocks={item} ctx={ctx} />
            </li>
          ))}
        </Tag>
      );
    }
    case 'code':
      return (
        <pre className="my-3 overflow-x-auto rounded-[3px] border border-line bg-side px-3 py-2 font-mono text-[12.5px] leading-[19px] text-fg">
          {highlight(b.text, ctx.query)}
        </pre>
      );
    case 'table':
      return (
        <div className="my-3 overflow-x-auto rounded-[3px] border border-line">
          <table className="w-full border-collapse text-left text-[13px]">
            <thead>
              <tr className="bg-side">
                {b.header.map((c, i) => (
                  <th key={i} className="px-3 py-1.5 align-bottom font-semibold whitespace-nowrap text-fg-strong">
                    {text(c, ctx)}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {b.rows.map((r, i) => (
                <tr key={i} className="border-t border-line align-top">
                  {r.map((c, j) => (
                    <td key={j} className={clsx('px-3 py-1.5', j === 0 && 'min-w-[7em]')}>
                      {text(c, ctx)}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      );
    case 'quote':
      return (
        <div className="my-3 rounded-[3px] border border-l-[3px] border-line border-l-accent bg-side px-4 py-1 [&>p:first-child>strong:only-child]:text-info">
          <Blocks blocks={b.blocks} ctx={ctx} />
        </div>
      );
    case 'hr':
      return <hr className="my-6 border-line" />;
  }
}

/**
 * ヘルプ（docs/guide.md）。左に目次と検索、右に本文。
 * anchor を渡すと、その見出しへ移動する（nonce が変わるたびに移動し直す）
 */
export function HelpView({ anchor, nonce }: { anchor?: string; nonce?: number }) {
  const [query, setQuery] = useState('');
  const [active, setActive] = useState(SECTIONS[0]?.id ?? '');
  const body = useRef<HTMLDivElement>(null);
  const q = query.trim().toLowerCase();

  const shown = useMemo(() => (q ? SECTIONS.filter((s) => s.search.includes(q)) : SECTIONS), [q]);

  const scrollTo = (id: string) => {
    const el = body.current?.querySelector<HTMLElement>(`[id="${CSS.escape(id)}"]`);
    if (el) {
      el.scrollIntoView({ block: 'start' });
      setActive(SECTIONS.find((s) => s.id === id || s.subs.some((x) => x.id === id))?.id ?? id);
    }
  };
  const jump = (id: string) => {
    // 検索で隠れている見出しへのリンクなら、検索を解除してから移動する
    if (!body.current?.querySelector(`[id="${CSS.escape(id)}"]`)) {
      setQuery('');
      requestAnimationFrame(() => scrollTo(id));
    } else scrollTo(id);
  };

  useEffect(() => {
    if (!anchor) return;
    setQuery('');
    requestAnimationFrame(() => scrollTo(anchor));
  }, [anchor, nonce]);

  // 本文のスクロールに合わせて、目次の現在位置を更新する
  const onScroll = () => {
    const el = body.current;
    if (!el) return;
    const top = el.getBoundingClientRect().top + 24;
    let current = shown[0]?.id ?? '';
    for (const h of el.querySelectorAll<HTMLElement>('[data-heading="1"],[data-heading="2"]')) {
      if (h.getBoundingClientRect().top <= top) current = h.id;
      else break;
    }
    if (current !== active) setActive(current);
  };

  const ctx: Ctx = { query: q, jump };

  return (
    <div className="flex h-full min-h-0">
      <aside className="flex w-[260px] shrink-0 flex-col border-r border-line bg-side" aria-label="目次">
        <div className="flex h-[35px] shrink-0 items-center px-5 text-[11px] tracking-wide text-muted uppercase">
          ヘルプ
        </div>
        <div className="px-3 pb-2">
          <div className="flex h-[26px] items-center gap-1.5 rounded-[2px] border border-[#3c3c3c] bg-input px-2 focus-within:border-accent">
            <Icon.Search size={13} className="shrink-0 text-muted" />
            <input
              aria-label="ヘルプを検索"
              className="min-w-0 flex-1 bg-transparent text-[13px] text-fg outline-none placeholder:text-faint"
              placeholder="検索（例: 衝突、DEL_、OneDrive）"
              value={query}
              onChange={(e) => {
                setQuery(e.target.value);
                body.current?.scrollTo({ top: 0 });
              }}
              onKeyDown={(e) => {
                if (e.key === 'Escape') setQuery('');
              }}
            />
            {query && (
              <button title="検索をやめる" aria-label="検索をやめる" onClick={() => setQuery('')}>
                <Icon.Close size={12} className="text-muted hover:text-fg" />
              </button>
            )}
          </div>
          {q && <div className="mt-1.5 px-0.5 text-[11px] text-muted">{shown.length} 件の章が見つかりました</div>}
        </div>
        <nav className="min-h-0 flex-1 overflow-auto pb-3">
          {shown.map((s) => (
            <div key={s.id}>
              <button
                className={clsx(
                  'flex w-full items-center truncate py-[3px] pr-3 pl-5 text-left text-[13px]',
                  active === s.id ? 'bg-focus text-fg-strong' : 'text-fg hover:bg-hover',
                )}
                onClick={() => scrollTo(s.id)}
              >
                {s.title}
              </button>
              {s.subs.map((x) => (
                <button
                  key={x.id}
                  className="flex w-full truncate py-[2px] pr-3 pl-9 text-left text-[12px] text-muted hover:bg-hover hover:text-fg"
                  onClick={() => scrollTo(x.id)}
                >
                  {x.title}
                </button>
              ))}
            </div>
          ))}
        </nav>
      </aside>
      <div ref={body} className="selectable min-w-0 flex-1 overflow-auto" onScroll={onScroll}>
        <article className="mx-auto max-w-[860px] px-8 py-6 text-[13.5px] leading-[1.75] text-fg">
          {shown.length === 0 ? (
            <div className="py-10 text-center text-muted">「{query}」は見つかりませんでした。</div>
          ) : (
            shown.map((s) => <Blocks key={s.id} blocks={s.blocks} ctx={ctx} />)
          )}
        </article>
      </div>
    </div>
  );
}
