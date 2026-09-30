import { readdir, readFile } from 'node:fs/promises';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { parseInline, parseMarkdown, slugify, toSections, type Block } from '../src/renderer/src/markdown';

const root = path.join(import.meta.dirname, '..');

function headingIds(blocks: Block[]): string[] {
  return blocks.flatMap((b) =>
    b.type === 'heading'
      ? [b.id]
      : b.type === 'quote'
        ? headingIds(b.blocks)
        : b.type === 'list'
          ? b.items.flatMap(headingIds)
          : [],
  );
}

describe('ヘルプ（docs/guide.md）', () => {
  it('ページ内リンクと、アプリから開く見出しがすべて存在する', async () => {
    const guide = await readFile(path.join(root, 'docs/guide.md'), 'utf8');
    const ids = new Set(headingIds(parseMarkdown(guide)));
    const inGuide = [...guide.matchAll(/\]\(#([^)]+)\)/g)].map((m) => m[1]);
    const dir = path.join(root, 'src/renderer/src');
    const files = (await readdir(dir, { recursive: true })).filter((f) => f.endsWith('.tsx'));
    const inApp: string[] = [];
    for (const f of files) {
      const src = await readFile(path.join(dir, f), 'utf8');
      for (const m of src.matchAll(/(?:onHelp|openHelp)\('([^']+)'\)|anchor="([^"]+)"/g)) inApp.push(m[1] ?? m[2]);
      // anchor={条件 ? '…' : '…'} のような書き方
      for (const m of src.matchAll(/anchor=\{([^}]*)\}/g))
        inApp.push(...[...m[1].matchAll(/[?:] *'([^']+)'/g)].map((x) => x[1]));
    }
    expect(inGuide.length).toBeGreaterThan(10);
    expect(inApp.length).toBeGreaterThan(3);
    expect([...inGuide, ...inApp].filter((id) => !ids.has(id))).toEqual([]);
  });

  it('大見出しごとに区切り、目次に小見出しを持つ', async () => {
    const guide = await readFile(path.join(root, 'docs/guide.md'), 'utf8');
    const sections = toSections(parseMarkdown(guide));
    expect(sections[0].title).toBe('xlCode の使い方');
    const build = sections.find((s) => s.id === '各操作の詳しい説明')!;
    expect(build.subs.map((s) => s.id)).toContain('build');
    expect(build.search).toContain('衝突シート');
  });
});

describe('Markdown', () => {
  it('見出しのアンカーは GitHub と同じ規則で作る', () => {
    expect(slugify('ルール（Agents.md / LocalAgents.md）')).toBe('ルールagentsmd--localagentsmd');
    expect(slugify('1. プロジェクトを開く')).toBe('1-プロジェクトを開く');
    expect(slugify('Excel でシートを編集するときの決まり')).toBe('excel-でシートを編集するときの決まり');
    const ids = parseMarkdown('## A\n## A').map((b) => (b.type === 'heading' ? b.id : ''));
    expect(ids).toEqual(['a', 'a-1']);
  });

  it('入れ子の箇条書き・表・引用・コードを読む', () => {
    const blocks = parseMarkdown(
      [
        '1. 一つ目',
        '   - 子',
        '2. 二つ目',
        '',
        '| a | b |',
        '|---|---|',
        '| `x|y` | **z** |',
        '',
        '> **注意**',
        '>',
        '> 本文',
        '',
        '```vb',
        'MsgBox 1',
        '```',
      ].join('\n'),
    );
    expect(blocks.map((b) => b.type)).toEqual(['list', 'table', 'quote', 'code']);
    const list = blocks[0] as Extract<Block, { type: 'list' }>;
    expect(list.ordered).toBe(true);
    expect(list.items).toHaveLength(2);
    expect(list.items[0].map((b) => b.type)).toEqual(['paragraph', 'list']);
    const table = blocks[1] as Extract<Block, { type: 'table' }>;
    expect(table.rows[0]).toEqual(['`x|y`', '**z**']);
    expect((blocks[2] as Extract<Block, { type: 'quote' }>).blocks.map((b) => b.type)).toEqual([
      'paragraph',
      'paragraph',
    ]);
  });

  it('行内の太字・コード・リンク', () => {
    expect(parseInline('a **b** `c` [d](#e)')).toEqual([
      { type: 'text', text: 'a ' },
      { type: 'strong', children: [{ type: 'text', text: 'b' }] },
      { type: 'text', text: ' ' },
      { type: 'code', text: 'c' },
      { type: 'text', text: ' ' },
      { type: 'link', href: '#e', children: [{ type: 'text', text: 'd' }] },
    ]);
  });
});
