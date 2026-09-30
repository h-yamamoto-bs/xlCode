/**
 * ヘルプ（docs/guide.md）を表示するための、小さな Markdown パーサー。
 * 使っている書き方だけに対応する: 見出し・段落・箇条書き（入れ子）・番号付きリスト・表・コード・引用・区切り線、
 * 行内の **太字**・`コード`・[リンク](#見出し)。
 */

export type Block =
  | { type: 'heading'; level: number; text: string; id: string }
  | { type: 'paragraph'; lines: string[] }
  | { type: 'list'; ordered: boolean; start: number; items: Block[][] }
  | { type: 'code'; lang: string; text: string }
  | { type: 'table'; header: string[]; rows: string[][] }
  | { type: 'quote'; blocks: Block[] }
  | { type: 'hr' };

export type Inline =
  | { type: 'text'; text: string }
  | { type: 'strong'; children: Inline[] }
  | { type: 'code'; text: string }
  | { type: 'link'; href: string; children: Inline[] };

/** 見出しのアンカー（GitHub と同じ規則: 小文字にし、記号を除き、空白を - にする） */
export function slugify(text: string): string {
  return plainText(parseInline(text))
    .trim()
    .toLowerCase()
    .replace(/[^\p{L}\p{N}\p{M}\s_-]/gu, '')
    .replace(/\s/g, '-');
}

const LIST_ITEM = /^( *)([-*]|\d+\.) +(.*)$/;
const indentOf = (l: string) => l.length - l.trimStart().length;
const isBlank = (l: string) => l.trim() === '';

export function parseMarkdown(src: string): Block[] {
  const used = new Map<string, number>();
  const uniqueId = (text: string) => {
    const base = slugify(text);
    const n = used.get(base) ?? 0;
    used.set(base, n + 1);
    return n === 0 ? base : `${base}-${n}`;
  };
  return parseBlocks(src.replace(/\r\n?/g, '\n').split('\n'), uniqueId);
}

function splitRow(line: string): string[] {
  let s = line.trim();
  if (s.startsWith('|')) s = s.slice(1);
  if (s.endsWith('|')) s = s.slice(0, -1);
  // `コード` の中の | では区切らない
  const cells: string[] = [];
  let cur = '';
  let inCode = false;
  for (const ch of s) {
    if (ch === '`') inCode = !inCode;
    if (ch === '|' && !inCode) {
      cells.push(cur.trim());
      cur = '';
    } else cur += ch;
  }
  cells.push(cur.trim());
  return cells;
}

function parseBlocks(lines: string[], uniqueId: (t: string) => string): Block[] {
  const blocks: Block[] = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (isBlank(line)) {
      i++;
      continue;
    }
    const fence = line.match(/^ *```(\w*)/);
    if (fence) {
      const body: string[] = [];
      i++;
      while (i < lines.length && !/^ *```/.test(lines[i])) body.push(lines[i++]);
      i++;
      blocks.push({ type: 'code', lang: fence[1], text: body.join('\n') });
      continue;
    }
    const h = line.match(/^(#{1,6}) +(.*)$/);
    if (h) {
      blocks.push({ type: 'heading', level: h[1].length, text: h[2].trim(), id: uniqueId(h[2].trim()) });
      i++;
      continue;
    }
    if (/^ *(-{3,}|\*{3,})\s*$/.test(line)) {
      blocks.push({ type: 'hr' });
      i++;
      continue;
    }
    if (line.trimStart().startsWith('>')) {
      const body: string[] = [];
      while (i < lines.length && lines[i].trimStart().startsWith('>')) {
        body.push(lines[i].trimStart().replace(/^> ?/, ''));
        i++;
      }
      blocks.push({ type: 'quote', blocks: parseBlocks(body, uniqueId) });
      continue;
    }
    if (line.trimStart().startsWith('|') && i + 1 < lines.length && /^ *\|? *:?-+/.test(lines[i + 1])) {
      const header = splitRow(line);
      i += 2;
      const rows: string[][] = [];
      while (i < lines.length && lines[i].trimStart().startsWith('|')) rows.push(splitRow(lines[i++]));
      blocks.push({ type: 'table', header, rows });
      continue;
    }
    const li = line.match(LIST_ITEM);
    if (li) {
      const base = li[1].length;
      const ordered = /\d/.test(li[2]);
      const items: Block[][] = [];
      while (i < lines.length) {
        const m = lines[i].match(LIST_ITEM);
        if (!m || m[1].length !== base || /\d/.test(m[2]) !== ordered) break;
        // 項目の続き（より深く字下げした行）を集めて、中身として読む
        const content = m[3];
        const inner: string[] = [content];
        const childIndent = base + m[2].length + 1;
        i++;
        while (i < lines.length) {
          const l = lines[i];
          if (isBlank(l)) {
            // 空行の次が字下げされた続きなら、同じ項目の中身
            if (i + 1 < lines.length && !isBlank(lines[i + 1]) && indentOf(lines[i + 1]) > base) {
              inner.push('');
              i++;
              continue;
            }
            break;
          }
          if (indentOf(l) <= base) break;
          inner.push(l.slice(Math.min(childIndent, indentOf(l))));
          i++;
        }
        items.push(parseBlocks(inner, uniqueId));
      }
      blocks.push({ type: 'list', ordered, start: ordered ? Number(li[2].slice(0, -1)) : 1, items });
      continue;
    }
    const para: string[] = [];
    while (
      i < lines.length &&
      !isBlank(lines[i]) &&
      !/^(#{1,6} |> ?| *```| *\|)/.test(lines[i].trimStart()) &&
      !LIST_ITEM.test(lines[i])
    ) {
      para.push(lines[i].trim());
      i++;
    }
    blocks.push({ type: 'paragraph', lines: para });
  }
  return blocks;
}

export function parseInline(src: string): Inline[] {
  const out: Inline[] = [];
  let text = '';
  const flush = () => {
    if (text) out.push({ type: 'text', text });
    text = '';
  };
  let i = 0;
  while (i < src.length) {
    if (src[i] === '`') {
      const end = src.indexOf('`', i + 1);
      if (end > i) {
        flush();
        out.push({ type: 'code', text: src.slice(i + 1, end) });
        i = end + 1;
        continue;
      }
    }
    if (src.startsWith('**', i)) {
      const end = src.indexOf('**', i + 2);
      if (end > i + 2) {
        flush();
        out.push({ type: 'strong', children: parseInline(src.slice(i + 2, end)) });
        i = end + 2;
        continue;
      }
    }
    if (src[i] === '[') {
      const m = src.slice(i).match(/^\[([^\]]+)\]\(([^)\s]+)\)/);
      if (m) {
        flush();
        out.push({ type: 'link', href: m[2], children: parseInline(m[1]) });
        i += m[0].length;
        continue;
      }
    }
    text += src[i++];
  }
  flush();
  return out;
}

export function plainText(inl: Inline[]): string {
  return inl.map((n) => (n.type === 'text' || n.type === 'code' ? n.text : plainText(n.children))).join('');
}

/** 検索用に、ブロックの文字をすべてつなげる */
export function blockText(b: Block): string {
  switch (b.type) {
    case 'heading':
      return plainText(parseInline(b.text));
    case 'paragraph':
      return b.lines.map((l) => plainText(parseInline(l))).join('\n');
    case 'list':
      return b.items.map((it) => it.map(blockText).join('\n')).join('\n');
    case 'code':
      return b.text;
    case 'table':
      return [b.header, ...b.rows].map((r) => r.map((c) => plainText(parseInline(c))).join(' ')).join('\n');
    case 'quote':
      return b.blocks.map(blockText).join('\n');
    case 'hr':
      return '';
  }
}

export interface Section {
  /** 見出し（h2）の id。最初の見出し前の部分は h1 の id */
  id: string;
  title: string;
  blocks: Block[];
  /** 小見出し（h3） */
  subs: { id: string; title: string }[];
  /** 検索用（小文字） */
  search: string;
}

/** 大見出し（h2）ごとに区切る */
export function toSections(blocks: Block[]): Section[] {
  const sections: Section[] = [];
  for (const b of blocks) {
    if ((b.type === 'heading' && b.level <= 2) || sections.length === 0) {
      const title = b.type === 'heading' ? plainText(parseInline(b.text)) : '';
      sections.push({ id: b.type === 'heading' ? b.id : 'top', title, blocks: [], subs: [], search: '' });
    }
    const s = sections[sections.length - 1];
    s.blocks.push(b);
    if (b.type === 'heading' && b.level === 3) s.subs.push({ id: b.id, title: plainText(parseInline(b.text)) });
    s.search += blockText(b).toLowerCase() + '\n';
  }
  return sections;
}
