/** 閲覧タブの「ファイル内を検索」（Ctrl+F）の計算。画面には依存しない */
export interface FindMatch {
  line: number;
  start: number;
  end: number;
}

/** 各行から検索語に一致する箇所を順に集める（重ならない、左から） */
export function findMatches(lines: readonly string[], query: string, matchCase: boolean): FindMatch[] {
  if (query === '') return [];
  const q = matchCase ? query : query.toLowerCase();
  const out: FindMatch[] = [];
  lines.forEach((raw, line) => {
    const l = matchCase ? raw : raw.toLowerCase();
    let from = 0;
    for (;;) {
      const at = l.indexOf(q, from);
      if (at < 0) break;
      out.push({ line, start: at, end: at + q.length });
      from = at + q.length;
    }
  });
  return out;
}

/** 1 行を、該当箇所とそれ以外に分ける（match は全体での通し番号。null は該当箇所以外） */
export function highlight(
  text: string,
  spans: readonly { start: number; end: number; index: number }[],
): { text: string; match: number | null }[] {
  const out: { text: string; match: number | null }[] = [];
  let pos = 0;
  for (const s of spans) {
    if (s.start > pos) out.push({ text: text.slice(pos, s.start), match: null });
    out.push({ text: text.slice(s.start, s.end), match: s.index });
    pos = s.end;
  }
  if (pos < text.length) out.push({ text: text.slice(pos), match: null });
  return out;
}
