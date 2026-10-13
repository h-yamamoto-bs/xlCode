/**
 * クイックオープン・コマンドパレット用のあいまい検索。
 * 検索語を空白で分け、それぞれが対象の中に順番どおりに現れるものを残す（VS Code と同じ感覚）。
 * 点数は、連続して一致する・語の先頭で一致するほど高い
 */
export interface Match {
  score: number;
  /** 一致した文字の位置（強調表示用） */
  positions: number[];
}

function matchToken(text: string, token: string): Match | null {
  const t = text.toLowerCase();
  // まず部分一致（いちばん分かりやすい）
  const at = t.indexOf(token);
  if (at >= 0) {
    const atWordStart = at === 0 || /[\s/\\._-]/.test(t[at - 1]);
    return {
      score: 100 + (atWordStart ? 30 : 0) - at * 0.5 + token.length * 2,
      positions: Array.from({ length: token.length }, (_, i) => at + i),
    };
  }
  // 次に順番どおりの一致（a.ts → app/App.tsx など）
  const positions: number[] = [];
  let from = 0;
  for (const ch of token) {
    const i = t.indexOf(ch, from);
    if (i < 0) return null;
    positions.push(i);
    from = i + 1;
  }
  let score = 10 + token.length;
  for (let i = 1; i < positions.length; i++) if (positions[i] === positions[i - 1] + 1) score += 3;
  score -= (positions[positions.length - 1] - positions[0] - positions.length + 1) * 0.5;
  return { score, positions };
}

export function fuzzyMatch(text: string, query: string): Match | null {
  const tokens = query.trim().toLowerCase().split(/\s+/).filter(Boolean);
  if (tokens.length === 0) return { score: 0, positions: [] };
  let score = 0;
  const positions = new Set<number>();
  for (const tok of tokens) {
    const m = matchToken(text, tok);
    if (!m) return null;
    score += m.score;
    for (const p of m.positions) positions.add(p);
  }
  return { score, positions: [...positions].sort((a, b) => a - b) };
}

/** 一致した位置を使って、文字列を [通常, 強調, 通常, ...] の交互の配列に分ける（偶数番目が通常） */
export function splitHighlights(text: string, positions: number[]): string[] {
  const set = new Set(positions);
  const out: string[] = [];
  let cur = '';
  let hl = false;
  for (let i = 0; i < text.length; i++) {
    const h = set.has(i);
    if (h !== hl) {
      out.push(cur);
      cur = '';
      hl = h;
    }
    cur += text[i];
  }
  out.push(cur);
  // 先頭は必ず「通常」から始める
  if (hl && out.length % 2 === 1) out.unshift('');
  return out;
}
