/** 行単位の差分。same は両側にある行、del は old だけ、add は new だけの行 */
export interface DiffLine {
  kind: 'same' | 'del' | 'add';
  text: string;
}

export interface DiffSummary {
  added: number;
  removed: number;
}

/** 両端の共通部分を除いた中央部分がこれより大きければ、LCS を諦めて全置換として扱う */
const MAX_CELLS = 4_000_000;

/**
 * 2 つの行配列の差分（LCS による最小編集）。
 * 共通の先頭・末尾を先に除くので、通常のソースファイルではごく小さな DP で済む。
 */
export function diffLines(oldLines: readonly string[], newLines: readonly string[]): DiffLine[] {
  let head = 0;
  while (head < oldLines.length && head < newLines.length && oldLines[head] === newLines[head]) head++;
  let tail = 0;
  while (
    tail < oldLines.length - head &&
    tail < newLines.length - head &&
    oldLines[oldLines.length - 1 - tail] === newLines[newLines.length - 1 - tail]
  )
    tail++;
  const a = oldLines.slice(head, oldLines.length - tail);
  const b = newLines.slice(head, newLines.length - tail);

  const out: DiffLine[] = oldLines.slice(0, head).map((text) => ({ kind: 'same', text }));
  if (a.length === 0 || b.length === 0 || a.length * b.length > MAX_CELLS) {
    for (const text of a) out.push({ kind: 'del', text });
    for (const text of b) out.push({ kind: 'add', text });
  } else {
    // lcs[i][j] = a[i..] と b[j..] の LCS 長
    const n = a.length;
    const m = b.length;
    const lcs = new Uint32Array((n + 1) * (m + 1));
    const at = (i: number, j: number) => i * (m + 1) + j;
    for (let i = n - 1; i >= 0; i--) {
      for (let j = m - 1; j >= 0; j--) {
        lcs[at(i, j)] = a[i] === b[j] ? lcs[at(i + 1, j + 1)] + 1 : Math.max(lcs[at(i + 1, j)], lcs[at(i, j + 1)]);
      }
    }
    let i = 0;
    let j = 0;
    while (i < n && j < m) {
      if (a[i] === b[j]) {
        out.push({ kind: 'same', text: a[i] });
        i++;
        j++;
      } else if (lcs[at(i + 1, j)] >= lcs[at(i, j + 1)]) {
        out.push({ kind: 'del', text: a[i] });
        i++;
      } else {
        out.push({ kind: 'add', text: b[j] });
        j++;
      }
    }
    for (; i < n; i++) out.push({ kind: 'del', text: a[i] });
    for (; j < m; j++) out.push({ kind: 'add', text: b[j] });
  }
  for (const text of oldLines.slice(oldLines.length - tail)) out.push({ kind: 'same', text });
  return out;
}

export function summarize(lines: readonly DiffLine[]): DiffSummary {
  let added = 0;
  let removed = 0;
  for (const l of lines) {
    if (l.kind === 'add') added++;
    else if (l.kind === 'del') removed++;
  }
  return { added, removed };
}

export interface Hunk {
  /** 変更前・変更後それぞれの開始行（1 始まり） */
  oldStart: number;
  newStart: number;
  lines: DiffLine[];
}

/** 変更行の前後 context 行だけを残してハンクにまとめる（画面表示用） */
export function toHunks(lines: readonly DiffLine[], context = 3): Hunk[] {
  const changed = lines.map((l) => l.kind !== 'same');
  const keep = new Array<boolean>(lines.length).fill(false);
  for (let i = 0; i < lines.length; i++) {
    if (!changed[i]) continue;
    for (let k = Math.max(0, i - context); k <= Math.min(lines.length - 1, i + context); k++) keep[k] = true;
  }
  const hunks: Hunk[] = [];
  let oldNo = 1;
  let newNo = 1;
  let cur: Hunk | null = null;
  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];
    if (keep[i]) {
      if (!cur) {
        cur = { oldStart: oldNo, newStart: newNo, lines: [] };
        hunks.push(cur);
      }
      cur.lines.push(l);
    } else cur = null;
    if (l.kind !== 'add') oldNo++;
    if (l.kind !== 'del') newNo++;
  }
  return hunks;
}
