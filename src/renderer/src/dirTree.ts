/** ディレクトリの階層（サイドバーの表示用） */
export interface DirNode<T> {
  /** ルートからのパス（'' はルート） */
  path: string;
  /** 表示名。1 つしか子ディレクトリを持たない階層は "a/b" のようにまとめる */
  label: string;
  dirs: DirNode<T>[];
  /** このディレクトリ直下の項目 */
  items: T[];
}

/**
 * 項目をディレクトリの階層にまとめる。
 * dirOf が '' の項目はルート直下。子ディレクトリ・項目は名前順
 */
export function buildDirTree<T>(items: T[], dirOf: (item: T) => string, nameOf: (item: T) => string): DirNode<T> {
  const root: DirNode<T> = { path: '', label: '', dirs: [], items: [] };
  const byPath = new Map<string, DirNode<T>>([['', root]]);
  const ensure = (p: string): DirNode<T> => {
    const hit = byPath.get(p);
    if (hit) return hit;
    const i = p.lastIndexOf('/');
    const parent = ensure(i < 0 ? '' : p.slice(0, i));
    const node: DirNode<T> = { path: p, label: p.slice(i + 1), dirs: [], items: [] };
    parent.dirs.push(node);
    byPath.set(p, node);
    return node;
  };
  for (const it of items) ensure(dirOf(it)).items.push(it);
  const cmp = (a: string, b: string) => a.localeCompare(b, 'ja', { numeric: true });
  const sort = (n: DirNode<T>) => {
    n.dirs.sort((a, b) => cmp(a.label, b.label));
    n.items.sort((a, b) => cmp(nameOf(a), nameOf(b)));
    n.dirs.forEach(sort);
  };
  sort(root);
  root.dirs = root.dirs.map(compact);
  return root;
}

/** 項目が無く子ディレクトリが 1 つだけの階層を、子とまとめる（VS Code の compact folders と同じ） */
function compact<T>(n: DirNode<T>): DirNode<T> {
  let cur = n;
  let label = n.label;
  while (cur.items.length === 0 && cur.dirs.length === 1) {
    cur = cur.dirs[0];
    label = `${label}/${cur.label}`;
  }
  return { ...cur, label, dirs: cur.dirs.map(compact) };
}

/** 階層に含まれる項目の数 */
export function countItems<T>(n: DirNode<T>): number {
  return n.items.length + n.dirs.reduce((s, d) => s + countItems(d), 0);
}
