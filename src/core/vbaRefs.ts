/**
 * 参照設定（VBE の「ツール → 参照設定」）。編集用ブックの #refs シートに 1 行 1 つ書く。
 *
 *   {420B2830-E718-11CF-893D-00A0C9054228} 1.0 Microsoft Scripting Runtime
 *
 * Build のときにビルド結果へ追加する（AddFromGuid）。' で始まる行と空行は無視する。
 */

export interface VbaReference {
  guid: string;
  major: number;
  minor: number;
  description: string;
}

export const REFS_HELP = [
  "' 参照設定（VBE の「ツール → 参照設定」）。Build のときにビルド結果へ追加します",
  "' 1 行に 1 つ: {GUID} 主バージョン.副バージョン 説明（バージョンを省くと、その PC にある最新版）",
  "' 先頭に ' を付けた行は無視します。よく使うもの:",
  "' {420B2830-E718-11CF-893D-00A0C9054228} 1.0 Microsoft Scripting Runtime（Dictionary・FileSystemObject）",
  "' {3F4DACA7-160D-11D2-A8E9-00104B365C9F} 5.5 Microsoft VBScript Regular Expressions 5.5",
  "' {B691E011-1797-432E-907A-4D8C69339129} 6.1 Microsoft ActiveX Data Objects 6.1 Library",
];

const LINE = /^\{?([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\}?(?:\s+(\d+)\.(\d+))?(?:\s+(.*))?$/i;

export function parseRefs(sheet: string, lines: readonly string[]): { refs: VbaReference[]; errors: string[] } {
  const refs: VbaReference[] = [];
  const errors: string[] = [];
  lines.forEach((raw, i) => {
    const l = raw.trim();
    if (l === '' || l.startsWith("'")) return;
    const m = LINE.exec(l);
    if (!m) {
      errors.push(`「${sheet}」${i + 1} 行目: 「{GUID} 1.0 説明」の形で書いてください: ${l}`);
      return;
    }
    const guid = `{${m[1].toUpperCase()}}`;
    if (refs.some((r) => r.guid === guid)) return;
    refs.push({ guid, major: Number(m[2] ?? 0), minor: Number(m[3] ?? 0), description: (m[4] ?? '').trim() });
  });
  return { refs, errors };
}

export function renderRefs(refs: readonly VbaReference[]): string[] {
  return [
    ...REFS_HELP,
    ...refs.map((r) => `${r.guid} ${r.major}.${r.minor}${r.description ? ` ${r.description}` : ''}`),
  ];
}
