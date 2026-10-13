import { normalizeBinding } from './keys';

/**
 * コマンド（コマンドパレットとショートカットの元になる一覧）。
 * 画面の操作はできるだけここを通し、キー・パレット・ボタンのどこからでも同じ動きになるようにする
 */
export interface Command {
  id: string;
  /** パレットに出す名前 */
  title: string;
  /** 分類（パレットで「ブック: Build」のように前に付ける） */
  category?: string;
  /** 割り当てるキー（複数可。"ctrl+k ctrl+s" のようなコードも可） */
  keys?: string[];
  /** 実行できるか（false ならパレットに出さず、キーも効かない） */
  when?: boolean;
  /** パレットに出さない（番号でタブを選ぶなど） */
  hidden?: boolean;
  run: () => void;
}

/** 1 段目のキーに続きがあるコード（ctrl+k など）を集める */
export function chordPrefixes(commands: Command[]): Set<string> {
  const out = new Set<string>();
  for (const c of commands)
    for (const k of c.keys ?? []) {
      const parts = normalizeBinding(k).split(' ');
      if (parts.length > 1) out.add(parts[0]);
    }
  return out;
}

/** キー（正規化済み。コードは "ctrl+k ctrl+s"）に割り当てられた、実行できるコマンド */
export function findCommand(commands: Command[], key: string): Command | undefined {
  const k = normalizeBinding(key);
  return commands.find((c) => c.when !== false && (c.keys ?? []).some((b) => normalizeBinding(b) === k));
}

/** パレット用の表示名 */
export function commandLabel(c: Command): string {
  return c.category ? `${c.category}: ${c.title}` : c.title;
}
