import type { DocKind } from '../../shared/api';

/** ドキュメントの相対パスから種類を決める */
export function docKind(rel: string): DocKind {
  const base = rel.split('/').pop()!;
  if (base === '.xlcodeignore') return 'ignore';
  if (rel === '.xlcode/notes.md') return 'notes';
  if (base === 'Agents.md' && !rel.includes('/')) return 'agents';
  if (base === 'LocalAgents.md') return 'local-agents';
  return 'markdown';
}

export const DOC_META: Record<DocKind, { title: string; hint: string; icon: 'rules' | 'ignore' }> = {
  agents: {
    title: 'プロジェクト共通のルール',
    hint: 'Copilot に守らせたいルール。保存後に Refresh Tree を実行すると、全ブックの Agents.md シートへ配布されます。',
    icon: 'rules',
  },
  'local-agents': {
    title: 'このフォルダ固有のルール',
    hint: '保存後に Sync すると、ブックの LocalAgents.md シートへ反映されます（Copilot に編集させて Build もできます）。',
    icon: 'rules',
  },
  notes: {
    title: '作業メモ',
    hint: 'このプロジェクトの作業メモ（.xlcode/notes.md）。この PC だけに置き、Git では管理しません。右のパネルの「メモ」と同じ内容です。',
    icon: 'rules',
  },
  markdown: {
    title: 'Markdown',
    hint: 'ブックのシートでもあるファイルです。保存後に Sync するとシートへ反映されます（Copilot に編集させて Build もできます）。',
    icon: 'rules',
  },
  ignore: {
    title: 'ブックにしないファイル・フォルダ',
    hint: '.gitignore と同じ書き方（1 行に 1 つ。docs/ や *.min.js など）。Git では管理したまま、シートにも #tree にも載せません。保存後に Sync と Refresh Tree で反映されます。',
    icon: 'ignore',
  },
};
