import type { FileStatus } from '../../core';

export interface StatusMeta {
  /** ソース管理ビュー風の1文字表示 */
  letter: string;
  label: string;
  /** 次に何をすればよいか */
  hint: string;
  color: string;
}

export const STATUS: Record<FileStatus, StatusMeta> = {
  clean: { letter: '', label: '同期済み', hint: '', color: 'text-muted' },
  'excel-changed': { letter: 'E', label: 'Excel側で編集中', hint: 'Build で反映', color: 'text-modified' },
  'excel-new': { letter: 'E+', label: 'Excel側で新規作成', hint: 'Build でファイル作成', color: 'text-added' },
  'source-changed': { letter: 'S', label: 'エディタ側で変更', hint: 'Sync で取り込み', color: 'text-info' },
  'source-new': { letter: 'S+', label: 'エディタ側で新規作成', hint: 'Sync でシート作成', color: 'text-info' },
  'source-deleted': { letter: 'D', label: 'エディタ側で削除', hint: 'Sync でシート削除', color: 'text-deleted' },
  'sheet-missing': { letter: '!', label: 'シートが削除された', hint: 'Sync で復元（削除は DEL_）', color: 'text-warn' },
  conflict: { letter: 'C', label: '両側で変更', hint: 'Build / Sync で衝突シート作成', color: 'text-conflict' },
  gone: { letter: '', label: '', hint: '', color: 'text-muted' },
};

export const EXCEL_SIDE: FileStatus[] = ['excel-changed', 'excel-new'];
export const SOURCE_SIDE: FileStatus[] = ['source-changed', 'source-new', 'source-deleted', 'sheet-missing'];
