import type { FileStatus } from '../../core';
import type { SyncState } from '../../shared/api';

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
  removed: { letter: 'D', label: 'シートを削除', hint: 'Build でモジュール削除', color: 'text-deleted' },
};

/** VBA モードでの表示（ソース側が無いため、Build したかどうかだけ） */
export const VBA_STATUS: Partial<Record<FileStatus, StatusMeta>> = {
  clean: { letter: '', label: 'ビルド済み', hint: '', color: 'text-muted' },
  'excel-changed': { letter: 'M', label: '変更あり', hint: 'Build で書き込み', color: 'text-modified' },
  'excel-new': { letter: '+', label: '新規', hint: 'Build でモジュール作成', color: 'text-added' },
};

export const EXCEL_SIDE: FileStatus[] = ['excel-changed', 'excel-new'];
export const SOURCE_SIDE: FileStatus[] = ['source-changed', 'source-new', 'source-deleted', 'sheet-missing'];

export interface SyncMeta {
  label: string;
  color: string;
  /** 転送中（終わるまで待つべき） */
  busy?: boolean;
  /** 利用者の対応が必要 */
  problem?: boolean;
}

export const SYNC_META: Record<SyncState, SyncMeta> = {
  unsupported: { label: '', color: 'text-faint' },
  outside: { label: 'OneDrive 外', color: 'text-faint' },
  synced: { label: '同期済み', color: 'text-added' },
  'online-only': { label: 'オンラインのみ', color: 'text-muted' },
  uploading: { label: 'アップロード待ち', color: 'text-info', busy: true },
  downloading: { label: 'ダウンロード待ち', color: 'text-info', busy: true },
  syncing: { label: '同期中', color: 'text-info', busy: true },
  paused: { label: '同期が一時停止中', color: 'text-warn', problem: true },
  error: { label: '同期エラー', color: 'text-deleted', problem: true },
  unknown: { label: '同期状態不明', color: 'text-faint' },
};
