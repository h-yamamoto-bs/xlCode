export type OpStatus =
  /** 完了 */
  | 'ok'
  /** エラーで中断（何も書き込んでいない） */
  | 'error'
  /** ユーザーの確認が必要（confirmed: true で再実行） */
  | 'confirm'
  /** 衝突シートを作成した（Build はブロック） */
  | 'conflict'
  /** Sync 時に未 Build のシート変更がある（5.1 の A/B/C を選ばせる） */
  | 'needs-decision';

export type ConfirmationKind = 'shrink' | 'delete' | 'uncommitted' | 'no-git';

export interface Confirmation {
  kind: ConfirmationKind;
  message: string;
  files: string[];
}

export type ChangeAction =
  'write-file' | 'delete-file' | 'write-sheet' | 'delete-sheet' | 'reformat-sheet' | 'conflict-sheet' | 'commit';

export interface Change {
  action: ChangeAction;
  target: string;
}

export interface OpResult {
  status: OpStatus;
  errors: string[];
  warnings: string[];
  confirmations: Confirmation[];
  changes: Change[];
  /** 作成した衝突シート */
  conflicts: { sheet: string; file: string }[];
  /** 未 Build のシート変更（needs-decision のとき） */
  unbuilt: string[];
}

export function newResult(): OpResult {
  return { status: 'ok', errors: [], warnings: [], confirmations: [], changes: [], conflicts: [], unbuilt: [] };
}
