import type {
  BookStatus,
  BuildOptions,
  CreateBookResult,
  GitSummary,
  OpResult,
  RefreshResult,
  SyncOptions,
  XlcodeConfig,
} from '../core';

/** どの Excel で編集するか */
export type ExcelMode = 'desktop' | 'web' | 'both';
export type OpenVia = 'desktop' | 'web';

/**
 * OneDrive の同期状態
 * - unsupported: Windows 以外 / outside: OneDrive の外 / online-only: このPCにダウンロードされていない
 * - uploading / downloading / syncing: 転送待ち・転送中 / paused: 一時停止 / error: エラー / unknown: 読み取れない
 */
export type SyncState =
  | 'unsupported'
  | 'outside'
  | 'synced'
  | 'online-only'
  | 'uploading'
  | 'downloading'
  | 'syncing'
  | 'paused'
  | 'error'
  | 'unknown';

export interface BookSync {
  state: SyncState;
  /** 診断用の生の値 */
  detail?: string;
}

export interface SyncReport {
  /** OneDrive が起動しているか（調べていない・分からない場合は null） */
  oneDriveRunning: boolean | null;
  books: Record<string, BookSync>;
}

export interface BookSummary extends Partial<BookStatus> {
  /** ルートからの相対パス */
  rel: string;
  dirRel: string;
  /** デスクトップ版 Excel で開かれている */
  open: boolean;
  openReason?: string;
  /** 読み込みに失敗した場合 */
  loadError?: string;
}

export interface ProjectInfo {
  root: string;
  name: string;
  books: BookSummary[];
  /** 現在のプロジェクトツリーのバージョン */
  treeVersion: string;
  git: GitSummary;
  hasAgents: boolean;
  /** ブックを持たないディレクトリ（ブック作成の候補） */
  dirsWithoutBook: string[];
}

export type Result<T> = { ok: true; value: T } | { ok: false; error: string };

export interface XlcodeApi {
  platform: string;
  pickProject(): Promise<string | null>;
  loadProject(root: string): Promise<Result<ProjectInfo>>;
  initProject(root: string): Promise<Result<string[]>>;
  refreshTree(root: string): Promise<Result<RefreshResult>>;
  createBook(root: string, dirRel: string): Promise<Result<CreateBookResult>>;
  build(root: string, bookRel: string, opts: BuildOptions): Promise<Result<OpResult>>;
  sync(root: string, bookRel: string, opts: SyncOptions): Promise<Result<OpResult>>;
  /** web の場合は開いた URL を返す */
  openInExcel(root: string, bookRel: string, via: OpenVia): Promise<Result<string | null>>;
  /** 各ブックがデスクトップ版 Excel で開かれているか（ポーリング用） */
  bookLocks(root: string, bookRels: string[]): Promise<Result<Record<string, boolean>>>;
  syncStatus(root: string, bookRels: string[]): Promise<Result<SyncReport>>;
  /** ブックの更新日時とサイズ */
  bookStamp(root: string, bookRel: string): Promise<Result<string>>;
  readConfig(root: string): Promise<Result<XlcodeConfig>>;
  writeConfig(root: string, config: XlcodeConfig): Promise<Result<void>>;
  openTerminal(root: string, dirRel: string): Promise<Result<void>>;
  revealInFolder(root: string, rel: string): Promise<Result<void>>;
  /** Agents.md / LocalAgents.md のみ読み書きできる */
  readRuleFile(root: string, rel: string): Promise<Result<string | null>>;
  writeRuleFile(root: string, rel: string, text: string): Promise<Result<void>>;
}
