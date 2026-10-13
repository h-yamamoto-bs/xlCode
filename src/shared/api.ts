import type {
  BookStatus,
  BuildOptions,
  CommitInfo,
  CreateBookResult,
  DirFilesPreview,
  FileContents,
  FileDiff,
  GitSummary,
  OpResult,
  RefreshResult,
  ProjectMode,
  RelocateOptions,
  RelocateResult,
  SyncOptions,
  UndoResult,
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
  /** ブックの置き場所（設定されていれば絶対パス。null ならソースの中） */
  bookRoot: string | null;
  /** プロジェクトの種類 */
  mode: ProjectMode;
  /** 種類が設定済みか（未設定のまま最初のブックを作るときに選ぶ） */
  modeSet: boolean;
  /**
   * xlCode で編集できるドキュメント（ルートからの相対パス）のうち、ファイルとして存在するもの。
   * Agents.md / <dir>/LocalAgents.md / .xlcodeignore
   */
  docs: string[];
  /** 除外ファイルの状態（ブック作成の前の確認に使う） */
  ignore: IgnoreInfo;
}

export interface IgnoreInfo {
  /** .xlcodeignore の有効な行の数（ファイルが無ければ null） */
  xlcode: number | null;
  /** .gitignore の有効な行の数（ファイルが無ければ null） */
  git: number | null;
  /** .gitignore の有効な行の先頭いくつか（表示用） */
  gitSample: string[];
}

/** xlCode で編集できるドキュメントの種類 */
export type DocKind = 'agents' | 'local-agents' | 'ignore' | 'notes' | 'markdown';

/** ファイルを開くアプリ */
export type OpenWith = 'default' | 'vscode';

export type Result<T> = { ok: true; value: T } | { ok: false; error: string };

export interface XlcodeApi {
  platform: string;
  pickProject(): Promise<string | null>;
  loadProject(root: string): Promise<Result<ProjectInfo>>;
  initProject(root: string): Promise<Result<string[]>>;
  refreshTree(root: string): Promise<Result<RefreshResult>>;
  createBook(root: string, dirRel: string): Promise<Result<CreateBookResult>>;
  /** ブックを作る前に、ディレクトリのファイルがどう扱われるか（シート / 除外 / 対象外）を見る */
  previewBook(root: string, dirRel: string): Promise<Result<DirFilesPreview>>;
  build(root: string, bookRel: string, opts: BuildOptions): Promise<Result<OpResult>>;
  sync(root: string, bookRel: string, opts: SyncOptions): Promise<Result<OpResult>>;
  /** web の場合は開いた URL を返す */
  openInExcel(root: string, bookRel: string, via: OpenVia): Promise<Result<string | null>>;
  /** 変更のあるファイルごとの差分（変更を確認する用） */
  bookDiff(root: string, bookRel: string): Promise<Result<FileDiff[]>>;
  /** 1 ファイルの Excel 側とソース側の内容（右のパネルの「内容」用） */
  fileContents(root: string, bookRel: string, name: string): Promise<Result<FileContents>>;
  /** 最近のコミット（dirRel が null ならプロジェクト全体。サイドバーの「Git」用） */
  recentCommits(root: string, dirRel: string | null): Promise<Result<CommitInfo[]>>;
  /** ソースファイルを既定のアプリ（または VS Code）で開く */
  openFile(root: string, rel: string, app: OpenWith): Promise<Result<void>>;
  /** ソースファイルをエクスプローラーで表示する */
  revealFile(root: string, rel: string): Promise<Result<void>>;
  /** 直前の Build / Sync を元に戻す */
  undoLast(root: string, bookRel: string): Promise<Result<UndoResult>>;
  /** 各ブックがデスクトップ版 Excel で開かれているか（ポーリング用） */
  bookLocks(root: string, bookRels: string[]): Promise<Result<Record<string, boolean>>>;
  syncStatus(root: string, bookRels: string[]): Promise<Result<SyncReport>>;
  /** ブックの更新日時とサイズ */
  bookStamp(root: string, bookRel: string): Promise<Result<string>>;
  readConfig(root: string): Promise<Result<XlcodeConfig>>;
  writeConfig(root: string, config: XlcodeConfig): Promise<Result<void>>;
  openTerminal(root: string, dirRel: string): Promise<Result<void>>;
  /** ブックをエクスプローラーで表示する */
  revealInFolder(root: string, bookRel: string): Promise<Result<void>>;
  pickFolder(title: string): Promise<string | null>;
  /** ブックの置き場所を変え、既存のブックを移動する（null ならソースの中へ戻す。opts.move が false なら設定だけ変える） */
  relocateBooks(root: string, newRoot: string | null, opts?: RelocateOptions): Promise<Result<RelocateResult>>;
  /** Markdown（*.md）・.xlcodeignore・.xlcode/notes.md のみ読み書きできる（無ければ null） */
  readDoc(root: string, rel: string): Promise<Result<string | null>>;
  writeDoc(root: string, rel: string, text: string): Promise<Result<void>>;
  /** プロジェクトの種類を決める（一度決めたら変えられない） */
  setProjectMode(root: string, mode: ProjectMode): Promise<Result<void>>;
  /** 取り込む Excel ツールを選ぶ */
  pickToolFile(): Promise<string | null>;
  /** 既存の Excel ツールから編集用ブックを作る（VBA モード） */
  importTool(root: string, dirRel: string, file: string): Promise<Result<CreateBookResult>>;
  /** VBA モードのビルド結果（.xlsm）を開く。reveal ならフォルダで表示 */
  openOutput(root: string, bookRel: string, reveal: boolean): Promise<Result<void>>;
  /** VBA モードの、上書き前のビルド結果のバックアップ（.xlcode/backup/）のフォルダを開く */
  openBackups(root: string, bookRel: string): Promise<Result<void>>;
}
