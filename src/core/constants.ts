/** 管理対象ブックの接尾辞 */
export const BOOK_SUFFIX = '.xlcode.xlsx';

/** 予約シート（接頭辞「#」） */
export const RESERVED_PREFIX = '#';
export const TREE_SHEET = '#tree';
export const META_SHEET = '#meta';
export const REFS_SHEET = '#refs';
export const CONFLICT_PREFIX = '#conflict_';
export const RESERVED_SHEETS: readonly string[] = [TREE_SHEET, META_SHEET, REFS_SHEET];

/** ルールシート */
export const AGENTS_SHEET = 'Agents.md';
export const LOCAL_AGENTS_SHEET = 'LocalAgents.md';

/** 削除マーク（未決 No.9 の暫定案） */
export const DELETE_PREFIX = 'DEL_';

/** Excel 外のツール用ディレクトリ */
export const XLCODE_DIR = '.xlcode';
export const STATE_FILE = 'state.json';
export const TREE_FILE = 'tree.txt';
export const CONFIG_FILE = 'config.json';

/** Excel の制約 */
export const MAX_SHEET_NAME = 31;
export const MAX_CELL_CHARS = 32767;
/** Excel が開けるファイルのフルパスの最大文字数（フォルダ＋ファイル名） */
export const EXCEL_MAX_PATH = 218;
export const FORBIDDEN_CHARS: readonly string[] = [':', '\\', '/', '?', '*', '[', ']'];

/** #tree 1行目の接頭辞 */
export const TREE_VERSION_PREFIX = '# tree-version: ';
