/**
 * ユーザーフォームのシート（UserForm1.frm）の読み取り。
 *
 * シートの先頭に VB6 の .frm に似た書き方で画面の配置を書き、その後にコードを書く。
 *
 *   Begin UserForm UserForm1
 *      Caption = "顧客登録"
 *      Width = 300
 *      Begin CommandButton btnOK
 *         Caption = "登録"
 *         Left = 150
 *      End
 *   End
 *
 *   Private Sub btnOK_Click()
 *   ...
 *
 * 実際のフォームは Build のときに Excel の VBE の機能（Designer.Controls.Add）で組み立てる。
 * .frx（バイナリ）を xlCode が直接作ることはしない。
 */

export type PropType = 'string' | 'int' | 'double' | 'bool';

export interface FormProp {
  /** プロパティ名（Font.Size のように . で区切ってもよい） */
  name: string;
  value: string | number | boolean;
  type: PropType;
  /** シートの行番号（1 始まり。エラー表示用） */
  line: number;
}

export interface FormControl {
  /** 種類（CommandButton など。Page は MultiPage のページ） */
  type: string;
  /** Controls.Add に渡す ProgID（Page は空） */
  progId: string;
  name: string;
  props: FormProp[];
  children: FormControl[];
  line: number;
}

export interface FormDef {
  name: string;
  props: FormProp[];
  children: FormControl[];
}

export interface ParsedForm {
  form: FormDef | null;
  /** 配置の後に書かれたコード（先頭の空行は除く） */
  code: string[];
  errors: string[];
}

/** 使えるコントロール（MSForms の標準のもの） */
export const FORM_CONTROLS: Readonly<Record<string, string>> = {
  Label: 'Forms.Label.1',
  TextBox: 'Forms.TextBox.1',
  CommandButton: 'Forms.CommandButton.1',
  ComboBox: 'Forms.ComboBox.1',
  ListBox: 'Forms.ListBox.1',
  CheckBox: 'Forms.CheckBox.1',
  OptionButton: 'Forms.OptionButton.1',
  ToggleButton: 'Forms.ToggleButton.1',
  Frame: 'Forms.Frame.1',
  MultiPage: 'Forms.MultiPage.1',
  TabStrip: 'Forms.TabStrip.1',
  ScrollBar: 'Forms.ScrollBar.1',
  SpinButton: 'Forms.SpinButton.1',
  Image: 'Forms.Image.1',
};

/** 子を持てる種類 */
const CONTAINERS = new Set(['UserForm', 'Frame', 'Page']);

/** VBA の識別子（モジュール名・コントロール名）。先頭は文字、以降は文字・数字・_ */
export function isVbaIdentifier(name: string): boolean {
  return name.length <= 31 && /^\p{L}[\p{L}\p{N}_]*$/u.test(name);
}

/** ' 以降のコメントを除く（文字列の中の ' は除かない） */
function stripComment(s: string): string {
  let inStr = false;
  for (let i = 0; i < s.length; i++) {
    if (s[i] === '"') inStr = !inStr;
    else if (s[i] === "'" && !inStr) return s.slice(0, i);
  }
  return s;
}

function parseValue(raw: string): { value: string | number | boolean; type: PropType } | null {
  const v = raw.trim();
  const str = /^"((?:[^"]|"")*)"$/.exec(v);
  if (str) return { value: str[1].replace(/""/g, '"'), type: 'string' };
  if (/^true$/i.test(v)) return { value: true, type: 'bool' };
  if (/^false$/i.test(v)) return { value: false, type: 'bool' };
  const hex = /^&H([0-9A-F]{1,8})&?$/i.exec(v);
  if (hex) {
    // 色（&H8000000F& など）は Long（符号付き 32 ビット）として渡す
    const n = parseInt(hex[1], 16);
    return { value: n > 0x7fffffff ? n - 0x100000000 : n, type: 'int' };
  }
  if (/^-?\d+$/.test(v)) {
    const n = Number(v);
    return Math.abs(n) <= 0x7fffffff ? { value: n, type: 'int' } : { value: n, type: 'double' };
  }
  if (/^-?\d*\.\d+$/.test(v)) return { value: Number(v), type: 'double' };
  return null;
}

function canonicalType(t: string): string | null {
  const bare = t.replace(/^MSForms\./i, '');
  if (/^userform$/i.test(bare)) return 'UserForm';
  if (/^page$/i.test(bare)) return 'Page';
  return Object.keys(FORM_CONTROLS).find((k) => k.toLowerCase() === bare.toLowerCase()) ?? null;
}

/**
 * フォームのシートを読む。
 * @param sheet シート名（UserForm1.frm）。フォーム名はシート名の拡張子を除いたものと一致させる
 */
export function parseFormSheet(sheet: string, lines: readonly string[]): ParsedForm {
  const errors: string[] = [];
  const at = (i: number, msg: string) => errors.push(`「${sheet}」${i + 1} 行目: ${msg}`);
  const formName = sheet.replace(/\.frm$/i, '');

  let i = 0;
  while (i < lines.length && (lines[i].trim() === '' || lines[i].trim().startsWith("'"))) i++;
  const head = i < lines.length ? /^Begin\s+(\S+)\s+(\S+)\s*$/i.exec(stripComment(lines[i]).trim()) : null;
  if (!head || canonicalType(head[1]) !== 'UserForm') {
    errors.push(`「${sheet}」: 先頭に「Begin UserForm ${formName}」〜「End」でフォームの配置を書いてください`);
    return { form: null, code: [], errors };
  }
  if (head[2].toLowerCase() !== formName.toLowerCase()) {
    at(i, `フォーム名「${head[2]}」がシート名と違います（「Begin UserForm ${formName}」にしてください）`);
  }

  const form: FormDef = { name: formName, props: [], children: [] };
  type Frame = { type: string; props: FormProp[]; children: FormControl[] };
  const stack: Frame[] = [{ type: 'UserForm', props: form.props, children: form.children }];
  const names = new Map<string, number>();
  let end = -1;

  for (i = i + 1; i < lines.length; i++) {
    const text = stripComment(lines[i]).trim();
    if (text === '') continue;
    const parent = stack[stack.length - 1];

    const begin = /^Begin\s+(\S+)\s+(\S+)$/i.exec(text);
    if (begin) {
      const type = canonicalType(begin[1]);
      const name = begin[2];
      const node: FormControl = {
        type: type ?? begin[1],
        progId: type && type !== 'Page' && type !== 'UserForm' ? FORM_CONTROLS[type] : '',
        name,
        props: [],
        children: [],
        line: i + 1,
      };
      if (!type || type === 'UserForm') {
        at(i, `「${begin[1]}」は使えません（使えるもの: ${Object.keys(FORM_CONTROLS).join(', ')}）`);
      } else if (type === 'Page' && parent.type !== 'MultiPage') {
        at(i, 'Page は MultiPage の中にだけ書けます');
      } else if (parent.type === 'MultiPage' && type !== 'Page') {
        at(i, `MultiPage の中には Page を書き、その中に ${type} を書いてください`);
      } else if (parent.type !== 'MultiPage' && !CONTAINERS.has(parent.type)) {
        at(i, `${parent.type} の中にはコントロールを置けません（置けるのは UserForm・Frame・Page の中）`);
      }
      if (!isVbaIdentifier(name)) {
        at(i, `名前「${name}」は使えません（先頭は文字、以降は文字・数字・_ で 31 文字以内）`);
      } else if (names.has(name.toLowerCase())) {
        at(i, `名前「${name}」は ${names.get(name.toLowerCase())} 行目でも使われています`);
      } else {
        names.set(name.toLowerCase(), i + 1);
      }
      parent.children.push(node);
      stack.push(node);
      continue;
    }

    if (/^End$/i.test(text)) {
      stack.pop();
      if (stack.length === 0) {
        end = i;
        break;
      }
      continue;
    }

    const prop = /^([A-Za-z_]\w*(?:\.[A-Za-z_]\w*)*)\s*=\s*(.*)$/.exec(text);
    if (prop) {
      const [, name, raw] = prop;
      if (/^name$/i.test(name)) {
        at(i, '名前は「Begin 種類 名前」の行で指定してください');
        continue;
      }
      const v = parseValue(raw);
      if (!v) {
        at(
          i,
          `値「${raw.trim()}」を読めません。文字列は "..." で囲み、定数（fmBorderStyleSingle など）は数値で書いてください`,
        );
        continue;
      }
      parent.props.push({ name, ...v, line: i + 1 });
      continue;
    }

    at(i, `読めない行です: ${text}（「プロパティ = 値」「Begin 種類 名前」「End」のどれかで書いてください）`);
  }

  if (end < 0) {
    errors.push(`「${sheet}」: フォームの配置の「End」が足りません（Begin と End の数を合わせてください）`);
    return { form: null, code: [], errors };
  }
  const code = lines.slice(end + 1);
  while (code.length > 0 && code[0].trim() === '') code.shift();
  return { form: errors.length > 0 ? null : form, code, errors };
}
