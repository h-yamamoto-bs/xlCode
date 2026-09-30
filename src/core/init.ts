import { access, appendFile, mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { AGENTS_SHEET, LOCAL_AGENTS_SHEET, REFS_SHEET, XLCODE_DIR } from './constants';
import { excelPathError, listSourceFiles, readTextFile } from './fsutil';
import { textToLines } from './normalize';
import { bookPathFor, bookRef, bookRootOf, booksInDir, openProject } from './project';
import { bookState, saveState } from './state';
import { loadConfig } from './config';
import { readCopies } from './vba';
import { renderRefs } from './vbaRefs';
import { applyTreeToBook, computeTree, readLocalAgents, readRootAgents } from './tree';
import { isCodeName, validateFileName } from './sheetName';
import { Book } from './workbook';

const GITIGNORE_ENTRIES = ['*.xlcode.xlsx', '~$*', `${XLCODE_DIR}/`];

const AGENTS_TEMPLATE = `# Agents.md

- TypeScript の strict モードを使う
- コメントは日本語で書く
- 1シート = 1ファイル。A列に1行ずつ書く
- シート名はファイル名（拡張子込み・31文字以内）
- 「#」で始まるシートは編集しない
`;

const LOCAL_AGENTS_TEMPLATE = `# LocalAgents.md

- このディレクトリ固有のルールを書く
`;

const VBA_AGENTS_TEMPLATE = `# Agents.md

- VBA のコードを書く。1シート = 1モジュール。A列に1行ずつ書く
- シート名はモジュール名＋拡張子（31文字以内）
  - 標準モジュール: Module1.bas / クラスモジュール: Class1.cls / ユーザーフォーム: UserForm1.frm
  - シートのイベント: 「シート名.cls」、ブックのイベント: ThisWorkbook.cls
- 各モジュールの先頭に Option Explicit を書く。Attribute 行は書かない
- 外部のライブラリ（Dictionary・正規表現・ADO など）は CreateObject で使う書き方をおすすめする
  （例: \`Dim d As Object: Set d = CreateObject("Scripting.Dictionary")\`）
  - 参照設定が必要な書き方（\`Dim d As New Scripting.Dictionary\`）をした場合は、#refs シートへの追加が必要だと利用者に伝える
- 拡張子のないシートは画面（UI）とデータ。指示が無い限り中身を変えない
- 「#」で始まるシート、Agents.md・LocalAgents.md シートは編集しない
- ユーザーフォームは、先頭に配置を書き、その後にコードを書く

\`\`\`vb
Begin UserForm UserForm1
   Caption = "顧客登録"
   Width = 300
   Height = 200
   Begin TextBox txtName
      Left = 80
      Top = 12
      Width = 150
   End
   Begin CommandButton btnOK
      Caption = "登録"
      Left = 150
      Top = 150
   End
End

Private Sub btnOK_Click()
    MsgBox txtName.Value
End Sub
\`\`\`

- 使えるコントロール: Label, TextBox, CommandButton, ComboBox, ListBox, CheckBox, OptionButton, ToggleButton, Frame, MultiPage（中に Page）, TabStrip, ScrollBar, SpinButton, Image
- 値は 文字列 "…"・数値・True / False・&H8000000F& のように書く。定数名（fmBorderStyleSingle など）は数値で書く
`;

const VBA_LOCAL_AGENTS_TEMPLATE = `# LocalAgents.md

- このツール固有のルールを書く
- UI シートの構成（どのシートの何のセルに何があるか、ボタンに登録するマクロ名など）を書いておくと、Copilot が正しいコードを書きやすくなる
`;

async function exists(p: string): Promise<boolean> {
  try {
    await access(p);
    return true;
  } catch {
    return false;
  }
}

/**
 * プロジェクトの初期設定。
 * - .gitignore に xlCode 用の除外（ブック・ロックファイル・.xlcode/）を追記（8.2）
 * - ルートに Agents.md が無ければ雛形を作成
 */
export async function initProject(root: string): Promise<string[]> {
  const done: string[] = [];
  const vba = (await loadConfig(root)).mode === 'vba';
  const gi = path.join(root, '.gitignore');
  const current = (await readFile(gi, 'utf8').catch(() => '')).split(/\r?\n/).map((l) => l.trim());
  const missing = GITIGNORE_ENTRIES.filter((e) => !current.includes(e));
  if (missing.length > 0) {
    const prefix = current.join('') === '' || current[current.length - 1] === '' ? '' : '\n';
    await appendFile(gi, `${prefix}# xlCode\n${missing.join('\n')}\n`);
    done.push(`.gitignore に追記: ${missing.join(', ')}`);
  }
  const agents = path.join(root, AGENTS_SHEET);
  if (!(await exists(agents))) {
    await writeFile(agents, vba ? VBA_AGENTS_TEMPLATE : AGENTS_TEMPLATE);
    done.push('Agents.md を作成');
  }
  return done;
}

export interface CreateBookResult {
  book: string;
  sheets: string[];
  skipped: string[];
}

/** ディレクトリに <dir_name>.xlcode.xlsx を作成し、既存ファイルをシートとして取り込む */
export async function createBook(root: string, dirAbs: string): Promise<CreateBookResult> {
  const ctx = await openProject(root);
  const bookRoot = bookRootOf(root, ctx.config);
  const bookAbs = bookPathFor(root, bookRoot, dirAbs);
  const existing = await booksInDir(path.dirname(bookAbs));
  if (existing.length > 0) throw new Error(`既にブックがあります: ${existing.join(', ')}`);
  const longPath = excelPathError(bookAbs);
  if (longPath) throw new Error(longPath);
  await mkdir(path.dirname(bookAbs), { recursive: true });
  const ref = bookRef(root, bookAbs, bookRoot);

  const vba = ctx.config.mode === 'vba';
  const localAgents = path.join(dirAbs, LOCAL_AGENTS_SHEET);
  if (!(await exists(localAgents)))
    await writeFile(localAgents, vba ? VBA_LOCAL_AGENTS_TEMPLATE : LOCAL_AGENTS_TEMPLATE);

  // 作成するブック自身も #tree に載るよう、先に空のブックを保存してからツリーを取る
  const book = Book.create();
  await book.save(bookAbs);
  const bs = bookState(ctx.state, ref.rel);
  const sheets: string[] = [];
  const skipped: string[] = [];
  if (vba) {
    // VBA モード: ルールのシートと、控え（vba/）があればそのコードをシートにする
    applyTreeToBook(book, await computeTree(root, ctx.ig), await readRootAgents(root), await readLocalAgents(dirAbs));
    for (const c of await readCopies(dirAbs)) {
      book.writeLines(c.sheet, c.lines);
      sheets.push(c.sheet);
    }
    if (!sheets.includes(REFS_SHEET)) book.writeLines(REFS_SHEET, renderRefs([]));
    if (!sheets.some((n) => n !== REFS_SHEET)) book.writeLines('Module1.bas', ['Option Explicit', '']);
    await book.save(bookAbs);
    // 控えから作ったシートは、次の Build で「未ビルド」として扱う（ビルド結果が控えと同じとは限らないため）
    bs.files = {};
    await saveState(root, ctx.state);
    return { book: ref.rel, sheets, skipped };
  }
  applyTreeToBook(book, await computeTree(root, ctx.ig), await readRootAgents(root));
  const files = await listSourceFiles(root, dirAbs, ctx.ig);
  // LocalAgents.md を先頭に
  files.sort((a, b) => Number(b.name === LOCAL_AGENTS_SHEET) - Number(a.name === LOCAL_AGENTS_SHEET));
  for (const f of files) {
    if (!isCodeName(f.name, ctx.config.extraCodeNames)) {
      skipped.push(`${f.name}（拡張子がないため対象外）`);
      continue;
    }
    const err = validateFileName(f.name);
    if (err) throw new Error(err);
    const read = await readTextFile(f.abs);
    if (read.kind === 'binary') {
      skipped.push(`${f.name}（${read.reason}）`);
      continue;
    }
    const c = await ctx.canon.canonical(f.name, f.abs, read.text);
    book.writeLines(f.name, textToLines(c.text));
    bs.files[f.name] = { hash: c.hash, lines: c.lines, chars: c.chars };
    sheets.push(f.name);
  }
  await book.save(bookAbs);
  bs.lastSyncAt = new Date().toISOString();
  await saveState(root, ctx.state);
  return { book: ref.rel, sheets, skipped };
}
