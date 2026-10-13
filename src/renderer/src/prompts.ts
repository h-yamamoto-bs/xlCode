import type { ProjectMode } from '../../core';

/** Excel の Copilot に貼り付ける指示のテンプレート（右のパネルの「指示」） */
export interface PromptTemplate {
  id: string;
  title: string;
  /** どんなときに使うか */
  when: string;
  /** 貼り付ける文。{file} などは選んでいるファイルで置き換える */
  text: string;
  /** VBA モードだけ / ソースコードモードだけ */
  mode?: ProjectMode;
}

export interface PromptGroup {
  title: string;
  items: PromptTemplate[];
}

export const PROMPT_GROUPS: PromptGroup[] = [
  {
    title: '基本（最初に伝える）',
    items: [
      {
        id: 'rules',
        title: 'ルールを読ませる',
        when: 'ブックを開いて最初に。Agents.md と LocalAgents.md のルールに従わせる',
        text: 'このブックの Agents.md シートと LocalAgents.md シートのルールを読んで、以後はそれに従ってください。#tree シートがプロジェクトの構成です。1 シート = 1 ファイルで、コードは A 列に 1 行ずつ書きます。',
      },
      {
        id: 'no-omit',
        title: '省略させない',
        when: '大きなファイルを編集させる前に',
        text: 'コードは省略せずにファイル全体を書いてください。「…既存のコード…」「// 省略」のような書き方はしないでください。',
      },
      {
        id: 'sheet-rules',
        title: 'シートの書き方を守らせる',
        when: 'Copilot が B 列に書いたり、数値に変換したりしたとき',
        text: 'A 列だけに、1 セル 1 行で書いてください。B 列以降には何も書かないでください。セルは文字列のままにして、数式や日付・数値に変換しないでください。',
      },
    ],
  },
  {
    title: '編集',
    items: [
      {
        id: 'edit-file',
        title: 'ファイルを直す',
        when: '選んでいるファイルを編集させる',
        text: '{file} シートの内容を次のように変更してください（ファイル全体を書き直す）: ',
      },
      {
        id: 'new-file',
        title: 'ファイルを追加する',
        when: '新しいファイルを作らせる（シート名 = ファイル名、拡張子込み 31 文字以内）',
        text: '新しいシート「{newfile}」を作り、A 列に次の内容のコードを 1 行ずつ書いてください: ',
      },
      {
        id: 'delete-file',
        title: 'ファイルを削除する',
        when: 'シートを消すだけでは削除されない。DEL_ を付ける',
        text: '{file} シートの名前を「DEL_{file}」に変えてください（中身は変えない）。',
      },
      {
        id: 'rename-file',
        title: 'ファイル名を変える',
        when: '新しい名前のシートに移して、古い方に DEL_ を付ける',
        text: '新しいシート「{newfile}」を作って {file} シートの内容をそのまま移し、{file} シートの名前を「DEL_{file}」に変えてください。',
      },
    ],
  },
  {
    title: '衝突・確認',
    items: [
      {
        id: 'merge',
        title: '衝突シートを統合する',
        when: 'Build / Sync で #conflict_NN ができたとき',
        text: '{conflict} シートの A 列（Excel 側）と B 列（ソース側）を統合して {file} シートに書き、{conflict} シートを削除してください。どちらの変更も失わないでください。',
      },
      {
        id: 'review',
        title: '変更点を説明させる',
        when: 'Build する前に、何を変えたか確かめる',
        text: '{file} シートで変更した箇所と理由を、箇条書きで説明してください。コードは書き直さないでください。',
      },
      {
        id: 'check',
        title: '自己確認させる',
        when: '文法の誤りや省略が無いか見直させる',
        text: '{file} シートのコードに文法の誤り、途中で切れている箇所、「省略」と書いた箇所が無いか確認し、あれば直してください。',
      },
    ],
  },
  {
    title: 'VBA',
    items: [
      {
        id: 'vba-module',
        title: 'モジュールを追加する',
        when: 'VBA モードで標準モジュールを作らせる',
        mode: 'vba',
        text: '新しいシート「Module2.bas」を作り、先頭に Option Explicit を書いて、次の処理の Sub を書いてください。Attribute 行は書かないでください: ',
      },
      {
        id: 'vba-form',
        title: 'ユーザーフォームを作る',
        when: 'VBA モードでフォームを作らせる（配置を Begin〜End で書く）',
        mode: 'vba',
        text: '新しいシート「UserForm1.frm」を作り、Agents.md シートの例の形式（先頭に Begin UserForm 〜 End の配置、その後にコード）で、次のフォームを書いてください: ',
      },
      {
        id: 'vba-ui',
        title: '画面のシートの構成を伝える',
        when: 'ボタンに登録するマクロ名やセルの位置を揃える',
        mode: 'vba',
        text: '拡張子の無いシートは画面・データ用です。指示が無い限り中身は変えないでください。ボタンに登録するマクロ名と、読み書きするセルの位置は LocalAgents.md シートに書いてあるとおりにしてください。',
      },
    ],
  },
];

export interface PromptContext {
  /** 選んでいるファイル名（無ければ占位のまま） */
  file?: string;
  /** 未解決の衝突シート（最初の 1 つ） */
  conflict?: string;
}

/** {file} などを、今の状態で置き換える */
export function fillPrompt(text: string, ctx: PromptContext): string {
  return text
    .replaceAll('{file}', ctx.file ?? '（ファイル名）')
    .replaceAll('{newfile}', '（新しいファイル名）')
    .replaceAll('{conflict}', ctx.conflict ?? '#conflict_01');
}
