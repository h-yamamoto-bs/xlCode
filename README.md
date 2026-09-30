# xlCode

Excel in Copilot（Web）を疑似的なコーディングエージェントとして使うため、Excel ブック（`*.xlcode.xlsx`）とソースコードを **Build / Sync / Refresh Tree** で連携させるツール。仕様は「xlCode 仕様書（2026-09-17 改訂）」に従う。

## 進捗

| 段階 | 内容 | 状態 |
|---|---|---|
| 1 | コア処理（Build / Sync / Refresh Tree / 衝突 / 省略検知 / Git 自動コミット）＋テスト | ✅ 完了 |
| 1 | 動作確認用 CLI | ✅ 完了 |
| 2 | Electron + React GUI（VS Code 風ダークテーマ） | ✅ 完了 |
| 3 | Windows 固有機能（Excel 起動・OneDrive 同期待ち・ターミナル起動） | 未着手（Windows での確認が必要） |
| 4 | VBA モード（.bas / .cls / .frm のシート → .xlsm） | 実装・テスト済み（Excel での書き込みは Windows での確認が必要） |

## npm install できない環境で使う

Windows で clone したあとの手順は [docs/windows-setup.md](docs/windows-setup.md)。

配布物は 2 種類。

| 配布物 | 中身 | 必要なもの |
|---|---|---|
| `xlCode-Setup-1.0.0.exe` | Windows 用インストーラー。スタートメニュー・「アプリと機能」に登録される | なし |
| `xlCode-src-win.zip` | ソース一式 + **Windows x64 用**の `node_modules` | Node.js 22 以上 |

`node_modules` は OS ごとに中身が違う（Electron 本体、esbuild・Rollup・Tailwind・lightningcss のネイティブ部品）。
Windows 用は Linux 上で次のように作る。

```sh
npm_config_platform=win32 npm_config_arch=x64 npm ci --os=win32 --cpu=x64
npm_config_platform=win32 npm_config_arch=x64 node node_modules/electron/install.js
```

ソース版の使い方（Windows、npm install 不要）:

```bat
npm run dev          :: 開発起動
npm test             :: テスト
npm run test:e2e     :: GUI の E2E テスト
npm run package:win  :: インストーラーを作る（electron-builder は NSIS などを初回にダウンロードする。オフラインでは不可）
```

## 使い方（GUI）

```sh
npm install
npm run dev              # 開発起動
npm run package:win      # Windows 用インストーラー（dist/xlCode-Setup-*.exe）。Linux では wine（32/64 ビット）と xvfb-run が必要
```

- **エクスプローラー**: ブック一覧と各ファイルの状態（E = Excel側で編集中、S = エディタ側で変更、C = 両側で変更）
- **次の操作**: ブック画面の先頭に、今すべきことを 1 つだけ示す（Build / Sync / Excel を閉じる / 衝突の統合 など）。対応するボタンだけを強調し、押せないボタンには理由をツールチップで出す
- **変更を確認**: ファイル一覧の行（または「変更を確認」）で、Build 後のソース・Sync 後のシートがどう変わるかを差分で見る
- **元に戻す**: 直前の Build / Sync 1 回分を、ソースファイル・ブック・`.xlcode/state.json` ごと操作前に戻す（`.xlcode/undo/` に控えを置く。Git の自動コミットは残る）
- **結果の表示**: Build / Sync の結果（出力・更新したもの、エラー、警告）をブック画面に出す。詳細は出力パネル
- **使う Excel**（設定。初回に確認）: デスクトップ版 / Web 版 / 両方
  - デスクトップ版で開いているかは常に自動検知（`~$` ロックファイル・Windows のファイルロック）し、開いていれば Build / Sync を止める。閉じたら自動で解除
  - Web 版で開いている状態は検知できないため、Build / Sync 前に閉じたかを確認する。「両方」では、xlCode から最後にデスクトップ版で開いたブックは確認を省く
- **デスクトップで開く / Webで開く**: Sync してから起動（5.5-2）。Web 版の URL は Windows の OneDrive 同期設定から自動で求める（求められなければ設定で指定）
- **Build / Sync**: 確認が必要な場合（未コミット変更・省略検知・削除）はダイアログを表示。Sync で未 Build 変更があれば「先に Build / 破棄 / 中止」を選択
- **ルール**: Agents.md（100 行超で警告）と各 LocalAgents.md を編集
- プロジェクトを開いたときに Refresh Tree を自動実行（6.4）。ウィンドウに戻ったときに状態を再読み込み
- ショートカット: Ctrl+O 開く / Ctrl+B Build / Ctrl+Shift+S Sync / Ctrl+J パネル

## 使い方（CLI）

```sh
npm install
npm test                                            # テスト
npm run test:e2e                                    # GUI の E2E テスト（Linux では xvfb-run npm run test:e2e）
npx tsx src/cli.ts init   <project>                 # .gitignore と Agents.md を用意
npx tsx src/cli.ts create <project> app             # app/app.xlcode.xlsx を作成
npx tsx src/cli.ts status <project> app/app.xlcode.xlsx
npx tsx src/cli.ts build  <project> app/app.xlcode.xlsx
npx tsx src/cli.ts sync   <project> app/app.xlcode.xlsx
npx tsx src/cli.ts tree   <project>                 # 全ブックの #tree を更新
```

## 構成

```
src/core/
  ops.ts        Build / Sync
  undo.ts       直前の Build / Sync を元に戻す（書き込み前の控え）
  diff.ts       行差分（変更を確認する表示用）
  scan.ts       ソース側・Excel側・前回値を比較してファイルごとの状態を決める
  tree.ts       Refresh Tree / #tree のバージョン確認
  init.ts       プロジェクト初期化・ブック作成
  canonical.ts  正規化 → Prettier → 正規化（比較・保存用の正準テキスト）
  workbook.ts   ExcelJS による読み書き（A列に1行1セル、書式 @）
  state.ts      .xlcode/state.json
  git.ts        自動コミット・未コミット変更の検出
  lock.ts       ブックが開かれているかの判定
src/main/       Electron メインプロセス（IPC・Excel 起動・ターミナル起動）
src/preload/    contextBridge（sandbox 有効）
src/renderer/   React + Tailwind の画面
src/shared/     メイン⇔画面の型
src/cli.ts      動作確認用 CLI
tests/          Vitest
```

## ファイルごとの状態と処理

前回 Build / Sync 時のハッシュ（`.xlcode/state.json`）と両側の正準テキストのハッシュを比較する。

| 状態 | Build | Sync |
|---|---|---|
| Excel 側だけ変更・新規 | ファイルへ出力 | 中断して選択（破棄 / 先に Build / 中止） |
| ソース側だけ変更・新規 | シートへ取り込み（強制 Sync） | シートへ反映 |
| ソース側で削除 | シートを削除 | シートを削除 |
| Excel 側でシート削除 | シートを復元（ファイルは消さない） | シートを復元 |
| 両側変更 | `#conflict_NN` を作成して中断 | `#conflict_NN` を作成 |
| `DEL_` 付きシート | 確認後、ファイルとシートを削除 | 何もしない |

衝突シートを作ったとき、前回値をソース側に合わせる。Copilot が統合して衝突シートを削除すれば、次の Build で Excel 側が採用される。

## 未決事項の暫定対応

| No. | 暫定対応 | 変更方法 |
|---|---|---|
| 5 | 行末空白は削除（Markdown は対象外） | `.xlcode/config.json` の `trimTrailingWhitespace` |
| 9 | シート名先頭の `DEL_` で削除。確認あり | — |
| 10 | ハッシュは Build / Sync 完了時のみ、正準テキストに対して保存 | — |
| 11 | `.xlcode/state.json` にブック単位 → ファイル単位で hash / 行数 / 文字数 | — |
| 21 | Agents.md シートは Build 対象外。Refresh Tree・ブック作成時にルートから配布。シート側の変更は警告 | — |
| 27 | 省略検知は前回 10 行未満のファイルでは行わない | `shrinkMinLines` |

`.xlcode/config.json` の既定値:

```json
{ "shrinkThreshold": 0.3, "shrinkMinLines": 10, "trimTrailingWhitespace": true, "format": true, "autoCommit": true }
```

## 仕様との差分・注意点

- **Web 版 Excel で開いている状態は検知できない。** 検知できるのはデスクトップ版 Excel（`~$` ロックファイル、Windows のファイルロック）のみ。GUI では「Web 版で閉じたか」の確認で補う。
- **OneDrive 上のブックをデスクトップ版で開いた場合（自動保存・共同編集）に検知できるかは未確認。** [docs/windows-checklist.md](docs/windows-checklist.md) の 1-4。
- **`DEL_` と 31 文字制限**: 28 文字を超えるファイル名には `DEL_` を付けられない（シート名が 31 文字を超える）。
- **整形されるのは変更時のみ**: ブック作成時、シートには整形後の内容が入るが、ソースファイルは次に Excel 側で変更されて Build されるまで元のまま。
- **Excel が作れないシート名**: `[id].tsx` のようなファイルはシートにできないため、Sync・ブック作成がエラーで止まる。対象外にするには `.gitignore` に追加する。
- **No.16（外部書き換えで失われる書式）** は未検証。ExcelJS で作成したブックの往復（数式・先行ゼロ・日付・指数・先頭空白・空行）は文字列のまま保持されることをテストで確認済み。Excel / Copilot が作成したブックでの確認は Windows で行う。

## 壊さないための仕組み

- **直前の Build / Sync は「元に戻す」で戻せる**。書き込む直前に、ブック・書き換えるソースファイル・state.json のそのブックの項目を `.xlcode/undo/` に控える（1 ブックにつき直前の 1 回分）

- **ブック・ソース・state.json は一時ファイルに書いてから置き換える**（途中で失敗しても元のファイルは壊れない。一時ファイル名は `~$xlcode-*.tmp`）
- **書き込み直前にもう一度「開かれていないか」を確認する**（確認〜保存の間に Excel で開かれた場合に備える）
- **state.json が壊れていたら止める**（黙って初期化すると全ファイルが衝突扱いになるため）。保存のたびに `state.json.bak` を残す
- **Markdown の保存していない編集はファイルや画面を切り替えても残す**。ほかの場所でファイルが変更されたら知らせる。閉じるときに確認する

## ブックの置き場所（1 人用）

Copilot を使うにはブックが OneDrive にある必要がある。一方、ソース・`.git`・`node_modules` を OneDrive に置くと同期が重くなり、`.git` が壊れやすい。
そこで、ブックだけを OneDrive 内の専用フォルダに、ソースと同じフォルダ構成で置ける（設定 →「ブックの置き場所」）。
最初のブックを作るときに置き場所を聞く（「OneDrive のフォルダを選ぶ…」か「プロジェクトの中に置く」）。あとから設定画面で変えると、既存のブックも移動する。

編集用ブックのバックアップは OneDrive のバージョン履歴に任せる（xlCode では取らない）。

```
C:\dev\shop\                       ← ソース（OneDrive の外）。正。Git・node_modules・.xlcode\state.json
  app\App.tsx
OneDrive - 会社\xlFolder\shop\     ← ブックだけ
  app\app.xlcode.xlsx              ↔ C:\dev\shop\app\
```

- 対応はフォルダの位置だけで決まる。正はソースのまま（仕様どおり）
- 置き場所を変えると、既存のブックを同じフォルダ構成のまま移動する（開いているブックや同名のファイルがあれば何も動かさない）
- 未設定なら、今までどおりソースの各フォルダの中に置く
- `.xlcode\state.json` は PC ごと。**1 台の PC で使う前提**（2 台で同じブックを編集すると衝突扱いになる）

## コードのシートとそれ以外のシート

- **拡張子のある名前のシートだけをコード**として扱う（`App.tsx`、`run.bat`、`.gitignore` など）。拡張子のない名前は `Makefile`・`Dockerfile`・`LICENSE` などの決まったものと、設定 `extraCodeNames` に追加したものだけ
- それ以外のシート（UI 用、データ表示用など）は一切触らない。画面にも出さない
- ソース側の拡張子のないファイルも対象外（問題パネルに警告）

## ブックの保存方法（UI 用シート・VBA を壊さない）

ブック全体を書き直さず、**変更したコードのシートの XML だけを差し替える**（xlsx は zip）。

- UI 用シートの図形・ボタン・グラフ・書式、`vbaProject.bin`（VBA）などはバイト単位でそのまま残る
- 文字列は共有文字列表の末尾に追加するだけ（既存の番号は変えない。使われなくなった項目は Excel の保存時に整理される）
- シートの削除・並べ替えに合わせて、印刷範囲などの名前の定義（`localSheetId`）と開いているタブを直す
- `calcChain.xml` は削除する（古いセル位置を指していると修復の対象になるため。Excel が作り直す）
- 読み込みは今までどおり ExcelJS

## 文字コード・改行コード

Excel のセルには文字しか入らないため、ファイルごとの形式を判定して、Build のときに元の形式で書き戻す。

- 読める形式: UTF-8 / UTF-8 BOM / UTF-16 LE（BOM あり）/ Shift_JIS（CP932）、改行は LF / CRLF
- 既存のファイルは元の形式を保つ（ブック画面の「形式」列に表示）
- Excel で新しく作ったファイルは拡張子で決める（Windows 11 日本語環境を想定）

| 拡張子 | 文字コード | 改行 |
|---|---|---|
| `.bat` `.cmd` `.vbs` | Shift_JIS | CRLF |
| `.ps1` `.psm1` `.psd1` | UTF-8 BOM | CRLF |
| `.reg` | UTF-16 LE | CRLF |
| それ以外 | UTF-8 | LF |

- ASCII だけのファイルは、日本語が加わったときに備えて拡張子の標準に合わせる
- Shift_JIS のファイルに Shift_JIS で表せない文字（絵文字など）が入ったら、化けさせずに Build を止める
- UTF-16 BE とバイナリは対象外（スキップして警告）

## OneDrive の同期状態（Windows のみ）

- ブックごとの同期状態（同期済み・アップロード待ち・ダウンロード待ち・同期中・一時停止・エラー・オンラインのみ）と OneDrive の起動状態を 10 秒ごとに表示
- Build / Sync / 開く の前に確認し、同期中なら終わるまで自動で待つ（最大 2 分）。一時停止・エラー・未起動なら続けるかを聞く
- 「Webで開く」は、Sync で書き換えたブックのアップロードが終わってから開く
- Web 版で開いたブックは、開いた時点の更新日時を覚えておき、Build / Sync の前にまだ変わっていなければ「Web 版での編集が届いていない可能性」を知らせる
- Refresh Tree は、#tree と Agents.md が最新のブックを保存しない（Web 版で編集中のブックとの同期の衝突を避ける）
- 判定は Windows のシェルプロパティによる。値の意味は未確認（[docs/windows-checklist.md](docs/windows-checklist.md) の 7）

## VBA モード

最初のブックを作るときに「ソースコード」か「VBA」かを選ぶ（`.xlcode/config.json` の `mode`。あとから変更不可）。

```
<ブックの置き場所>/販売管理.xlcode.xlsx   編集用。画面・データのシートと、Module1.bas などのコードのシート
        │ Build ①（ソースコードモードと同じ）          ▲ Sync（エディタで直したソースをシートへ）
        ▼                                              │
<プロジェクト>/Module1.bas など            ソースコード（正本。Git 管理。Shift_JIS・CRLF）
        │ Build ②（Windows のデスクトップ版 Excel で VBE に書き込む）
        ▼
<プロジェクト>/販売管理.xlsm               ビルド結果（画面・データのシート + VBA。Git 管理しない）
```

- ①はソースコードモードの Build そのもの。衝突・省略検知・`DEL_` による削除・自動コミットも同じ。Sync も使える
- ソースコードの中身はシートと同じ（`Attribute VB_Name` などの見出しは付けない。VBE の「ファイルのインポート」ではなく xlCode が書き込む）
- シート名（= ファイル名）と書き込み先
  - `Module1.bas` → 標準モジュール、`Class1.cls` → クラスモジュール、`UserForm1.frm` → ユーザーフォーム
  - `ThisWorkbook.cls` → ブックのコード、`<画面のシート名>.cls` → そのシートのコード（`Sheet1.cls` のようにコード名でもよい）
  - `References.refs` → 参照設定
  - 拡張子のないシートは画面・データとしてビルド結果に残す。`Agents.md`・`#tree` などはビルド結果から除く
- ユーザーフォームは、先頭に `Begin UserForm UserForm1` 〜 `End` で配置を書き、その後にコードを書く（書き方は Agents.md の雛形を参照）。フォームは Build のときに VBE の機能（`Designer.Controls.Add`）で組み立てる
- ②の流れ: 編集用ブックから拡張子付きのシートを除いたコピーを作る → PowerShell から画面に出ない Excel を起動（マクロ無効・確認ダイアログなし・イベント停止・自動保存オフ）→ ソースコードから VBA を書き込んで .xlsm で保存 → ビルド結果を置き換える。失敗したらビルド結果は変えない（①のソースコードは出力済み）。3 分で終わらなければ Excel を強制終了する
- 必要なもの: Windows、デスクトップ版 Excel、「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」
- Git 管理するのはソースコードだけ。`.xlsm` は `.gitignore` に `*.xlsm` を入れて外す。画面・データのシートは編集用ブック（OneDrive）にだけあるので、戻すときは OneDrive のバージョン履歴を使う

### ビルド結果の使い方

- ビルド結果（`販売管理.xlsm`）は Build のたびに作り直す**ひな形**。実際に使うときはコピーして使う（例: `販売管理_2026年10月.xlsm`）。ビルド結果に直接データを入力すると、次の Build で消える（変更されていれば確認し、`.xlcode/backup/` に 5 世代残す）
- 形式は `.xlsm` だけ（アドイン `.xlam`・テンプレート `.xltm` は必要になったら追加する）
- メールやダウンロードで受け取った `.xlsm` は、Windows の安全機能でマクロが止められることがある。ファイルを右クリック →「プロパティ」→「全般」タブの下の「セキュリティ: 許可する」にチェック →「OK」で解除できる

### 参照設定

外部のライブラリ（Dictionary・正規表現・ADO など）の使い方は 2 通りあり、どちらでもよい。**おすすめは B**。

| | 書き方 | 必要なこと |
|---|---|---|
| A（参照設定） | `Dim d As New Scripting.Dictionary` | `References.refs` にライブラリを書く（Build のときにビルド結果へ追加する） |
| B（CreateObject） | `Dim d As Object: Set d = CreateObject("Scripting.Dictionary")` | なし。ほかの PC でも「参照不可」になりにくい |

`References.refs` は 1 行に 1 つ `{GUID} 主.副 説明` の形で書く（`'` で始まる行は無視）。よく使うものの例はファイルの先頭のコメントにある。

## 取り込み（既存のソース・Excel ツールから編集用ブックを作る）

| モード | 取り込み元 | 操作 |
|---|---|---|
| ソースコード | ディレクトリのファイル | サイドバーの「ブック未作成のディレクトリ」で＋（1 つずつ）、見出しの＋で全ディレクトリをまとめて作成 |
| VBA | 既存の Excel ツール（.xlsm / .xlsb / .xls / .xlsx / .xltm） | ＋ →「既存の Excel ツールから作成…」でファイルを選ぶ |
| VBA | ソースコード（`.bas` / `.cls` / `.frm` / `References.refs`） | ＋ →「空のブックを作成」（フォルダのファイルをシートとして取り込む。無ければ雛形を置く） |

VBA モードの取り込み（Windows・デスクトップ版 Excel が必要）:

- 元のツールはコピーしてから開くため変更しない（開いたままでも取り込める）
- シート（画面・データ）は .xlsx として保存し直したものをそのまま使う（図形・ボタン・書式・入力規則が残る）
- VBA はソースコードのファイルとして書き出してから、シートとして取り込む（同じ名前のファイルがあれば何もせずに止める）
- 標準・クラスモジュール → `.bas` / `.cls`、ブック・シートのコード → `ThisWorkbook.cls` / `<シート名>.cls`（中身が無いものは省く）
- ユーザーフォーム → `.frm`。配置を `Begin` 〜 `End` の形に書き起こす。既定値と同じプロパティは省く。画像・標準以外のコントロール・文字列の改行は取り込めない（注意として出す）
- 参照設定 → `References.refs`（ほかのブックへの参照は取り込めない）
- 最初の Build では、元のツールがビルド結果の場所にあれば上書きの確認が出る（元のファイルは `.xlcode/backup/` に残る）

## 技術スタックの補足

- Vite は 7 系（electron-vite が Vite 8 未対応のため）
- TypeScript は 6.0 系（typescript-eslint が 7 系未対応のため）

## 受け入れ基準（13章）

`tests/acceptance.test.ts`: 20 ファイル × 500 行を全件変更した Build が約 0.2 秒（基準 3 秒）、Build → Sync の往復で差分なし。
