# xlCode

Excel in Copilot（Web）を疑似的なコーディングエージェントとして使うため、Excel ブック（`*.xlcode.xlsx`）とソースコードを **Build / Sync / Refresh Tree** で連携させるツール。仕様は「xlCode 仕様書（2026-09-17 改訂）」に従う。

## 進捗

| 段階 | 内容 | 状態 |
|---|---|---|
| 1 | コア処理（Build / Sync / Refresh Tree / 衝突 / 省略検知 / Git 自動コミット）＋テスト | ✅ 完了 |
| 1 | 動作確認用 CLI | ✅ 完了 |
| 2 | Electron + React GUI（VS Code 風ダークテーマ） | ✅ 完了 |
| 3 | Windows 固有機能（Excel 起動・OneDrive 同期待ち・ターミナル起動） | 未着手（Windows での確認が必要） |

## npm install できない環境で使う

配布物は 2 種類。

| 配布物 | 中身 | 必要なもの |
|---|---|---|
| `xlCode 1.0.0.exe` | そのまま起動できる Windows 版（portable） | なし |
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
npm run package:win  :: exe を作る（electron-builder は NSIS などを初回にダウンロードする。オフラインでは不可）
```

## 使い方（GUI）

```sh
npm install
npm run dev              # 開発起動
npm run package:win      # Windows 用にパッケージ（dist/ に portable と installer）
```

- **エクスプローラー**: ブック一覧と各ファイルの状態（E = Excel側で編集中、S = エディタ側で変更、C = 両側で変更）
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

- **ブック・ソース・state.json は一時ファイルに書いてから置き換える**（途中で失敗しても元のファイルは壊れない。一時ファイル名は `~$xlcode-*.tmp`）
- **書き込み直前にもう一度「開かれていないか」を確認する**（確認〜保存の間に Excel で開かれた場合に備える）
- **state.json が壊れていたら止める**（黙って初期化すると全ファイルが衝突扱いになるため）。保存のたびに `state.json.bak` を残す
- **Markdown の保存していない編集はファイルや画面を切り替えても残す**。ほかの場所でファイルが変更されたら知らせる。閉じるときに確認する

## 技術スタックの補足

- Vite は 7 系（electron-vite が Vite 8 未対応のため）
- TypeScript は 6.0 系（typescript-eslint が 7 系未対応のため）

## 受け入れ基準（13章）

`tests/acceptance.test.ts`: 20 ファイル × 500 行を全件変更した Build が約 0.2 秒（基準 3 秒）、Build → Sync の往復で差分なし。
