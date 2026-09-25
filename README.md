# xlCode

Excel in Copilot（Web）を疑似的なコーディングエージェントとして使うため、Excel ブック（`*.xlcode.xlsx`）とソースコードを **Build / Sync / Refresh Tree** で連携させるツール。仕様は「xlCode 仕様書（2026-09-17 改訂）」に従う。

## 進捗

| 段階 | 内容 | 状態 |
|---|---|---|
| 1 | コア処理（Build / Sync / Refresh Tree / 衝突 / 省略検知 / Git 自動コミット）＋テスト | ✅ 完了 |
| 1 | 動作確認用 CLI | ✅ 完了 |
| 2 | Electron + React GUI | 未着手 |
| 3 | Windows 固有機能（Excel 起動・OneDrive 同期待ち・ターミナル起動） | 未着手（Windows での確認が必要） |

## 使い方（CLI）

```sh
npm install
npm test                                            # テスト
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

- **Web 版 Excel で開いている状態は検知できない。** 検知できるのはデスクトップ版 Excel（`~$` ロックファイル、Windows のファイルロック）のみ。GUI では「Web 版で閉じたか」の確認で補う予定。
- **`DEL_` と 31 文字制限**: 28 文字を超えるファイル名には `DEL_` を付けられない（シート名が 31 文字を超える）。
- **整形されるのは変更時のみ**: ブック作成時、シートには整形後の内容が入るが、ソースファイルは次に Excel 側で変更されて Build されるまで元のまま。
- **Excel が作れないシート名**: `[id].tsx` のようなファイルはシートにできないため、Sync・ブック作成がエラーで止まる。対象外にするには `.gitignore` に追加する。
- **No.16（外部書き換えで失われる書式）** は未検証。ExcelJS で作成したブックの往復（数式・先行ゼロ・日付・指数・先頭空白・空行）は文字列のまま保持されることをテストで確認済み。Excel / Copilot が作成したブックでの確認は Windows で行う。

## 受け入れ基準（13章）

`tests/acceptance.test.ts`: 20 ファイル × 500 行を全件変更した Build が約 0.2 秒（基準 3 秒）、Build → Sync の往復で差分なし。
