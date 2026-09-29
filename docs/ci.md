# Windows での自動テスト（GitHub Actions）— 未導入

開発環境（Linux）では Windows 固有の動作（ファイルロック、パス、インストーラー）を確認できない。
GitHub Actions の Windows ランナーで、push のたびにテストとインストーラー作成を自動で行う。

**状態: 未導入。** 導入するときは、下の YAML を `.github/workflows/windows.yml` として追加する。

## 何が確認できるか

| 項目 | 確認できる | 備考 |
|---|---|---|
| 単体・結合テスト（`npm test`） | ○ | パス区切り・大文字小文字・`EBUSY` などが本物の Windows で動くか |
| GUI の E2E テスト（`npm run test:e2e`） | ○ | Windows ランナーは画面があるので xvfb 不要 |
| インストーラーの作成（`npm run package:win`） | ○ | wine 不要。成果物としてダウンロードできる |
| インストーラーでのインストール | ○ | `/S`（サイレント）で実行して、ファイルとアンインストーラーを確認 |
| OneDrive の同期状態の読み取り | △ | ランナーに OneDrive は無い。「OneDrive 外」と判定されることだけ確認できる |
| 実際の Excel・Copilot・OneDrive | × | `docs/windows-checklist.md` に沿って手で確認する |

## ワークフロー

```yaml
name: windows
on:
  push:
  workflow_dispatch:

jobs:
  test:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: 22
          cache: npm
      - run: npm ci
      - run: npm run typecheck
      - run: npm run lint
      - run: npm test
      - run: npm run test:e2e

  package:
    needs: test
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: 22
          cache: npm
      - run: npm ci
      - run: npm run package:win
      # サイレントインストールして、インストール先とアンインストーラーを確認する
      - name: install
        shell: pwsh
        run: |
          $setup = Get-ChildItem dist -Filter 'xlCode-Setup-*.exe' | Select-Object -First 1
          Start-Process $setup.FullName -ArgumentList '/S', '/currentuser' -Wait
          $dir = Join-Path $env:LOCALAPPDATA 'Programs\xlCode'
          if (-not (Test-Path (Join-Path $dir 'xlCode.exe'))) { throw 'xlCode.exe がインストールされていない' }
          if (-not (Test-Path (Join-Path $dir 'Uninstall xlCode.exe'))) { throw 'アンインストーラーが無い' }
      - name: e2e (installed app)
        shell: pwsh
        run: |
          $env:XLCODE_E2E_EXE = Join-Path $env:LOCALAPPDATA 'Programs\xlCode\xlCode.exe'
          npx vitest run -c vitest.e2e.config.ts
      - uses: actions/upload-artifact@v4
        with:
          name: xlCode-Setup
          path: dist/xlCode-Setup-*.exe
```

## 注意

- 公開リポジトリなら Windows ランナーは無料。非公開リポジトリは無料枠（月 2,000 分、Windows は 2 倍で計算）を消費する。1 回あたり 10〜15 分程度の見込み。
- `release/` の分割 zip は Git の履歴に残り続けるため、更新するたびにリポジトリが約 250MB ずつ大きくなる。
  更新は区切りのよいときだけにするか、GitHub Releases（履歴に残らない）に置くことを検討する。
