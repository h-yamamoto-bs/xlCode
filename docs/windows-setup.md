# Windows での導入手順（clone 後）

## 0. 前提

- **Git for Windows**（clone と、xlCode の自動コミットに使う）
- 分割ファイルが合計約 360MB あるため、clone に時間がかかる
- clone 先は `C:\dev\xlCode` のような**短いパス**で、**OneDrive の外**にする
  - `node_modules` はパスが深く、260 文字を超えると展開に失敗することがある
  - `.git` を OneDrive に置くと壊れやすい

```bat
cd C:\dev
git clone https://github.com/h-yamamoto-bs/xlCode.git
cd xlCode
```

## A. アプリとして使う（インストーラー。Node.js 不要）

1. `release\installer\join.bat` をダブルクリック → `xlCode-Setup-1.5.0.exe` ができる
2. （任意）ハッシュを確認する

   ```bat
   certutil -hashfile release\installer\xlCode-Setup-1.5.0.exe SHA256
   ```

   `1e340281c39c93b3bc675a7264af373b2ef2d81c42a5ff1ed6bdbf04a3900362` と一致すれば OK
3. `xlCode-Setup-1.5.0.exe` を実行し、「自分だけ（管理者権限不要）」か「すべてのユーザー」を選ぶ
   - SmartScreen の警告が出たら「詳細情報 → 実行」（署名がないため）
   - 会社の AppLocker などで起動できない場合は「すべてのユーザー」（Program Files）で試す
4. スタートメニュー・デスクトップの xlCode から起動する

## B. ソース版で動かす（npm install 不要。Node.js 22 以上が必要）

1. `node -v` で 22 以上か確認する
2. `release\join.bat` をダブルクリック → `release\xlCode-src-win.zip` ができる
3. （任意）ハッシュを確認する

   ```bat
   certutil -hashfile release\xlCode-src-win.zip SHA256
   ```

   `1e851a309af46b8564cd45935e3e9802b56e908a036f54dcf3f8d1f2e9d9cdb4` と一致すれば OK
4. 展開する。エクスプローラーの展開は遅く、長いパスで失敗しやすいので `tar` を使う

   ```bat
   mkdir C:\dev\xlCode-src
   tar -xf release\xlCode-src-win.zip -C C:\dev\xlCode-src
   ```

   `C:\dev\xlCode-src\xlCode\` にソースと Windows 用 `node_modules` が出る
5. 起動する

   ```bat
   cd C:\dev\xlCode-src\xlCode
   npm run dev
   ```

### Git の履歴付きで開発したい場合

zip には `.git` が入っていない。展開した `node_modules` フォルダだけを clone した `C:\dev\xlCode\` 直下へ移せば、clone 側でそのまま `npm run dev` できる（`node_modules` は `.gitignore` 済み）。
zip はコミット 56dacb3 時点のもの。clone 側の `package.json` の依存が変わっていれば作り直しが必要。

### よく使うコマンド

```bat
npm run dev          :: 開発起動
npm test             :: テスト
npm run test:e2e     :: GUI の E2E テスト
npm run package:win  :: インストーラーを作る（初回に NSIS などをダウンロードする。オフラインでは不可）
```

## C. 動作確認

[windows-checklist.md](windows-checklist.md) の項目を確認し、「結果」欄に書く。優先度の高いもの:

- **6** インストーラー
- **1-4** OneDrive 内のブックをデスクトップ版 Excel で開いたときに検知できるか
- **7** OneDrive の同期状態（ステータスバーの「OneDrive: 〜」をクリックして出る `attrs` / `sts` / `sps` の値を記録）
- **10・11・12**

1.5.0 のインストーラーは Windows 上ではまだ動作を確認していない。
