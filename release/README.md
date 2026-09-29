# release

npm install できない環境向けの一式を、GitHub の 1 ファイル 100MB 制限に収まるよう分割した zip。

## 中身（xlCode-src-win.zip）

- ソース一式と Git 履歴（コミット ff685d1 時点）
- **Windows x64 用**の `node_modules`（Electron 本体、esbuild・Rollup・Tailwind・lightningcss の win32-x64 版を含む）
- `package.json` / `package-lock.json`

Linux / macOS では `node_modules` の中身が合わないため、そのままでは動かない。

## 取り出し方

1. このフォルダを取得する（`git clone` か、GitHub の「Code → Download ZIP」）
2. `join.bat` をダブルクリック（Linux / macOS は `sh join.sh`）→ `xlCode-src-win.zip` ができる
3. 展開して、`xlCode` フォルダで `npm run dev`（Node.js 22 以上が必要。npm install は不要）

## 確認用

xlCode-src-win.zip の SHA-256:
`904790f452f4f9b8bd14d158e91130bb174f0c651b353bbbdd4f406a0f1b2571`

Windows では `certutil -hashfile xlCode-src-win.zip SHA256` で確認できる。
