# release

| 場所 | 中身 |
|---|---|
| `installer/` | Windows 用インストーラー（1.5.1、コミット 4f5227e） |
| このフォルダの `xlCode-src-win.zip.part*` | ソース一式＋Windows 用 node_modules（コミット 4f5227e） |

どちらも GitHub の 1 ファイル 100MB 制限に収まるよう分割している。

## xlCode-src-win.zip

npm install できない環境向けの一式。

- ソース一式（コミット 4f5227e・main。`release/` は除く）
- **Windows x64 用**の `node_modules`（Electron 本体、esbuild・Rollup・Tailwind・lightningcss の win32-x64 版、iconv-lite を含む）
- `node_modules\.bin` の Windows 用の起動ファイル（`.cmd` / `.ps1`）。`npm run dev` などがそのまま動く
- `package.json` / `package-lock.json`
- **Git の履歴（`.git`）は含まない。** 履歴はこのリポジトリから取得する

Linux / macOS では `node_modules` の中身が合わないため、そのままでは動かない。

### 取り出し方

1. このフォルダを取得する（`git clone` か、GitHub の「Code → Download ZIP」）
2. `join.bat` をダブルクリック（Linux / macOS は `sh join.sh`）→ `xlCode-src-win.zip` ができる
3. 展開して、`xlCode` フォルダで `npm run dev`（Node.js 22 以上が必要。npm install は不要）

### 確認用

xlCode-src-win.zip の SHA-256:
`8b5820c5883ec6e61100d51b3ab9d0ed2fe66c201bd9e7339d5b060e8f4d6740`

Windows では `certutil -hashfile xlCode-src-win.zip SHA256` で確認できる。
