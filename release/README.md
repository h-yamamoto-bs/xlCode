# release

| 場所 | 中身 |
|---|---|
| `installer/` | Windows 用インストーラー（1.7.0。コミットは `installer/README.md` を参照） |
| このフォルダの `xlCode-src-win.zip.part*` | ソース一式＋Windows 用 node_modules（1.6.0、コミット aadc66a。依存パッケージは 1.7.0 と同じなので、`node_modules` はそのまま使える） |

どちらも GitHub の 1 ファイル 100MB 制限に収まるよう分割している。

## xlCode-src-win.zip

npm install できない環境向けの一式。

- ソース一式（コミット aadc66a・main、1.6.0。`release/` は除く）
- **Windows x64 用**の `node_modules`（Electron 本体、esbuild・Rollup・Tailwind・lightningcss の win32-x64 版、iconv-lite を含む）
- `node_modules\.bin` の Windows 用の起動ファイル（`.cmd` / `.ps1`。npm が Windows で作るものと同じ cmd-shim で生成）。`npm run dev` などがそのまま動く
- `package.json` / `package-lock.json`
- **Git の履歴（`.git`）は含まない。** 履歴はこのリポジトリから取得する

Linux / macOS では `node_modules` の中身が合わないため、そのままでは動かない。

### 取り出し方

1. このフォルダを取得する（`git clone` か、GitHub の「Code → Download ZIP」）
2. `join.bat` をダブルクリック（Linux / macOS は `sh join.sh`）→ `xlCode-src-win.zip` ができる
3. 展開して、`xlCode` フォルダで `npm run dev`（Node.js 22 以上が必要。npm install は不要）

### 確認用

xlCode-src-win.zip の SHA-256:
`51b3819a0400b984634c940033505efb0777558065f78b7e951882324a180473`

Windows では `certutil -hashfile xlCode-src-win.zip SHA256` で確認できる。
