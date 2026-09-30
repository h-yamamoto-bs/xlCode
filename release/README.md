# release

| 場所 | 中身 |
|---|---|
| `installer/` | Windows 用インストーラー（1.1.0 preview 2、コミット d439f73） |
| このフォルダの `xlCode-src-win.zip.part*` | ソース一式＋Windows 用 node_modules（コミット d439f73） |

どちらも GitHub の 1 ファイル 100MB 制限に収まるよう分割している。

## xlCode-src-win.zip

npm install できない環境向けの一式。

- ソース一式（コミット d439f73・vba-mode ブランチ。`release/` は除く）
- **Windows x64 用**の `node_modules`（Electron 本体、esbuild・Rollup・Tailwind・lightningcss の win32-x64 版、iconv-lite を含む）
- `package.json` / `package-lock.json`
- **Git の履歴（`.git`）は含まない。** 履歴はこのリポジトリから取得する

Linux / macOS では `node_modules` の中身が合わないため、そのままでは動かない。

### 取り出し方

1. このフォルダを取得する（`git clone` か、GitHub の「Code → Download ZIP」）
2. `join.bat` をダブルクリック（Linux / macOS は `sh join.sh`）→ `xlCode-src-win.zip` ができる
3. 展開して、`xlCode` フォルダで `npm run dev`（Node.js 22 以上が必要。npm install は不要）

### 確認用

xlCode-src-win.zip の SHA-256:
`3763a154f3a43fe4fbfc8e7884928d6567d2a5a3f765fd0fc12c8cb6efb4e003`

Windows では `certutil -hashfile xlCode-src-win.zip SHA256` で確認できる。
