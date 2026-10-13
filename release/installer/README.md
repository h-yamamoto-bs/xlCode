# xlCode インストーラー（1.6.0）

GitHub の 1 ファイル 100MB 制限のため分割している。**Windows 上での動作はまだ確認していない**（確認項目は [docs/windows-checklist.html](../../docs/windows-checklist.html)）。

## 使い方

1. このフォルダを取得する（`git clone` か「Code → Download ZIP」）
2. `join.bat` をダブルクリック → `xlCode-Setup-1.6.0.exe` ができる
3. `xlCode-Setup-1.6.0.exe` を実行して、「自分だけ（管理者権限不要）」か「すべてのユーザー」を選ぶ
4. スタートメニュー・デスクトップの xlCode から起動

以前のバージョンを入れている場合は、そのまま上書きでインストールできる。

署名がないため SmartScreen の警告が出る場合は「詳細情報 → 実行」。

## 確認用

作られるソース: コミット f6a3d59（main。1.6.0: タブ、VS Code 準拠のショートカットとコマンドパレット、エクスプローラーの階層にルールと .xlcodeignore、ブックにしないファイル（.xlcodeignore）とブック作成前の確認、ブックの置き場所の変更の改良。VBA モード・取り込み・アプリ内ヘルプを含む）

xlCode-Setup-1.6.0.exe の SHA-256:
`04f155b470975ffe45f72fc5a07827c1c3e9b9d0d4fa5d2f15bc66d54bfe9050`

Windows では `certutil -hashfile xlCode-Setup-1.6.0.exe SHA256` で確認できる。

確認してほしい項目は [docs/windows-checklist.html](../../docs/windows-checklist.html)（インストーラー・メニュー削除・ショートカット・IME・置き場所・OneDrive を優先）。
