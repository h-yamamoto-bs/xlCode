# xlCode インストーラー（1.8.0）

GitHub の 1 ファイル 100MB 制限のため分割している。**Windows 上での動作はまだ確認していない**（確認項目は [docs/windows-checklist.html](../../docs/windows-checklist.html)）。

## 使い方

1. このフォルダを取得する（`git clone` か「Code → Download ZIP」）
2. `join.bat` をダブルクリック → `xlCode-Setup-1.8.0.exe` ができる
3. `xlCode-Setup-1.8.0.exe` を実行して、「自分だけ（管理者権限不要）」か「すべてのユーザー」を選ぶ
4. スタートメニュー・デスクトップの xlCode から起動

以前のバージョンを入れている場合は、そのまま上書きでインストールできる。

署名がないため SmartScreen の警告が出る場合は「詳細情報 → 実行」。

## 確認用

作られるソース: コミット 85c6a4d（main。1.8.0: ブックのテーマ（ダーク / ライト）と xlCode の表示モード（ダーク / ライト / OS）、閲覧タブのファイル内検索（Ctrl+F）、再読み込み（F5）で OneDrive から最新を取り直す。1.7.x の左右のサイドバー、コードの閲覧タブ、指示の作成欄とメモ、1.6.0 のタブ・ショートカット・コマンドパレット・.xlcodeignore・置き場所の変更、VBA モード・取り込み・アプリ内ヘルプを含む）

xlCode-Setup-1.8.0.exe の SHA-256:
`c4c99747be8fef73be675c7776ffcbbef439d5f8cc0cffc8df5648df4e2fab54`

Windows では `certutil -hashfile xlCode-Setup-1.8.0.exe SHA256` で確認できる。

確認してほしい項目は [docs/windows-checklist.html](../../docs/windows-checklist.html)（ブックのテーマ H-1〜H-12、OneDrive から最新を取り直す D-9〜D-14、インストーラーを優先）。
