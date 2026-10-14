# xlCode インストーラー（1.7.1）

GitHub の 1 ファイル 100MB 制限のため分割している。**Windows 上での動作はまだ確認していない**（確認項目は [docs/windows-checklist.html](../../docs/windows-checklist.html)）。

## 使い方

1. このフォルダを取得する（`git clone` か「Code → Download ZIP」）
2. `join.bat` をダブルクリック → `xlCode-Setup-1.7.1.exe` ができる
3. `xlCode-Setup-1.7.1.exe` を実行して、「自分だけ（管理者権限不要）」か「すべてのユーザー」を選ぶ
4. スタートメニュー・デスクトップの xlCode から起動

以前のバージョンを入れている場合は、そのまま上書きでインストールできる。

署名がないため SmartScreen の警告が出る場合は「詳細情報 → 実行」。

## 確認用

作られるソース: コミット e405090（main。1.7.1: ドキュメントを開いただけでは保存を聞かない・CRLF を保つ・サイドバーの階層が閉じない修正。1.7.0 の左右のサイドバー（開閉・幅、OneDrive の同期状態、Git）、コードの閲覧タブ（既定のアプリ / VS Code で開く）、指示の作成欄とメモ、1.6.0 のタブ・ショートカット・コマンドパレット・.xlcodeignore・置き場所の変更、VBA モード・取り込み・アプリ内ヘルプを含む）

xlCode-Setup-1.7.1.exe の SHA-256:
`67fc7800a2573e8eefbf9dea039369970a7ec3807716115fd81ce921b7ad5457`

Windows では `certutil -hashfile xlCode-Setup-1.7.1.exe SHA256` で確認できる。

確認してほしい項目は [docs/windows-checklist.html](../../docs/windows-checklist.html)（インストーラー・メニュー削除・ショートカット・IME・置き場所・OneDrive を優先）。
