# xlCode インストーラー（1.1.0 preview 2）

GitHub の 1 ファイル 100MB 制限のため 2 つに分割している。**Windows 上での動作はまだ確認していない。**

## 使い方

1. このフォルダを取得する（`git clone` か「Code → Download ZIP」）
2. `join.bat` をダブルクリック → `xlCode-Setup-1.1.0.exe` ができる
3. `xlCode-Setup-1.1.0.exe` を実行して、「自分だけ（管理者権限不要）」か「すべてのユーザー」を選ぶ
4. スタートメニュー・デスクトップの xlCode から起動

1.0.0 を入れている場合は、そのまま上書きでインストールできる。

署名がないため SmartScreen の警告が出る場合は「詳細情報 → 実行」。

## 確認用

作られるソース: コミット d439f73（vba-mode ブランチ。VBA モード・UI 用シートの保護を含む）

xlCode-Setup-1.1.0.exe の SHA-256:
`3437535ad736b20a397976cb00565b11da92d71f42d5b95c204f8a63978e3dbb`

Windows では `certutil -hashfile xlCode-Setup-1.1.0.exe SHA256` で確認できる。

確認してほしい項目は [docs/windows-checklist.md](../../docs/windows-checklist.md)（特に 6・1-4・7・10・11）。
