# xlCode インストーラー（1.5.1）

GitHub の 1 ファイル 100MB 制限のため 2 つに分割している。**Windows 上での動作はまだ確認していない。**

## 使い方

1. このフォルダを取得する（`git clone` か「Code → Download ZIP」）
2. `join.bat` をダブルクリック → `xlCode-Setup-1.5.1.exe` ができる
3. `xlCode-Setup-1.5.1.exe` を実行して、「自分だけ（管理者権限不要）」か「すべてのユーザー」を選ぶ
4. スタートメニュー・デスクトップの xlCode から起動

以前のバージョンを入れている場合は、そのまま上書きでインストールできる。

署名がないため SmartScreen の警告が出る場合は「詳細情報 → 実行」。

## 確認用

作られるソース: コミット 4f5227e（main。サイドバーの階層表示、変更が無ければ確認なしで開く、Web 版で開けない問題の修正。アプリ内ヘルプ（左端の ？ / F1）を追加。UI/UX 改良・VBA モード・取り込み・参照設定を含む）

xlCode-Setup-1.5.1.exe の SHA-256:
`5a808c175332c53bf53907afabae529ce8ab32aa4bdca23da7929e82c92b9bc5`

Windows では `certutil -hashfile xlCode-Setup-1.5.1.exe SHA256` で確認できる。

確認してほしい項目は [docs/windows-checklist.md](../../docs/windows-checklist.md)（特に 6・1-4・7・10・11・12）。
