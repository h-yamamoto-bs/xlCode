# xlCode インストーラー（1.5.2）

GitHub の 1 ファイル 100MB 制限のため 2 つに分割している。**Windows 上での動作はまだ確認していない。**

## 使い方

1. このフォルダを取得する（`git clone` か「Code → Download ZIP」）
2. `join.bat` をダブルクリック → `xlCode-Setup-1.5.2.exe` ができる
3. `xlCode-Setup-1.5.2.exe` を実行して、「自分だけ（管理者権限不要）」か「すべてのユーザー」を選ぶ
4. スタートメニュー・デスクトップの xlCode から起動

以前のバージョンを入れている場合は、そのまま上書きでインストールできる。

署名がないため SmartScreen の警告が出る場合は「詳細情報 → 実行」。

## 確認用

作られるソース: コミット dacc9f2（main。個人用 OneDrive のブックを Web 版で開くと 404 になる問題の修正。サイドバーの階層表示、変更が無ければ確認なしで開く、Web 版で開けない問題の修正。アプリ内ヘルプ（左端の ？ / F1）を追加。UI/UX 改良・VBA モード・取り込み・参照設定を含む）

xlCode-Setup-1.5.2.exe の SHA-256:
`894028068bad58ea30416721c4da60a42affc5acec626dbe0cfa9f3fc39932c8`

Windows では `certutil -hashfile xlCode-Setup-1.5.2.exe SHA256` で確認できる。

確認してほしい項目は [docs/windows-checklist.md](../../docs/windows-checklist.md)（特に 6・1-4・7・10・11・12）。
