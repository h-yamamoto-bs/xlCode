@echo off
rem 分割した zip を結合して xlCode-src-win.zip を作る
cd /d "%~dp0"
copy /b xlCode-src-win.zip.part00+xlCode-src-win.zip.part01+xlCode-src-win.zip.part02 "xlCode-src-win.zip"
if errorlevel 1 (
  echo 結合に失敗しました。パートがすべて揃っているか確認してください。
) else (
  echo xlCode-src-win.zip を作成しました。
)
pause
