@echo off
rem 分割したインストーラーを結合して xlCode-Setup-1.3.1.exe を作る
cd /d "%~dp0"
copy /b xlCode-Setup-1.3.1.exe.part00+xlCode-Setup-1.3.1.exe.part01 "xlCode-Setup-1.3.1.exe"
if errorlevel 1 (
  echo 結合に失敗しました。パートがすべて揃っているか確認してください。
) else (
  echo xlCode-Setup-1.3.1.exe を作成しました。ダブルクリックでインストールできます。
)
pause
