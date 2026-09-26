@echo off
chcp 65001 > nul
rem フォルダ容量ビューアの .exe を作るバッチファイル
rem Python が入った PC でダブルクリックして実行してください。

cd /d "%~dp0"

echo PyInstaller を準備しています...
python -m pip install --upgrade pyinstaller
if errorlevel 1 (
  echo.
  echo Python が見つかりません。Python をインストールしてから再実行してください。
  pause
  exit /b 1
)

echo .exe を作成しています...
python -m PyInstaller --onefile --windowed --clean --name "フォルダ容量ビューア" folder_size_viewer.py
if errorlevel 1 (
  echo.
  echo .exe の作成に失敗しました。上のメッセージを確認してください。
  pause
  exit /b 1
)

echo.
echo 完成しました: dist\フォルダ容量ビューア.exe
pause
