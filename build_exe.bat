@echo off
chcp 65001 > nul
rem フォルダ容量ビューアの .exe を作るバッチファイル
rem Python が入った PC でダブルクリックして実行してください。
rem
rem 2 種類を作ります。
rem   1. dist\フォルダ容量ビューア.exe
rem        ふだん配る版。.exe 1 つだけで動く。
rem   2. dist\フォルダ版\フォルダ容量ビューア\フォルダ容量ビューア.exe
rem        予備。ウイルス対策ソフトや管理設定で 1 の版が止められたときに試す。
rem        起動のたびに一時フォルダへ展開しないので、止められにくく、起動も速い。
rem        「フォルダ容量ビューア」フォルダごとコピーして使う(中の .exe だけでは動かない)。

cd /d "%~dp0"

echo PyInstaller を準備しています...
python -m pip install --upgrade pyinstaller
if errorlevel 1 (
  echo.
  echo Python が見つかりません。Python をインストールしてから再実行してください。
  pause
  exit /b 1
)

echo .exe(1 ファイル版)を作成しています...
python -m PyInstaller --onefile --windowed --clean --noconfirm --name "フォルダ容量ビューア" folder_size_viewer.py
if errorlevel 1 goto failed

echo .exe(フォルダ版)を作成しています...
python -m PyInstaller --onedir --windowed --clean --noconfirm --name "フォルダ容量ビューア" --distpath "dist\フォルダ版" folder_size_viewer.py
if errorlevel 1 goto failed

rem 作業用に作られたファイルを片付ける(dist の中身だけ残す)
rmdir /s /q build 2> nul
del /q *.spec 2> nul

echo.
echo 完成しました:
echo   dist\フォルダ容量ビューア.exe
echo   dist\フォルダ版\フォルダ容量ビューア\  (フォルダ版)
pause
exit /b 0

:failed
echo.
echo .exe の作成に失敗しました。上のメッセージを確認してください。
pause
exit /b 1
