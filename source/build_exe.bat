@echo off
chcp 65001 > nul
rem ================================================================
rem  【開発者用】フォルダ容量ビューアの .exe を作るバッチファイル
rem  アプリを使う先生は、これを開く必要はありません。
rem  (アプリは、ひとつ上のフォルダにある「フォルダ容量ビューア.exe」です)
rem ================================================================
rem Python が入った PC でダブルクリックして実行してください。
rem
rem 2 種類を作ります(どちらも、このバッチがある source フォルダのひとつ上にできる)。
rem   1. フォルダ容量ビューア.exe
rem        ふだん使う版。.exe 1 つだけで動く。
rem   2. spare-version\フォルダ容量ビューア\フォルダ容量ビューア.exe
rem        予備。ウイルス対策ソフトや管理設定で 1 の版が止められたときに試す。
rem        起動のたびに一時フォルダへ展開しないので、止められにくく、起動も速い。
rem        「フォルダ容量ビューア」フォルダごとコピーして使う(中の .exe だけでは動かない)。
rem
rem 開発用のファイルを source フォルダに分けているのは、先生が間違えて
rem このバッチを開かないようにするため(校務用PCでは黒い画面が一瞬出て消える)。

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
python -m PyInstaller --onefile --windowed --clean --noconfirm --name "フォルダ容量ビューア" --distpath .. --workpath build --specpath build folder_size_viewer.py
if errorlevel 1 goto failed

echo .exe(予備のフォルダ版)を作成しています...
python -m PyInstaller --onedir --windowed --clean --noconfirm --name "フォルダ容量ビューア" --distpath ..\spare-version --workpath build --specpath build folder_size_viewer.py
if errorlevel 1 goto failed

rem 作業用に作られたファイルを片付ける
rmdir /s /q build 2> nul
rmdir /s /q __pycache__ 2> nul

echo.
echo 完成しました(ひとつ上のフォルダ):
echo   フォルダ容量ビューア.exe
echo   spare-version\フォルダ容量ビューア\  (予備のフォルダ版)
pause
exit /b 0

:failed
echo.
echo .exe の作成に失敗しました。上のメッセージを確認してください。
pause
exit /b 1
