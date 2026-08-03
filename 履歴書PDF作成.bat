@echo off
rem 履歴書PDF作成ツールを起動する（このファイルをダブルクリック）
chcp 65001 >nul
cd /d "%~dp0"

set PY=
where py >nul 2>nul && set PY=py -3
if "%PY%"=="" where python >nul 2>nul && set PY=python
if "%PY%"=="" (
  echo Python が見つかりません。
  echo https://www.python.org/ から Python をインストールしてから、もう一度実行してください。
  echo ※インストール時に「Add Python to PATH」にチェックを入れてください。
  pause
  exit /b 1
)

%PY% gui.py
if errorlevel 1 (
  echo.
  echo 起動できませんでした。上のメッセージを確認してください。
  echo 必要な部品が入っていない場合は、次のコマンドで入れられます:
  echo     %PY% -m pip install -r requirements.txt
  pause
)
