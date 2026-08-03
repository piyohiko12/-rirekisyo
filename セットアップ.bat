@echo off
rem 必要な部品（PyMuPDF・openpyxl）を入れる。最初に1回だけ実行すればOK。
chcp 65001 >nul
cd /d "%~dp0"

set PY=
where py >nul 2>nul && set PY=py -3
if "%PY%"=="" where python >nul 2>nul && set PY=python
if "%PY%"=="" (
  echo Python が見つかりません。https://www.python.org/ からインストールしてください。
  pause
  exit /b 1
)

echo 必要な部品を入れています。しばらくお待ちください...
%PY% -m pip install -r requirements.txt
echo.
echo 終わりました。「履歴書PDF作成.bat」をダブルクリックして起動してください。
pause
