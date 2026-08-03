#!/bin/sh
# 履歴書PDF作成ツールを起動する（このファイルをダブルクリック）
cd "$(dirname "$0")" || exit 1

if command -v python3 >/dev/null 2>&1; then
  python3 gui.py && exit 0
else
  echo "Python 3 が見つかりません。https://www.python.org/ からインストールしてください。"
  printf "Enterキーで閉じます: "
  read -r _
  exit 1
fi

echo
echo "起動できませんでした。上のメッセージを確認してください。"
echo "必要な部品が入っていない場合は、次のコマンドで入れられます:"
echo "    python3 -m pip install -r requirements.txt"
printf "Enterキーで閉じます: "
read -r _
