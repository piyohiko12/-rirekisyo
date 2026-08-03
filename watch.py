#!/usr/bin/env python3
"""入力シートの保存を見張って、変わるたびにPDFを作り直す。

    python watch.py                     # 入力シート.xlsx を監視
    python watch.py 入力シート.xlsx -o 出力/履歴書.pdf

Ctrl+C で終了。
"""

from __future__ import annotations

import argparse
import time
from pathlib import Path

from build_pdf import DEFAULT_INPUT, DEFAULT_OUTPUT, build
from rirekisho import DEFAULT_TEMPLATE
from rirekisho.model import InputError


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="入力シートを監視してPDFを自動更新する")
    parser.add_argument("input", nargs="?", default=DEFAULT_INPUT)
    parser.add_argument("-o", "--output", default=DEFAULT_OUTPUT)
    parser.add_argument("-t", "--template", default=DEFAULT_TEMPLATE)
    parser.add_argument("--font", help="日本語フォントのパス")
    parser.add_argument("--png", action="store_true", help="確認用のPNGも書き出す")
    parser.add_argument("--interval", type=float, default=1.0, help="確認の間隔（秒）")
    args = parser.parse_args(argv)

    src = Path(args.input)
    print(f"監視中: {src} （Ctrl+C で終了）")
    last: float | None = None
    while True:
        try:
            if src.exists():
                stamp = src.stat().st_mtime
                if stamp != last:
                    last = stamp
                    time.sleep(0.3)  # 保存の途中で読まないよう少し待つ
                    try:
                        out, warnings = build(
                            src, args.output, template=args.template,
                            font=args.font, png=args.png,
                        )
                        for w in warnings:
                            print(f"  注意: {w}")
                        print(f"[{time.strftime('%H:%M:%S')}] 更新しました → {out}")
                    except (InputError, ValueError, OSError) as exc:
                        print(f"[{time.strftime('%H:%M:%S')}] エラー: {exc}")
            time.sleep(args.interval)
        except KeyboardInterrupt:
            print("\n終了しました。")
            return 0


if __name__ == "__main__":
    raise SystemExit(main())
