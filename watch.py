#!/usr/bin/env python3
"""入力シートの保存を見張って、変わるたびに全員分のPDFを作り直す。

    python watch.py                     # 入力シート.xlsx を監視
    python watch.py 入力シート.xlsx -o 出力

Ctrl+C で終了。
"""

from __future__ import annotations

import argparse
import time
from pathlib import Path

from build_pdf import DEFAULT_INPUT, DEFAULT_OUTDIR, build_all
from rirekisho import DEFAULT_TEMPLATE
from rirekisho.model import InputError


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="入力シートを監視してPDFを自動更新する")
    parser.add_argument("input", nargs="?", default=DEFAULT_INPUT)
    parser.add_argument("-o", "--outdir", default=DEFAULT_OUTDIR)
    parser.add_argument("-t", "--template", default=DEFAULT_TEMPLATE)
    parser.add_argument("--no", type=int, help="この No. の生徒だけ作る")
    parser.add_argument("--name", help="氏名に この文字を含む生徒だけ作る")
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
                    now = time.strftime("%H:%M:%S")
                    try:
                        results = build_all(
                            src, args.outdir, template=args.template, font=args.font,
                            only_no=args.no, only_name=args.name, png=args.png,
                        )
                        for result in results:
                            for w in result.warnings:
                                print(f"  注意: {result.student.name}: {w}")
                        print(f"[{now}] {len(results)}人分を更新しました → {args.outdir}/")
                    except (InputError, ValueError, OSError) as exc:
                        print(f"[{now}] エラー: {exc}")
            time.sleep(args.interval)
        except KeyboardInterrupt:
            print("\n終了しました。")
            return 0


if __name__ == "__main__":
    raise SystemExit(main())
