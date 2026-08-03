#!/usr/bin/env python3
"""入力シート（.xlsx / .csv）から履歴書PDFを作る。

    python build_pdf.py                          # 入力シート.xlsx → 出力/履歴書.pdf
    python build_pdf.py 入力シート.xlsx -o 出力/履歴書.pdf
    python build_pdf.py シート.csv --png         # 確認用のPNGも出力する
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from rirekisho import DEFAULT_TEMPLATE
from rirekisho.inputs import read_values
from rirekisho.model import InputError, build_resume
from rirekisho.render import render_preview_png, render_resume

DEFAULT_INPUT = "入力シート.xlsx"
DEFAULT_OUTPUT = "出力/履歴書.pdf"


def build(
    input_path: str | Path = DEFAULT_INPUT,
    output_path: str | Path = DEFAULT_OUTPUT,
    *,
    template: str | Path = DEFAULT_TEMPLATE,
    font: str | None = None,
    png: bool = False,
) -> tuple[Path, list[str]]:
    """入力ファイルを読んでPDFを書き出し、(出力パス, 注意メッセージ) を返す。"""
    values = read_values(input_path)
    resume = build_resume(values)
    out = render_resume(resume, template=template, out_path=output_path, font_path=font)
    if png:
        render_preview_png(out, out.with_suffix(".png"))
    return out, resume.warnings


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="入力シートから履歴書PDFを作成する")
    parser.add_argument("input", nargs="?", default=DEFAULT_INPUT, help="入力シート(.xlsx/.csv)")
    parser.add_argument("-o", "--output", default=DEFAULT_OUTPUT, help="出力PDF")
    parser.add_argument("-t", "--template", default=DEFAULT_TEMPLATE, help="用紙のテンプレートPDF")
    parser.add_argument("--font", help="日本語フォントのパス（省略時は自動検出）")
    parser.add_argument("--png", action="store_true", help="確認用のPNGも書き出す")
    args = parser.parse_args(argv)

    try:
        out, warnings = build(
            args.input, args.output, template=args.template, font=args.font, png=args.png
        )
    except (InputError, FileNotFoundError, ValueError) as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        return 1

    for w in warnings:
        print(f"注意: {w}")
    print(f"作成しました: {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
