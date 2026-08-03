#!/usr/bin/env python3
"""入力用のExcelファイル（入力シート.xlsx）を作る。

    python make_xlsx.py                 # 入力シート.xlsx を作る
    python make_xlsx.py --with-sample   # 記入例入りで作る
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from rirekisho.inputs import write_template

SAMPLE = {
    "name_kana": "さの たろう",
    "name": "佐野 太郎",
    "birth": "2008-05-12",
    "as_of": "2026-09-01",
    "zip": "5980001",
    "address_kana": "おおさかふ いずみさのし いちばひがし",
    "address": "大阪府泉佐野市市場東1丁目2-3 サンプルハイツ101",
    "license.1.ym": "2024-06",
    "license.1.name": "第二種電気工事士",
    "license.2.ym": "令和7年3月",
    "license.2.name": "危険物取扱者乙種第4類",
    "license.3.ym": "2025-11",
    "license.3.name": "技能検定3級 機械加工（普通旋盤作業）",
    "activities": (
        "・1年〜3年 機械研究部に所属（3年次は副部長）\n"
        "・2年 ものづくりコンテスト旋盤作業部門 校内代表\n"
        "・3年 体育大会 応援団リーダー"
    ),
    "desired_job": "機械加工・生産技術",
    "appeal": "旋盤とフライス盤の実習で、寸法公差を守る段取りと確認を身につけました。分からないことはその場で質問し、最後まで丁寧にやり切ることを大切にしています。",
    "motivation": (
        "実習で金属を削って形にする面白さを知り、ものづくりの現場で働きたいと考えるようになりました。"
        "貴社は多品種少量の精密加工に強みがあり、若いうちから幅広い機械に触れられると伺っています。"
        "学校で学んだ機械加工の基礎を早く現場で活かし、確かな技術を身につけて長く貢献したいと思い志望しました。"
    ),
    "remarks": "普通自動車免許は卒業後に取得予定です。",
}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="履歴書の入力シート(.xlsx)を作成する")
    parser.add_argument("-o", "--output", default="入力シート.xlsx", help="出力先（既定: 入力シート.xlsx）")
    parser.add_argument("--with-sample", action="store_true", help="記入例を入れた状態で作る")
    parser.add_argument("--force", action="store_true", help="既存ファイルを上書きする")
    args = parser.parse_args(argv)

    out = Path(args.output)
    if out.exists() and not args.force:
        print(f"エラー: {out} はすでにあります。上書きするなら --force を付けてください。", file=sys.stderr)
        return 1

    write_template(out, defaults=SAMPLE if args.with_sample else None)
    print(f"作成しました: {out}")
    print("「入力」シートの黄色いセルに入力して保存したあと、python build_pdf.py を実行してください。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
