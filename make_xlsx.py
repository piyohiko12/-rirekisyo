#!/usr/bin/env python3
"""入力用のExcelファイル（入力シート.xlsx・1行＝1生徒の名列順）を作る。

    python make_xlsx.py                 # 40人分の空シートを作る
    python make_xlsx.py --with-sample   # 記入例（3人分）入りで作る
    python make_xlsx.py --rows 45       # 行数を変える
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from rirekisho.inputs import ROSTER_ROWS, write_template

SAMPLE_STUDENTS = [
    {
        "meta.no": 1,
        "meta.class": "3年2組",
        "meta.number": 1,
        "name": "佐野 太郎",
        "name_kana": "さの たろう",
        "birth": "2008-05-12",
        "zip": "5980001",
        "address": "大阪府泉佐野市市場東1丁目2-3 サンプルハイツ101",
        "address_kana": "おおさかふ いずみさのし いちばひがし",
        "license.1.name": "第二種電気工事士",
        "license.1.ym": "2024-06",
        "license.2.name": "危険物取扱者乙種第4類",
        "license.2.ym": "令和7年3月",
        "license.3.name": "技能検定3級 機械加工（普通旋盤作業）",
        "license.3.ym": "2025-11",
        "activities": (
            "・1年〜3年 機械研究部に所属（3年次は副部長）\n"
            "・2年 ものづくりコンテスト旋盤作業部門 校内代表\n"
            "・3年 体育大会 応援団リーダー"
        ),
        "desired_job": "機械加工・生産技術",
        "appeal": (
            "旋盤とフライス盤の実習で、寸法公差を守る段取りと確認を身につけました。"
            "分からないことはその場で質問し、最後まで丁寧にやり切ることを大切にしています。"
        ),
        "motivation": (
            "実習で金属を削って形にする面白さを知り、ものづくりの現場で働きたいと考えるようになりました。"
            "貴社は多品種少量の精密加工に強みがあり、若いうちから幅広い機械に触れられると伺っています。"
            "学校で学んだ機械加工の基礎を早く現場で活かし、確かな技術を身につけて長く貢献したいと思い志望しました。"
        ),
        "remarks": "普通自動車免許は卒業後に取得予定です。",
    },
    {
        "meta.no": 2,
        "meta.class": "3年2組",
        "meta.number": 2,
        "name": "近畿 花子",
        "name_kana": "きんき はなこ",
        "birth": "平成20年11月3日",
        "zip": "5900001",
        "address": "大阪府堺市堺区南瓦町3-1 テストレジデンス2201号室",
        "address_kana": "おおさかふ さかいし さかいく みなみかわらまち",
        "contact_zip": "5980048",
        "contact_address": "大阪府泉佐野市りんくう往来北1-1 保護者 近畿一郎方",
        "contact_kana": "おおさかふ いずみさのし りんくうおうらいきた",
        "license.1.name": "計算技術検定3級",
        "license.1.ym": "2024-11",
        "license.2.name": "基礎製図検定",
        "license.2.ym": "2024-05",
        "activities": "・生徒会 会計（2年）\n・吹奏楽部 パートリーダー（3年）\n・海岸清掃活動に参加（1年〜3年）",
        "desired_job": "生産管理・事務",
        "appeal": "3年間、無遅刻無欠席で続けてきました。人と話すことが好きで、部活動では初心者の後輩の指導を担当しました。",
        "motivation": (
            "工場見学で、たくさんの部品が計画どおりに動いて製品になる仕組みに興味を持ちました。"
            "簿記と情報の授業で学んだことを活かして、現場を支える仕事に就きたいと考え志望しました。"
        ),
    },
    {
        "meta.no": 3,
        "meta.class": "3年2組",
        "meta.number": 3,
        "name": "泉州 一郎",
        "name_kana": "せんしゅう いちろう",
        "birth": "2008-01-20",
        "zip": "5960076",
        "address": "大阪府岸和田市野田町1-1-1",
        "address_kana": "おおさかふ きしわだし のだちょう",
        "license.1.name": "小型フォークリフト運転特別教育修了",
        "license.1.ym": "2025-08",
        "activities": "・ソフトテニス部 主将（3年）\n・泉州オープンファクトリー参加（2年）",
        "desired_job": "機械オペレーター",
        "appeal": "部活動で主将を務め、練習メニューを考えて部員をまとめました。体力に自信があります。",
        "motivation": "インターンシップで貴社の現場を見学し、機械を扱う正確さに憧れました。基礎から学んで一人前になりたいと思い志望しました。",
    },
]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="履歴書の入力シート(.xlsx)を作成する")
    parser.add_argument("-o", "--output", default="入力シート.xlsx", help="出力先（既定: 入力シート.xlsx）")
    parser.add_argument("--with-sample", action="store_true", help="記入例（3人分）を入れて作る")
    parser.add_argument("--rows", type=int, default=ROSTER_ROWS, help=f"名簿の行数（既定: {ROSTER_ROWS}）")
    parser.add_argument("--force", action="store_true", help="既存ファイルを上書きする")
    args = parser.parse_args(argv)

    out = Path(args.output)
    if out.exists() and not args.force:
        print(f"エラー: {out} はすでにあります。上書きするなら --force を付けてください。", file=sys.stderr)
        return 1

    write_template(out, students=SAMPLE_STUDENTS if args.with_sample else None, rows=args.rows)
    print(f"作成しました: {out}")
    print("「入力」シートに名列順で入力して保存したあと、python build_pdf.py を実行してください。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
