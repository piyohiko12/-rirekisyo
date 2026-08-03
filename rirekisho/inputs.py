"""入力シート（Excel / Googleスプレッドシート）の定義・生成・読み取り。

シートの構造（列）:
    A: 項目名        B: 入力欄        C: 入力欄2（資格・職歴の表のみ）
    D: 記入例・説明  E: Bのキー(非表示)  F: Cのキー(非表示)

E・F列のキーだけを頼りに読み取るので、行を増減しても、
Googleスプレッドシートで作って .xlsx や .csv で書き出しても同じように動く。
"""

from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path

KEY_COL_B = 5  # E列（1始まり）
KEY_COL_C = 6  # F列

MAX_LICENSE_ROWS = 8
MAX_JOB_ROWS = 4


@dataclass
class Field:
    label: str
    key: str
    help: str = ""
    height: float | None = None   # 行の高さ（複数行入力欄用）
    key_c: str = ""               # C列も入力欄にする場合のキー
    label_c: str = ""             # C列の見出し（表形式のとき）


@dataclass
class Section:
    title: str
    fields: list[Field]
    note: str = ""


def sheet_definition() -> list[Section]:
    """入力シートに並べる項目の定義。"""
    licenses = [
        Field(
            label=f"資格 {i}",
            key=f"license.{i}.ym",
            key_c=f"license.{i}.name",
            help="取得年月 → 左、名称 → 右" if i == 1 else "",
        )
        for i in range(1, MAX_LICENSE_ROWS + 1)
    ]
    jobs = [
        Field(
            label=f"職歴 {i}",
            key=f"job.{i}.ym",
            key_c=f"job.{i}.text",
            help="年月 → 左、内容 → 右（新卒で職歴なしの場合は空のまま）" if i == 1 else "",
        )
        for i in range(1, MAX_JOB_ROWS + 1)
    ]
    return [
        Section(
            "基本情報",
            [
                Field("ふりがな（名前）", "name_kana", "例: さの たろう"),
                Field("名前", "name", "例: 佐野 太郎"),
                Field("生年月日", "birth", "例: 2008-05-12 / 平成20年5月12日（元号と満年齢は自動）"),
                Field("基準日（満○歳の計算日）", "as_of", "用紙に印字された「令和8年9月1日現在」に合わせています"),
            ],
        ),
        Section(
            "現住所",
            [
                Field("郵便番号", "zip", "例: 598-0001（〒は不要、7桁だけでも可）"),
                Field("ふりがな（住所）", "address_kana", "例: おおさかふ いずみさのし ..."),
                Field("住所", "address", "例: 大阪府泉佐野市市場東1-2-3 サンプルハイツ101", height=30),
            ],
        ),
        Section(
            "連絡先",
            [
                Field("郵便番号", "contact_zip", ""),
                Field("ふりがな（連絡先）", "contact_kana", ""),
                Field("連絡先住所", "contact_address", "", height=30),
            ],
            note="空欄のままにすると、PDFには自動で「同上」と印字されます。",
        ),
        Section(
            "資格等",
            licenses,
            note="取得年月は 2024-06 / 令和6年6月 のどちらでもOK。PDFには和暦で印字されます。",
        ),
        Section(
            "校内外の諸活動",
            [Field("校内外の諸活動", "activities", "改行で箇条書きにできます", height=90)],
        ),
        Section(
            "志望の動機・希望の職種・アピールポイント",
            [
                Field("希望の職種", "desired_job", "例: 機械オペレーター", height=30),
                Field("アピールポイント", "appeal", "", height=90),
                Field("志望の動機", "motivation", "", height=110),
            ],
            note="この3つは用紙では1つの大きな欄です。見出し付きでまとめて印字します。",
        ),
        Section("職歴", jobs),
        Section("備考", [Field("備考", "remarks", "", height=60)]),
    ]


# ------------------------------------------------------------------ 生成
def write_template(path: str | Path, *, defaults: dict[str, object] | None = None) -> Path:
    """入力用の .xlsx を生成する。"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    defaults = defaults or {}
    path = Path(path)

    wb = Workbook()
    ws = wb.active
    ws.title = "入力"

    thin = Side(style="thin", color="B0B0B0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    input_fill = PatternFill("solid", fgColor="FFFCE8")
    section_fill = PatternFill("solid", fgColor="DCE6F1")
    title_font = Font(size=14, bold=True)
    section_font = Font(size=11, bold=True)
    help_font = Font(size=9, color="666666")
    wrap_top = Alignment(vertical="top", wrap_text=True)

    ws["A1"] = "履歴書 入力シート（近畿高等学校統一用紙 その2・令和7年度改定）"
    ws["A1"].font = title_font
    ws["A2"] = "黄色いセルに入力して保存すると、build_pdf.py（または watch.py）でPDFに反映されます。"
    ws["A2"].font = help_font

    row = 4
    for section in sheet_definition():
        ws.cell(row=row, column=1, value=section.title).font = section_font
        for col in range(1, 5):
            ws.cell(row=row, column=col).fill = section_fill
        if section.note:
            ws.cell(row=row, column=4, value=section.note).font = help_font
        row += 1

        is_table = any(f.key_c for f in section.fields)
        if is_table:
            headers = ("取得年月", "資格等の名称") if section.title == "資格等" else ("年月", "内容")
            ws.cell(row=row, column=2, value=headers[0]).font = help_font
            ws.cell(row=row, column=3, value=headers[1]).font = help_font
            row += 1

        for f in section.fields:
            ws.cell(row=row, column=1, value=f.label).alignment = wrap_top
            b = ws.cell(row=row, column=2, value=defaults.get(f.key))
            b.fill, b.border, b.alignment = input_fill, border, wrap_top
            if f.key_c:
                c = ws.cell(row=row, column=3, value=defaults.get(f.key_c))
                c.fill, c.border, c.alignment = input_fill, border, wrap_top
            else:
                ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
                ws.cell(row=row, column=3).fill = input_fill
                ws.cell(row=row, column=3).border = border
            if f.help:
                h = ws.cell(row=row, column=4, value=f.help)
                h.font, h.alignment = help_font, wrap_top
            ws.cell(row=row, column=KEY_COL_B, value=f.key)
            if f.key_c:
                ws.cell(row=row, column=KEY_COL_C, value=f.key_c)
            if f.height:
                ws.row_dimensions[row].height = f.height
            row += 1
        row += 1

    widths = {"A": 24, "B": 34, "C": 34, "D": 44}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for col in (KEY_COL_B, KEY_COL_C):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].hidden = True
        ws.column_dimensions[letter].width = 20

    ws.freeze_panes = "A4"

    guide = wb.create_sheet("使い方")
    for i, line in enumerate(USAGE_LINES, start=1):
        cell = guide.cell(row=i, column=1, value=line)
        cell.alignment = Alignment(vertical="top")
        if line.startswith("■"):
            cell.font = Font(bold=True)
    guide.column_dimensions["A"].width = 100

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


USAGE_LINES = [
    "■ このファイルについて",
    "「入力」シートの黄色いセルに入力すると、履歴書PDF（近畿高等学校統一用紙 その2）に自動で反映されます。",
    "",
    "■ 使い方（Excel）",
    "1. 「入力」シートに入力して上書き保存する。",
    "2. コマンドで  python build_pdf.py  を実行する。→ 出力/履歴書.pdf ができる。",
    "3. 入力しながら自動で作り直したいときは  python watch.py  を起動しておく（保存するたびPDFが更新されます）。",
    "",
    "■ 使い方（Googleスプレッドシート）",
    "1. このファイルをGoogleドライブにアップロードし、スプレッドシートとして開いて入力する。",
    "2. ファイル → ダウンロード → Microsoft Excel(.xlsx) または CSV で書き出す。",
    "3. python build_pdf.py 書き出したファイル  を実行する。",
    "※ 非表示のE・F列（キー）は消さないでください。読み取りに使っています。",
    "",
    "■ 自動で入る項目",
    "・満○歳 … 生年月日と基準日から自動計算します。",
    "・元号（昭和／平成／令和） … 生年月日から自動判定します。",
    "・連絡先 … 空欄なら「同上」と印字します。",
    "・郵便番号 … 7桁の数字だけでも 123-4567 の形に整えます。",
    "・資格の取得年月 … 西暦で入力しても和暦（例: 令和6年6月）で印字します。",
    "",
    "■ 注意",
    "・行を増やしたいときは、E・F列のキー（license.9.ym など）も合わせて入れてください。",
    "・写真は用紙の「写真をはる位置」に貼ってください（PDFには合成しません）。",
]


# ------------------------------------------------------------------ 読み取り
def read_values(path: str | Path) -> dict[str, object]:
    """入力シート（.xlsx / .csv）からキー→値の辞書を読み取る。"""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"入力ファイルが見つかりません: {path}")
    if path.suffix.lower() == ".csv":
        return _read_csv(path)
    return _read_xlsx(path)


def _read_xlsx(path: Path) -> dict[str, object]:
    from openpyxl import load_workbook

    wb = load_workbook(path, data_only=True)
    ws = wb["入力"] if "入力" in wb.sheetnames else wb.worksheets[0]
    values: dict[str, object] = {}
    for row in ws.iter_rows():
        cells = {c.column: c.value for c in row}
        for key_col, value_col in ((KEY_COL_B, 2), (KEY_COL_C, 3)):
            key = cells.get(key_col)
            if isinstance(key, str) and key.strip():
                values[key.strip()] = cells.get(value_col)
    if not values:
        raise ValueError(
            f"{path} からキー列（E・F列）が読み取れませんでした。"
            "make_xlsx.py で作った入力シートを使ってください。"
        )
    return values


def _read_csv(path: Path) -> dict[str, object]:
    values: dict[str, object] = {}
    with path.open(encoding="utf-8-sig", newline="") as fh:
        for row in csv.reader(fh):
            def cell(idx: int):
                return row[idx] if len(row) > idx else None

            for key_idx, value_idx in ((KEY_COL_B - 1, 1), (KEY_COL_C - 1, 2)):
                key = cell(key_idx)
                if isinstance(key, str) and key.strip():
                    values[key.strip()] = cell(value_idx)
    if not values:
        raise ValueError(
            f"{path} からキー列（E・F列）が読み取れませんでした。"
            "CSVで書き出すときは非表示のE・F列も含めてください。"
        )
    return values
