#!/usr/bin/env python3
"""Excelだけで使える履歴書ブック「履歴書作成.xlsx」を作る（開発用スクリプト）。

先生はこのスクリプトを動かす必要はない。出来上がった 履歴書作成.xlsx を
Excelで開いて使うだけで、PDFまで作れる。

作りかた:
    公式様式（近畿高等学校統一用紙 その2 のExcel版）の履歴書シートを土台にして、
    - 「入力」（1行＝1生徒の名簿）
    - 「設定」「反映項目」「学科マスタ」「資格マスタ」「資格集約」
    - 生徒40人分の履歴書シート（01〜40）
    を組み立てる。履歴書シートの記入欄は、すべて数式で「入力」シートを参照する。

    罫線が画像で描かれているため、openpyxl では図が保存できない。
    そこで、いったん openpyxl で保存したあと、元ファイルの図（drawing）と
    画像をZIPに入れ直している（_attach_drawings）。

    python make_excel_book.py --official <公式様式.xlsx> -o 履歴書作成.xlsx
"""

from __future__ import annotations

import argparse
import datetime as dt
import re
import shutil
import sys
import zipfile
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.pagebreak import Break
from openpyxl.workbook.defined_name import DefinedName

from rirekisho.inputs import DEFAULT_MASTER, OUTPUT_FIELDS, columns

FORM_SHEET_SRC = "履歴書元データ"
ROSTER = "入力"
SETTINGS = "設定"
FIELDS = "反映項目"
COURSES = "学科マスタ"
MASTER = "資格マスタ"
GATHER = "資格集約"
CALC = "計算"
CALC_FIRST_ROW = 3
PASTE = "資格取込"
PASTE_FIRST_ROW = 6      # 貼り付けを始める行
PASTE_ROWS = 400         # 貼り付けられる件数
IMPORT_SLOTS = 20        # 1人が取り込める件数
GUIDE = "使い方"
ZIPCODES = "郵便番号"
FORM_SHEET = "履歴書"
BLOCK_ROWS = 91          # 履歴書1人分の行数（1ページ）
ZIP_PREFECTURES = ("大阪府", "和歌山県", "奈良県")

ROSTER_FIRST_ROW = 5      # 生徒1人目の行
STUDENTS = 40             # 履歴書シートの枚数
LICENSE_SLOTS = 6         # 「入力」シートの資格の枠数
GATHER_SLOTS = LICENSE_SLOTS + IMPORT_SLOTS   # 資格集約の1人分の行数（手入力＋取込）
LICENSE_ROWS_ON_FORM = GATHER_SLOTS  # 用紙の資格欄に流し込む件数（全件）
CALC_LICENSE_COL = 15                                          # 計算シートの資格欄の開始列
CALC_LICENSE_MAX = LICENSE_ROWS_ON_FORM                           # 計算シートが持つ資格の件数
CALC_AFTER_LICENSE = CALC_LICENSE_COL + CALC_LICENSE_MAX * 2      # 諸活動から先の開始列
GATHER_FIRST_ROW = 3

# 学科（公式様式の6シートから。電気系電気技術科は原本の誤記を直してある）
COURSE_LIST = [
    "産業創造系製品開発科（専科）",
    "産業創造系ﾃｷｽﾀｲﾙﾃﾞｻﾞｲﾝ科（専科）",
    "機械系機械技術科（専科）",
    "機械系機械設計科（専科）",
    "電気系電気技術科（専科）",
    "電気系電子制御科（専科）",
]

# 用紙（履歴書シート）の記入欄。(セル範囲, 配置, 文字サイズ, 折り返し)
NAME_KANA_CELL = "L11:AT14"
NAME_CELL = "L15:AT20"
BIRTH_YEAR_CELL = "V21:W24"
BIRTH_MONTH_CELL = "Z21:AB24"
BIRTH_DAY_CELL = "AE21:AG24"
BIRTH_AGE_CELL = "AP21:AQ24"
ADDR_KANA_CELL = "L25:BI27"
ADDR_ZIP_CELL = "P28:AJ30"
ADDR_CELL = "L31:BI33"
CONTACT_KANA_CELL = "L34:BI36"
CONTACT_ZIP_CELL = "P37:AJ39"
CONTACT_CELL = "L40:BI41"
SCHOOL_CELL = "Y54:AV59"          # 学校名＋学科（元から結合済み）
GRAD_YEAR_CELL = "P55:R57"        # 卒業（見込）年
GRAD_MONTH_CELL = "U55:V57"       # 同 月
TODAY_YEAR_CELL = "AD8:AF10"      # 「令和　年　月　日現在」
TODAY_MONTH_CELL = "AI8:AK10"
TODAY_DAY_CELL = "AN8:AO10"
ACTIVITIES_CELL = "BW30:DU46"
MOTIVATION_CELL = "BW47:DU75"
REMARKS_CELL = "BW76:DU87"
LICENSE_YM_COLS = ("BW", "CG")
LICENSE_NAME_COLS = ("CH", "DU")
LICENSE_FONT_SIZE = 11.0
# 1人分を必ず1ページに収めるための印刷設定。
# 縮小率は公式様式のまま（97%）＝枠の大きさは正式なものから変えない。
# 用紙に収める分は、何も描かれていない下2行（90・91行目）を高さ0にし、
# 上下の余白を詰めることでまかなう（余白は枠の大きさに影響しない）。
PRINT_SCALE = 97             # 縮小率(%)。公式様式と同じ＝枠は原寸
BLANK_TAIL_ROWS = 2          # 1人分の下端の空き行（90・91行目）。高さ0にする
PRINT_MARGIN_TOP = 0.08      # 上の余白(inch)
PRINT_MARGIN_BOTTOM = 0.16   # 下の余白(inch)
PRINT_MARGIN_HEADER = 0.05   # ヘッダーの余白(inch)。上の余白以下にする
PRINT_MARGIN_FOOTER = 0.10   # フッターの余白(inch)。下の余白以下にする
BODY_FONT_SIZE = 11.0        # 校内外の諸活動・志望の動機・備考

# 用紙の行は左側の欄（氏名・生年月日・現住所）と共有しているため、
# 資格欄だけ行の高さを変えると様式全体が崩れる。行の高さは一切変えない。
LICENSE_AREA_ROWS = (9, 29)          # 資格欄（この範囲を1つの高いセルとして使う）
LICENSE_AREA_HEIGHT = 6.75 * 21      # 資格欄の高さ(pt)
JOB_ROW_BANDS = [(64, 69), (70, 75), (76, 81), (82, 87)]
JOB_YEAR_COLS = ("P", "R")
JOB_MONTH_COLS = ("U", "V")
JOB_TEXT_COLS = ("Y", "BI")

FORM_FONT = "ＭＳ Ｐ明朝"
# 用紙の数字欄はとても狭い（2列＝約16ピクセル）。大きい文字だと Excel で ### になるため、
# 欄の幅に合わせて文字サイズを決める。
NARROW_NUM_SIZE = 8.0    # 2列分の欄（生年の和暦・満年齢・卒業月・職歴の月・日付の日）
WIDE_NUM_SIZE = 9.5      # 3列分の欄（生月・生日・年・月）


# ------------------------------------------------------------------ 便利関数
def roster_index(key: str) -> int:
    """「入力」シートで、そのキーの列番号を返す。"""
    for i, col in enumerate(columns(), start=1):
        if col.key == key:
            return i
    raise KeyError(key)


def roster_col(key: str) -> str:
    """「入力」シートで、そのキーの列文字を返す。"""
    return get_column_letter(roster_index(key))


def master_lookup(name_ref: str) -> str:
    """資格名を正式名称に直す式（空白の有無は無視して照合する）。"""
    key = f'ASC(SUBSTITUTE(SUBSTITUTE({name_ref}," ",""),"　",""))'
    return (f'IFERROR(INDEX({MASTER}!$B:$B,MATCH({key},{MASTER}!$D:$D,0)),{name_ref})')


def field_cell(key: str) -> str:
    """「反映項目」シートで、その項目の○×セル（例: $B$4）を返す。"""
    for i, (field_key, _label, _default) in enumerate(OUTPUT_FIELDS, start=4):
        if field_key == key:
            return f"{FIELDS}!$B${i}"
    raise KeyError(key)


def guard(field_key: str, value_formula: str, empty_check: str | None = None) -> str:
    """反映項目が×のとき・値が空のときは空欄にする数式にする。"""
    check = empty_check if empty_check is not None else f'{value_formula}=""'
    return f'=IF(OR({field_cell(field_key)}="×",{check}),"",{value_formula})'


def set_cell(ws, ref: str, value, *, size=10.5, align="left", valign="center",
             wrap=False, shrink=False, indent=0, number_format="General"):
    """結合セルに値（数式）と書式を設定する。

    表示形式は必ず指定する。Excel は数式に DATE() などが含まれると結果を
    日付書式にしてしまい、「10」が日付として表示されて ### になるため。
    """
    first = ref.split(":")[0]
    if ":" in ref and ref not in {str(r) for r in ws.merged_cells.ranges}:
        try:
            ws.merge_cells(ref)
        except ValueError:
            pass
    cell = ws[first]
    cell.value = value
    cell.number_format = number_format
    cell.font = Font(name=FORM_FONT, size=size)
    cell.alignment = Alignment(
        horizontal=align, vertical=valign, wrap_text=wrap, shrink_to_fit=shrink, indent=indent
    )
    return cell


# ------------------------------------------------------------------ 各シート
# 画面の色（黄=入力する / 水色=自動で入る / 灰=さわらない）
FILL_INPUT = "FFFDE7"
FILL_AUTO = "D9EDF7"
FILL_META = "EFEFEF"
FILL_HEAD = "E8EEF4"
FILL_GROUP = "D6E2EF"
FILL_LINK = "DCE6F1"
TAB_COLORS = {}


def note_font() -> Font:
    return Font(size=9, color="666666")


def link_button(ws, ref: str, text: str, target: str, *, big: bool = False) -> None:
    """他のシートへ飛ぶボタン（ハイパーリンク）。"""
    cell = ws[ref]
    cell.value = text
    cell.hyperlink = Hyperlink(ref=ref, location=f"{target}!A1", display=text)
    if big:
        cell.font = Font(size=14, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2C6FBB")
        cell.border = Border(*(Side(style="medium", color="1F4E79"),) * 4)
    else:
        cell.font = Font(size=11, bold=True, color="1F4E79")
        cell.fill = PatternFill("solid", fgColor=FILL_LINK)
        cell.border = Border(*(Side(style="thin", color="9DB7D4"),) * 4)
    cell.alignment = Alignment(horizontal="center", vertical="center")


def back_to_guide(ws, ref: str = "A1", *, shift_col: int = 0) -> None:
    """右上に「使い方へ」の小さなリンクを置く。"""
    link_button(ws, ref, "◀ 使い方", GUIDE)


def color_tabs(wb) -> None:
    """シートのタブに色を付けて、役割をひと目で分かるようにする。"""
    colors = {
        GUIDE: "4472C4", ROSTER: "FFC000", FORM_SHEET: "70AD47",
        SETTINGS: "ED7D31", PASTE: "FFD966", FIELDS: "A9D08E",
        COURSES: "BFBFBF", MASTER: "BFBFBF", MACRO_SHEET: "8FAADC",
    }
    for name, color in colors.items():
        if name in wb.sheetnames:
            wb[name].sheet_properties.tabColor = color


def build_roster(wb, students: list[dict] | None) -> None:
    """「入力」シート（1行＝1生徒の名簿）。"""
    ws = wb.create_sheet(ROSTER, 1)
    cols = columns()
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    input_fill = PatternFill("solid", fgColor=FILL_INPUT)
    auto_fill = PatternFill("solid", fgColor=FILL_AUTO)
    meta_fill = PatternFill("solid", fgColor=FILL_META)
    header_fill = PatternFill("solid", fgColor=FILL_HEAD)
    group_fill = PatternFill("solid", fgColor=FILL_GROUP)
    small = note_font()

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)
    back_to_guide(ws)
    ws.cell(row=1, column=3, value="入力（1行＝1生徒・名列順）").font = Font(size=13, bold=True)
    ws.cell(
        row=1, column=6,
        value="■ 黄色 ＝ 打ち込むところ　　■ 水色 ＝ 自動で入るところ（上から書けば手入力が優先）"
              "　　名前が空の行はグレーになります",
    ).font = small
    ws.row_dimensions[1].height = 20

    start = 1
    for i, col in enumerate(cols, start=1):
        if i == len(cols) or cols[i].group != col.group:
            cell = ws.cell(row=2, column=start, value=col.group)
            cell.font, cell.alignment = Font(size=10, bold=True), Alignment(horizontal="center")
            if i > start:
                ws.merge_cells(start_row=2, start_column=start, end_row=2, end_column=i)
            for c in range(start, i + 1):
                ws.cell(row=2, column=c).fill = group_fill
            start = i + 1

    for i, col in enumerate(cols, start=1):
        head = ws.cell(row=3, column=i, value=col.label)
        head.fill, head.border = header_fill, border
        head.font = Font(size=10, bold=True)
        head.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.cell(row=4, column=i, value=col.key)
        ws.column_dimensions[get_column_letter(i)].width = col.width
    ws.row_dimensions[3].height = 30
    ws.row_dimensions[4].hidden = True

    course_dv = DataValidation(
        type="list", formula1=f"={COURSES}!$A$3:$A$8", allow_blank=True,
        prompt="6種類から選びます", promptTitle="学科",
    )
    ws.add_data_validation(course_dv)

    students = students or []
    for r in range(STUDENTS):
        row = ROSTER_FIRST_ROW + r
        data = students[r] if r < len(students) else {}
        for i, col in enumerate(cols, start=1):
            value = data.get(col.key)
            if col.key == "meta.no" and value is None:
                value = r + 1
            cell = ws.cell(row=row, column=i, value=value)
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if col.key.startswith("meta."):
                cell.fill = meta_fill
            elif col.key in ("address_kana", "contact_kana"):
                cell.fill = auto_fill
            else:
                cell.fill = input_fill
            if col.key in ("zip", "contact_zip"):
                cell.number_format = "000-0000"
            if col.kind == "date":
                cell.number_format = "yyyy/mm/dd"
        course_dv.add(ws.cell(row=row, column=cols.index(next(c for c in cols if c.key == "course")) + 1))
        ws.row_dimensions[row].height = 28

    # 郵便番号から住所のふりがなを自動で入れる（上から書き込めば手入力が優先される）
    auto_kana = [
        (roster_col("address_kana"), roster_col("zip")),
        (roster_col("contact_kana"), roster_col("contact_zip")),
    ]
    for kana_col, zip_col in auto_kana:
        for r in range(STUDENTS):
            row = ROSTER_FIRST_ROW + r
            ws[f"{kana_col}{row}"] = (
                f'=IF({zip_col}{row}="","",IFERROR(VLOOKUP(TEXT(VALUE('
                f'SUBSTITUTE({zip_col}{row},"-","")),"0000000"),{ZIPCODES}!$A:$C,3,FALSE),""))'
            )

    match_col = roster_col_match()
    ws[f"{match_col}3"] = "照合用（自動）"
    ws[f"{match_col}3"].font = Font(size=8, color="999999")
    for r in range(STUDENTS):
        row = ROSTER_FIRST_ROW + r
        ws[f"{match_col}{row}"] = (
            f'=SUBSTITUTE(SUBSTITUTE({roster_col("name")}{row}," ",""),"　","")'
        )
    ws.column_dimensions[match_col].hidden = True
    ws.freeze_panes = ws.cell(row=ROSTER_FIRST_ROW, column=5)

    # 資格と職歴はふだん見ないので、まとめて折りたためるようにする（職歴は最初から閉じる）
    first_license = get_column_letter(roster_index("license.1.name"))
    last_license = get_column_letter(roster_index(f"license.{LICENSE_SLOTS}.ym"))
    first_job = get_column_letter(roster_index("job.1.ym"))
    last_job = get_column_letter(roster_index("job.2.text"))
    ws.column_dimensions.group(first_license, last_license, outline_level=1, hidden=False)
    ws.column_dimensions.group(first_job, last_job, outline_level=1, hidden=True)
    ws.sheet_properties.outlinePr.summaryRight = True

    # 名前が入っていない行は薄いグレーにして、使っていないことを分かるようにする
    last_row = ROSTER_FIRST_ROW + STUDENTS - 1
    body_range = f"A{ROSTER_FIRST_ROW}:{get_column_letter(len(cols))}{last_row}"
    ws.conditional_formatting.add(body_range, FormulaRule(
        formula=[f'${roster_col("name")}{ROSTER_FIRST_ROW}=""'],
        fill=PatternFill(bgColor="F5F5F5"), stopIfTrue=True))

    # セルを選んだときに出る入力のヒント
    hints = {
        "birth": ("生年月日", "2008/5/12 と入れてください。元号と満○歳は自動です。"),
        "zip": ("郵便番号", "5980001 のように7桁でOK。住所のふりがなが自動で入ります。"),
        "contact_address": ("連絡先", "空欄にすると、履歴書には「同上」と入ります。"),
        "license.1.ym": ("取得年月", "令和6年6月 / 2024/6 / 2024/6/10 のどれでもOKです。"),
        "activities": ("校内外の諸活動", "改行（Alt+Enter）で箇条書きにできます。"),
    }
    for key, (title, message) in hints.items():
        dv = DataValidation(type=None, allow_blank=True, showInputMessage=True,
                            promptTitle=title, prompt=message)
        ws.add_data_validation(dv)
        letter = get_column_letter(roster_index(key))
        dv.add(f"{letter}{ROSTER_FIRST_ROW}:{letter}{last_row}")


def build_settings(wb) -> None:
    ws = wb.create_sheet(SETTINGS)
    input_fill = PatternFill("solid", fgColor="FFFDE7")
    small = Font(size=9, color="666666")
    back_to_guide(ws)
    ws["B1"] = "設定（全員に共通）"
    ws["B1"].font = Font(size=13, bold=True)
    bar_font = Font(size=12, bold=True, color="FFFFFF")
    bar_fill = PatternFill("solid", fgColor="2C6FBB")

    def bar(row: int, text: str) -> None:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        cell = ws.cell(row=row, column=1, value=text)
        cell.font, cell.fill = bar_font, bar_fill
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row].height = 28

    bar(2, "学校のこと")
    rows = [
        ("基準日（満○歳を計算する日）", dt.date(2026, 9, 1),
         "用紙の「令和　年　月　日現在」にも入ります"),
        ("学校名", "大阪府立佐野工科高等学校", "履歴書の在籍校欄に出ます"),
        ("既定の学科", COURSE_LIST[2], "「入力」の学科が空欄のとき、この学科になります"),
        ("卒業（見込）年（令和）", 9, "在籍校欄の「令和　年」"),
        ("卒業（見込）月", 3, "在籍校欄の「　月」"),
    ]
    for i, (label, value, note) in enumerate(rows, start=3):
        ws.cell(row=i, column=1, value=label)
        cell = ws.cell(row=i, column=2, value=value)
        cell.fill = input_fill
        cell.border = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
        if i == 3:
            cell.number_format = "yyyy/mm/dd"
        ws.cell(row=i, column=3, value=note).font = small
    dv = DataValidation(type="list", formula1=f"={COURSES}!$A$3:$A$8", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(ws["B5"])

    # ---- 印刷する範囲（No.○ から No.○ まで）
    bar(8, "印刷する範囲")
    ws["A9"], ws["A10"] = "開始No（この生徒から）", "終了No（この生徒まで）"
    for ref_, value in (("B9", 1), ("B10", STUDENTS)):
        cell = ws[ref_]
        cell.value = value
        cell.fill = input_fill
        cell.font = Font(size=12, bold=True)
        cell.alignment = Alignment(horizontal="center")
        cell.border = Border(*(Side(style="thin"),) * 4)
    ws["C9"] = "「入力」シートのNo（1〜40）"
    ws["C10"] = "1人だけのときは、開始と終了に同じ番号を入れます"
    for ref_ in ("C9", "C10"):
        ws[ref_].font = small
    no_dv = DataValidation(type="whole", operator="between", formula1=1, formula2=STUDENTS,
                           allow_blank=False, showErrorMessage=True,
                           errorTitle="No", error=f"1〜{STUDENTS} の数字を入れてください")
    ws.add_data_validation(no_dv)
    no_dv.add(ws["B9"])
    no_dv.add(ws["B10"])

    ws.merge_cells("A12:C13")
    link_button(ws, "A12", "▶ この範囲を印刷する（押したあと Ctrl + P）", FORM_SHEET, big=True)
    for row in ws["A12:C13"]:
        for cell in row:
            cell.border = Border(*(Side(style="medium", color="1F4E79"),) * 4)
    ws["A14"] = (
        "※ 範囲がうまく効かないときは、印刷画面の「ページ指定」に同じ番号を入れてください"
        "（1ページ＝生徒1人・ページ番号＝No）。"
    )
    ws["A14"].font = small

    for ref, text, target in (("A16", "① 入力シートへ", ROSTER),
                              ("B16", "資格をまとめて取込", PASTE),
                              ("C16", "履歴書に出す項目を選ぶ", FIELDS)):
        link_button(ws, ref, text, target)
    ws.row_dimensions[12].height = 24
    ws.row_dimensions[13].height = 24
    ws.row_dimensions[16].height = 26

    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 46
    ws.sheet_view.showGridLines = False


def build_fields(wb) -> None:
    ws = wb.create_sheet(FIELDS)
    input_fill = PatternFill("solid", fgColor="FFFDE7")
    back_to_guide(ws)
    ws["B1"] = "反映項目（履歴書に出す項目を選びます）"
    ws["B1"].font = Font(size=13, bold=True)
    ws["A2"] = "「×」にすると、入力してあってもその欄は空欄のまま印刷されます。"
    ws["A2"].font = Font(size=9, color="666666")
    ws["A3"], ws["B3"] = "項目", "反映する（○／×）"
    for cell in (ws["A3"], ws["B3"]):
        cell.font = Font(size=10, bold=True)
    dv = DataValidation(type="list", formula1='"○,×"', allow_blank=True)
    ws.add_data_validation(dv)
    for i, (key, label, default) in enumerate(OUTPUT_FIELDS, start=4):
        ws.cell(row=i, column=1, value=label)
        cell = ws.cell(row=i, column=2, value="○" if default else "×")
        cell.fill, cell.alignment = input_fill, Alignment(horizontal="center")
        dv.add(cell)
        ws.cell(row=i, column=3, value=key)
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].hidden = True


def build_courses(wb) -> None:
    ws = wb.create_sheet(COURSES)
    ws["A1"] = "学科マスタ（在籍校欄に出す学科・6種類）"
    ws["A1"].font = Font(size=12, bold=True)
    ws["A2"] = "学科名"
    ws["A2"].font = Font(size=10, bold=True)
    input_fill = PatternFill("solid", fgColor="FFFDE7")
    for i, name in enumerate(COURSE_LIST, start=3):
        cell = ws.cell(row=i, column=1, value=name)
        cell.fill = input_fill
    ws.cell(row=10, column=1, value="※ 学科名を変えたいときは、上の6行を書き換えてください。").font = Font(
        size=9, color="666666"
    )
    ws.column_dimensions["A"].width = 40


def build_master(wb) -> None:
    ws = wb.create_sheet(MASTER)
    ws["A1"] = "資格マスタ（入力した名称 → 履歴書に印字する正式名称）"
    ws["A1"].font = Font(size=12, bold=True)
    ws["A2"], ws["B2"] = "入力での名称", "履歴書での正式名称"
    for cell in (ws["A2"], ws["B2"]):
        cell.font = Font(size=10, bold=True)
    input_fill = PatternFill("solid", fgColor="FFFDE7")
    ws["C2"] = "メモ"
    ws["C2"].font = Font(size=10, bold=True)
    ws["D2"] = "照合キー（自動）"
    ws["D2"].font = Font(size=8, color="999999")
    for i, (src, dest) in enumerate(DEFAULT_MASTER, start=3):
        ws.cell(row=i, column=1, value=src).fill = input_fill
        ws.cell(row=i, column=2, value=dest).fill = input_fill
    for i in range(3, 3 + max(len(DEFAULT_MASTER), 200)):
        ws.cell(row=i, column=4,
                value=f'=IF(A{i}="","",ASC(SUBSTITUTE(SUBSTITUTE(A{i}," ",""),"　","")))')
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["C"].width = 24
    ws.column_dimensions["D"].hidden = True
    ws.freeze_panes = "A3"


def _gather_common(ws, row: int, top: int, bottom: int) -> None:
    """資格集約の共通列（正式名称・並び順・順位・表示）。"""
    ws.cell(row=row, column=5,
            value=f'=IF(D{row}="","",{master_lookup(f"D{row}")})')
    ws.cell(row=row, column=6, value=f'=IF(E{row}="","",IF(C{row}="",DATE(9999,1,1),C{row}))')
    ws.cell(row=row, column=7, value=(
        f'=IF(E{row}="","",COUNTIFS($F${top}:$F${bottom},"<"&F{row},'
        f'$F${top}:$F${bottom},"<>")+COUNTIFS($F${top}:$F${bottom},F{row},'
        f'$B${top}:$B${bottom},"<"&B{row})+1)'))
    ws.cell(row=row, column=8, value=(
        f'=IF(OR(E{row}="",C{row}=""),"",'
        f'IF(C{row}>=DATE(2019,5,1),"令和"&(YEAR(C{row})-2018),'
        f'IF(C{row}>=DATE(1989,1,8),"平成"&(YEAR(C{row})-1988),'
        f'"昭和"&(YEAR(C{row})-1925)))&"年"&MONTH(C{row})&"月")'))


def build_gather(wb) -> None:
    """資格を「取得年月順」に並べ替えるための下ごしらえシート。"""
    ws = wb.create_sheet(GATHER)
    ws["A1"] = "資格集約（自動計算・さわらないでください）"
    ws["A1"].font = Font(size=12, bold=True)
    headers = ["生徒No", "枠", "取得年月", "入力した名称", "正式名称", "並び順キー", "順位", "表示（年月）"]
    for i, label in enumerate(headers, start=1):
        ws.cell(row=2, column=i, value=label).font = Font(size=9, bold=True)

    name_cols = [roster_col(f"license.{k}.name") for k in range(1, LICENSE_SLOTS + 1)]
    ym_cols = [roster_col(f"license.{k}.ym") for k in range(1, LICENSE_SLOTS + 1)]

    paste_first, paste_last = PASTE_FIRST_ROW, PASTE_FIRST_ROW + PASTE_ROWS - 1
    for i in range(1, STUDENTS + 1):
        r = ROSTER_FIRST_ROW + i - 1
        top = GATHER_FIRST_ROW + (i - 1) * GATHER_SLOTS
        bottom = top + GATHER_SLOTS - 1
        for k in range(1, GATHER_SLOTS + 1):
            row = top + k - 1
            ws.cell(row=row, column=1, value=i)
            ws.cell(row=row, column=2, value=k)
            if k > LICENSE_SLOTS:   # 資格取込から取り込む分  (最大 IMPORT_SLOTS 件)
                j = k - LICENSE_SLOTS
                key = f'{i}&"_"&{j}'
                pos = f'MATCH({key},{PASTE}!$Z${paste_first}:$Z${paste_last},0)'
                ws.cell(row=row, column=3, value=(
                    f'=IFERROR(INDEX({PASTE}!$S${paste_first}:$S${paste_last},{pos}),"")'))
                ws.cell(row=row, column=4, value=(
                    f'=IFERROR(INDEX({PASTE}!$R${paste_first}:$R${paste_last},{pos}),"")'))
                _gather_common(ws, row, top, bottom)
                continue
            ws.cell(row=row, column=3,
                    value=f"={to_date(f'{ROSTER}!{ym_cols[k-1]}{r}')}")
            ws.cell(row=row, column=4, value=f'=IF({ROSTER}!{name_cols[k-1]}{r}="","",{ROSTER}!{name_cols[k-1]}{r})')
            _gather_common(ws, row, top, bottom)
        ws.cell(row=top, column=3).number_format = "yyyy/mm/dd"
    for col, width in (("A", 8), ("B", 6), ("C", 12), ("D", 30), ("E", 30), ("F", 12), ("G", 8), ("H", 14)):
        ws.column_dimensions[col].width = width
    ws.sheet_state = "hidden"


def build_paste(wb, paste_lines: list[str] | None = None) -> None:
    """「資格取込」シート: 貼り付けた資格一覧を数式で自動的に振り分ける。"""
    ws = wb.create_sheet(PASTE)
    first, last = PASTE_FIRST_ROW, PASTE_FIRST_ROW + PASTE_ROWS - 1
    input_fill = PatternFill("solid", fgColor="FFFDE7")
    small = Font(size=9, color="666666")

    ws["P1"] = "資格取込（下のA列に、1行＝1件で貼り付けるだけ）"
    ws["P1"].font = Font(size=13, bold=True)
    back_to_guide(ws)
    ws["A2"] = (
        "書式: 3-2-15〔タブか空白〕山田 太郎　基礎製図検定　令和6年7月10日"
        "　… 学年-組-出席番号 → 氏名 → 資格名 → 取得日"
    )
    ws["A2"].font = Font(size=10, name="ＭＳ ゴシック")
    ws["A3"] = (
        "先頭の番号がない行は氏名で照合します。取得日は 令和6年7月10日 / 2024/7/10 のどちらでも。"
    )
    ws["A3"].font = small
    ws["A4"] = (
        "右の「状態」列を見て、赤い行（要確認）だけ直してください。緑＝反映済み、灰＝重複。"
        f"　1人あたりの取込は{IMPORT_SLOTS}件までです"
        f"（「入力」シートの手入力{LICENSE_SLOTS}件と合わせて最大{LICENSE_SLOTS + IMPORT_SLOTS}件）。"
    )
    ws["A4"].font = small

    headers = {
        1: "貼付原文", 2: "整形", 3: "語数",
        4: "語1", 5: "語2", 6: "語3", 7: "語4", 8: "語5", 9: "語6", 10: "語7", 11: "語8",
        12: "ID", 13: "ID有", 14: "組", 15: "出席番号", 16: "氏名",
        17: "資格名（貼付）", 18: "正式名称", 19: "取得日",
        20: "名簿No", 21: "名簿の氏名", 22: "判定", 23: "有効", 24: "行", 25: "順位", 26: "キー",
        31: "状態",
    }
    for col, label in headers.items():
        cell = ws.cell(row=first - 1, column=col, value=label)
        cell.font = Font(size=9, bold=True)
        cell.fill = PatternFill("solid", fgColor="E8EEF4")

    name_cols = [roster_col(f"license.{k}.name") for k in range(1, LICENSE_SLOTS + 1)]
    match_col = roster_col_match()

    paste_lines = paste_lines or []
    for row in range(first, last + 1):
        pasted = paste_lines[row - first] if row - first < len(paste_lines) else None
        ws.cell(row=row, column=1, value=pasted).fill = input_fill
        raw, fmt, count = f"A{row}", f"B{row}", f"C{row}"
        ws.cell(row=row, column=2, value=(
            f'=TRIM(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({raw},CHAR(9)," "),"　"," "),CHAR(160)," "))'
        ))
        ws.cell(row=row, column=3, value=f'=IF({fmt}="",0,LEN({fmt})-LEN(SUBSTITUTE({fmt}," ",""))+1)')
        for k in range(1, 9):   # 語1〜語8
            ws.cell(row=row, column=3 + k, value=(
                f'=IF({count}>={k},TRIM(MID(SUBSTITUTE({fmt}," ",REPT(" ",200)),{(k-1)*200+1},200)),"")'
            ))
        w = {k: f"{get_column_letter(3 + k)}{row}" for k in range(1, 9)}
        ident, has_id = f"L{row}", f"M{row}"
        ws.cell(row=row, column=12, value=(
            f'=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({w[1]},"年","-"),"組","-"),"番","")'
        ))
        ws.cell(row=row, column=13, value=(
            f'=IF({w[1]}="",0,IF(ISNUMBER(VALUE(SUBSTITUTE({ident},"-",""))),1,0))'
        ))
        ws.cell(row=row, column=14, value=(
            f'=IF({has_id}=0,"",IFERROR(VALUE(TRIM(MID(SUBSTITUTE({ident},"-",REPT(" ",100)),101,100))),""))'
        ))
        ws.cell(row=row, column=15, value=(
            f'=IF({has_id}=0,"",IFERROR(VALUE(TRIM(RIGHT(SUBSTITUTE({ident},"-",REPT(" ",100)),100))),""))'
        ))
        # 氏名が「姓 名」と分かれていても正しく切り出せるよう、名簿と突き合わせる
        span = f"$D{row}:$K{row}"
        first_word = f"IF({has_id}=1,2,1)"
        cand1 = f"INDEX({span},{first_word})"
        cand2 = f"{cand1}&INDEX({span},{first_word}+1)"
        cand3 = f"{cand2}&INDEX({span},{first_word}+2)"
        roster_names = (f'{ROSTER}!${match_col}${ROSTER_FIRST_ROW}:'
                        f'${match_col}${ROSTER_FIRST_ROW + STUDENTS - 1}')
        by_number = (f'IFERROR(MATCH(O{row},{ROSTER}!$C${ROSTER_FIRST_ROW}:'
                     f'$C${ROSTER_FIRST_ROW + STUDENTS - 1},0),"")')
        ws.cell(row=row, column=28, value=f"={by_number}")           # 番号での照合
        ws.cell(row=row, column=29, value=(                           # 名簿の氏名（空白なし）
            f'=IF(AB{row}="","",INDEX({roster_names},AB{row}))'))
        ws.cell(row=row, column=30, value=(                           # 氏名の語数
            f'=IF({cand1}=AC{row},1,IF({cand2}=AC{row},2,IF({cand3}=AC{row},3,'
            f'IF(AB{row}<>"",2,'
            f'IF(ISNUMBER(MATCH({cand1},{roster_names},0)),1,'
            f'IF(ISNUMBER(MATCH({cand2},{roster_names},0)),2,'
            f'IF(ISNUMBER(MATCH({cand3},{roster_names},0)),3,2)))))))'))
        words = f"AD{row}"
        ws.cell(row=row, column=16, value=(
            f'=IF({fmt}="","",TRIM(INDEX({span},{first_word})'
            f'&IF({words}>=2," "&INDEX({span},{first_word}+1),"")'
            f'&IF({words}>=3," "&INDEX({span},{first_word}+2),"")))'))

        start = f"({first_word}+{words})"
        end = f'{count}-IF(S{row}="",0,1)'   # 取得日が読めた行は、最後の語を日付として外す
        parts = "&".join(
            f'IF(AND({k}>={start},{k}<={end}),{w[k]}&" ","")' for k in range(2, 9)
        )
        ws.cell(row=row, column=17, value=f'=IF({count}<2,"",TRIM({parts}))')
        licence = f"Q{row}"
        ws.cell(row=row, column=18, value=f'=IF({licence}="","",{master_lookup(licence)})')

        tail = f'INDEX({get_column_letter(4)}{row}:{get_column_letter(11)}{row},MIN(8,MAX(1,{count})))'
        norm = f"AA{row}"
        ws.cell(row=row, column=27, value=(
            f'=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({tail},'
            f'"年","/"),"月","/"),"日",""),"-","/"),".","/")'
        ))
        tok = [f'TRIM(MID(SUBSTITUTE({norm},"/",REPT(" ",50)),{k*50+1},50))' for k in range(3)]
        era = f"LEFT({tail},2)"
        wareki = (
            f'DATE(VALUE(MID({tail},3,FIND("年",{tail})-3))'
            f'+IF({era}="令和",2018,IF({era}="平成",1988,1925)),'
            f'VALUE(MID({tail},FIND("年",{tail})+1,FIND("月",{tail})-FIND("年",{tail})-1)),'
            f'IFERROR(VALUE(SUBSTITUTE(MID({tail},FIND("月",{tail})+1,10),"日","")),1))'
        )
        seireki = f'DATE(VALUE({tok[0]}),VALUE({tok[1]}),IFERROR(VALUE({tok[2]}),1))'
        ws.cell(row=row, column=19, value=(
            f'=IF({count}<2,"",IFERROR(IF(OR({era}="令和",{era}="平成",{era}="昭和"),'
            f'{wareki},{seireki}),IFERROR(DATEVALUE({tail}),"")))'
        ))

        number, name = f"O{row}", f"P{row}"
        ws.cell(row=row, column=20, value=(
            f'=IF({fmt}="","",IF(AB{row}<>"",AB{row},'
            f'IFERROR(MATCH(SUBSTITUTE(SUBSTITUTE({name}," ",""),"　",""),'
            f'{roster_names},0),"")))'
        ))
        no, official, date = f"T{row}", f"R{row}", f"S{row}"
        ws.cell(row=row, column=21, value=(
            f'=IF({no}="","",INDEX({ROSTER}!${roster_col("name")}${ROSTER_FIRST_ROW}:'
            f'${roster_col("name")}${ROSTER_FIRST_ROW + STUDENTS - 1},{no}))'
        ))
        dup_manual = "+".join(
            f'IF(INDEX({ROSTER}!${col}${ROSTER_FIRST_ROW}:${col}${ROSTER_FIRST_ROW + STUDENTS - 1},'
            f'{no})={official},1,0)' for col in name_cols
        )
        dup_paste = (
            f'COUNTIFS($T${first}:$T${last},{no},$R${first}:$R${last},{official},'
            f'$X${first}:$X${last},"<"&X{row})'
        )
        ws.cell(row=row, column=22, value=(
            f'=IF({fmt}="","",IF({no}="","要確認（名簿と照合できません）",'
            f'IF({official}="","要確認（資格名がありません）",'
            f'IF({date}="","要確認（取得日が読めません）",'
            f'IF(({dup_manual})+({dup_paste})>0,"重複","反映")))))'
        ))
        ws.cell(row=row, column=23, value=f'=IF(V{row}="反映",1,0)')
        ws.cell(row=row, column=24, value=row)
        ws.cell(row=row, column=25, value=(
            f'=IF(W{row}=0,"",COUNTIFS($T${first}:$T${last},{no},$W${first}:$W${last},1,'
            f'$S${first}:$S${last},"<"&{date})'
            f'+COUNTIFS($T${first}:$T${last},{no},$W${first}:$W${last},1,'
            f'$S${first}:$S${last},{date},$X${first}:$X${last},"<"&X{row})+1)'
        ))
        ws.cell(row=row, column=26, value=f'=IF(Y{row}="","",{no}&"_"&Y{row})')
        # 表示用の「状態」。1人あたりの取込枠を超えた行は、反映されないことを知らせる
        ws.cell(row=row, column=31, value=(
            f'=IF(V{row}="","",IF(AND(V{row}="反映",N(Y{row})>{IMPORT_SLOTS}),'
            f'"要確認（この生徒の取込が{IMPORT_SLOTS}件を超えました）",V{row}))'
        ))
        ws.cell(row=row, column=19).number_format = "yyyy/mm/dd"

    # 途中の計算列は隠して、見るのは「貼付原文・氏名・資格名・取得日・状態」だけにする
    for col in list(range(2, 16)) + [17, 20] + list(range(22, 31)):
        ws.column_dimensions[get_column_letter(col)].hidden = True
    for col, width in ((1, 54), (16, 14), (18, 30), (19, 13), (21, 14), (31, 34)):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = f"A{first}"

    # 「状態」の色分け（反映＝緑・重複＝灰・要確認＝赤）
    state = f"AE{first}:AE{last}"
    for formula, fg, bg in (
        (f'$AE{first}="反映"', "1E6B2F", "DFF3E2"),
        (f'$AE{first}="重複"', "666666", "EDEDED"),
        (f'LEFT($AE{first},3)="要確認"', "9C0006", "FFC7CE"),
    ):
        ws.conditional_formatting.add(state, FormulaRule(
            formula=[formula], font=Font(color=fg, bold=True),
            fill=PatternFill(bgColor=bg), stopIfTrue=True))
    ws.sheet_view.showGridLines = False


def roster_col_match() -> str:
    """「入力」シートの照合用（空白を抜いた氏名）列。"""
    return get_column_letter(len(columns()) + 2)


KATAKANA_TO_HIRAGANA = {chr(c): chr(c - 0x60) for c in range(0x30A1, 0x30F7)}


def to_hiragana(text: str) -> str:
    """カタカナの読みをひらがなにする（履歴書のふりがなはひらがなのため）。"""
    return "".join(KATAKANA_TO_HIRAGANA.get(ch, ch) for ch in text)


def zipcode_rows() -> list[tuple[str, str, str]]:
    """(郵便番号, 住所, ふりがな) の一覧。posuto の郵便番号データから作る。"""
    import json
    import sqlite3

    import posuto

    db_path = Path(posuto.__file__).with_name("postaldata.db")
    rows = []
    with sqlite3.connect(db_path) as db:
        marks = ",".join("?" * len(ZIP_PREFECTURES))
        for code, data in db.execute(
            f"select code, data from postal_data where prefecture in ({marks}) order by code",
            ZIP_PREFECTURES,
        ):
            d = json.loads(data)
            town = d["neighborhood"]
            town_kana = d["neighborhood_kana"]
            if town in ("以下に掲載がない場合",):
                town, town_kana = "", ""
            address = f'{d["prefecture"]}{d["city"]}{town}'
            kana = " ".join(
                to_hiragana(part)
                for part in (d["prefecture_kana"], d["city_kana"], town_kana) if part
            )
            rows.append((code, address, kana))
    return rows


def build_zipcodes(wb) -> None:
    """「郵便番号」シート（非表示）: 郵便番号 → 住所・ふりがな の対応表。"""
    ws = wb.create_sheet(ZIPCODES)
    ws["A1"] = "郵便番号データ（自動・さわらないでください）"
    ws["A1"].font = Font(size=12, bold=True)
    ws["A2"] = f"収録: {'・'.join(ZIP_PREFECTURES)}（日本郵便の郵便番号データより）"
    ws["A2"].font = Font(size=9, color="666666")
    for i, label in enumerate(("郵便番号", "住所", "ふりがな"), start=1):
        ws.cell(row=3, column=i, value=label).font = Font(size=9, bold=True)
    try:
        rows = zipcode_rows()
    except Exception as exc:   # データが無くてもブックは作れる
        print(f"郵便番号データを入れられませんでした: {exc}", file=sys.stderr)
        rows = []
    for i, (code, address, kana) in enumerate(rows, start=4):
        ws.cell(row=i, column=1, value=code)
        ws.cell(row=i, column=2, value=address)
        ws.cell(row=i, column=3, value=kana)
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 34
    ws.sheet_state = "hidden"
    print(f"郵便番号: {len(rows)}件")


MACRO_SHEET = "マクロ"


def vba_box(ref: str) -> str:
    """'BW30:DU46' → '30, 46, 75, 125'（VBAに渡す 上,下,左,右 の数字）。"""
    from openpyxl.utils import column_index_from_string as ci

    (c1, r1), (c2, r2) = (re.fullmatch(r"([A-Z]+)(\d+)", x).groups() for x in ref.split(":"))
    return f"{r1}, {r2}, {ci(c1)}, {ci(c2)}"


def macro_code() -> list[str]:
    """履歴書の文字を、欄に収まる大きさへ自動でそろえるVBA。

    行数を文字数から見積もるのではなく、作業用シートで Excel 自身に
    折り返させて高さを実測する（禁則処理や字幅の違いもそのまま反映される）。
    """
    ym = vba_box(f"{LICENSE_YM_COLS[0]}{LICENSE_AREA_ROWS[0]}:"
                 f"{LICENSE_YM_COLS[1]}{LICENSE_AREA_ROWS[1]}")
    name = vba_box(f"{LICENSE_NAME_COLS[0]}{LICENSE_AREA_ROWS[0]}:"
                   f"{LICENSE_NAME_COLS[1]}{LICENSE_AREA_ROWS[1]}")
    return f"""Option Explicit

' ============================================================
'  履歴書の文字を整える
'
'  資格の件数や文章の長さに合わせて、文字の大きさだけを自動で調整します。
'  行の高さ・列の幅・枠（画像）は一切変えないので、様式は崩れません。
'  基本（最大）は11ポイント。欄に入りきらないときだけ小さくします。
'
'  行数は見積もりではなく、作業用シートで Excel 自身に折り返させて
'  高さを実測します（禁則処理や字幅の違いもそのまま反映されます）。
'
'  使い方: Alt + F8 →「履歴書の文字を整える」→ 実行
' ============================================================

Private Const SHEET_FORM As String = "{FORM_SHEET}"
Private Const SHEET_WORK As String = "文字の大きさ作業用"
Private Const BLOCK_ROWS As Long = {BLOCK_ROWS}      ' 1人分の行数（1ページ）
Private Const STUDENT_COUNT As Long = {STUDENTS}     ' 名簿の人数
Private Const MAX_PT As Double = {BODY_FONT_SIZE}    ' 基本（最大）の文字の大きさ
Private Const MIN_PT As Double = 6         ' これより小さくはしない
Private Const STEP_PT As Double = 0.5      ' 大きさの刻み
Private Const YOHAKU As Double = 2         ' 欄の高さに対する余裕(pt)

Private ws測定 As Worksheet
Private 測定幅 As Double

Public Sub 履歴書の文字を整える()
    If 文字を整える実行() Then
        MsgBox "履歴書の文字を整えました。" & vbLf & _
               "入力を変えたら、もう一度実行してください。", vbInformation
    End If
End Sub

Public Function 文字を整える実行() As Boolean
    Dim frm As Worksheet, 元シート As Object
    Dim i As Long, pt As Double
    Dim pt年月(1 To STUDENT_COUNT) As Double
    Dim pt名称(1 To STUDENT_COUNT) As Double
    Dim 元計算 As Long, 元更新 As Boolean, 理由 As String

    ' 途中で失敗しても元に戻せるよう、先に安全な値を入れておく
    元計算 = xlCalculationAutomatic
    元更新 = True

    On Error GoTo エラー
    Set frm = ThisWorkbook.Worksheets(SHEET_FORM)
    Set 元シート = ActiveSheet
    元更新 = Application.ScreenUpdating
    元計算 = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    測定開始

    ' --- 資格等。幅ごとにまとめて測ると速い（測定用の列幅を作り直さずに済む）
    For i = 1 To STUDENT_COUNT
        pt年月(i) = 収まる大きさ(欄(frm, i, {ym}))
    Next i
    For i = 1 To STUDENT_COUNT
        pt名称(i) = 収まる大きさ(欄(frm, i, {name}))
    Next i
    ' 取得年月と名称は、行がずれないよう小さいほうにそろえる
    For i = 1 To STUDENT_COUNT
        pt = pt年月(i)
        If pt名称(i) < pt Then pt = pt名称(i)
        欄(frm, i, {ym}).Font.Size = pt
        欄(frm, i, {name}).Font.Size = pt
    Next i

    ' --- 校内外の諸活動・志望の動機・備考（3つとも同じ幅）
    For i = 1 To STUDENT_COUNT
        文字を合わせる 欄(frm, i, {vba_box(ACTIVITIES_CELL)})
    Next i
    For i = 1 To STUDENT_COUNT
        文字を合わせる 欄(frm, i, {vba_box(MOTIVATION_CELL)})
    Next i
    For i = 1 To STUDENT_COUNT
        文字を合わせる 欄(frm, i, {vba_box(REMARKS_CELL)})
    Next i

    測定終了
    元シート.Activate
    Application.Calculation = 元計算
    Application.ScreenUpdating = 元更新
    文字を整える実行 = True
    Exit Function

エラー:
    理由 = Err.Description          ' 後始末の前に控えておく（Err は消えてしまう）
    On Error Resume Next
    測定終了
    If Not 元シート Is Nothing Then 元シート.Activate
    Application.Calculation = 元計算
    Application.ScreenUpdating = True
    On Error GoTo 0
    MsgBox "うまくいきませんでした: " & 理由, vbExclamation
End Function

' 1人分の欄（用紙の 上行,下行,左列,右列 で指定する）
Private Function 欄(frm As Worksheet, i As Long, _
                    上 As Long, 下 As Long, 左 As Long, 右 As Long) As Range
    Dim off As Long
    off = (i - 1) * BLOCK_ROWS
    Set 欄 = frm.Range(frm.Cells(上 + off, 左), frm.Cells(下 + off, 右))
End Function

Private Sub 文字を合わせる(ByVal 対象 As Range)
    対象.Font.Size = 収まる大きさ(対象)
End Sub

' 欄に文章がちょうど収まる文字の大きさを返す
Private Function 収まる大きさ(ByVal 対象 As Range) As Double
    Dim v As Variant, s As String
    Dim pt As Double, 高さ As Double, 幅 As Double
    Dim フォント As String, 字下げ As Long

    収まる大きさ = MAX_PT
    v = 対象.Cells(1, 1).Value
    If IsError(v) Then Exit Function
    s = CStr(v)
    If Len(s) = 0 Then Exit Function

    高さ = 対象.Height - YOHAKU
    幅 = 対象.Width
    フォント = 対象.Cells(1, 1).Font.Name
    字下げ = 対象.Cells(1, 1).IndentLevel

    pt = MAX_PT
    Do While pt > MIN_PT
        If 測る(s, 幅, フォント, pt, 字下げ) <= 高さ Then Exit Do
        pt = pt - STEP_PT
    Loop
    収まる大きさ = pt
End Function

' 同じ幅・同じフォントで Excel に折り返させ、必要な高さ(pt)を実測する
Private Function 測る(s As String, 幅 As Double, フォント As String, _
                      pt As Double, 字下げ As Long) As Double
    幅を合わせる 幅
    With ws測定.Cells(1, 1)
        .ClearContents
        .NumberFormat = "@"          ' 日付などに変換されないよう文字として扱う
        .WrapText = True
        .IndentLevel = 字下げ
        .Font.Name = フォント
        .Font.Size = pt
        .Value = s
    End With
    ws測定.Rows(1).AutoFit
    測る = ws測定.Rows(1).RowHeight
End Function

' 測定用の列を、目標の幅（ポイント）以下でいちばん近い幅にする
Private Sub 幅を合わせる(幅 As Double)
    Dim i As Long, w As Double, cw As Double

    If Abs(測定幅 - 幅) < 0.4 Then Exit Sub
    ws測定.Columns(1).ColumnWidth = 10
    For i = 1 To 20
        w = ws測定.Columns(1).Width
        If w > 0 And w <= 幅 And 幅 - w < 1 Then Exit For
        If w <= 0 Then Exit For
        cw = ws測定.Columns(1).ColumnWidth * 幅 / w
        If cw < 0.05 Then cw = 0.05
        If cw > 250 Then cw = 250
        ws測定.Columns(1).ColumnWidth = cw
    Next i
    For i = 1 To 40                    ' 目標より広いときは少しずつ狭める
        If ws測定.Columns(1).Width <= 幅 Then Exit For
        cw = ws測定.Columns(1).ColumnWidth - 0.05
        If cw < 0.05 Then Exit For
        ws測定.Columns(1).ColumnWidth = cw
    Next i
    測定幅 = 幅
End Sub

Private Sub 測定開始()
    Dim 元警告 As Boolean
    元警告 = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_WORK).Delete
    On Error GoTo 0
    Set ws測定 = ThisWorkbook.Worksheets.Add
    ws測定.Name = SHEET_WORK
    Application.DisplayAlerts = 元警告
    測定幅 = -1
End Sub

Private Sub 測定終了()
    Dim 元警告 As Boolean
    If ws測定 Is Nothing Then Exit Sub
    元警告 = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    ws測定.Delete
    On Error GoTo 0
    Application.DisplayAlerts = 元警告
    Set ws測定 = Nothing
End Sub
""".splitlines()


THISWORKBOOK_CODE = """Private Sub Workbook_BeforePrint(Cancel As Boolean)
    ' 印刷・PDF出力の直前に、文字の大きさを自動でそろえる
    Application.EnableEvents = False
    On Error Resume Next
    文字を整える実行
    On Error GoTo 0
    Application.EnableEvents = True
End Sub
""".splitlines()


BAS_MODULE_NAME = "履歴書マクロ"
BAS_FILE_NAME = "履歴書マクロ.bas"


def write_bas(path: Path) -> Path:
    """VBEに「ファイル → ファイルのインポート」で読み込ませる標準モジュール。

    貼り付けだとExcelが二重引用符を壊してしまうため、ファイルで渡す。
    日本語のためVBEが読める Shift_JIS(cp932) で書く。
    """
    body = "\r\n".join([f'Attribute VB_Name = "{BAS_MODULE_NAME}"'] + macro_code())
    path.write_bytes(body.encode("cp932"))
    return path


MACRO_STEPS = [
    "■ これを入れると、資格の件数や文章の長さに合わせて文字の大きさが自動で決まります",
    "　・基本（最大）は11ポイント。欄に入りきらないときだけ小さくします（最小6ポイント）。",
    "　・変えるのは文字の大きさだけです。行の高さ・列の幅・枠は触らないので様式は崩れません。",
    "　・資格が多い生徒も、諸活動や志望の動機が長い生徒も、欄の中に全部おさまります。",
    "",
    "■ 入れかた（1回だけ・3分ほど）",
    "",
    "① このファイル（履歴書作成.xlsm）を開いたまま、Alt + F11 を押す",
    "　　黒っぽい別画面（VBAの画面）が開きます。",
    "",
    "② その画面のメニューで「ファイル」→「ファイルのインポート」をクリック",
    "　　" + BAS_FILE_NAME + " を選んで「開く」（このファイルと同じ場所に入っています）",
    "　　左側の一覧に「標準モジュール」→「" + BAS_MODULE_NAME + "」が増えれば成功です。",
    "",
    "③ Alt + Q で元のExcelに戻り、Ctrl + S で上書き保存",
    "",
    "④ Alt + F8 →「履歴書の文字を整える」を選んで「実行」",
    "　　数秒で全員分の文字の大きさがそろいます。",
    "",
    "※ ①〜③は最初の1回だけです。次からは ④ だけでOKです。",
    "※ 入力を変えたら、印刷の前に ④ をもう一度実行してください。",
    "",
    "◆ 印刷の前に自動で実行させる（おすすめ・これをすると ④ が要らなくなります）",
    "　　※ 必ず ②のインポートを先に済ませてから行ってください（順番が逆だとエラーになります）",
    "　　Alt + F11 → 左の一覧の「ThisWorkbook」をダブルクリック →",
    "　　このシートの E列 を（見出しの「E」をクリックして）まるごとコピーし、",
    "　　開いた白い画面に貼り付け →",
    "　　Alt + Q → Ctrl + S で上書き保存。",
    "　　以後、印刷・PDF出力の直前に自動で文字がそろいます。",
    "　　（E列は二重引用符を含まないので、コピー＆貼り付けでも壊れません）",
    "",
    "■ うまくいかないとき",
    "・「マクロが無効」と出る → 上部の黄色い帯の「コンテンツの有効化」を押してください。",
    "・「ファイルのインポート」に " + BAS_FILE_NAME + " が見えない →",
    "　　ファイルの種類を「すべてのファイル」にするか、履歴書作成.xlsm と同じ場所に置いてください。",
    "",
    "★ " + BAS_FILE_NAME + " をなくしてしまったとき（予備・C列を使う方法）",
    "　　C列を（見出しの「C」をクリックして）まるごとコピーし、②の代わりに",
    "　　「挿入」→「標準モジュール」で開いた白い画面に貼り付け、",
    "　　そのあと VBAの画面で Ctrl + H を押し、" + chr(0x201D) + " を \" に「すべて置換」してください。",
    "　　※ Excelはコピーのときに \" を二重にしてしまうため、C列では " + chr(0x201D) + " に置き換えてあります。",
    "　　　この置換をしないと必ずエラーになります。ふだんは ② のインポートを使ってください。",
    "",
    "※ これを入れなくても、これまでどおり11ポイントで印刷できます",
    "　（そのときは、欄に入りきらない分は表示されません）。",
]


def build_macro_sheet(wb) -> None:
    """「マクロ」シート: 入れかたの手順と、予備のVBAコード。"""
    ws = wb.create_sheet(MACRO_SHEET)
    back_to_guide(ws)
    ws["B1"] = "文字の大きさを自動でそろえる（1回だけ設定します）"
    ws["B1"].font = Font(size=13, bold=True)
    heads = ("■", "①", "②", "③", "④", "⑤", "◆", "★")
    for i, line in enumerate(MACRO_STEPS, start=3):
        cell = ws.cell(row=i, column=1, value=line)
        if line[:1] in heads:
            cell.font = Font(bold=True)
    mono = Font(name="ＭＳ ゴシック", size=9)
    ws["C1"] = ("' 【予備】★の方法で使う列。貼り付けたあとVBEのCtrl+Hで "
                + chr(0x201D) + " を半角の二重引用符に置換してください")
    ws["C1"].font = mono
    # 空行は " " にしておく（空文字だとセルが消え、貼り付けたときに詰まってしまう）
    # 二重引用符は全角にしておく（Excelはコピーのとき " を二重にしてしまうため）
    for i, line in enumerate(macro_code(), start=2):
        ws.cell(row=i, column=3, value=line.replace('"', "\u201d") or " ").font = mono
    ws["E1"] = "' ◆で ThisWorkbook に貼り付ける列（列ごとコピーしても壊れません）"
    ws["E1"].font = mono
    for i, line in enumerate(THISWORKBOOK_CODE, start=2):
        ws.cell(row=i, column=5, value=line or " ").font = mono
    ws.column_dimensions["A"].width = 84
    ws.column_dimensions["C"].width = 82
    ws.column_dimensions["E"].width = 56
    ws.sheet_view.showGridLines = False


def build_guide(wb) -> None:
    """「使い方」シート: 最初に開く案内。要点だけを短く。"""
    ws = wb.create_sheet(GUIDE, 0)
    ws.sheet_view.showGridLines = False

    title = Font(size=16, bold=True, color="1F3864")
    step = Font(size=12, bold=True, color="FFFFFF")
    head = Font(size=11, bold=True, color="1F3864")
    body = Font(size=11)
    note = note_font()
    step_fill = PatternFill("solid", fgColor="2C6FBB")
    name_font = Font(size=11, bold=True)

    ws["A1"] = "履歴書 自動作成（近畿高等学校統一用紙 その2）"
    ws["A1"].font = title
    ws["A2"] = "「入力」に打ち込む → 「履歴書」が出来上がる → 「設定」で範囲を決めて印刷。この3つだけです。"
    ws["A2"].font = note

    for ref, text, target in (("A4", "① 入力シートへ", ROSTER),
                              ("B4", "③ 設定・印刷へ", SETTINGS),
                              ("C4", "資格をまとめて取込", PASTE),
                              ("D4", "文字の大きさをそろえる", MACRO_SHEET)):
        link_button(ws, ref, text, target)
    ws.row_dimensions[4].height = 26

    def put(row: int, text: str, font: Font) -> None:
        ws.cell(row=row, column=1, value=text).font = font

    def step_bar(row: int, text: str) -> None:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        cell = ws.cell(row=row, column=1, value=text)
        cell.font, cell.fill = step, step_fill
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row].height = 28

    step_bar(6, "STEP 1　「入力」シートに打ち込む（黄色いセルだけ・1行＝1生徒）")
    for i, text in enumerate([
        "生年月日は 2008/5/12 と入れるだけ　→　元号と満○歳は自動",
        "郵便番号を入れるだけ　→　住所のふりがなが自動（直したいときは上から書けばOK）",
        "連絡先を空欄にする　→　履歴書には「同上」",
        "学科はドロップダウンから選ぶ（空欄なら「設定」の学科）",
    ], start=7):
        put(i, "　・" + text, body)

    step_bar(12, "STEP 2　「履歴書」シートを見る（1人＝1ページ・Noの順）")
    put(13, "　　打ち込んだ内容がそのまま用紙に入っています。スクロールで全員分を確認できます。", body)

    step_bar(15, "STEP 3　印刷・PDFにする")
    put(16, "　① 「設定」シートで【開始No】と【終了No】を入れる（1人だけなら同じ番号）", body)
    put(17, "　②【▶ 印刷する】を押す　→　③ そのまま Ctrl + P", body)
    put(18, "　　PDFにするときは、印刷画面のプリンターで「Microsoft Print to PDF」を選びます。", note)
    put(19, "　　※ 1人＝1ページ・枠は公式様式と同じ大きさ（拡大縮小97%）です。", note)

    put(20, "■ 資格をまとめて取り込む", head)
    put(21, "　「資格取込」シートのA6以降に、資格の一覧を1行1件で貼り付けるだけです。", body)
    put(22, "　　3-2-15〔タブ〕山田 太郎　基礎製図検定　令和6年7月10日", Font(size=10, name="ＭＳ ゴシック"))
    put(23, "　名簿と照合して自動で振り分けます。「状態」の列が赤い行だけ直してください。", body)

    put(25, "■ 文字が欄に入りきらないとき", head)
    put(26, "　「マクロ」シートの手順（1回だけ・3分）を行うと、資格・諸活動・志望の動機・備考の", body)
    put(27, "　文字の大きさが、欄に収まるよう自動でそろいます（最大11ポイント）。", body)

    put(29, "■ シートの役割", head)
    sheets = [
        (ROSTER, "生徒の情報を打ち込む（1行＝1生徒・40人分）"),
        (FORM_SHEET, "印刷する用紙。さわらなくて大丈夫です"),
        (SETTINGS, "印刷する範囲・学校名・基準日・卒業年月"),
        (PASTE, "資格の一覧を貼り付けて自動で振り分ける"),
        (FIELDS, "履歴書に出す項目を○×で選ぶ"),
        (COURSES, "学科の一覧（6種類・書き換え可）"),
        (MASTER, "資格名を正式名称に直す変換表"),
        (MACRO_SHEET, "文字の大きさをそろえる仕組みの入れかた"),
    ]
    for i, (name, what) in enumerate(sheets, start=30):
        ws.cell(row=i, column=1, value="　" + name).font = name_font
        ws.cell(row=i, column=2, value=what).font = body

    put(39, "■ 印刷したあとに手で書き足すところ", head)
    put(40, "　・写真を貼る", body)
    put(41, "　・生年月日の「昭和・平成」、職歴の「平成・令和」に丸を付ける", body)

    for col, width in (("A", 26), ("B", 30), ("C", 24), ("D", 26)):
        ws.column_dimensions[col].width = width


# ------------------------------------------------------------------ 履歴書シート
def shift(ref: str, offset: int) -> str:
    """セル範囲の行番号をずらす（例: L11:AT14 → L102:AT105）。"""
    return re.sub(r"([A-Z]+)(\d+)", lambda m: f"{m.group(1)}{int(m.group(2)) + offset}", ref)


def fill_form(ws, i: int, offset: int = 0, *, license_size=None) -> None:
    """履歴書シートの記入欄に、「計算」シートを参照する数式を入れる。"""
    license_size = license_size or LICENSE_FONT_SIZE
    row = CALC_FIRST_ROW + i - 1

    def calc(col: str) -> str:
        return f"={CALC}!${col}${row}"

    set_cell(ws, shift(NAME_KANA_CELL, offset), calc("C"), size=10, align="center", shrink=True)
    set_cell(ws, shift(NAME_CELL, offset), calc("B"), size=16, align="center", shrink=True)
    set_cell(ws, shift(BIRTH_YEAR_CELL, offset), calc("D"), align="center", size=NARROW_NUM_SIZE)
    set_cell(ws, shift(BIRTH_MONTH_CELL, offset), calc("E"), align="center", size=WIDE_NUM_SIZE)
    set_cell(ws, shift(BIRTH_DAY_CELL, offset), calc("F"), align="center", size=WIDE_NUM_SIZE)
    set_cell(ws, shift(BIRTH_AGE_CELL, offset), calc("G"), align="center", size=NARROW_NUM_SIZE)

    set_cell(ws, shift(ADDR_ZIP_CELL, offset), calc("H"), size=BODY_FONT_SIZE)
    set_cell(ws, shift(ADDR_CELL, offset), calc("I"), size=BODY_FONT_SIZE, wrap=True, indent=1)
    set_cell(ws, shift(ADDR_KANA_CELL, offset), calc("J"), size=9, indent=1)

    set_cell(ws, shift(CONTACT_ZIP_CELL, offset), calc("K"), size=BODY_FONT_SIZE)
    set_cell(ws, shift(CONTACT_CELL, offset), calc("L"), size=BODY_FONT_SIZE,
             align="center", wrap=True)
    set_cell(ws, shift(CONTACT_KANA_CELL, offset), calc("M"), size=9, indent=1)

    set_cell(ws, shift(SCHOOL_CELL, offset), calc("N"), size=BODY_FONT_SIZE, wrap=True)

    # 名簿が空の行（使っていない生徒）は、日付や卒業年月も出さない
    used = f"{CALC}!$B${row}"

    def only_if_used(formula: str) -> str:
        return f'=IF({used}="","",{formula})&""'

    set_cell(ws, shift(GRAD_YEAR_CELL, offset), only_if_used(f"{SETTINGS}!$B$6"),
             align="center", size=WIDE_NUM_SIZE)
    set_cell(ws, shift(GRAD_MONTH_CELL, offset), only_if_used(f"{SETTINGS}!$B$7"),
             align="center", size=NARROW_NUM_SIZE)

    base = f"{SETTINGS}!$B$3"
    set_cell(ws, shift(TODAY_YEAR_CELL, offset),
             only_if_used(f'IF(N({base})=0,"",YEAR({base})-2018)'),
             align="center", size=WIDE_NUM_SIZE)
    set_cell(ws, shift(TODAY_MONTH_CELL, offset),
             only_if_used(f'IF(N({base})=0,"",MONTH({base}))'),
             align="center", size=WIDE_NUM_SIZE)
    set_cell(ws, shift(TODAY_DAY_CELL, offset),
             only_if_used(f'IF(N({base})=0,"",DAY({base}))'),
             align="center", size=NARROW_NUM_SIZE)

    # 資格は「取得年月」「名称」それぞれ1つの高いセルに、改行でつないで流し込む。
    # 行の高さを触らないので様式は崩れず、文字の大きさだけで件数に対応できる。
    top_row, bottom_row = LICENSE_AREA_ROWS
    list_ym = get_column_letter(CALC_AFTER_LICENSE + 13)
    list_name = get_column_letter(CALC_AFTER_LICENSE + 14)
    set_cell(ws, shift(f"{LICENSE_YM_COLS[0]}{top_row}:{LICENSE_YM_COLS[1]}{bottom_row}", offset),
             calc(list_ym), size=license_size, align="center", valign="top", wrap=True)
    set_cell(ws, shift(f"{LICENSE_NAME_COLS[0]}{top_row}:{LICENSE_NAME_COLS[1]}{bottom_row}", offset),
             calc(list_name), size=license_size, valign="top", wrap=True, indent=1)

    set_cell(ws, shift(ACTIVITIES_CELL, offset), calc(get_column_letter(CALC_AFTER_LICENSE)),
             size=BODY_FONT_SIZE, valign="top", wrap=True, indent=1)
    set_cell(ws, shift(MOTIVATION_CELL, offset), calc(get_column_letter(CALC_AFTER_LICENSE + 1)),
             size=BODY_FONT_SIZE, valign="top", wrap=True, indent=1)
    set_cell(ws, shift(REMARKS_CELL, offset), calc(get_column_letter(CALC_AFTER_LICENSE + 2)),
             size=BODY_FONT_SIZE, valign="top", wrap=True, indent=1)

    for k, (row_top, row_bottom) in enumerate(JOB_ROW_BANDS[:2], start=1):
        year_col = get_column_letter(CALC_AFTER_LICENSE + 3 + (k - 1) * 3)
        month_col = get_column_letter(CALC_AFTER_LICENSE + 4 + (k - 1) * 3)
        text_col = get_column_letter(CALC_AFTER_LICENSE + 5 + (k - 1) * 3)
        set_cell(ws, shift(f"{JOB_YEAR_COLS[0]}{row_top+2}:{JOB_YEAR_COLS[1]}{row_top+3}", offset),
                 calc(year_col), align="center", size=WIDE_NUM_SIZE)
        set_cell(ws, shift(f"{JOB_MONTH_COLS[0]}{row_top+2}:{JOB_MONTH_COLS[1]}{row_top+3}", offset),
                 calc(month_col), align="center", size=NARROW_NUM_SIZE)
        set_cell(ws, shift(f"{JOB_TEXT_COLS[0]}{row_top}:{JOB_TEXT_COLS[1]}{row_bottom}", offset),
                 calc(text_col), size=BODY_FONT_SIZE, indent=1, wrap=True)


def build_form_sheet(wb, src, title: str = FORM_SHEET, *, license_size=None) -> None:
    """1枚の履歴書シートに、生徒40人分を縦に並べる（1人＝1ページ）。

    用紙の行の高さは変えない（左の欄と行を共有しているため）。
    """
    from copy import copy

    ws = wb.create_sheet(title)
    ws.sheet_format = copy(src.sheet_format)
    ws.sheet_view.showGridLines = False
    ws.page_setup = copy(src.page_setup)
    ws.page_margins = copy(src.page_margins)
    ws.print_options = copy(src.print_options)

    # 1人分（1ブロック）が必ず1ページに収まるようにする。
    # 元の様式は 97% ・上下余白 0.354inch で、用紙の高さ(544pt)に対して
    # 中身(602pt×0.97=584pt)がはみ出し、1人が2ページに分かれていた。
    # 縮小率は 97% のまま（枠は原寸）。下端の空き行と余白だけで収める。
    ws.page_setup.scale = PRINT_SCALE
    ws.page_setup.fitToWidth = None
    ws.page_setup.fitToHeight = None
    ws.sheet_properties.pageSetUpPr.fitToPage = False
    ws.page_margins.top = PRINT_MARGIN_TOP
    ws.page_margins.bottom = PRINT_MARGIN_BOTTOM
    ws.page_margins.header = PRINT_MARGIN_HEADER
    ws.page_margins.footer = PRINT_MARGIN_FOOTER
    for key, dim in src.column_dimensions.items():
        new = copy(dim)
        new.worksheet = ws
        ws.column_dimensions[key] = new

    src_rows = list(src.iter_rows(min_row=1, max_row=BLOCK_ROWS))
    merges = [str(r) for r in src.merged_cells.ranges]
    default_height = src.sheet_format.defaultRowHeight

    for i in range(1, STUDENTS + 1):
        offset = (i - 1) * BLOCK_ROWS
        for row in src_rows:
            for cell in row:
                if cell.value is None and not cell.has_style:
                    continue
                new = ws.cell(row=cell.row + offset, column=cell.column, value=cell.value)
                if cell.has_style:
                    new._style = copy(cell._style)
        for r in range(1, BLOCK_ROWS + 1):
            dim = src.row_dimensions.get(r)
            height = dim.height if dim is not None and dim.height else default_height
            ws.row_dimensions[r + offset].height = height
            if r > BLOCK_ROWS - BLANK_TAIL_ROWS:
                # 何も描かれていない下端の行。非表示にして印刷の高さから外す
                ws.row_dimensions[r + offset].hidden = True
        for rng in merges:
            ws.merge_cells(shift(rng, offset))
        fill_form(ws, i, offset, license_size=license_size)
        if i > 1:
            ws.row_breaks.append(Break(id=offset))

    # 「設定」の開始No〜終了No だけを印刷する（数式の印刷範囲）
    start = f"MAX(1,MIN({STUDENTS},N({SETTINGS}!$B$9)))"
    end = f"MAX({start},MIN({STUDENTS},N({SETTINGS}!$B$10)))"
    area = (
        f"OFFSET('{title}'!$A$1,({start}-1)*{BLOCK_ROWS},0,"
        f"({end}-{start}+1)*{BLOCK_ROWS},{src.max_column})"
    )
    ws.defined_names.add(DefinedName("_xlnm.Print_Area", attr_text=area))


def to_date(ref: str) -> str:
    """セルを日付にする式。日付・西暦文字・和暦文字（平成10年12月20日）に対応。"""
    era = f"LEFT({ref},2)"
    # 「日」が無い（令和6年6月）ときは1日として読む
    wareki = (
        f'DATE(VALUE(MID({ref},3,FIND("年",{ref})-3))'
        f'+IF({era}="令和",2018,IF({era}="平成",1988,IF({era}="昭和",1925,IF({era}="大正",1911,0)))),'
        f'VALUE(MID({ref},FIND("年",{ref})+1,FIND("月",{ref})-FIND("年",{ref})-1)),'
        f'IFERROR(VALUE(SUBSTITUTE(MID({ref},FIND("月",{ref})+1,10),"日","")),1))'
    )
    # 2008/5/12 や 2008-5-12、2008年5月12日 のような西暦の文字列
    norm = (f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({ref},'
            f'"年","/"),"月","/"),"日",""),"-","/"),".","/")')
    tok = [f'TRIM(MID(SUBSTITUTE({norm},"/",REPT(" ",50)),{k * 50 + 1},50))' for k in range(3)]
    seireki = f'DATE(VALUE({tok[0]}),VALUE({tok[1]}),IFERROR(VALUE({tok[2]}),1))'
    return (
        f'IF({ref}="","",IF(ISNUMBER({ref}),{ref},'
        f'IFERROR(IF(OR({era}="令和",{era}="平成",{era}="昭和",{era}="大正"),{wareki},{seireki}),'
        f'IFERROR(DATEVALUE({ref}),""))))'
    )


def build_calc(wb) -> None:
    """「計算」シート: 履歴書に出す文字を1行＝1生徒で組み立てる（非表示）。"""
    ws = wb.create_sheet(CALC)
    headers = [
        "No", "氏名", "ふりがな", "生年(和暦)", "生月", "生日", "満年齢",
        "郵便番号", "住所", "住所ふりがな", "連絡先〒", "連絡先", "連絡先ふりがな",
        "在籍校",
    ]
    for k in range(1, CALC_LICENSE_MAX + 1):
        headers += [f"資格{k}年月", f"資格{k}名称"]
    headers += [
        "諸活動", "志望の動機など", "備考",
        "職歴1年", "職歴1月", "職歴1内容", "職歴2年", "職歴2月", "職歴2内容",
    ]
    for i, label in enumerate(("資格一覧(年月)", "資格一覧(名称)", "資格件数"), start=0):
        ws.cell(row=2, column=CALC_AFTER_LICENSE + 13 + i, value=label).font = Font(size=9, bold=True)
    ws.cell(row=1, column=1, value="計算（自動・さわらないでください）").font = Font(size=12, bold=True)
    for i, label in enumerate(headers, start=1):
        ws.cell(row=2, column=i, value=label).font = Font(size=9, bold=True)

    def dateval(ref: str) -> str:
        """文字で入力されていても日付として読む。"""
        return f'IFERROR(IF(ISNUMBER({ref}),{ref},DATEVALUE({ref})),"")'

    def zip_fmt(ref: str) -> str:
        return f'IFERROR(TEXT(VALUE(SUBSTITUTE({ref},"-","")),"000-0000"),{ref})'

    for i in range(1, STUDENTS + 1):
        r = ROSTER_FIRST_ROW + i - 1
        row = CALC_FIRST_ROW + i - 1
        top = GATHER_FIRST_ROW + (i - 1) * GATHER_SLOTS
        bottom = top + GATHER_SLOTS - 1

        def ref(key: str) -> str:
            return f"{ROSTER}!${roster_col(key)}${r}"

        birth = f"${get_column_letter(CALC_AFTER_LICENSE + 10)}${row}"  # 変換済みの生年月日

        # 用紙の数字欄はとても狭いので、数値ではなく文字として出す（### 対策）
        as_text = {4, 5, 6, 7,
                   CALC_AFTER_LICENSE + 3, CALC_AFTER_LICENSE + 4,
                   CALC_AFTER_LICENSE + 6, CALC_AFTER_LICENSE + 7}

        def put(col, value):
            """氏名が空の行（＝使っていない生徒）は、すべて空欄にする。"""
            if isinstance(value, str) and value.startswith("=") and col >= 2:
                value = f'=IF({ROSTER}!${roster_col("name")}${r}="","",{value[1:]})'
                if col in as_text:
                    value = f'{value}&""'
            return ws.cell(row=row, column=col, value=value)

        put(1, i)
        put(2, guard("name", ref("name")))
        put(3, guard("name_kana", ref("name_kana")))
        put(4, guard("birth",
                     f'YEAR({birth})-IF({birth}>=DATE(2019,5,1),2018,'
                     f'IF({birth}>=DATE(1989,1,8),1988,1925))', f'N({birth})=0'))
        put(5, guard("birth", f"MONTH({birth})", f'N({birth})=0'))
        put(6, guard("birth", f"DAY({birth})", f'N({birth})=0'))
        put(7, guard("birth", f'DATEDIF({birth},{SETTINGS}!$B$3,"Y")',
                     f'OR(N({birth})=0,N({SETTINGS}!$B$3)=0)'))
        put(8, guard("zip", zip_fmt(ref("zip")), f'{ref("zip")}=""'))
        put(9, guard("address", ref("address")))
        zip_key = f'TEXT(VALUE(SUBSTITUTE({ref("zip")},"-","")),"0000000")'
        lookup = f'IFERROR(VLOOKUP({zip_key},{ZIPCODES}!$A:$C,3,FALSE),"")'
        put(10, f'=IF({field_cell("address_kana")}="×","",'
                f'IF({ref("address_kana")}<>"",{ref("address_kana")},'
                f'IF({ref("zip")}="","",{lookup})))')
        put(11, guard("contact", zip_fmt(ref("contact_zip")), f'{ref("contact_zip")}=""'))
        put(12, f'=IF({field_cell("contact")}="×","",'
                f'IF(AND({ref("contact_zip")}="",{ref("contact_address")}=""),"同上",'
                f'IF({ref("contact_address")}="","",{ref("contact_address")})))')
        put(13, f'=IF(OR({field_cell("contact")}="×",{ref("contact_kana")}="",'
                f'AND({ref("contact_zip")}="",{ref("contact_address")}="")),"",{ref("contact_kana")})')
        put(14, f'={SETTINGS}!$B$4&CHAR(10)&IF({ref("course")}="",{SETTINGS}!$B$5,{ref("course")})')

        for k in range(1, CALC_LICENSE_MAX + 1):
            rank = f'MATCH({k},{GATHER}!$G${top}:$G${bottom},0)'
            col = CALC_LICENSE_COL + (k - 1) * 2
            put(col, f'=IF({field_cell("licenses")}="×","",'
                     f'IFERROR(INDEX({GATHER}!$H${top}:$H${bottom},{rank}),""))')
            put(col + 1, f'=IF({field_cell("licenses")}="×","",'
                         f'IFERROR(INDEX({GATHER}!$E${top}:$E${bottom},{rank}),""))')

        put(CALC_AFTER_LICENSE, guard("activities", ref("activities")))
        put(CALC_AFTER_LICENSE + 1, f'=IF({field_cell("motivation")}="×","",'
                f'IF({ref("motivation")}="","",{ref("motivation")}&CHAR(10)))'
                f'&IF({field_cell("desired_job")}="×","",'
                f'IF({ref("desired_job")}="","",{ref("desired_job")}&CHAR(10)))'
                f'&IF({field_cell("appeal")}="×","",'
                f'IF({ref("appeal")}="","",{ref("appeal")}))')
        put(CALC_AFTER_LICENSE + 2, guard("remarks", ref("remarks")))

        work = CALC_AFTER_LICENSE + 10         # 作業列（生年月日・職歴年月）
        for k in (1, 2):
            job_ym = f"${get_column_letter(work + k)}${row}"
            base_col = CALC_AFTER_LICENSE + 3 + (k - 1) * 3
            put(base_col, guard("jobs",
                                f'YEAR({job_ym})-IF({job_ym}>=DATE(2019,5,1),2018,1988)',
                                f'N({job_ym})=0'))
            put(base_col + 1, guard("jobs", f"MONTH({job_ym})", f'N({job_ym})=0'))
            put(base_col + 2, guard("jobs", ref(f"job.{k}.text")))

        # 資格の一覧（改行でつないだもの）と件数。用紙では1つの高いセルに流し込む
        ym_cells = [get_column_letter(CALC_LICENSE_COL + (k - 1) * 2) + str(row)
                    for k in range(1, CALC_LICENSE_MAX + 1)]
        name_cells = [get_column_letter(CALC_LICENSE_COL + (k - 1) * 2 + 1) + str(row)
                      for k in range(1, CALC_LICENSE_MAX + 1)]
        # 行がずれないよう、どちらの列も「名称が入っているか」で改行を決める
        join_ym = "&".join([f"${ym_cells[0]}"] + [
            f'IF(${n}="","",CHAR(10)&${c})' for c, n in zip(ym_cells[1:], name_cells[1:])])
        join_name = "&".join([f"${name_cells[0]}"] + [
            f'IF(${c}="","",CHAR(10)&${c})' for c in name_cells[1:]])
        count = "+".join(f'IF(${c}="",0,1)' for c in name_cells)
        put(work + 3, f"={join_ym}")
        put(work + 4, f"={join_name}")
        put(work + 5, f"={count}")

        # 作業列（日付に変換したもの。和暦の文字で入力されていても読む）
        put(work, f"={to_date(ref('birth'))}")
        put(work + 1, f"={to_date(ref('job.1.ym'))}")
        put(work + 2, f"={to_date(ref('job.2.ym'))}")

    # 和暦年・月・日・満年齢・職歴の年月は数値として表示する（### 対策）
    for i in range(1, STUDENTS + 1):
        row = CALC_FIRST_ROW + i - 1
        for col in (4, 5, 6, 7, CALC_AFTER_LICENSE + 3, CALC_AFTER_LICENSE + 4,
                    CALC_AFTER_LICENSE + 6, CALC_AFTER_LICENSE + 7):
            ws.cell(row=row, column=col).number_format = "General"

    for i in range(CALC_AFTER_LICENSE + 10, CALC_AFTER_LICENSE + 13):
        ws.column_dimensions[get_column_letter(i)].hidden = True
    ws.sheet_state = "hidden"


# ------------------------------------------------------------------ 組み立て
def build(official: Path, out: Path, students: list[dict] | None = None,
          sheets: int = STUDENTS, paste_lines: list[str] | None = None) -> Path:
    wb = load_workbook(official)
    form = wb[FORM_SHEET_SRC]
    form._images = []
    form._charts = []

    build_form_sheet(wb, form)
    del wb[FORM_SHEET_SRC]
    for name in list(wb.sheetnames):
        if name != FORM_SHEET:
            del wb[name]

    build_roster(wb, students)
    build_settings(wb)
    build_fields(wb)
    build_courses(wb)
    build_master(wb)
    build_paste(wb, paste_lines)
    build_gather(wb)
    build_calc(wb)
    build_zipcodes(wb)
    build_macro_sheet(wb)
    build_guide(wb)

    order = [GUIDE, ROSTER, PASTE, SETTINGS, FIELDS, COURSES, MASTER,
             FORM_SHEET, MACRO_SHEET, GATHER, CALC, ZIPCODES]
    wb._sheets.sort(key=lambda ws: order.index(ws.title) if ws.title in order else 99)
    color_tabs(wb)
    wb.active = 0
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    _attach_drawings(official, out, [FORM_SHEET])
    write_bas(Path(BAS_FILE_NAME) if out.parent == Path("samples")
              else out.parent / BAS_FILE_NAME)
    return out


def _sheet_parts(zf: zipfile.ZipFile) -> dict[str, str]:
    """シート名 → パート名（xl/worksheets/sheetN.xml）の対応を返す。"""
    import xml.etree.ElementTree as ET

    ns_r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    targets = {rel.get("Id"): rel.get("Target") for rel in rels}
    book = ET.fromstring(zf.read("xl/workbook.xml"))
    parts = {}
    for sheet in book.iter():
        if not sheet.tag.endswith("}sheet"):
            continue
        rid = sheet.get(f"{ns_r}id")
        if rid is None:
            continue
        target = targets[rid].lstrip("/")
        parts[sheet.get("name")] = target if target.startswith("xl/") else "xl/" + target
    return parts


ANCHOR_RE = re.compile(rb"<xdr:(oneCellAnchor|twoCellAnchor)\b.*?</xdr:\1>", re.S)


def _repeat_drawing(drawing_xml: bytes, blocks: int, block_rows: int) -> bytes:
    """罫線の図を、ページの数だけ行をずらして並べ直す。"""
    anchors = ANCHOR_RE.findall(drawing_xml)
    matches = list(ANCHOR_RE.finditer(drawing_xml))
    if not matches:
        return drawing_xml
    head = drawing_xml[: matches[0].start()]
    tail = drawing_xml[matches[-1].end():]
    body = bytearray()
    ident = 1000
    for i in range(blocks):
        offset = i * block_rows
        for m in matches:
            chunk = m.group(0)
            chunk = re.sub(
                rb"<xdr:row>(\d+)</xdr:row>",
                lambda mm: b"<xdr:row>%d</xdr:row>" % (int(mm.group(1)) + offset),
                chunk,
            )
            ident += 1
            chunk = re.sub(rb'(<xdr:cNvPr[^>]*\bid=")\d+(")',
                           rb"\g<1>%d\g<2>" % ident, chunk)
            body += chunk
    return head + bytes(body) + tail


def _attach_drawings(official: Path, out: Path, form_sheets: list[str]) -> None:
    """罫線の図（drawing）と画像を、出来上がったブックの各履歴書シートに入れ直す。

    openpyxl は読み込んだ図をそのまま書き戻せないので、元ファイルから
    drawing と media をコピーし、シートXMLに <drawing> を差し込む。
    """
    with zipfile.ZipFile(official) as src:
        src_parts = _sheet_parts(src)
        src_sheet = src_parts[FORM_SHEET_SRC]
        sheet_rels = f"xl/worksheets/_rels/{Path(src_sheet).name}.rels"
        rels_xml = src.read(sheet_rels).decode("utf-8")
        drawing_target = re.search(r'Target="([^"]*drawings/drawing\d+\.xml)"', rels_xml).group(1)
        drawing_part = "xl/" + drawing_target.replace("../", "")
        drawing_xml = src.read(drawing_part)
        drawing_rels_part = f"xl/drawings/_rels/{Path(drawing_part).name}.rels"
        drawing_rels = src.read(drawing_rels_part)
        media = {n: src.read(n) for n in src.namelist() if n.startswith("xl/media/")}

    tmp = out.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(out) as zin:
        parts = _sheet_parts(zin)
        names = zin.namelist()
        data = {n: zin.read(n) for n in names}

    for index, sheet_name in enumerate(form_sheets, start=1):
        part = parts[sheet_name]
        drawing_name = f"drawing{index}.xml"
        data[f"xl/drawings/{drawing_name}"] = _repeat_drawing(drawing_xml, STUDENTS, BLOCK_ROWS)
        data[f"xl/drawings/_rels/{drawing_name}.rels"] = drawing_rels

        rels_name = f"xl/worksheets/_rels/{Path(part).name}.rels"
        rel_id = "rIdDraw1"
        rel = (f'<Relationship Id="{rel_id}" Type="http://schemas.openxmlformats.org/'
               f'officeDocument/2006/relationships/drawing" Target="../drawings/{drawing_name}"/>')
        if rels_name in data:
            xml = data[rels_name].decode("utf-8")
            data[rels_name] = xml.replace("</Relationships>", rel + "</Relationships>").encode("utf-8")
        else:
            data[rels_name] = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f"{rel}</Relationships>"
            ).encode("utf-8")

        sheet_xml = data[part].decode("utf-8")
        sheet_xml = re.sub(r'\s*<drawing[^>]*/>', "", sheet_xml)
        if "xmlns:r=" not in sheet_xml.split(">", 1)[0] + ">":
            # openpyxl のシートXMLには r: の名前空間が無いことがあるので足す
            sheet_xml = re.sub(
                r"(<worksheet\b[^>]*)",
                r'\1 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
                sheet_xml, count=1,
            )
        sheet_xml = sheet_xml.replace("</worksheet>", f'<drawing r:id="{rel_id}"/></worksheet>')
        data[part] = sheet_xml.encode("utf-8")

    data.update(media)

    ct = data["[Content_Types].xml"].decode("utf-8")
    if out.suffix.lower() == ".xlsm":
        # マクロ有効ブックとして開けるようにする（vbaProject.bin はまだ入っていない）
        ct = ct.replace(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
        )
    if 'Extension="png"' not in ct:
        ct = ct.replace(
            "<Types ", '<Types ', 1
        ).replace(
            ">", '><Default Extension="png" ContentType="image/png"/>', 1
        )
    overrides = "".join(
        f'<Override PartName="/xl/drawings/drawing{i}.xml" '
        f'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
        for i in range(1, len(form_sheets) + 1)
    )
    ct = ct.replace("</Types>", overrides + "</Types>")
    data["[Content_Types].xml"] = ct.encode("utf-8")

    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, payload in data.items():
            zout.writestr(name, payload)
    shutil.move(tmp, out)


def _as_date(value):
    """記入例の日付文字列を本物の日付にする。"""
    from rirekisho.model import parse_date, parse_year_month

    text = str(value)
    try:
        return parse_date(text)
    except Exception:
        try:
            year, month = parse_year_month(text)
            return dt.date(year, month, 1)
        except Exception:
            return value


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Excelだけで使える履歴書ブックを作る")
    parser.add_argument("--official", required=True, help="公式様式のExcelファイル")
    parser.add_argument("-o", "--output", default="履歴書作成.xlsx")
    parser.add_argument("--sheets", type=int, default=STUDENTS, help="履歴書シートの枚数")
    parser.add_argument("--with-sample", action="store_true", help="記入例を入れる")
    args = parser.parse_args(argv)

    students = None
    if args.with_sample:
        from make_xlsx import SAMPLE_STUDENTS

        students = [dict(s) for s in SAMPLE_STUDENTS]
        for s in students:          # 住所の読みは郵便番号から自動で入るので消しておく
            s.pop("address_kana", None)
            s.pop("contact_kana", None)
        for s in students:
            for key, value in list(s.items()):
                if key == "birth" or key.endswith(".ym") or key.endswith("job.1.ym"):
                    s[key] = _as_date(value)
        for s, course in zip(students, (COURSE_LIST[2], COURSE_LIST[0], COURSE_LIST[4])):
            s["course"] = course

    paste = None
    if args.with_sample:
        paste = [
            "3-2-1\t佐野 太郎\t計算技術検定3級\t令和6年11月15日",
            "3-2-2 近畿 花子 実用英語検定2級 2025/6/8",
            "3-2-3　泉州 一郎\t危険物取扱者乙4\t令和7年3月14日",
            "3-2-2 近畿 花子 技能検定　機械加工（普通旋盤作業）3級 令和7年7月13日",
        ]
    out = build(Path(args.official), Path(args.output), students, args.sheets, paste)
    print(f"作成しました: {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
