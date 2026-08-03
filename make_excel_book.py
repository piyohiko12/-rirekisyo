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
PASTE_ROWS = 200         # 貼り付けられる件数
IMPORT_SLOTS = 8         # 1人が取り込める件数
GUIDE = "使い方"
ZIPCODES = "郵便番号"
FORM_SHEET = "履歴書"
BLOCK_ROWS = 91          # 履歴書1人分の行数（1ページ）
ZIP_PREFECTURES = ("大阪府", "和歌山県", "奈良県")

ROSTER_FIRST_ROW = 5      # 生徒1人目の行
STUDENTS = 40             # 履歴書シートの枚数
LICENSE_SLOTS = 6         # 「入力」シートの資格の枠数
GATHER_SLOTS = 14         # 資格集約の1人分の行数（手入力6＋取込8）
LICENSE_ROWS_ON_FORM = 10  # 用紙の資格欄に出す行数（欄からあふれる分は印字しない）
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
# 資格欄は用紙の行9〜29。用紙の行の高さ（6.75pt）は変えずに2行ずつ使う
# （1件あたり13.5pt）。10件を超える分は欄に入らないので印字しない。
LICENSE_ROW_BANDS = [(9 + i * 2, 10 + i * 2) for i in range(LICENSE_ROWS_ON_FORM)]
LICENSE_FONT_SIZE = 11.0
BODY_FONT_SIZE = 11.0        # 校内外の諸活動・志望の動機・備考

# 用紙の行は左側の欄（氏名・生年月日・現住所）と共有しているため、
# 資格欄だけ行の高さを変えると様式全体が崩れる。行の高さは変えない。
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
def roster_col(key: str) -> str:
    """「入力」シートで、そのキーの列文字を返す。"""
    for i, col in enumerate(columns(), start=1):
        if col.key == key:
            return get_column_letter(i)
    raise KeyError(key)


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
def build_roster(wb, students: list[dict] | None) -> None:
    """「入力」シート（1行＝1生徒の名簿）。"""
    ws = wb.create_sheet(ROSTER, 1)
    cols = columns()
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    input_fill = PatternFill("solid", fgColor="FFFDE7")
    meta_fill = PatternFill("solid", fgColor="EFEFEF")
    header_fill = PatternFill("solid", fgColor="E8EEF4")
    group_fill = PatternFill("solid", fgColor="D6E2EF")
    small = Font(size=9, color="666666")

    ws.cell(row=1, column=1, value="入力（1行＝1生徒・名列順）").font = Font(size=13, bold=True)
    ws.cell(
        row=1, column=4,
        value="黄色いセルに入力すると、「履歴書」シートに自動で反映されます（1人＝1ページ・Noの順）。"
              "印刷は「設定」シートで No.○ 〜 No.○ を指定してください。",
    ).font = small

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
            cell.fill = meta_fill if col.key.startswith("meta.") else input_fill
            if col.kind == "date":
                cell.number_format = "yyyy/mm/dd"
        course_dv.add(ws.cell(row=row, column=cols.index(next(c for c in cols if c.key == "course")) + 1))
        ws.row_dimensions[row].height = 22

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


def build_settings(wb) -> None:
    ws = wb.create_sheet(SETTINGS)
    input_fill = PatternFill("solid", fgColor="FFFDE7")
    small = Font(size=9, color="666666")
    ws["A1"] = "設定（全員に共通）"
    ws["A1"].font = Font(size=12, bold=True)
    rows = [
        ("基準日（満○歳の計算日・用紙の「令和　年　月　日現在」）", dt.date(2026, 9, 1),
         "用紙の日付欄にもこの日付が入ります"),
        ("学校名", "大阪府立佐野工科高等学校", "履歴書の在籍校欄に出ます"),
        ("既定の学科", COURSE_LIST[2], "「入力」の学科が空欄のとき、この学科になります"),
        ("卒業（見込）年（令和）", 9, "在籍校欄の「令和　年」"),
        ("卒業（見込）月", 3, "在籍校欄の「　月」"),
    ]
    for i, (label, value, note) in enumerate(rows, start=3):
        ws.cell(row=i, column=1, value=label)
        cell = ws.cell(row=i, column=2, value=value)
        cell.fill = input_fill
        if i == 3:
            cell.number_format = "yyyy/mm/dd"
        ws.cell(row=i, column=3, value=note).font = small
    dv = DataValidation(type="list", formula1=f"={COURSES}!$A$3:$A$8", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(ws["B5"])

    # ---- 印刷する範囲（No.○ から No.○ まで）
    ws["A8"] = "印刷する範囲"
    ws["A8"].font = Font(size=12, bold=True)
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
    button = ws["A12"]
    button.value = "▶ 印刷する（クリック → 履歴書シートへ移動 → Ctrl+P）"
    button.font = Font(size=14, bold=True, color="FFFFFF")
    button.fill = PatternFill("solid", fgColor="2C6FBB")
    button.alignment = Alignment(horizontal="center", vertical="center")
    button.hyperlink = Hyperlink(ref="A12", location=f"{FORM_SHEET}!A1", display="印刷する")
    for row in ws["A12:C13"]:
        for cell in row:
            cell.border = Border(*(Side(style="medium", color="1F4E79"),) * 4)
    ws["A14"] = (
        "上の番号を入れてからボタンを押すと履歴書シートへ移動します。"
        "そのまま Ctrl+P（ファイル → 印刷／PDFで保存）を押すと、指定した範囲だけが出ます。"
    )
    ws["A14"].font = small
    ws["A15"] = (
        "※ 範囲がうまく反映されないときは、印刷画面の「ページ指定」に同じ番号を入れてください"
        "（1ページ＝生徒1人・ページ番号＝No）。"
    )
    ws["A15"].font = small
    ws.row_dimensions[12].height = 22
    ws.row_dimensions[13].height = 22

    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 46


def build_fields(wb) -> None:
    ws = wb.create_sheet(FIELDS)
    input_fill = PatternFill("solid", fgColor="FFFDE7")
    ws["A1"] = "反映項目（履歴書に出す項目を選びます）"
    ws["A1"].font = Font(size=12, bold=True)
    ws["A2"] = "B列を「×」にすると、入力してあってもその欄は空欄のまま印刷されます。"
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

    ws["A1"] = "資格取込（1行＝1件で貼り付けてください）"
    ws["A1"].font = Font(size=12, bold=True)
    ws["A2"] = (
        f"A{first} 以降に、資格取得の一覧をそのまま貼り付けます。"
        "「入力」シートの名簿と照合して、各生徒の資格欄に自動で追加します。"
    )
    ws["A2"].font = small
    ws["A3"] = (
        "書式: 3-2-15〔空白またはタブ〕山田 太郎 基礎製図検定 令和6年7月10日"
        "　…学年-組-出席番号 → 氏名 → 資格名 → 取得日 の順。"
    )
    ws["A3"].font = small
    ws["A4"] = (
        "先頭の番号がないときは氏名で照合します。取得日は 令和6年7月10日 / 2024/7/10 のどちらでも。"
        "すでに「入力」シートに手入力してある資格と同じものは「重複」として飛ばします。"
    )
    ws["A4"].font = small

    headers = {
        1: "貼付原文", 2: "整形", 3: "語数",
        4: "語1", 5: "語2", 6: "語3", 7: "語4", 8: "語5", 9: "語6", 10: "語7", 11: "語8",
        12: "ID", 13: "ID有", 14: "組", 15: "出席番号", 16: "氏名",
        17: "資格名（貼付）", 18: "正式名称", 19: "取得日",
        20: "名簿No", 21: "名簿の氏名", 22: "状態", 23: "有効", 24: "行", 25: "順位", 26: "キー",
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
            f'VALUE(SUBSTITUTE(MID({tail},FIND("月",{tail})+1,10),"日","")))'
        )
        seireki = f'DATE(VALUE({tok[0]}),VALUE({tok[1]}),VALUE({tok[2]}))'
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
        ws.cell(row=row, column=19).number_format = "yyyy/mm/dd"

    for col in (2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 23, 24, 25, 26, 27, 28, 29, 30):
        ws.column_dimensions[get_column_letter(col)].hidden = True
    for col, width in ((1, 52), (14, 6), (15, 9), (16, 14), (17, 28), (18, 28), (19, 12),
                       (20, 8), (21, 14), (22, 26)):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = f"A{first}"


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


def build_guide(wb) -> None:
    ws = wb.create_sheet(GUIDE, 0)
    lines = [
        "■ このファイルの使い方（Excelだけで完結します。マクロもPythonも使いません）",
        "",
        "1.「入力」シートに、名列順で生徒の情報を入力します（黄色いセル）。",
        "   ・学科は6種類からドロップダウンで選べます（空欄なら「設定」の既定の学科）。",
        "   ・郵便番号を入れると、ふりがな（住所）が自動で入ります（大阪府・和歌山県・奈良県）。",
        "     連絡先のふりがなも、連絡先の郵便番号から同じように入ります。",
        "     読みを直したいときは、そのセルに上から書き込んでください（そのセルだけ自動が外れます）。",
        "     郵便番号では町名までしか分からないので、番地の読みが要るときは書き足してください。",
        "   ・連絡先を空欄にすると、履歴書には自動で「同上」と入ります。",
        "   ・生年月日は「2008/5/12」のように日付で入力してください（元号と満○歳は自動）。",
        "",
        "2.「履歴書」シートが、そのまま印刷する用紙です（1人＝1ページ・上から入力シートのNo順）。",
        "   画面をスクロールすれば、印刷前に全員分を確認できます。",
        "   ・資格は取得年月の古い順に10件まで、11ポイントで印字します。",
        "     （用紙の欄の高さで決まっているため、11件以上は印字できません）",
        "",
        "3. 印刷・PDFにする（No.○ から No.○ まで）",
        "   ①「設定」シートの【開始No】【終了No】に番号を入れます（1人だけなら同じ番号）。",
        "   ②【▶ 印刷する】ボタンを押すと「履歴書」シートに移動します。",
        "   ③ そのまま Ctrl+P（ファイル → 印刷）→ 指定した範囲だけが印刷されます。",
        "      PDFにするときは、印刷画面のプリンターで「Microsoft Print to PDF」を選ぶか、",
        "      ファイル → 名前を付けて保存 → ファイルの種類で「PDF」を選びます。",
        "   ※ 範囲がうまく効かないときは、印刷画面の「ページ指定」に同じ番号を入れてください。",
        "     1ページ＝生徒1人なので、ページ番号＝入力シートのNo です。",
        "",
        "■ 資格をまとめて取り込む（貼り付けるだけ）",
        "「資格取込」シートのA6以降に、資格取得の一覧を1行1件で貼り付けてください。",
        "   例) 3-2-15〔タブ〕山田 太郎 基礎製図検定 令和6年7月10日",
        "   ・氏名は「山田 太郎」のように姓と名が離れていてもかまいません。",
        "   ・学年-組-出席番号 → 氏名 → 資格名 → 取得日 の順（区切りは空白でもタブでも可）",
        "   ・先頭の番号がないときは氏名で照合します。",
        "   ・「状態」の列に 反映／重複／要確認 が出るので、要確認の行だけ直してください。",
        "   ・手入力した資格と合わせて、取得年月の古い順に並べて印字します（最大10件）。",
        "",
        "■ 各シートの役割",
        "・入力　　　… 生徒の情報（1行＝1生徒）",
        "・履歴書　　… 印刷する用紙（1人＝1ページ）",
        "・設定　　　… 印刷する範囲、基準日、学校名、既定の学科、卒業（見込）年月",
        "・反映項目　… 履歴書に出す項目を○×で選ぶ",
        "・学科マスタ… 在籍校欄に出る学科（6種類）",
        "・資格取込　… 資格一覧を貼り付けると、名簿と照合して自動で振り分けます",
        "・資格マスタ… 入力した資格名を正式名称に直す変換表",
        "・資格集約／計算／郵便番号 … 自動計算用（非表示・さわらないでください）",
        "",
        "■ 自動で入るもの",
        "・満○歳　　… 生年月日と「設定」の基準日から計算します。",
        "・元号の年　… 生年月日から自動で計算します（昭和・平成の丸は手で付けてください）。",
        "・郵便番号　… 7桁の数字だけでも 123-4567 の形にします。",
        "・住所のふりがな … 郵便番号から自動で入ります（手で入れた場合はそちらが優先）。",
        "・連絡先　　… 空欄なら「同上」。",
        "・資格　　　… 取得年月の古い順に並べ、資格マスタの正式名称で印字します（最大10件）。",
        "・在籍校　　… 「設定」の学校名と、生徒ごとの学科を組み合わせます。",
        "",
        "■ 注意",
        "・写真は印刷した用紙に貼ってください。",
        "・職歴の「平成／令和」の丸は、印刷後に手で付けてください。",
        "・非表示のシートは変えないでください。",
    ]
    for i, line in enumerate(lines, start=1):
        cell = ws.cell(row=i, column=1, value=line)
        if line.startswith("■"):
            cell.font = Font(bold=True, size=11)
    ws.column_dimensions["A"].width = 96
    ws.sheet_view.showGridLines = False


# ------------------------------------------------------------------ 履歴書シート
def shift(ref: str, offset: int) -> str:
    """セル範囲の行番号をずらす（例: L11:AT14 → L102:AT105）。"""
    return re.sub(r"([A-Z]+)(\d+)", lambda m: f"{m.group(1)}{int(m.group(2)) + offset}", ref)


def fill_form(ws, i: int, offset: int = 0, *, bands=None, license_size=None) -> None:
    """履歴書シートの記入欄に、「計算」シートを参照する数式を入れる。

    bands / license_size で、資格欄の行の割り付けと文字サイズを変えられる。
    """
    bands = bands or LICENSE_ROW_BANDS
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
    set_cell(ws, shift(GRAD_YEAR_CELL, offset), f'={SETTINGS}!$B$6&""',
             align="center", size=WIDE_NUM_SIZE)
    set_cell(ws, shift(GRAD_MONTH_CELL, offset), f'={SETTINGS}!$B$7&""',
             align="center", size=NARROW_NUM_SIZE)

    base = f"{SETTINGS}!$B$3"
    set_cell(ws, shift(TODAY_YEAR_CELL, offset), f'=IF(N({base})=0,"",YEAR({base})-2018)&""',
             align="center", size=WIDE_NUM_SIZE)
    set_cell(ws, shift(TODAY_MONTH_CELL, offset), f'=IF(N({base})=0,"",MONTH({base}))&""',
             align="center", size=WIDE_NUM_SIZE)
    set_cell(ws, shift(TODAY_DAY_CELL, offset), f'=IF(N({base})=0,"",DAY({base}))&""',
             align="center", size=NARROW_NUM_SIZE)

    for k, (row_top, row_bottom) in enumerate(bands, start=1):
        ym_col = get_column_letter(CALC_LICENSE_COL + (k - 1) * 2)
        name_col = get_column_letter(CALC_LICENSE_COL + (k - 1) * 2 + 1)
        set_cell(ws, shift(f"{LICENSE_YM_COLS[0]}{row_top}:{LICENSE_YM_COLS[1]}{row_bottom}", offset),
                 calc(ym_col), size=license_size, align="center")
        set_cell(ws, shift(f"{LICENSE_NAME_COLS[0]}{row_top}:{LICENSE_NAME_COLS[1]}{row_bottom}", offset),
                 calc(name_col), size=license_size, indent=1, shrink=True)

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


def build_form_sheet(wb, src, title: str = FORM_SHEET, *, bands=None, license_size=None) -> None:
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
            ws.row_dimensions[r + offset].height = (
                dim.height if dim is not None and dim.height else default_height
            )
        for rng in merges:
            ws.merge_cells(shift(rng, offset))
        fill_form(ws, i, offset, bands=bands, license_size=license_size)
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
    wareki = (
        f'DATE(VALUE(MID({ref},3,FIND("年",{ref})-3))'
        f'+IF({era}="令和",2018,IF({era}="平成",1988,IF({era}="昭和",1925,IF({era}="大正",1911,0)))),'
        f'VALUE(MID({ref},FIND("年",{ref})+1,FIND("月",{ref})-FIND("年",{ref})-1)),'
        f'VALUE(SUBSTITUTE(MID({ref},FIND("月",{ref})+1,10),"日","")))'
    )
    # 2008/5/12 や 2008-5-12、2008年5月12日 のような西暦の文字列
    norm = (f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({ref},'
            f'"年","/"),"月","/"),"日",""),"-","/"),".","/")')
    tok = [f'TRIM(MID(SUBSTITUTE({norm},"/",REPT(" ",50)),{k * 50 + 1},50))' for k in range(3)]
    seireki = f'DATE(VALUE({tok[0]}),VALUE({tok[1]}),VALUE({tok[2]}))'
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
    build_guide(wb)

    order = [GUIDE, ROSTER, PASTE, SETTINGS, FIELDS, COURSES, MASTER,
             FORM_SHEET, GATHER, CALC, ZIPCODES]
    wb._sheets.sort(key=lambda ws: order.index(ws.title) if ws.title in order else 99)
    wb.active = 0
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    _attach_drawings(official, out, [FORM_SHEET])
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
