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
GUIDE = "使い方"

ROSTER_FIRST_ROW = 5      # 生徒1人目の行
STUDENTS = 40             # 履歴書シートの枚数
LICENSE_SLOTS = 6         # 「入力」シートの資格の枠数
LICENSE_ROWS_ON_FORM = 5  # 用紙の資格欄の行数
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
LICENSE_ROW_BANDS = [(9, 12), (13, 16), (17, 20), (21, 24), (25, 29)]
JOB_ROW_BANDS = [(64, 69), (70, 75), (76, 81), (82, 87)]
JOB_YEAR_COLS = ("P", "R")
JOB_MONTH_COLS = ("U", "V")
JOB_TEXT_COLS = ("Y", "BI")

FORM_FONT = "ＭＳ Ｐ明朝"


# ------------------------------------------------------------------ 便利関数
def roster_col(key: str) -> str:
    """「入力」シートで、そのキーの列文字を返す。"""
    for i, col in enumerate(columns(), start=1):
        if col.key == key:
            return get_column_letter(i)
    raise KeyError(key)


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
             wrap=False, shrink=False, indent=0):
    """結合セルに値（数式）と書式を設定する。"""
    first = ref.split(":")[0]
    if ":" in ref and ref not in {str(r) for r in ws.merged_cells.ranges}:
        try:
            ws.merge_cells(ref)
        except ValueError:
            pass
    cell = ws[first]
    cell.value = value
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
        value="黄色いセルに入力すると、履歴書シート（01〜40）に自動で反映されます。"
              "Noの数字が、そのままシート名になります。",
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
    for i, (src, dest) in enumerate(DEFAULT_MASTER, start=3):
        ws.cell(row=i, column=1, value=src).fill = input_fill
        ws.cell(row=i, column=2, value=dest).fill = input_fill
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 36
    ws.freeze_panes = "A3"


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

    for i in range(1, STUDENTS + 1):
        r = ROSTER_FIRST_ROW + i - 1
        top = GATHER_FIRST_ROW + (i - 1) * LICENSE_SLOTS
        bottom = top + LICENSE_SLOTS - 1
        for k in range(1, LICENSE_SLOTS + 1):
            row = top + k - 1
            ws.cell(row=row, column=1, value=i)
            ws.cell(row=row, column=2, value=k)
            ws.cell(row=row, column=3, value=(
                f'=IFERROR(IF(ISNUMBER({ROSTER}!{ym_cols[k-1]}{r}),{ROSTER}!{ym_cols[k-1]}{r},'
                f'DATEVALUE({ROSTER}!{ym_cols[k-1]}{r})),"")'))
            ws.cell(row=row, column=4, value=f'=IF({ROSTER}!{name_cols[k-1]}{r}="","",{ROSTER}!{name_cols[k-1]}{r})')
            ws.cell(row=row, column=5,
                    value=f'=IF(D{row}="","",IFERROR(VLOOKUP(D{row},{MASTER}!$A:$B,2,FALSE),D{row}))')
            ws.cell(row=row, column=6, value=f'=IF(E{row}="","",IF(C{row}="",DATE(9999,1,1),C{row}))')
            ws.cell(
                row=row, column=7,
                value=(f'=IF(E{row}="","",COUNTIFS($F${top}:$F${bottom},"<"&F{row},'
                       f'$F${top}:$F${bottom},"<>")+COUNTIFS($F${top}:$F${bottom},F{row},'
                       f'$B${top}:$B${bottom},"<"&B{row})+1)'),
            )
            ws.cell(row=row, column=8, value=(
                f'=IF(OR(E{row}="",C{row}=""),"",'
                f'IF(C{row}>=DATE(2019,5,1),"令和"&(YEAR(C{row})-2018),'
                f'IF(C{row}>=DATE(1989,1,8),"平成"&(YEAR(C{row})-1988),'
                f'"昭和"&(YEAR(C{row})-1925)))&"年"&MONTH(C{row})&"月")'
            ))
        ws.cell(row=top, column=3).number_format = "yyyy/mm/dd"
    for col, width in (("A", 8), ("B", 6), ("C", 12), ("D", 30), ("E", 30), ("F", 12), ("G", 8), ("H", 14)):
        ws.column_dimensions[col].width = width
    ws.sheet_state = "hidden"


def build_guide(wb) -> None:
    ws = wb.create_sheet(GUIDE, 0)
    lines = [
        "■ このファイルの使い方（Excelだけで完結します）",
        "",
        "1.「入力」シートに、名列順で生徒の情報を入力します（黄色いセル）。",
        "   ・学科は6種類からドロップダウンで選べます（空欄なら「設定」の既定の学科）。",
        "   ・連絡先を空欄にすると、履歴書には自動で「同上」と入ります。",
        "   ・生年月日は「2008/5/12」のように日付で入力してください（元号と満○歳は自動）。",
        "",
        "2.「01」〜「40」のシートが、そのまま履歴書になります（数字は入力シートのNo）。",
        "   入力すると自動で反映されるので、印刷前に見て確認してください。",
        "",
        "3. PDFにする（1人分）",
        "   その生徒のシートを開く → ファイル → 名前を付けて保存 → ファイルの種類で「PDF」を選ぶ",
        "   （または ファイル → エクスポート → PDF/XPS ドキュメントの作成）",
        "",
        "4. PDFにする（全員分・1つのファイル）",
        "   ファイル → エクスポート → PDF/XPS → 「オプション」→「ブック全体」→ 発行",
        "   ※ 必要な生徒のシートだけ出したいときは、シート見出しを Ctrl キーを押しながら選び、",
        "     オプションで「選択したシート」を選びます。",
        "",
        "5. 印刷するときは、そのままシートを印刷してください（A4横・1ページに収まります）。",
        "",
        "■ 各シートの役割",
        "・入力　　　… 生徒の情報（1行＝1生徒）",
        "・設定　　　… 基準日（満○歳の計算日）、学校名、既定の学科、卒業（見込）年月",
        "・反映項目　… 履歴書に出す項目を○×で選ぶ",
        "・学科マスタ… 在籍校欄に出る学科（6種類）",
        "・資格マスタ… 入力した資格名を正式名称に直す変換表",
        "・資格集約　… 資格を取得年月順に並べる計算用（非表示・さわらないでください）",
        "",
        "■ 自動で入るもの",
        "・満○歳　　… 生年月日と「設定」の基準日から計算します。",
        "・元号の年　… 生年月日から自動で計算します（昭和・平成の丸は手で付けてください）。",
        "・郵便番号　… 7桁の数字だけでも 123-4567 の形にします。",
        "・連絡先　　… 空欄なら「同上」。",
        "・資格　　　… 取得年月の古い順に並べ、資格マスタの正式名称で印字します（上から5件）。",
        "・在籍校　　… 「設定」の学校名と、生徒ごとの学科を組み合わせます。",
        "",
        "■ 注意",
        "・写真は印刷した用紙に貼ってください。",
        "・職歴の「平成／令和」の丸は、印刷後に手で付けてください。",
        "・シート名（01〜40）や、非表示のシートは変えないでください。",
    ]
    for i, line in enumerate(lines, start=1):
        cell = ws.cell(row=i, column=1, value=line)
        if line.startswith("■"):
            cell.font = Font(bold=True, size=11)
    ws.column_dimensions["A"].width = 96
    ws.sheet_view.showGridLines = False


# ------------------------------------------------------------------ 履歴書シート
def fill_form(ws, i: int) -> None:
    """履歴書シートの記入欄に、「計算」シートを参照する数式を入れる。"""
    row = CALC_FIRST_ROW + i - 1

    def calc(col: str) -> str:
        return f"={CALC}!${col}${row}"

    set_cell(ws, NAME_KANA_CELL, calc("C"), size=10, align="center", shrink=True)
    set_cell(ws, NAME_CELL, calc("B"), size=16, align="center", shrink=True)
    set_cell(ws, BIRTH_YEAR_CELL, calc("D"), align="center")
    set_cell(ws, BIRTH_MONTH_CELL, calc("E"), align="center")
    set_cell(ws, BIRTH_DAY_CELL, calc("F"), align="center")
    set_cell(ws, BIRTH_AGE_CELL, calc("G"), align="center")

    set_cell(ws, ADDR_ZIP_CELL, calc("H"), size=10)
    set_cell(ws, ADDR_CELL, calc("I"), size=10, wrap=True, indent=1)
    set_cell(ws, ADDR_KANA_CELL, calc("J"), size=9, indent=1)

    set_cell(ws, CONTACT_ZIP_CELL, calc("K"), size=10)
    set_cell(ws, CONTACT_CELL, calc("L"), size=10, align="center", wrap=True)
    set_cell(ws, CONTACT_KANA_CELL, calc("M"), size=9, indent=1)

    set_cell(ws, SCHOOL_CELL, calc("N"), size=10, wrap=True)
    set_cell(ws, GRAD_YEAR_CELL, f"={SETTINGS}!$B$6", align="center")
    set_cell(ws, GRAD_MONTH_CELL, f"={SETTINGS}!$B$7", align="center")

    base = f"{SETTINGS}!$B$3"
    set_cell(ws, TODAY_YEAR_CELL, f'=IF(N({base})=0,"",YEAR({base})-2018)', align="center")
    set_cell(ws, TODAY_MONTH_CELL, f'=IF(N({base})=0,"",MONTH({base}))', align="center")
    set_cell(ws, TODAY_DAY_CELL, f'=IF(N({base})=0,"",DAY({base}))', align="center")

    for k, (row_top, row_bottom) in enumerate(LICENSE_ROW_BANDS, start=1):
        ym_col = get_column_letter(15 + (k - 1) * 2)      # O,Q,S,U,W
        name_col = get_column_letter(16 + (k - 1) * 2)    # P,R,T,V,X
        set_cell(ws, f"{LICENSE_YM_COLS[0]}{row_top}:{LICENSE_YM_COLS[1]}{row_bottom}",
                 calc(ym_col), size=10, align="center")
        set_cell(ws, f"{LICENSE_NAME_COLS[0]}{row_top}:{LICENSE_NAME_COLS[1]}{row_bottom}",
                 calc(name_col), size=10, indent=1, shrink=True)

    set_cell(ws, ACTIVITIES_CELL, calc("Y"), size=10, valign="top", wrap=True, indent=1)
    set_cell(ws, MOTIVATION_CELL, calc("Z"), size=10, valign="top", wrap=True, indent=1)
    set_cell(ws, REMARKS_CELL, calc("AA"), size=10, valign="top", wrap=True, indent=1)

    for k, (row_top, row_bottom) in enumerate(JOB_ROW_BANDS[:2], start=1):
        year_col = get_column_letter(28 + (k - 1) * 3)    # AB, AE
        month_col = get_column_letter(29 + (k - 1) * 3)   # AC, AF
        text_col = get_column_letter(30 + (k - 1) * 3)    # AD, AG
        set_cell(ws, f"{JOB_YEAR_COLS[0]}{row_top+2}:{JOB_YEAR_COLS[1]}{row_top+3}",
                 calc(year_col), align="center")
        set_cell(ws, f"{JOB_MONTH_COLS[0]}{row_top+2}:{JOB_MONTH_COLS[1]}{row_top+3}",
                 calc(month_col), align="center")
        set_cell(ws, f"{JOB_TEXT_COLS[0]}{row_top}:{JOB_TEXT_COLS[1]}{row_bottom}",
                 calc(text_col), size=10, indent=1, wrap=True)


def build_calc(wb) -> None:
    """「計算」シート: 履歴書に出す文字を1行＝1生徒で組み立てる（非表示）。"""
    ws = wb.create_sheet(CALC)
    headers = [
        "No", "氏名", "ふりがな", "生年(和暦)", "生月", "生日", "満年齢",
        "郵便番号", "住所", "住所ふりがな", "連絡先〒", "連絡先", "連絡先ふりがな",
        "在籍校",
        "資格1年月", "資格1名称", "資格2年月", "資格2名称", "資格3年月", "資格3名称",
        "資格4年月", "資格4名称", "資格5年月", "資格5名称",
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
        top = GATHER_FIRST_ROW + (i - 1) * LICENSE_SLOTS
        bottom = top + LICENSE_SLOTS - 1

        def ref(key: str) -> str:
            return f"{ROSTER}!${roster_col(key)}${r}"

        birth = f"$AJ${row}"   # 変換済みの生年月日（右端の作業列）

        def put(col, value):
            """氏名が空の行（＝使っていない生徒）は、すべて空欄にする。"""
            if isinstance(value, str) and value.startswith("=") and col >= 2:
                value = f'=IF({ROSTER}!${roster_col("name")}${r}="","",{value[1:]})'
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
        put(10, guard("address_kana", ref("address_kana")))
        put(11, guard("contact", zip_fmt(ref("contact_zip")), f'{ref("contact_zip")}=""'))
        put(12, f'=IF({field_cell("contact")}="×","",'
                f'IF(AND({ref("contact_zip")}="",{ref("contact_address")}=""),"同上",'
                f'IF({ref("contact_address")}="","",{ref("contact_address")})))')
        put(13, f'=IF(OR({field_cell("contact")}="×",{ref("contact_kana")}="",'
                f'AND({ref("contact_zip")}="",{ref("contact_address")}="")),"",{ref("contact_kana")})')
        put(14, f'={SETTINGS}!$B$4&CHAR(10)&IF({ref("course")}="",{SETTINGS}!$B$5,{ref("course")})')

        for k in range(1, LICENSE_ROWS_ON_FORM + 1):
            rank = f'MATCH({k},{GATHER}!$G${top}:$G${bottom},0)'
            put(13 + k * 2, f'=IF({field_cell("licenses")}="×","",'
                            f'IFERROR(INDEX({GATHER}!$H${top}:$H${bottom},{rank}),""))')
            put(14 + k * 2, f'=IF({field_cell("licenses")}="×","",'
                            f'IFERROR(INDEX({GATHER}!$E${top}:$E${bottom},{rank}),""))')

        put(25, guard("activities", ref("activities")))
        put(26, f'=IF({field_cell("motivation")}="×","",'
                f'IF({ref("motivation")}="","",{ref("motivation")}&CHAR(10)))'
                f'&IF({field_cell("desired_job")}="×","",'
                f'IF({ref("desired_job")}="","",{ref("desired_job")}&CHAR(10)))'
                f'&IF({field_cell("appeal")}="×","",'
                f'IF({ref("appeal")}="","",{ref("appeal")}))')
        put(27, guard("remarks", ref("remarks")))

        for k in (1, 2):
            job_ym = f"${get_column_letter(36 + k)}${row}"   # AK, AL に変換済みの年月
            base_col = 28 + (k - 1) * 3
            put(base_col, guard("jobs",
                                f'YEAR({job_ym})-IF({job_ym}>=DATE(2019,5,1),2018,1988)',
                                f'N({job_ym})=0'))
            put(base_col + 1, guard("jobs", f"MONTH({job_ym})", f'N({job_ym})=0'))
            put(base_col + 2, guard("jobs", ref(f"job.{k}.text")))

        # 作業列（日付に変換したもの）
        put(36, f'=IFERROR(IF(ISNUMBER({ref("birth")}),{ref("birth")},DATEVALUE({ref("birth")})),"")')
        put(37, f'=IFERROR(IF(ISNUMBER({ref("job.1.ym")}),{ref("job.1.ym")},'
                f'DATEVALUE({ref("job.1.ym")})),"")')
        put(38, f'=IFERROR(IF(ISNUMBER({ref("job.2.ym")}),{ref("job.2.ym")},'
                f'DATEVALUE({ref("job.2.ym")})),"")')

    for col in ("AJ", "AK", "AL"):
        ws.column_dimensions[col].hidden = True
    ws.sheet_state = "hidden"


# ------------------------------------------------------------------ 組み立て
def build(official: Path, out: Path, students: list[dict] | None = None,
          sheets: int = STUDENTS) -> Path:
    wb = load_workbook(official)
    form = wb[FORM_SHEET_SRC]
    for name in list(wb.sheetnames):
        if name != FORM_SHEET_SRC:
            del wb[name]

    forms = [form]
    for _ in range(sheets - 1):
        forms.append(wb.copy_worksheet(form))
    for i, ws in enumerate(forms, start=1):
        ws.title = f"{i:02d}"
        ws._images = []   # 図はあとで元ファイルからそのまま入れ直す
        ws._charts = []
        ws.sheet_view.showGridLines = False
        ws.print_area = form.print_area
        fill_form(ws, i)

    build_roster(wb, students)
    build_settings(wb)
    build_fields(wb)
    build_courses(wb)
    build_master(wb)
    build_gather(wb)
    build_calc(wb)
    build_guide(wb)

    order = [GUIDE, ROSTER, SETTINGS, FIELDS, COURSES, MASTER, GATHER, CALC]
    wb._sheets.sort(key=lambda ws: order.index(ws.title) if ws.title in order else 100 + int(ws.title))
    wb.active = 0
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    _attach_drawings(official, out, [f"{i:02d}" for i in range(1, sheets + 1)])
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
        data[f"xl/drawings/{drawing_name}"] = drawing_xml
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
        for s in students:
            for key, value in list(s.items()):
                if key == "birth" or key.endswith(".ym") or key.endswith("job.1.ym"):
                    s[key] = _as_date(value)
        for s, course in zip(students, (COURSE_LIST[2], COURSE_LIST[0], COURSE_LIST[4])):
            s["course"] = course

    out = build(Path(args.official), Path(args.output), students, args.sheets)
    print(f"作成しました: {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
