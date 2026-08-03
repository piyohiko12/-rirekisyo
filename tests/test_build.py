"""入力シート生成 → 読み取り → 生徒ごとのPDF出力 まで一通り動くことを確認する。"""

import csv

import fitz
import pytest

from build_pdf import build_all, merge_pdfs
from make_xlsx import SAMPLE_STUDENTS
from rirekisho import DEFAULT_TEMPLATE
from rirekisho.inputs import read_students, write_template


def pdf_text(path) -> str:
    """PDFの文字を取り出す。フォントの都合で空白やハイフンが別の文字コードに
    なることがあるので、比較しやすいようにそろえる。"""
    return fitz.open(path)[0].get_text().replace("\xa0", " ").replace("\xad", "-")


@pytest.fixture()
def sheet(tmp_path):
    path = tmp_path / "入力シート.xlsx"
    write_template(path, students=SAMPLE_STUDENTS)
    return path


def test_roster_roundtrip(sheet):
    data = read_students(sheet)
    assert [s.name for s in data.students] == ["佐野 太郎", "近畿 花子", "泉州 一郎"]
    assert data.students[0].values["license.1.name"] == "電気工事士第二種"
    assert data.students[0].class_name == "3年2組"
    assert data.settings["as_of"] == "2026-09-01"
    # 空の行は読み飛ばす（名簿は40行ある）
    assert len(data.students) == 3


def test_build_all_writes_one_pdf_per_student(sheet, tmp_path):
    results = build_all(sheet, tmp_path / "出力", template=DEFAULT_TEMPLATE)
    assert [r.path.name for r in results] == [
        "3年2組_01_佐野太郎.pdf",
        "3年2組_02_近畿花子.pdf",
        "3年2組_03_泉州一郎.pdf",
    ]
    text = pdf_text(results[0].path)
    for expected in ["佐野 太郎", "さの たろう", "598-0001", "同上", "電気工事士第二種", "18"]:
        assert expected in text, expected
    assert "大阪府立佐野工科高等学校" in text  # 用紙の印字が消えていない
    # 連絡先を入力した生徒は「同上」にならない
    assert "同上" not in pdf_text(results[1].path)


def test_filter_by_number_and_name(sheet, tmp_path):
    only = build_all(sheet, tmp_path / "a", template=DEFAULT_TEMPLATE, only_no=2)
    assert [r.student.name for r in only] == ["近畿 花子"]
    by_name = build_all(sheet, tmp_path / "b", template=DEFAULT_TEMPLATE, only_name="泉州")
    assert [r.student.name for r in by_name] == ["泉州 一郎"]


def test_merge_makes_one_file_per_page(sheet, tmp_path):
    results = build_all(sheet, tmp_path / "出力", template=DEFAULT_TEMPLATE)
    merged = merge_pdfs(results, tmp_path / "まとめ.pdf")
    with fitz.open(merged) as doc:
        assert len(doc) == len(results)


def test_build_from_csv(sheet, tmp_path):
    """Googleスプレッドシートから書き出したCSVでも同じように作れる。"""
    from openpyxl import load_workbook

    ws = load_workbook(sheet)["入力"]
    csv_path = tmp_path / "sheet.csv"
    with csv_path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.writer(fh)
        for row in ws.iter_rows(values_only=True):
            writer.writerow(["" if c is None else c for c in row])

    results = build_all(csv_path, tmp_path / "csv", template=DEFAULT_TEMPLATE)
    assert "佐野 太郎" in pdf_text(results[0].path)


def test_key_row_is_required(tmp_path):
    from openpyxl import Workbook

    path = tmp_path / "壊れたシート.xlsx"
    wb = Workbook()
    wb.active.title = "入力"
    wb.active["A1"] = "No"
    wb.save(path)
    with pytest.raises(ValueError, match="キー行"):
        read_students(path)


def test_old_single_person_sheet_still_works(tmp_path):
    """以前の「1人1枚」形式のシートも読み込める。"""
    from openpyxl import Workbook

    path = tmp_path / "旧形式.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "入力"
    for i, (key, value) in enumerate(
        [("name", "旧 形式"), ("name_kana", "きゅう けいしき"), ("birth", "2008-05-12")], start=1
    ):
        ws.cell(row=i, column=2, value=value)
        ws.cell(row=i, column=5, value=key)
    wb.save(path)

    results = build_all(path, tmp_path / "旧", template=DEFAULT_TEMPLATE, as_of="2026-09-01")
    assert len(results) == 1
    assert "旧 形式" in pdf_text(results[0].path)


def test_empty_roster_produces_nothing(tmp_path):
    path = tmp_path / "空.xlsx"
    write_template(path)
    assert read_students(path).students == []
    assert build_all(path, tmp_path / "空", template=DEFAULT_TEMPLATE) == []


def test_key_row_is_hidden_and_marked(sheet):
    from openpyxl import load_workbook

    ws = load_workbook(sheet)["入力"]
    assert ws["A4"].value == "meta.no"
    assert ws.row_dimensions[4].hidden is True
    assert ws.freeze_panes == "E5"
