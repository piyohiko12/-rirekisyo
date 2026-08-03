"""入力シート生成 → 読み取り → PDF出力 まで一通り動くことを確認する。"""

import csv

import fitz
import pytest

from build_pdf import build
from make_xlsx import SAMPLE
from rirekisho import DEFAULT_TEMPLATE
from rirekisho.inputs import KEY_COL_B, KEY_COL_C, read_values, write_template


def pdf_text(path) -> str:
    """PDFの文字を取り出す。フォントの都合で空白やハイフンが別の文字コードに
    なることがあるので、比較しやすいようにそろえる。"""
    return fitz.open(path)[0].get_text().replace("\xa0", " ").replace("\xad", "-")


@pytest.fixture()
def sheet(tmp_path):
    path = tmp_path / "入力シート.xlsx"
    write_template(path, defaults=SAMPLE)
    return path


def test_template_roundtrip(sheet):
    values = read_values(sheet)
    assert values["name"] == "佐野 太郎"
    assert values["license.1.name"] == "第二種電気工事士"
    assert values["job.1.ym"] is None


def test_build_pdf_contains_input_text(sheet, tmp_path):
    out, warnings = build(sheet, tmp_path / "履歴書.pdf", template=DEFAULT_TEMPLATE)
    assert out.exists()
    assert warnings == []
    text = pdf_text(out)
    for expected in ["佐野 太郎", "さの たろう", "598-0001", "同上", "第二種電気工事士", "18"]:
        assert expected in text, expected
    # 用紙にもともと印刷されている内容が消えていないこと
    assert "大阪府立佐野工科高等学校" in text


def test_build_from_csv(sheet, tmp_path):
    """Googleスプレッドシートから書き出したCSVでも同じように作れる。"""
    from openpyxl import load_workbook

    ws = load_workbook(sheet).worksheets[0]
    csv_path = tmp_path / "sheet.csv"
    with csv_path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.writer(fh)
        for row in ws.iter_rows(values_only=False):
            cells = {c.column: c.value for c in row}
            writer.writerow(
                [
                    cells.get(1) or "",
                    cells.get(2) or "",
                    cells.get(3) or "",
                    cells.get(4) or "",
                    cells.get(KEY_COL_B) or "",
                    cells.get(KEY_COL_C) or "",
                ]
            )

    out, _ = build(csv_path, tmp_path / "csv.pdf", template=DEFAULT_TEMPLATE)
    assert "佐野 太郎" in pdf_text(out)


def test_empty_sheet_still_builds(tmp_path):
    path = tmp_path / "空.xlsx"
    write_template(path)
    out, warnings = build(path, tmp_path / "空.pdf", template=DEFAULT_TEMPLATE)
    assert out.exists()
    assert any("名前" in w for w in warnings)
    assert "同上" in pdf_text(out)
