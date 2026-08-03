"""資格の一括取り込みと、資格名マスタの確認。"""

import datetime as dt

import pytest

from build_pdf import build_all
from make_xlsx import SAMPLE_STUDENTS
from rirekisho import DEFAULT_TEMPLATE
from rirekisho.inputs import DEFAULT_MASTER, read_students, write_template
from rirekisho.licenses import (
    STATUS_APPLIED,
    STATUS_DUPLICATE,
    STATUS_ERROR,
    build_master,
    import_licenses,
    official_name,
    parse_line,
    write_log,
)
from rirekisho.model import build_resume

MASTER = build_master(DEFAULT_MASTER)


@pytest.fixture()
def students(tmp_path):
    path = tmp_path / "入力シート.xlsx"
    write_template(path, students=SAMPLE_STUDENTS)
    return read_students(path).students


@pytest.mark.parametrize(
    "raw,expected",
    [
        ("実用英語検定2級", "実用英語検定2級"),
        ("実用英語検定 準2級", "実用英語検定準２級"),   # 空白と全角半角をそろえる
        ("ＱＣ検定3級", "QC検定3級"),
        ("アーク溶接安全衛生教育修了", "ア－ク溶接安全衛生教育修了"),  # 長音記号のゆれも吸収
        ("アーク溶接技能者適格性証明書A-2F", "アーク溶接技能者適格性証明書Ａ－２Ｆ"),
        ("だれも知らない検定", "だれも知らない検定"),  # マスタに無ければそのまま
    ],
)
def test_official_name(raw, expected):
    assert official_name(raw, MASTER) == expected


def test_master_has_no_conflicting_keys():
    """変換表のキーがぶつかっていない（同じ名前とみなされる行が無い）こと。"""
    assert len(MASTER) == len(DEFAULT_MASTER)


def test_master_applies_to_hand_typed_licenses():
    resume = build_resume(
        {"license.1.name": "日本漢字能力検定 2級", "license.1.ym": "2025-06"}, master=MASTER
    )
    assert resume.licenses[0].name == "日本漢字能力検定2級"


def test_parse_line(students):
    row = parse_line("3-2-15\t山田太郎\t基礎製図検定\t令和6年7月10日", students=students, master=MASTER)
    assert (row.grade, row.class_no, row.number) == ("3", "2", "15")
    assert row.src_name == "山田太郎"
    assert row.license_name == "基礎製図検定"
    assert row.date == dt.date(2024, 7, 10)


def test_parse_line_handles_name_with_space(students):
    """氏名に空白が入っていても、名簿と一致する切り方を選ぶ。"""
    row = parse_line("3-2-1 佐野 太郎 計算技術検定3級 2024/11/15", students=students, master=MASTER)
    assert row.src_name == "佐野 太郎"
    assert row.license_name == "計算技術検定3級"


def test_import_matches_by_number_and_name(students):
    report = import_licenses(
        [
            "3-2-1 佐野太郎 計算技術検定3級 令和6年11月15日",  # 番号で照合
            "近畿花子 実用英語検定2級 2025/6/8",                # 氏名だけで照合
        ],
        students,
        master=MASTER,
    )
    assert [r.status for r in report.rows] == [STATUS_APPLIED, STATUS_APPLIED]
    hanako = build_resume(students[1].values, master=MASTER)
    assert "実用英語検定2級" in [lic.name for lic in hanako.licenses]


def test_import_skips_duplicates_and_reports_problems(students):
    report = import_licenses(
        [
            "",  # 空行は数えない
            "3-2-1 佐野太郎 電気工事士第二種 2024-06-20",   # 手入力ずみ
            "3-2-9 いない生徒 日本漢字能力検定3級 令和6年10月1日",  # 名簿にいない
            "3-2-2 近畿花子 基礎製図検定",                   # 取得日なし
        ],
        students,
        master=MASTER,
    )
    assert [r.status for r in report.rows] == [STATUS_DUPLICATE, STATUS_ERROR, STATUS_ERROR]
    assert report.applied == 0 and report.duplicated == 1 and report.errors == 2
    assert "すでに入力済み" in report.rows[0].note
    assert "照合できません" in report.rows[1].note
    assert "取得日" in report.rows[2].note


def test_imported_licenses_are_sorted_by_date(students):
    import_licenses(["3-2-1 佐野太郎 計算技術検定3級 令和6年11月15日"], students, master=MASTER)
    resume = build_resume(students[0].values, master=MASTER)
    ym = [(lic.year, lic.month) for lic in resume.licenses]
    assert ym == sorted(ym)
    assert [lic.name for lic in resume.licenses][:2] == ["電気工事士第二種", "計算技術検定3級"]


def test_input_order_can_be_kept(students):
    resume = build_resume(students[0].values, master=MASTER, license_order="入力順")
    assert [lic.ym_text for lic in resume.licenses] == ["令和6年6月", "令和7年3月", "令和7年11月"]


def test_paste_sheet_is_imported_during_build(tmp_path):
    path = tmp_path / "入力シート.xlsx"
    write_template(
        path,
        students=SAMPLE_STUDENTS,
        paste_lines=["3-2-3 泉州一郎 危険物取扱者乙4 令和7年3月14日"],
    )
    results = build_all(path, tmp_path / "出力", template=DEFAULT_TEMPLATE)
    text = fitz_text(results[2].path)
    assert "危険物取扱者乙4" in text
    assert (tmp_path / "出力" / "資格取込ログ.csv").exists()


def fitz_text(path) -> str:
    import fitz

    return fitz.open(path)[0].get_text().replace("\xa0", " ")


def test_write_log(students, tmp_path):
    report = import_licenses(["3-2-1 佐野太郎 計算技術検定3級 令和6年11月15日"], students, master=MASTER)
    log = tmp_path / "ログ.csv"
    write_log(report, log)
    body = log.read_text(encoding="utf-8-sig")
    assert "反映" in body and "計算技術検定3級" in body
