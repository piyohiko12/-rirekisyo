import datetime as dt

import pytest

from rirekisho.model import (
    InputError,
    build_resume,
    calc_age,
    format_wareki_ym,
    format_zip,
    parse_date,
    parse_year_month,
    to_wareki,
)


@pytest.mark.parametrize(
    "value,expected",
    [
        ("2008-05-12", dt.date(2008, 5, 12)),
        ("2008/5/12", dt.date(2008, 5, 12)),
        ("2008年5月12日", dt.date(2008, 5, 12)),
        ("平成20年5月12日", dt.date(2008, 5, 12)),
        ("Ｈ20.5.12", dt.date(2008, 5, 12)),  # 全角も可
        ("令和元年5月1日", dt.date(2019, 5, 1)),
        ("昭和60年11月3日", dt.date(1985, 11, 3)),
        ("20080512", dt.date(2008, 5, 12)),
        (dt.datetime(2008, 5, 12, 9, 0), dt.date(2008, 5, 12)),
    ],
)
def test_parse_date(value, expected):
    assert parse_date(value) == expected


def test_parse_date_rejects_garbage():
    with pytest.raises(InputError):
        parse_date("だいたい2008年ごろ")


@pytest.mark.parametrize(
    "date,expected",
    [
        (dt.date(2019, 4, 30), ("平成", 31)),
        (dt.date(2019, 5, 1), ("令和", 1)),
        (dt.date(1989, 1, 7), ("昭和", 64)),
        (dt.date(1989, 1, 8), ("平成", 1)),
    ],
)
def test_to_wareki_boundaries(date, expected):
    assert to_wareki(date) == expected


@pytest.mark.parametrize(
    "birth,on,expected",
    [
        (dt.date(2008, 5, 12), dt.date(2026, 9, 1), 18),
        (dt.date(2008, 9, 1), dt.date(2026, 9, 1), 18),   # 誕生日当日
        (dt.date(2008, 9, 2), dt.date(2026, 9, 1), 17),   # 誕生日前日
        (dt.date(2008, 2, 29), dt.date(2026, 2, 28), 17),  # うるう年生まれ
    ],
)
def test_calc_age(birth, on, expected):
    assert calc_age(birth, on) == expected


@pytest.mark.parametrize(
    "value,expected", [("2024-06", (2024, 6)), ("令和6年6月", (2024, 6)), ("202406", (2024, 6))]
)
def test_parse_year_month(value, expected):
    assert parse_year_month(value) == expected


def test_format_wareki_ym():
    assert format_wareki_ym(2024, 6) == "令和6年6月"
    assert format_wareki_ym(2019, 5) == "令和元年5月"


@pytest.mark.parametrize(
    "value,expected",
    [("5980001", "598-0001"), ("598-0001", "598-0001"), ("〒598-0001", "598-0001"), ("", "")],
)
def test_format_zip(value, expected):
    assert format_zip(value) == expected


def test_contact_defaults_to_same_as_above():
    resume = build_resume({"name": "佐野 太郎", "birth": "2008-05-12", "as_of": "2026-09-01"})
    assert resume.contact_is_same
    assert resume.contact_text == "同上"
    assert resume.age == "18"
    assert resume.birth_era == "平成"
    assert resume.birth_wareki_year == "20"


def test_contact_kept_when_filled():
    resume = build_resume({"contact_address": "大阪府堺市堺区南瓦町3-1", "contact_zip": "5900001"})
    assert not resume.contact_is_same
    assert resume.contact_text == "大阪府堺市堺区南瓦町3-1"
    assert resume.contact_zip == "590-0001"


def test_licenses_and_jobs_are_collected_in_order():
    resume = build_resume(
        {
            "license.1.ym": "2024-06",
            "license.1.name": "第二種電気工事士",
            "license.3.ym": "令和7年3月",
            "license.3.name": "危険物取扱者乙種第4類",
            "job.1.ym": "2019-04",
            "job.1.text": "株式会社サンプル製作所 入社",
        }
    )
    assert [lic.ym_text for lic in resume.licenses] == ["令和6年6月", "令和7年3月"]
    assert resume.jobs[0].era == "平成" and resume.jobs[0].year == 31
