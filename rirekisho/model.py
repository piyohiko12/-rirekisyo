"""入力値の正規化・自動計算（和暦変換、満年齢、連絡先の「同上」など）。"""

from __future__ import annotations

import datetime as dt
import re
import unicodedata
from dataclasses import dataclass, field

# 元号の開始日（この日以降がその元号）
ERAS = [
    ("令和", dt.date(2019, 5, 1), 2018),
    ("平成", dt.date(1989, 1, 8), 1988),
    ("昭和", dt.date(1926, 12, 25), 1925),
    ("大正", dt.date(1912, 7, 30), 1911),
]
ERA_ALIASES = {
    "R": "令和", "r": "令和", "令": "令和",
    "H": "平成", "h": "平成", "平": "平成",
    "S": "昭和", "s": "昭和", "昭": "昭和",
    "T": "大正", "t": "大正", "大": "大正",
}

SAME_AS_ABOVE = "同上"


class InputError(ValueError):
    """入力シートの値が解釈できないときに送出する。"""


def normalize(value) -> str:
    """全角英数字などをそろえて前後の空白を落とした文字列にする。"""
    if value is None:
        return ""
    if isinstance(value, str):
        text = value
    elif isinstance(value, float) and value.is_integer():
        text = str(int(value))
    else:
        text = str(value)
    text = unicodedata.normalize("NFKC", text)
    return text.strip()


def to_wareki(date: dt.date) -> tuple[str, int]:
    """西暦の日付を (元号, 和暦年) にする。元年は 1 を返す。"""
    for name, start, offset in ERAS:
        if date >= start:
            return name, date.year - offset
    raise InputError(f"対応していない年です: {date}")


def from_wareki(era: str, year: int, month: int, day: int) -> dt.date:
    era = ERA_ALIASES.get(era, era)
    for name, _start, offset in ERAS:
        if name == era:
            return dt.date(offset + year, month, day)
    raise InputError(f"元号を認識できません: {era}")


_WAREKI_RE = re.compile(
    r"^(令和|平成|昭和|大正|[RHSTrhst令平昭大])\s*(元|\d{1,2})\s*[年.\-/]\s*"
    r"(\d{1,2})\s*[月.\-/]\s*(\d{1,2})\s*日?$"
)
_SEIREKI_RE = re.compile(r"^(\d{4})\s*[年.\-/]\s*(\d{1,2})\s*[月.\-/]\s*(\d{1,2})\s*日?$")


def parse_date(value) -> dt.date:
    """日付セルを date にする。西暦・和暦・Excelの日付型のいずれでも可。"""
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    text = normalize(value)
    if not text:
        raise InputError("日付が空です")
    m = _SEIREKI_RE.match(text)
    if m:
        y, mo, d = (int(x) for x in m.groups())
        return dt.date(y, mo, d)
    m = _WAREKI_RE.match(text)
    if m:
        era, y, mo, d = m.groups()
        year = 1 if y == "元" else int(y)
        return from_wareki(era, year, int(mo), int(d))
    if re.fullmatch(r"\d{8}", text):  # 20080512
        return dt.date(int(text[:4]), int(text[4:6]), int(text[6:]))
    raise InputError(f"日付として読み取れません: {value!r}（例: 2008-05-12 / 平成20年5月12日）")


_WAREKI_YM_RE = re.compile(
    r"^(令和|平成|昭和|大正|[RHSTrhst令平昭大])\s*(元|\d{1,2})\s*[年.\-/]\s*(\d{1,2})\s*月?$"
)
_SEIREKI_YM_RE = re.compile(r"^(\d{4})\s*[年.\-/]\s*(\d{1,2})\s*月?$")


def parse_year_month(value) -> tuple[int, int]:
    """「取得年月」などのセルを (西暦年, 月) にする。"""
    if isinstance(value, dt.datetime):
        return value.year, value.month
    if isinstance(value, dt.date):
        return value.year, value.month
    text = normalize(value)
    if not text:
        raise InputError("年月が空です")
    m = _SEIREKI_YM_RE.match(text)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = _WAREKI_YM_RE.match(text)
    if m:
        era, y, mo = m.groups()
        year = 1 if y == "元" else int(y)
        return from_wareki(era, year, int(mo), 1).year, int(mo)
    if re.fullmatch(r"\d{6}", text):  # 202406
        return int(text[:4]), int(text[4:])
    raise InputError(f"年月として読み取れません: {value!r}（例: 2024-06 / 令和6年6月）")


def format_wareki_ym(year: int, month: int) -> str:
    era, wy = to_wareki(dt.date(year, month, 1))
    return f"{era}{'元' if wy == 1 else wy}年{month}月"


def calc_age(birth: dt.date, on: dt.date) -> int:
    """満年齢。"""
    age = on.year - birth.year - ((on.month, on.day) < (birth.month, birth.day))
    if age < 0:
        raise InputError("生年月日が基準日より後になっています")
    return age


def format_zip(value) -> str:
    """郵便番号を 123-4567 の形にそろえる。"""
    text = normalize(value)
    if not text:
        return ""
    digits = re.sub(r"\D", "", text)
    if len(digits) == 7:
        return f"{digits[:3]}-{digits[3:]}"
    return text.lstrip("〒 ")


@dataclass
class License:
    """資格等の1行。"""

    year: int
    month: int
    name: str

    @property
    def ym_text(self) -> str:
        if not self.year:
            return ""
        return format_wareki_ym(self.year, self.month)


@dataclass
class Job:
    """職歴の1行。"""

    era: str
    year: int
    month: int
    text: str


@dataclass
class Resume:
    """PDFに書き込む内容。すべて描画用に整形済みの値。"""

    name: str = ""
    name_kana: str = ""
    birth: dt.date | None = None
    as_of: dt.date | None = None
    zip_code: str = ""
    address: str = ""
    address_kana: str = ""
    contact_zip: str = ""
    contact_address: str = ""
    contact_kana: str = ""
    licenses: list[License] = field(default_factory=list)
    activities: str = ""
    motivation: str = ""
    desired_job: str = ""
    appeal: str = ""
    remarks: str = ""
    jobs: list[Job] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    # ---- 自動計算 -----------------------------------------------------
    @property
    def birth_era(self) -> str:
        return to_wareki(self.birth)[0] if self.birth else ""

    @property
    def birth_wareki_year(self) -> str:
        if not self.birth:
            return ""
        wy = to_wareki(self.birth)[1]
        return "元" if wy == 1 else str(wy)

    @property
    def age(self) -> str:
        if not (self.birth and self.as_of):
            return ""
        return str(calc_age(self.birth, self.as_of))

    @property
    def contact_is_same(self) -> bool:
        """連絡先が未入力なら「同上」扱い。"""
        return not (self.contact_address or self.contact_zip)

    @property
    def contact_text(self) -> str:
        return SAME_AS_ABOVE if self.contact_is_same else self.contact_address


def build_resume(values: dict[str, object], *, as_of: dt.date | None = None) -> Resume:
    """入力シートから読んだ生の値（キー→値）を Resume に組み立てる。"""
    warnings: list[str] = []

    def text(key: str) -> str:
        return normalize(values.get(key))

    resume = Resume(warnings=warnings)
    resume.name = text("name")
    resume.name_kana = text("name_kana")
    resume.address = text("address")
    resume.address_kana = text("address_kana")
    resume.zip_code = format_zip(values.get("zip"))
    resume.contact_address = text("contact_address")
    resume.contact_kana = text("contact_kana")
    resume.contact_zip = format_zip(values.get("contact_zip"))
    resume.activities = text("activities")
    resume.motivation = text("motivation")
    resume.desired_job = text("desired_job")
    resume.appeal = text("appeal")
    resume.remarks = text("remarks")

    if not resume.name:
        warnings.append("名前が入力されていません。")

    birth_raw = values.get("birth")
    if normalize(birth_raw):
        resume.birth = parse_date(birth_raw)
    else:
        warnings.append("生年月日が入力されていません。")

    as_of_raw = values.get("as_of")
    if as_of is not None:
        resume.as_of = as_of
    elif normalize(as_of_raw):
        resume.as_of = parse_date(as_of_raw)
    else:
        resume.as_of = dt.date.today()

    if resume.contact_is_same and resume.contact_kana:
        warnings.append("連絡先が未入力のため「同上」にします（ふりがなは印字しません）。")

    for i in range(1, 21):
        ym_raw = values.get(f"license.{i}.ym")
        name_raw = normalize(values.get(f"license.{i}.name"))
        if not normalize(ym_raw) and not name_raw:
            continue
        if not name_raw:
            warnings.append(f"資格{i}: 名称が空のため取得年月だけになります。")
        if normalize(ym_raw):
            year, month = parse_year_month(ym_raw)
        else:
            year, month = 0, 0
            warnings.append(f"資格{i}: 取得年月が空です。")
        resume.licenses.append(License(year, month, name_raw))

    for i in range(1, 5):
        ym_raw = values.get(f"job.{i}.ym")
        job_text = normalize(values.get(f"job.{i}.text"))
        if not normalize(ym_raw) and not job_text:
            continue
        if normalize(ym_raw):
            year, month = parse_year_month(ym_raw)
            era, wy = to_wareki(dt.date(year, month, 1))
            resume.jobs.append(Job(era, wy, month, job_text))
        else:
            resume.jobs.append(Job("", 0, 0, job_text))

    return resume
