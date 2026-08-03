"""資格の一括取り込みと、資格名の正式名称への変換（マスタ）。

「資格取込」シートに

    3-2-15	山田太郎	基礎製図検定	令和6年7月10日
    （学年-組-出席番号　氏名　資格名　取得日）

の形で1行1件を貼り付けると、名簿と照合して各生徒の資格欄に入れる。
資格名は「資格マスタ」シートの変換表で正式名称に直してから使う。
"""

from __future__ import annotations

import datetime as dt
import re
import unicodedata
from dataclasses import dataclass, field

from .inputs import Student
from .model import InputError, normalize, parse_date

# 貼り付け行の先頭にある「学年-組-出席番号」
_ID_PATTERNS = [
    re.compile(r"^(\d)\s*[-年]\s*(\d{1,2})\s*[-組]\s*(\d{1,2})\s*番?$"),  # 3-2-15 / 3年2組15番
    re.compile(r"^(\d)\s*[-]\s*(\d{1,2})\s*[-]\s*(\d{1,2})$"),
]
_NUMBER_ONLY = re.compile(r"^(\d{1,2})\s*番?$")

STATUS_APPLIED = "反映"
STATUS_DUPLICATE = "重複"
STATUS_ERROR = "要確認"


@dataclass
class ImportRow:
    """貼り付け1行の解析結果。"""

    raw: str
    grade: str = ""
    class_no: str = ""
    number: str = ""
    src_name: str = ""
    src_license: str = ""
    license_name: str = ""
    date: dt.date | None = None
    matched: Student | None = None
    status: str = STATUS_ERROR
    note: str = ""

    @property
    def matched_name(self) -> str:
        return self.matched.name if self.matched else ""


@dataclass
class ImportReport:
    rows: list[ImportRow] = field(default_factory=list)

    def count(self, status: str) -> int:
        return sum(1 for r in self.rows if r.status == status)

    @property
    def applied(self) -> int:
        return self.count(STATUS_APPLIED)

    @property
    def duplicated(self) -> int:
        return self.count(STATUS_DUPLICATE)

    @property
    def errors(self) -> int:
        return self.count(STATUS_ERROR)

    def summary(self) -> str:
        return (
            f"資格取込: {len(self.rows)}件 中 反映{self.applied} / "
            f"重複{self.duplicated} / 要確認{self.errors}"
        )


# ------------------------------------------------------------------ マスタ
def official_name(name: str, master: dict[str, str] | None) -> str:
    """資格名を正式名称に直す。マスタに無ければそのまま返す。"""
    text = normalize(name)
    if not text or not master:
        return text
    return master.get(_master_key(text), text)


def _master_key(text: str) -> str:
    """表記ゆれを吸収した照合用のキー。

    全角半角をそろえたうえで、空白・記号・長音記号（ー／－）を落とす。
    「ア－ク溶接…」と「アーク溶接…」、「Ａ－２Ｆ」と「A-2F」を同じものとして扱う。
    """
    text = unicodedata.normalize("NFKC", str(text)).lower()
    return re.sub(r"[\s　・（）()【】\[\]．.,、。/／ーｰ-]", "", text)


def build_master(pairs: list[tuple[object, object]]) -> dict[str, str]:
    """(変換前, 正式名称) の並びから変換表を作る。

    正式名称は **マスタに書かれたとおり** に使う（「色彩検定３級」のように
    全角で登録してあれば全角のまま印字する）。
    """
    master: dict[str, str] = {}
    for src, dest in pairs:
        src_text = normalize(src)
        dest_text = "" if dest is None else str(dest).strip()
        if src_text and dest_text:
            master[_master_key(src_text)] = dest_text
    return master


# ------------------------------------------------------------------ 解析
def _split_tokens(line: str) -> list[str]:
    return [t for t in re.split(r"[\s　\t,、]+", normalize(line)) if t]


def _parse_id(token: str) -> tuple[str, str, str] | None:
    for pattern in _ID_PATTERNS:
        m = pattern.match(token)
        if m:
            return m.group(1), m.group(2), m.group(3)
    m = _NUMBER_ONLY.match(token)
    if m:
        return "", "", m.group(1)
    return None


def _take_date(tokens: list[str]) -> tuple[list[str], dt.date | None]:
    """末尾から日付を取り出す（「令和6年 7月 10日」のように分かれていても拾う）。"""
    for take in (1, 2, 3):
        if len(tokens) <= take:
            break
        tail = "".join(tokens[-take:])
        try:
            return tokens[:-take], parse_date(tail)
        except InputError:
            continue
    return tokens, None


def parse_line(line: str, *, students: list[Student], master: dict[str, str] | None) -> ImportRow:
    row = ImportRow(raw=str(line).strip())
    tokens = _split_tokens(line)
    if not tokens:
        row.note = "空行"
        return row

    ident = _parse_id(tokens[0])
    if ident:
        row.grade, row.class_no, row.number = ident
        tokens = tokens[1:]

    tokens, row.date = _take_date(tokens)
    if not tokens:
        row.note = "氏名と資格名が読み取れません"
        return row

    # 氏名は空白入りのこともあるので、名簿と一致する切り方を優先して探す
    name_tokens = 1
    for take in range(1, min(3, len(tokens)) + 1):
        if _find_by_name("".join(tokens[:take]), students):
            name_tokens = take
            break
    row.src_name = " ".join(tokens[:name_tokens])
    row.src_license = " ".join(tokens[name_tokens:])
    row.license_name = official_name(row.src_license, master)

    if not row.src_license:
        row.note = "資格名がありません"
        return row
    if row.date is None:
        row.note = "取得日が読み取れません"
        return row
    if master and _master_key(row.src_license) not in master:
        row.note = "マスタ未登録（そのままの名称で反映）"
    return row


def _norm_name(name: str) -> str:
    return re.sub(r"[\s　]", "", normalize(name))


def _find_by_name(name: str, students: list[Student]) -> Student | None:
    target = _norm_name(name)
    if not target:
        return None
    for student in students:
        if _norm_name(student.name) == target:
            return student
    return None


def _class_parts(text: str) -> tuple[str, str]:
    """「3年2組」→ ("3","2")、「3-2」→ ("3","2")。"""
    digits = re.findall(r"\d+", normalize(text))
    if len(digits) >= 2:
        return digits[0], digits[1]
    if len(digits) == 1:
        return "", digits[0]
    return "", ""


def _find_by_number(row: ImportRow, students: list[Student]) -> Student | None:
    if not row.number:
        return None
    for student in students:
        grade, class_no = _class_parts(student.class_name)
        if student.number.lstrip("0") != row.number.lstrip("0"):
            continue
        if row.class_no and class_no and row.class_no.lstrip("0") != class_no.lstrip("0"):
            continue
        if row.grade and grade and row.grade != grade:
            continue
        return student
    return None


# ------------------------------------------------------------------ 反映
def _existing_licenses(student: Student) -> list[tuple[str, object]]:
    out = []
    for i in range(1, 41):
        name = student.values.get(f"license.{i}.name")
        if normalize(name):
            out.append((_master_key(normalize(name)), student.values.get(f"license.{i}.ym")))
    return out


def _next_slot(student: Student) -> int:
    used = [
        i for i in range(1, 41)
        if normalize(student.values.get(f"license.{i}.name"))
        or normalize(student.values.get(f"license.{i}.ym"))
    ]
    return (max(used) + 1) if used else 1


def import_licenses(
    lines: list[str],
    students: list[Student],
    *,
    master: dict[str, str] | None = None,
) -> ImportReport:
    """貼り付けた行を解析して、各生徒の資格欄に追加する。"""
    report = ImportReport()
    for line in lines:
        if not normalize(line):
            continue
        row = parse_line(line, students=students, master=master)
        if row.note in ("空行",):
            continue
        if not row.license_name or row.date is None:
            report.rows.append(row)
            continue

        student = _find_by_number(row, students) or _find_by_name(row.src_name, students)
        if student is None:
            row.status = STATUS_ERROR
            row.note = "名簿と照合できません（組・出席番号・氏名を確認してください）"
            report.rows.append(row)
            continue
        row.matched = student

        key = _master_key(row.license_name)
        if any(existing == key for existing, _ in _existing_licenses(student)):
            row.status = STATUS_DUPLICATE
            row.note = f"すでに入力済み（{student.name}）"
            report.rows.append(row)
            continue

        slot = _next_slot(student)
        student.values[f"license.{slot}.name"] = row.license_name
        student.values[f"license.{slot}.ym"] = f"{row.date.year}-{row.date.month:02d}"
        row.status = STATUS_APPLIED
        report.rows.append(row)
    return report


def write_log(report: ImportReport, path) -> None:
    """取り込み結果をCSVに残す。"""
    import csv
    from pathlib import Path

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(
            ["状態", "貼付原文", "組", "出席番号", "抽出元氏名", "名簿氏名",
             "資格名(変換前)", "資格名(正式名称)", "取得日", "備考"]
        )
        for row in report.rows:
            writer.writerow(
                [
                    row.status, row.raw, row.class_no, row.number, row.src_name,
                    row.matched_name, row.src_license, row.license_name,
                    row.date.isoformat() if row.date else "", row.note,
                ]
            )
