#!/usr/bin/env python3
"""入力シート（1行＝1生徒）から、生徒ごとの履歴書PDFを作る。

    python build_pdf.py                                  # 全員分 → 出力/
    python build_pdf.py --no 3                           # No.3 の生徒だけ
    python build_pdf.py --name 佐野                       # 氏名で絞り込み
    python build_pdf.py --merge 3年2組_履歴書.pdf          # 全員を1つのPDFにまとめる
    python build_pdf.py 入力シート.csv -o 出力            # スプレッドシートのCSVから
"""

from __future__ import annotations

import argparse
import datetime as dt
import sys
from dataclasses import dataclass
from pathlib import Path

from rirekisho import DEFAULT_TEMPLATE
from rirekisho.inputs import OUTPUT_FIELDS, Student, read_students
from rirekisho.licenses import ImportReport, build_master, import_licenses, write_log
from rirekisho.model import InputError, build_resume, normalize, parse_date
from rirekisho.render import render_preview_png, render_resume

DEFAULT_INPUT = "入力シート.xlsx"
DEFAULT_OUTDIR = "出力"
LOG_NAME = "資格取込ログ.csv"


@dataclass
class Result:
    student: Student
    path: Path
    warnings: list[str]


def build_all(
    input_path: str | Path = DEFAULT_INPUT,
    out_dir: str | Path = DEFAULT_OUTDIR,
    *,
    template: str | Path = DEFAULT_TEMPLATE,
    font: str | None = None,
    only_no: int | None = None,
    only_name: str | None = None,
    as_of: str | dt.date | None = None,
    png: bool = False,
    license_lines: list[str] | None = None,
    report: list[ImportReport] | None = None,
    fields: dict[str, bool] | None = None,
) -> list[Result]:
    """入力シートを読んで、生徒ごとにPDFを書き出す。

    「資格取込」シート（または --licenses で渡したファイル）の内容は、
    「資格マスタ」で正式名称に直したうえで各生徒の資格欄に追加する。
    """
    data = read_students(input_path)
    settings = data.settings
    selection = {**data.fields, **(fields or {})}
    class_label = normalize(settings.get("class_label"))
    license_order = normalize(settings.get("license_order")) or "取得年月順"
    master = build_master(data.master_pairs)

    lines = [*data.paste_lines, *(license_lines or [])]
    import_report = import_licenses(lines, data.students, master=master) if lines else ImportReport()
    if report is not None:
        report.append(import_report)

    as_of_value = as_of if as_of is not None else settings.get("as_of")
    as_of_date = parse_date(as_of_value) if normalize(as_of_value) else dt.date.today()

    out_dir = Path(out_dir)
    if import_report.rows:
        write_log(import_report, out_dir / LOG_NAME)

    results: list[Result] = []
    for student in data.students:
        if only_no is not None and student.no != only_no:
            continue
        if only_name and only_name not in student.name:
            continue
        resume = build_resume(
            student.values, as_of=as_of_date, master=master, license_order=license_order
        )
        out = out_dir / f"{student.file_stem(class_label)}.pdf"
        render_resume(resume, template=template, out_path=out, font_path=font, fields=selection)
        if png:
            render_preview_png(out, out.with_suffix(".png"))
        results.append(Result(student, out, resume.warnings))
    return results


def merge_pdfs(results: list[Result], merged_path: str | Path) -> Path:
    """作ったPDFを名列順に1つのファイルへまとめる（印刷用）。"""
    import fitz

    merged_path = Path(merged_path)
    merged_path.parent.mkdir(parents=True, exist_ok=True)
    doc = fitz.open()
    for result in results:
        with fitz.open(result.path) as src:
            doc.insert_pdf(src)
    doc.save(str(merged_path), garbage=4, deflate=True)
    doc.close()
    return merged_path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="入力シートから生徒ごとの履歴書PDFを作成する")
    parser.add_argument("input", nargs="?", default=DEFAULT_INPUT, help="入力シート(.xlsx/.csv)")
    parser.add_argument("-o", "--outdir", default=DEFAULT_OUTDIR, help="出力先フォルダ")
    parser.add_argument("-t", "--template", default=DEFAULT_TEMPLATE, help="用紙のテンプレートPDF")
    parser.add_argument("--no", type=int, help="この No. の生徒だけ作る")
    parser.add_argument("--name", help="氏名に この文字を含む生徒だけ作る")
    parser.add_argument("--merge", help="全員分を1つのPDFにまとめて保存する")
    parser.add_argument("--as-of", help="満○歳の基準日（既定: 設定シートの値）")
    parser.add_argument(
        "--licenses",
        help="資格一覧を1行1件で書いたテキスト/CSVを追加で取り込む（「資格取込」シートと同じ書式）",
    )
    parser.add_argument("--font", help="日本語フォントのパス（省略時は自動検出）")
    parser.add_argument("--png", action="store_true", help="確認用のPNGも書き出す")
    parser.add_argument(
        "--skip",
        help="PDFに反映しない項目をカンマ区切りで指定する（例: contact,jobs,remarks）。"
        f"指定できる項目: {', '.join(key for key, _l, _d in OUTPUT_FIELDS)}",
    )
    args = parser.parse_args(argv)

    skip_fields = None
    if args.skip:
        names = [name.strip() for name in args.skip.split(",") if name.strip()]
        known = {key for key, _l, _d in OUTPUT_FIELDS}
        unknown = [name for name in names if name not in known]
        if unknown:
            print(f"エラー: 知らない項目です: {', '.join(unknown)}", file=sys.stderr)
            return 1
        skip_fields = {name: False for name in names}

    license_lines = None
    if args.licenses:
        license_lines = Path(args.licenses).read_text(encoding="utf-8-sig").splitlines()

    reports: list[ImportReport] = []
    try:
        results = build_all(
            args.input,
            args.outdir,
            template=args.template,
            font=args.font,
            only_no=args.no,
            only_name=args.name,
            as_of=args.as_of,
            png=args.png,
            license_lines=license_lines,
            report=reports,
            fields=skip_fields,
        )
    except (InputError, FileNotFoundError, ValueError) as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        return 1

    for report in reports:
        if not report.rows:
            continue
        print(report.summary())
        for row in report.rows:
            if row.status != "反映":
                print(f"  [{row.status}] {row.raw} … {row.note}")
        print(f"  取り込みの明細: {Path(args.outdir) / LOG_NAME}")

    if not results:
        print("作成対象の生徒がいません（氏名が入力された行がありません）。", file=sys.stderr)
        return 1

    for result in results:
        print(f"作成しました: {result.path}")
        for w in result.warnings:
            print(f"  注意: {w}")

    if args.merge:
        merged = merge_pdfs(results, args.merge)
        print(f"まとめました: {merged}（{len(results)}人分）")

    print(f"合計 {len(results)} 人分を作成しました。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
