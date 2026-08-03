"""履歴書PDFの描画。

テンプレートPDF（罫線と印字済みの文言が入った近畿統一用紙）の上に、
入力内容を文字として重ねて書き込む。
"""

from __future__ import annotations

from pathlib import Path

import fitz

from . import layout as L
from .fonts import find_japanese_font
from .inputs import default_fields
from .model import SAME_AS_ABOVE, Resume

FONT_NAME = "jpgothic"

# 行頭に置かない文字（簡易な禁則処理）
NO_LINE_START = "、。，．）］｝」』】〉》’”ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶーぐ！？!?,.:;"


class Renderer:
    def __init__(self, template: str | Path, font_path: str | Path | None = None):
        self.doc = fitz.open(str(template))
        self.page = self.doc[0]
        self.font_path = Path(find_japanese_font(font_path))
        self.font = fitz.Font(fontfile=str(self.font_path))
        self.page.insert_font(fontname=FONT_NAME, fontfile=str(self.font_path))
        # テンプレートは /Rotate 90 の縦ページ。見た目の座標→ページ座標へ変換する行列。
        self._derotate = self.page.derotation_matrix

    # ---------------------------------------------------------- 低レベル描画
    def _point(self, x: float, y: float) -> fitz.Point:
        return fitz.Point(x, y) * self._derotate

    def _rect(self, x0: float, y0: float, x1: float, y1: float) -> fitz.Rect:
        return fitz.Rect(x0, y0, x1, y1) * self._derotate

    def text_width(self, text: str, size: float) -> float:
        return self.font.text_length(text, size)

    def _baseline(self, y0: float, y1: float, size: float) -> float:
        """文字の見た目が枠の上下中央に来るベースライン位置。"""
        ascender = getattr(self.font, "ascender", 0.88)
        descender = getattr(self.font, "descender", -0.12)
        return (y0 + y1) / 2 + (ascender + descender) / 2 * size

    def draw_line(
        self,
        text: str,
        band: tuple[float, float, float, float],
        size: float,
        *,
        align: str = "left",
        pad: float = 3.0,
        min_size: float = L.MIN_FONT_SIZE,
    ) -> None:
        """1行だけの欄に、上下中央そろえで書く。幅に収まらなければ自動で縮小する。"""
        if not text:
            return
        x0, y0, x1, y1 = band
        width = x1 - x0 - pad * 2
        while size > min_size and self.text_width(text, size) > width:
            size -= 0.25
        w = self.text_width(text, size)
        if align == "center":
            x = x0 + (x1 - x0 - w) / 2
        elif align == "right":
            x = x1 - pad - w
        else:
            x = x0 + pad
        self.page.insert_text(
            self._point(x, self._baseline(y0, y1, size)),
            text,
            fontname=FONT_NAME,
            fontsize=size,
            rotate=90,
        )

    def draw_balanced_line(
        self,
        text: str,
        band: tuple[float, float, float, float],
        *,
        max_size: float,
        fill: float = 0.6,
        min_size: float = L.MIN_FONT_SIZE,
        pad: float = 6.0,
    ) -> None:
        """氏名のように、枠に対して見た目のバランスをとって中央に置く。

        文字数が少なければ大きく、多ければ小さくして、枠幅の ``fill`` 倍あたりに
        収まるようにする（上下も中央そろえ）。
        """
        if not text:
            return
        x0, y0, x1, y1 = band
        width = x1 - x0
        height = y1 - y0
        unit = self.text_width(text, 1.0) or 1.0
        size = min(max_size, width * fill / unit, height * 0.62)
        while size > min_size and self.text_width(text, size) > width - pad * 2:
            size -= 0.25
        size = max(size, min_size)
        x = x0 + (width - self.text_width(text, size)) / 2
        self.page.insert_text(
            self._point(x, self._baseline(y0, y1, size)),
            text,
            fontname=FONT_NAME,
            fontsize=size,
            rotate=90,
        )

    def wrap(self, text: str, size: float, width: float) -> list[str]:
        """日本語向けの折り返し（任意の文字位置で折り返し、簡易禁則あり）。"""
        lines: list[str] = []
        for para in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
            if not para:
                lines.append("")
                continue
            cur = ""
            for ch in para:
                if cur and self.text_width(cur + ch, size) > width:
                    if ch in NO_LINE_START:  # 行頭禁則: はみ出しても現在行に残す
                        cur += ch
                        continue
                    lines.append(cur)
                    cur = ch
                else:
                    cur += ch
            lines.append(cur)
        return lines

    def draw_block(
        self,
        text: str,
        rect: tuple[float, float, float, float],
        size: float,
        *,
        pad: float = 4.0,
        line_gap: float = 1.45,
        min_size: float = L.MIN_FONT_SIZE,
        valign: str = "top",
    ) -> float:
        """複数行の欄に折り返して書く。収まらなければ文字を自動で小さくする。

        戻り値は実際に使った文字サイズ（縮小しても収まらなかった場合は min_size）。
        """
        if not text.strip():
            return size
        x0, y0, x1, y1 = rect
        width = x1 - x0 - pad * 2
        height = y1 - y0 - pad * 2
        while True:
            lines = self.wrap(text, size, width)
            if len(lines) * size * line_gap <= height or size <= min_size:
                break
            size -= 0.25
        used = len(lines) * size * line_gap
        top = y0 + pad
        if valign == "center" and used < height:
            top += (height - used) / 2
        for i, line in enumerate(lines):
            if not line:
                continue
            baseline = top + size * line_gap * i + size * 0.92
            if baseline > y1:  # 欄からあふれる分は書かない
                break
            self.page.insert_text(
                self._point(x0 + pad, baseline),
                line,
                fontname=FONT_NAME,
                fontsize=size,
                rotate=90,
            )
        return size

    def draw_circle(self, rect: tuple[float, float, float, float], width: float = 1.1) -> None:
        """元号などを丸で囲む。"""
        self.page.draw_oval(self._rect(*rect), color=(0, 0, 0), width=width)

    def save(self, out_path: str | Path) -> Path:
        out_path = Path(out_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            # 使った文字だけを埋め込む（そのままだと日本語フォント全体で数MBになる）
            self.doc.subset_fonts(verbose=False)
        except Exception:  # サブセット化に失敗してもPDF自体は作れる
            pass
        self.doc.save(str(out_path), garbage=4, deflate=True)
        return out_path


# ------------------------------------------------------------------ 各欄の描画
def _draw_personal(r: Renderer, resume: Resume, on) -> None:
    fs = L.FONT_SIZES
    if on("name_kana"):
        r.draw_balanced_line(resume.name_kana, L.NAME_KANA, max_size=fs["kana"], fill=0.55)
    if on("name"):
        r.draw_balanced_line(resume.name, L.NAME, max_size=fs["name"], fill=L.NAME_FILL)

    if resume.birth and on("birth"):
        y0, y1 = L.BIRTH_ROW
        era = resume.birth_era
        if era != L.BIRTH_ERA_PREPRINTED:
            if era in L.BIRTH_ERA_CIRCLE:
                r.draw_circle(L.BIRTH_ERA_CIRCLE[era])
                resume.warnings.append(
                    f"生年月日は「{era}」なので丸を付けましたが、用紙には"
                    f"「{L.BIRTH_ERA_PREPRINTED}」の丸が最初から印刷されています。"
                    "印刷後に手で消してください。"
                )
            else:
                resume.warnings.append(
                    f"生年月日の元号「{era}」は用紙に印刷されていないため、丸囲みは付けていません。"
                )
        band = lambda xr: (xr[0], y0, xr[1], y1)  # noqa: E731
        r.draw_line(resume.birth_wareki_year, band(L.BIRTH_YEAR), fs["birth"], align="center")
        r.draw_line(str(resume.birth.month), band(L.BIRTH_MONTH), fs["birth"], align="center")
        r.draw_line(str(resume.birth.day), band(L.BIRTH_DAY), fs["birth"], align="center")
        r.draw_line(resume.age, band(L.BIRTH_AGE), fs["birth"], align="center")

    if on("zip"):
        r.draw_line(resume.zip_code, L.ADDR_ZIP, fs["zip"])
    if on("address_kana"):
        r.draw_line(resume.address_kana, L.ADDR_KANA, fs["kana"], pad=8)
    if on("address"):
        r.draw_block(resume.address, L.ADDR_TEXT, fs["address"], valign="center")

    if not on("contact"):
        pass
    elif resume.contact_is_same:
        # 「同上」だけのときは連絡先の枠の真ん中に置く
        r.draw_line(SAME_AS_ABOVE, L.CONTACT_SAME, fs["address"], align="center")
    else:
        r.draw_line(resume.contact_zip, L.CONTACT_ZIP, fs["zip"])
        r.draw_line(resume.contact_kana, L.CONTACT_KANA, fs["kana"], pad=8)
        r.draw_block(resume.contact_address, L.CONTACT_TEXT, fs["address"], valign="center")


def _draw_licenses(r: Renderer, resume: Resume) -> None:
    if not resume.licenses:
        return
    top, bottom = L.LICENSE_AREA
    rows = max(L.LICENSE_ROWS, len(resume.licenses))
    row_h = (bottom - top) / rows
    if len(resume.licenses) > rows:
        resume.warnings.append("資格の行数が欄に収まりません。")
    for i, lic in enumerate(resume.licenses):
        y0 = top + row_h * i
        y1 = y0 + row_h
        r.draw_line(
            lic.ym_text, (L.LICENSE_YM[0], y0, L.LICENSE_YM[1], y1),
            L.FONT_SIZES["license"], align="center",
        )
        r.draw_line(
            lic.name, (L.LICENSE_NAME[0], y0, L.LICENSE_NAME[1], y1),
            L.FONT_SIZES["license"], pad=6,
        )


def _draw_jobs(r: Renderer, resume: Resume) -> None:
    for i, job in enumerate(resume.jobs):
        if i >= len(L.JOB_ROWS):
            resume.warnings.append("職歴が4行に収まりません。")
            break
        top, bottom = L.JOB_ROWS[i]
        ym_band = (top + L.JOB_YM_BAND[0], top + L.JOB_YM_BAND[1])
        if job.era in L.JOB_ERA_CIRCLE:
            x0, dy0, x1, dy1 = L.JOB_ERA_CIRCLE[job.era]
            r.draw_circle((x0, top + dy0, x1, top + dy1))
        if job.year:
            r.draw_line(
                str(job.year), (L.JOB_YEAR[0], ym_band[0], L.JOB_YEAR[1], ym_band[1]),
                L.FONT_SIZES["job"], align="center",
            )
            r.draw_line(
                str(job.month), (L.JOB_MONTH[0], ym_band[0], L.JOB_MONTH[1], ym_band[1]),
                L.FONT_SIZES["job"], align="center",
            )
        r.draw_block(
            job.text, (L.JOB_TEXT[0], top + 2, L.JOB_TEXT[1], bottom - 2),
            L.FONT_SIZES["job"], valign="center",
        )


def compose_motivation(resume: Resume, on=None) -> str:
    """志望の動機・希望の職種・アピールポイントを1つの欄用にまとめる。

    見出しは付けず、用紙の項目名（志望の動機／希望の職種／アピールポイント）の
    並び順どおりに、入力のあるものだけを続けて書く。
    """
    if on is None:
        def on(_key: str) -> bool:  # 既定はすべて反映
            return True

    parts = [
        (resume.motivation if on("motivation") else ""),
        (resume.desired_job if on("desired_job") else ""),
        (resume.appeal if on("appeal") else ""),
    ]
    return "\n".join(part for part in parts if part)


def render_resume(
    resume: Resume,
    *,
    template: str | Path,
    out_path: str | Path,
    font_path: str | Path | None = None,
    fields: dict[str, bool] | None = None,
) -> Path:
    """Resume の内容をテンプレートPDFに書き込んで保存する。

    fields で「反映する項目」を指定できる（False にした項目は書き込まない）。
    """
    selection = {**default_fields(), **(fields or {})}

    def on(key: str) -> bool:
        return selection.get(key, True)

    r = Renderer(template, font_path)
    _draw_personal(r, resume, on)
    if on("licenses"):
        _draw_licenses(r, resume)
    if on("jobs"):
        _draw_jobs(r, resume)
    if on("activities"):
        r.draw_block(resume.activities, L.ACTIVITIES, L.FONT_SIZES["body"])
    r.draw_block(compose_motivation(resume, on), L.MOTIVATION, L.FONT_SIZES["body"])
    if on("remarks"):
        r.draw_block(resume.remarks, L.REMARKS, L.FONT_SIZES["body"])
    return r.save(out_path)


def render_preview_png(pdf_path: str | Path, png_path: str | Path, dpi: int = 150) -> Path:
    """出来上がったPDFを確認用のPNGにする。"""
    doc = fitz.open(str(pdf_path))
    png_path = Path(png_path)
    doc[0].get_pixmap(dpi=dpi).save(str(png_path))
    return png_path
