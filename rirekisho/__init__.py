"""履歴書（近畿高等学校統一用紙 その2）自動作成システム。

Excel / Googleスプレッドシートに入力した内容を、そのままPDFの用紙に反映する。
"""

from .inputs import Student, read_students, write_template
from .model import Resume, build_resume
from .render import render_preview_png, render_resume

__all__ = [
    "Resume",
    "Student",
    "build_resume",
    "read_students",
    "write_template",
    "render_resume",
    "render_preview_png",
]

DEFAULT_TEMPLATE = "templates/rirekisho_kinki_r7.pdf"
