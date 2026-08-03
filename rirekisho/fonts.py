"""日本語フォントの検出。

環境変数 RIREKISHO_FONT にフォントファイルのパスを設定すると、それを最優先で使う。
"""

from __future__ import annotations

import os
from pathlib import Path

# OSごとの代表的な明朝体。上から順に探す（履歴書は明朝体で印字する）。
CANDIDATES = [
    # Windows
    r"C:\Windows\Fonts\yumin.ttf",       # 游明朝
    r"C:\Windows\Fonts\YuMincho.ttc",
    r"C:\Windows\Fonts\msmincho.ttc",    # MS明朝
    r"C:\Windows\Fonts\HGRME.TTC",       # HG明朝E
    # macOS
    "/System/Library/Fonts/ヒラギノ明朝 ProN.ttc",
    "/System/Library/Fonts/Hiragino Mincho ProN.ttc",
    "/Library/Fonts/ヒラギノ明朝 Pro W3.otf",
    # Linux
    "/usr/share/fonts/opentype/ipafont-mincho/ipam.ttf",
    "/usr/share/fonts/truetype/fonts-japanese-mincho.ttf",
    "/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc",
    "/usr/share/fonts/truetype/noto/NotoSerifCJKjp-Regular.otf",
    # 明朝体が1つも無いときの最後の手段（ゴシック体）
    r"C:\Windows\Fonts\meiryo.ttc",
    r"C:\Windows\Fonts\msgothic.ttc",
    "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc",
    "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf",
    "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf",
]


class FontNotFoundError(RuntimeError):
    pass


def find_japanese_font(explicit: str | os.PathLike[str] | None = None) -> Path:
    """使用する日本語フォントファイルを返す。"""
    for cand in [explicit, os.environ.get("RIREKISHO_FONT"), *CANDIDATES]:
        if not cand:
            continue
        path = Path(cand)
        if path.is_file():
            return path
    raise FontNotFoundError(
        "日本語フォントが見つかりませんでした。\n"
        "環境変数 RIREKISHO_FONT に日本語フォント（.ttf / .ttc）のパスを指定してください。\n"
        "  例) Windows: set RIREKISHO_FONT=C:\\Windows\\Fonts\\msmincho.ttc\n"
        "      macOS:   export RIREKISHO_FONT='/System/Library/Fonts/ヒラギノ明朝 ProN.ttc'"
    )
