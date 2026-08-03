"""日本語フォントの検出。

環境変数 RIREKISHO_FONT にフォントファイルのパスを設定すると、それを最優先で使う。
"""

from __future__ import annotations

import os
from pathlib import Path

# OSごとの代表的なゴシック体。上から順に探す。
CANDIDATES = [
    # Windows
    r"C:\Windows\Fonts\YuGothM.ttc",
    r"C:\Windows\Fonts\yugothm.ttc",
    r"C:\Windows\Fonts\meiryo.ttc",
    r"C:\Windows\Fonts\msgothic.ttc",
    # macOS
    "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc",
    "/System/Library/Fonts/Hiragino Sans GB.ttc",
    "/Library/Fonts/Arial Unicode.ttf",
    # Linux
    "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf",
    "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/noto/NotoSansCJKjp-Regular.ttf",
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
        "  例) Windows: set RIREKISHO_FONT=C:\\Windows\\Fonts\\meiryo.ttc\n"
        "      macOS:   export RIREKISHO_FONT='/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc'"
    )
