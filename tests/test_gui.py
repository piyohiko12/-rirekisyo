"""画面（gui.py）の確認。

tkinter が使えない環境や画面のない環境では自動的に飛ばす。
"""

import os
import time

import pytest

pytest.importorskip("tkinter", reason="tkinter が入っていません")
if not (os.environ.get("DISPLAY") or os.name == "nt"):
    pytest.skip("画面がないため飛ばします", allow_module_level=True)

import tkinter as tk  # noqa: E402

import gui  # noqa: E402
from make_xlsx import SAMPLE_STUDENTS  # noqa: E402
from rirekisho.inputs import write_template  # noqa: E402


@pytest.fixture()
def app():
    root = tk.Tk()
    application = gui.App(root)
    root.update()
    yield application
    root.destroy()


def test_layout_fits_in_the_window(app):
    root = app.master
    root.geometry(f"{root.winfo_reqwidth()}x{root.winfo_reqheight()}")
    root.update()
    width, height = root.winfo_width(), root.winfo_height()

    def check(widget):
        for child in widget.winfo_children():
            if child.winfo_ismapped():
                x = child.winfo_rootx() - root.winfo_rootx()
                y = child.winfo_rooty() - root.winfo_rooty()
                assert x + child.winfo_width() <= width + 1, child
                assert y + child.winfo_height() <= height + 1, child
            check(child)

    check(root)


def test_build_button_creates_pdfs(app, tmp_path):
    sheet = tmp_path / "入力シート.xlsx"
    write_template(sheet, students=SAMPLE_STUDENTS)
    app.input_path.set(str(sheet))
    app.out_dir.set(str(tmp_path / "出力"))
    app.merge.set(True)

    app.on_build()
    for _ in range(400):  # 作成が終わるまで画面を回す
        app.master.update()
        if not app._busy:
            break
        time.sleep(0.05)
    app.master.update()

    made = sorted(p.name for p in (tmp_path / "出力").iterdir())
    assert "3年2組_01_佐野太郎.pdf" in made
    assert "履歴書まとめ.pdf" in made
    assert "人分を作成しました" in app.status.get()


def test_missing_input_is_reported_without_crashing(app, tmp_path, monkeypatch):
    monkeypatch.setattr(gui.messagebox, "showerror", lambda *a, **k: None)
    app.input_path.set(str(tmp_path / "ない.xlsx"))
    app.on_build()
    for _ in range(200):
        app.master.update()
        if not app._busy:
            break
        time.sleep(0.05)
    app.master.update()
    assert "エラー" in app.status.get()
