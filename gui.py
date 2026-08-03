#!/usr/bin/env python3
"""履歴書PDF作成ツール（ボタンで操作する画面）。

コマンドを打たずに、ボタンだけでPDFを作れるようにしたもの。

    python gui.py

Windows は「履歴書PDF作成.bat」、macOS は「履歴書PDF作成.command」を
ダブルクリックしても起動する。
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import time
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from build_pdf import DEFAULT_INPUT, DEFAULT_OUTDIR, LOG_NAME, build_all, merge_pdfs
from rirekisho import DEFAULT_TEMPLATE
from rirekisho.inputs import OUTPUT_FIELDS, default_fields, read_students
from rirekisho.fonts import FontNotFoundError
from rirekisho.model import InputError

BASE_DIR = Path(__file__).resolve().parent
WATCH_INTERVAL = 1.0


@dataclass
class Options:
    """画面の設定を写し取ったもの。

    tkinter は別スレッドから触れないので、作成処理には画面の値ではなく
    この写しだけを渡す。
    """

    input_path: str
    out_dir: str
    target: str
    only_no: str
    only_name: str
    merge: bool
    merge_name: str
    png: bool
    fields: dict[str, bool]


def open_in_explorer(path: Path) -> None:
    """ファイルやフォルダをOSの標準アプリで開く。"""
    path = path.resolve()
    if sys.platform.startswith("win"):
        os.startfile(path)  # type: ignore[attr-defined]  # noqa: S606
    elif sys.platform == "darwin":
        subprocess.run(["open", str(path)], check=False)
    else:
        subprocess.run(["xdg-open", str(path)], check=False)


class App(ttk.Frame):
    def __init__(self, master: tk.Tk):
        super().__init__(master, padding=12)
        master.title("履歴書PDF作成")
        master.minsize(700, 620)
        self.grid(sticky="nsew")
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)

        self.input_path = tk.StringVar(value=str(BASE_DIR / DEFAULT_INPUT))
        self.out_dir = tk.StringVar(value=str(BASE_DIR / DEFAULT_OUTDIR))
        self.target = tk.StringVar(value="all")
        self.only_no = tk.StringVar()
        self.only_name = tk.StringVar()
        self.merge = tk.BooleanVar(value=False)
        self.merge_name = tk.StringVar(value="履歴書まとめ.pdf")
        self.png = tk.BooleanVar(value=False)
        self.status = tk.StringVar(value="準備できました。")
        self.field_vars = {
            key: tk.BooleanVar(value=default) for key, _label, default in OUTPUT_FIELDS
        }

        self._messages: queue.Queue[tuple[str, str]] = queue.Queue()
        self._snapshot: Options | None = None
        self._busy = False
        self._watching = False
        self._watch_stop = threading.Event()

        self._after_ids: list[str] = []
        self._build_widgets()
        self.on_load_fields(quiet=True)
        self._refresh_options()
        self._after_ids.append(self.after(120, self._drain_messages))
        master.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Destroy>", self._cancel_timers)

    # ------------------------------------------------------------ 画面
    def _build_widgets(self) -> None:
        row = 0
        ttk.Label(self, text="入力シート").grid(row=row, column=0, sticky="w", pady=3)
        ttk.Entry(self, textvariable=self.input_path).grid(row=row, column=1, sticky="ew", padx=6)
        ttk.Button(self, text="参照…", command=self._choose_input).grid(row=row, column=2)

        row += 1
        ttk.Label(self, text="出力先フォルダ").grid(row=row, column=0, sticky="w", pady=3)
        ttk.Entry(self, textvariable=self.out_dir).grid(row=row, column=1, sticky="ew", padx=6)
        ttk.Button(self, text="参照…", command=self._choose_outdir).grid(row=row, column=2)

        row += 1
        target = ttk.LabelFrame(self, text="作成する生徒", padding=8)
        target.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(10, 4))
        ttk.Radiobutton(target, text="全員", variable=self.target, value="all").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Radiobutton(target, text="No.", variable=self.target, value="no").grid(
            row=0, column=1, sticky="w", padx=(16, 2)
        )
        ttk.Entry(target, textvariable=self.only_no, width=6).grid(row=0, column=2, sticky="w")
        ttk.Radiobutton(target, text="氏名に含む", variable=self.target, value="name").grid(
            row=0, column=3, sticky="w", padx=(16, 2)
        )
        ttk.Entry(target, textvariable=self.only_name, width=14).grid(row=0, column=4, sticky="w")

        row += 1
        option = ttk.LabelFrame(self, text="オプション", padding=8)
        option.grid(row=row, column=0, columnspan=3, sticky="ew", pady=4)
        ttk.Checkbutton(
            option, text="全員を1つのPDFにまとめる（印刷用）", variable=self.merge
        ).grid(row=0, column=0, sticky="w")
        ttk.Entry(option, textvariable=self.merge_name, width=24).grid(
            row=0, column=1, sticky="w", padx=6
        )
        ttk.Checkbutton(option, text="確認用のPNGも作る", variable=self.png).grid(
            row=1, column=0, sticky="w", pady=(4, 0)
        )

        row += 1
        fields = ttk.LabelFrame(self, text="PDFに反映する項目（チェックを外すとその欄は空欄になります）", padding=8)
        fields.grid(row=row, column=0, columnspan=3, sticky="ew", pady=4)
        columns = 4
        for i, (key, label, _default) in enumerate(OUTPUT_FIELDS):
            ttk.Checkbutton(fields, text=label, variable=self.field_vars[key]).grid(
                row=i // columns, column=i % columns, sticky="w", padx=(0, 10)
            )
        edge = ttk.Frame(fields)
        edge.grid(row=(len(OUTPUT_FIELDS) - 1) // columns + 1, column=0, columnspan=columns,
                  sticky="w", pady=(6, 0))
        ttk.Button(edge, text="すべて入れる", width=12,
                   command=lambda: self._set_all_fields(True)).grid(row=0, column=0)
        ttk.Button(edge, text="すべて外す", width=12,
                   command=lambda: self._set_all_fields(False)).grid(row=0, column=1, padx=6)
        ttk.Button(edge, text="シートの設定を読み込む", width=22,
                   command=self.on_load_fields).grid(row=0, column=2)

        row += 1
        buttons = ttk.Frame(self)
        buttons.grid(row=row, column=0, columnspan=3, sticky="ew", pady=10)
        self.build_button = tk.Button(
            buttons,
            text="PDF作成",
            command=self.on_build,
            font=("", 14, "bold"),
            height=2,
            width=16,
            bg="#2c6fbb",
            fg="white",
            activebackground="#255f9f",
            activeforeground="white",
        )
        self.build_button.grid(row=0, column=0, rowspan=2, padx=(0, 12))
        self.watch_button = ttk.Button(
            buttons, text="自動更新を開始", width=18, command=self.on_toggle_watch
        )
        self.watch_button.grid(row=0, column=1, sticky="ew", pady=(0, 4))
        ttk.Button(buttons, text="入力シートを開く", width=18, command=self.on_open_input).grid(
            row=1, column=1, sticky="ew"
        )
        ttk.Button(buttons, text="出力フォルダを開く", width=18, command=self.on_open_outdir).grid(
            row=0, column=2, sticky="ew", padx=6, pady=(0, 4)
        )
        ttk.Button(buttons, text="取込ログを開く", width=18, command=self.on_open_log).grid(
            row=1, column=2, sticky="ew", padx=6
        )

        row += 1
        ttk.Label(self, textvariable=self.status, foreground="#333333").grid(
            row=row, column=0, columnspan=3, sticky="w"
        )

        row += 1
        self.log = tk.Text(self, height=10, wrap="word")
        self.log.grid(row=row, column=0, columnspan=3, sticky="nsew", pady=(6, 0))
        scroll = ttk.Scrollbar(self, command=self.log.yview)
        scroll.grid(row=row, column=3, sticky="ns", pady=(6, 0))
        self.log.configure(yscrollcommand=scroll.set, state="disabled")
        self.log.tag_configure("error", foreground="#b00020")
        self.log.tag_configure("warn", foreground="#a06000")
        self.log.tag_configure("ok", foreground="#1a6b32")
        self.rowconfigure(row, weight=1)

        self._write("「PDF作成」を押すと、入力シートの内容から生徒ごとの履歴書PDFを作ります。")

    # ------------------------------------------------------------ 部品
    def _choose_input(self) -> None:
        path = filedialog.askopenfilename(
            title="入力シートを選ぶ",
            initialdir=str(Path(self.input_path.get()).parent),
            filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.csv"), ("すべて", "*.*")],
        )
        if path:
            self.input_path.set(path)
            self.on_load_fields(quiet=True)

    def _choose_outdir(self) -> None:
        path = filedialog.askdirectory(title="出力先フォルダを選ぶ", initialdir=self.out_dir.get())
        if path:
            self.out_dir.set(path)

    def _write(self, text: str, tag: str = "") -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _drain_messages(self) -> None:
        """別スレッドからのメッセージを画面に出す。"""
        while True:
            try:
                tag, text = self._messages.get_nowait()
            except queue.Empty:
                break
            if tag == "status":
                self.status.set(text)
            elif tag == "dialog":
                messagebox.showerror("履歴書PDF作成", text)
            elif tag == "done":
                self._busy = False
                self.build_button.configure(state="normal")
            else:
                self._write(text, tag)
        self._after_ids.append(self.after(120, self._drain_messages))

    def _post(self, tag: str, text: str) -> None:
        self._messages.put((tag, text))

    def _cancel_timers(self, event=None) -> None:
        """閉じたあとにタイマーが動いてエラーにならないよう、後片付けをする。"""
        if event is not None and event.widget is not self:
            return
        self._watch_stop.set()
        for after_id in self._after_ids:
            try:
                self.after_cancel(after_id)
            except tk.TclError:
                pass
        self._after_ids.clear()

    def _on_close(self) -> None:
        self._cancel_timers()
        self.master.destroy()

    # ------------------------------------------------------------ 実行
    def _collect_options(self) -> Options:
        """画面の値を写し取る（メインスレッドから呼ぶこと）。"""
        return Options(
            input_path=self.input_path.get(),
            out_dir=self.out_dir.get(),
            target=self.target.get(),
            only_no=self.only_no.get().strip(),
            only_name=self.only_name.get().strip(),
            merge=self.merge.get(),
            merge_name=self.merge_name.get().strip(),
            png=self.png.get(),
            fields={key: var.get() for key, var in self.field_vars.items()},
        )

    def _refresh_options(self) -> None:
        """自動更新スレッドが使えるように、画面の値をこまめに写しておく。"""
        self._snapshot = self._collect_options()
        self._after_ids.append(self.after(300, self._refresh_options))

    def on_build(self) -> None:
        if self._busy:
            return
        self._busy = True
        self.build_button.configure(state="disabled")
        self._write("")
        options = self._collect_options()
        threading.Thread(
            target=self._run_build, args=(options,), kwargs={"manual": True}, daemon=True
        ).start()

    def _run_build(self, options: Options, *, manual: bool) -> None:
        """PDFを作る。ここは別スレッドで動くので、画面には触らずメッセージだけ送る。"""
        stamp = time.strftime("%H:%M:%S")
        try:
            only_no = None
            only_name = None
            if options.target == "no":
                if not options.only_no.isdigit():
                    raise InputError("No. には数字を入れてください。")
                only_no = int(options.only_no)
            elif options.target == "name":
                if not options.only_name:
                    raise InputError("氏名の一部を入れてください。")
                only_name = options.only_name

            out_dir = Path(options.out_dir)
            reports: list = []
            results = build_all(
                options.input_path,
                out_dir,
                template=BASE_DIR / DEFAULT_TEMPLATE,
                only_no=only_no,
                only_name=only_name,
                png=options.png,
                report=reports,
                fields=options.fields,
            )

            for report in reports:
                if report.rows:
                    self._post("", f"[{stamp}] {report.summary()}")
                    for r in report.rows:
                        if r.status != "反映":
                            self._post("warn", f"    [{r.status}] {r.raw} … {r.note}")

            if not results:
                self._post("warn", f"[{stamp}] 作成対象の生徒がいません（氏名が空の行だけです）。")
                self._post("status", "作成対象なし")
                return

            for result in results:
                self._post("", f"    {result.path.name}")
                for w in result.warnings:
                    self._post("warn", f"      注意: {result.student.name}: {w}")

            if options.merge:
                name = options.merge_name or "履歴書まとめ.pdf"
                if not name.lower().endswith(".pdf"):
                    name += ".pdf"
                merged = merge_pdfs(results, out_dir / name)
                self._post("", f"    まとめ: {merged.name}（{len(results)}人分）")

            self._post("ok", f"[{stamp}] {len(results)}人分を作成しました → {out_dir}")
            self._post("status", f"{len(results)}人分を作成しました（{stamp}）")
        except FileNotFoundError as exc:
            self._fail(stamp, f"ファイルが見つかりません: {exc}", manual)
        except FontNotFoundError as exc:
            self._fail(stamp, str(exc), manual)
        except (InputError, ValueError) as exc:
            self._fail(stamp, str(exc), manual)
        except PermissionError:
            self._fail(
                stamp,
                "PDFに書き込めませんでした。同じPDFを開いたままになっていないか確認してください。",
                manual,
            )
        except Exception as exc:  # 想定外でも画面を落とさない
            self._fail(stamp, f"エラー: {exc}", manual)
        finally:
            if manual:
                self._post("done", "")

    def _fail(self, stamp: str, message: str, manual: bool) -> None:
        self._post("error", f"[{stamp}] {message}")
        self._post("status", "エラーが出ました。内容を確認してください。")
        if manual:
            self._post("dialog", message)

    def _set_all_fields(self, value: bool) -> None:
        for var in self.field_vars.values():
            var.set(value)

    def on_load_fields(self, quiet: bool = False) -> None:
        """入力シートの「反映項目」シートの○×をチェックボックスに読み込む。"""
        try:
            fields = read_students(self.input_path.get()).fields
        except Exception as exc:  # 読めなくても画面は動かす
            if not quiet:
                messagebox.showinfo("履歴書PDF作成", f"反映項目を読み込めませんでした: {exc}")
            return
        for key, value in {**default_fields(), **fields}.items():
            if key in self.field_vars:
                self.field_vars[key].set(value)
        if not quiet:
            self._write("入力シートの「反映項目」を読み込みました。")

    # ------------------------------------------------------------ 自動更新
    def on_toggle_watch(self) -> None:
        if self._watching:
            self._watch_stop.set()
            self._watching = False
            self.watch_button.configure(text="自動更新を開始")
            self._write("自動更新を止めました。")
            self.status.set("自動更新: 停止")
            return
        self._watch_stop.clear()
        self._watching = True
        self.watch_button.configure(text="自動更新を停止")
        self._write("自動更新を開始しました。入力シートを保存するたびに作り直します。")
        self.status.set("自動更新: 監視中")
        threading.Thread(target=self._watch_loop, daemon=True).start()

    def _watch_loop(self) -> None:
        last: float | None = None
        while not self._watch_stop.is_set():
            options = self._snapshot
            if options is None:
                self._watch_stop.wait(WATCH_INTERVAL)
                continue
            path = Path(options.input_path)
            if path.exists():
                stamp = path.stat().st_mtime
                if last is None:
                    last = stamp  # 開始直後は作り直さない
                elif stamp != last:
                    last = stamp
                    time.sleep(0.3)  # 保存の途中で読まないよう少し待つ
                    self._run_build(options, manual=False)
            self._watch_stop.wait(WATCH_INTERVAL)

    # ------------------------------------------------------------ 開く
    def on_open_input(self) -> None:
        self._open(Path(self.input_path.get()), "入力シートが見つかりません。")

    def on_open_outdir(self) -> None:
        out_dir = Path(self.out_dir.get())
        out_dir.mkdir(parents=True, exist_ok=True)
        self._open(out_dir, "出力フォルダが見つかりません。")

    def on_open_log(self) -> None:
        self._open(Path(self.out_dir.get()) / LOG_NAME, "取込ログはまだありません。")

    def _open(self, path: Path, missing_message: str) -> None:
        if not path.exists():
            messagebox.showinfo("履歴書PDF作成", missing_message)
            return
        open_in_explorer(path)


def main() -> int:
    root = tk.Tk()
    try:
        root.call("tk", "scaling", 1.2)
    except tk.TclError:
        pass
    App(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
