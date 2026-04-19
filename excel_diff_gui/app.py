"""メインウィンドウ。"""
import tkinter as tk
from tkinter import ttk

from . import settings as cfg
from .tab_dir_diff import TabDirDiff
from .tab_file_diff import TabFileDiff
from .tab_patterns import TabPatterns
from .tab_split import TabSplit
from .widgets import LogArea

try:
    from tkinterdnd2 import TkinterDnD
    _AppBase = TkinterDnD.Tk
except Exception:
    _AppBase = tk.Tk


class App(_AppBase):

    def __init__(self) -> None:
        super().__init__()
        self.title("excel-diff GUI")
        self.minsize(700, 640)
        self.update_idletasks()
        self.geometry("900x700")
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"900x700+{(sw - 900) // 2}+{(sh - 700) // 2}")
        self.lift()
        self.focus_force()

        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        # ノートブックとログを縦分割ペイン（仕切りをドラッグしてログ幅を調整可能）
        paned = ttk.PanedWindow(self, orient="vertical")
        paned.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)

        nb = ttk.Notebook(paned)
        paned.add(nb, weight=5)

        self._log_area = LogArea(paned, height=7)
        paned.add(self._log_area, weight=1)

        # タブ①: フォルダ比較（条件設定）
        tab_dir = TabDirDiff(nb, self._log)  # switch_to_pair_build は後で設定
        self._tab_dir = tab_dir

        # タブ②: フォルダ比較（ペアリング・比較実行）
        tab_patterns = TabPatterns(
            nb, self._log,
            get_compare_options=tab_dir.get_compare_options,
        )
        self._tab_patterns = tab_patterns

        # タブ①のナビゲーションボタンがタブ②を指すよう後から設定
        tab_dir._switch_to_pair_build = lambda: nb.select(tab_patterns)

        # タブ③: ファイル比較
        self._tab_file = TabFileDiff(nb, self._log)

        # タブ④: シート分解
        self._tab_split = TabSplit(nb, self._log)

        nb.add(tab_dir,           text="フォルダ比較（条件設定）")
        nb.add(tab_patterns,      text="フォルダ比較（ペアリング・比較実行）")
        nb.add(self._tab_file,    text="ファイル比較")
        nb.add(self._tab_split,   text="シート分解")

        self.protocol("WM_DELETE_WINDOW", self._quit)

    def _log(self, msg: str) -> None:
        self._log_area.log(msg)

    def _quit(self) -> None:
        # 各タブの現在UI値を _data に書き戻してから保存
        for tab in (self._tab_file, self._tab_dir, self._tab_patterns, self._tab_split):
            tab.save_state()
        cfg.save()
        self.destroy()
