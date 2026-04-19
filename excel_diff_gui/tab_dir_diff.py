"""タブ① フォルダ比較（条件設定）。"""
import tkinter as tk
from tkinter import ttk
from typing import Callable, Optional

from . import settings as cfg
from .widgets import FileSelectRow


class TabDirDiff(tk.Frame):

    def __init__(
        self,
        parent,
        log: Callable[[str], None],
        switch_to_pair_build: Optional[Callable] = None,
    ) -> None:
        super().__init__(parent)
        self._log = log
        self._switch_to_pair_build = switch_to_pair_build

        self._out_dir  = tk.StringVar(value=cfg.get("dir_diff", "output_dir"))
        self._sheet    = tk.StringVar(value=cfg.get("dir_diff", "sheet"))
        self._cols     = tk.StringVar(value=cfg.get("dir_diff", "include_cols"))
        self._matchers = tk.StringVar(value=cfg.get("dir_diff", "matchers"))
        self._key_cols = tk.StringVar(value=cfg.get("dir_diff", "key_cols"))
        self._strike   = tk.BooleanVar(value=cfg.get("dir_diff", "strikethrough"))
        self._open_br  = tk.BooleanVar(value=cfg.get("dir_diff", "open_browser", True))
        self._mode     = tk.StringVar(value=cfg.get("dir_diff", "diff_mode", "lcs"))

        self._build()

    # ------------------------------------------------------------------ レイアウト

    def _build(self) -> None:
        pad = {"padx": 6, "pady": 3}

        # 差分モード（比較オプションより先に表示）
        grp_mode = tk.LabelFrame(self, text="差分モード")
        grp_mode.pack(fill="x", **pad)

        tk.Radiobutton(
            grp_mode, text="LCS（行の出現順で比較・デフォルト）",
            variable=self._mode, value="lcs", command=self._on_mode,
        ).pack(anchor="w", padx=8)

        fr_key = tk.Frame(grp_mode)
        fr_key.pack(anchor="w", padx=8, pady=(0, 4))
        tk.Radiobutton(
            fr_key, text="キーJOIN（キー列の値で行を対応付ける）",
            variable=self._mode, value="key", command=self._on_mode,
        ).pack(side="left")
        tk.Label(fr_key, text="キー列").pack(side="left", padx=(12, 2))
        self._entry_key = tk.Entry(fr_key, textvariable=self._key_cols, width=18)
        self._entry_key.pack(side="left")
        tk.Label(fr_key, text="例: C  または  B,C", fg="gray").pack(side="left", padx=6)

        # 比較オプション
        grp_opt = tk.LabelFrame(self, text="比較オプション")
        grp_opt.pack(fill="x", **pad)

        FileSelectRow(
            grp_opt, "出力フォルダ", self._out_dir, mode="dir",
        ).pack(fill="x", padx=6, pady=2)

        fr = tk.Frame(grp_opt)
        fr.pack(fill="x", padx=6, pady=2)
        tk.Label(fr, text="比較シート", width=14, anchor="w").pack(side="left")
        tk.Entry(fr, textvariable=self._sheet).pack(side="left", fill="x", expand=True)
        tk.Label(fr, text="空=全シート", fg="gray").pack(side="left", padx=4)

        fr2 = tk.Frame(grp_opt)
        fr2.pack(fill="x", padx=6, pady=2)
        tk.Label(fr2, text="比較列", width=14, anchor="w").pack(side="left")
        tk.Entry(fr2, textvariable=self._cols).pack(side="left", fill="x", expand=True)
        tk.Label(fr2, text="例: A:C,E", fg="gray").pack(side="left", padx=4)

        FileSelectRow(
            grp_opt, "マッチャーJSON", self._matchers,
            filetypes=[("JSON", "*.json"), ("All", "*.*")],
        ).pack(fill="x", padx=6, pady=2)

        tk.Checkbutton(
            grp_opt, text="取り消し線も差分として扱う", variable=self._strike,
        ).pack(anchor="w", padx=6, pady=2)
        tk.Checkbutton(
            grp_opt, text="完了後ブラウザで開く", variable=self._open_br,
        ).pack(anchor="w", padx=6, pady=(2, 6))

        # ナビゲーションボタン
        nav_row = tk.Frame(self)
        nav_row.pack(fill="x", padx=6, pady=(12, 4))
        tk.Label(
            nav_row,
            text="比較対象フォルダ・ペアリング方法の設定は次のタブで行います",
            fg="gray", font=("", 8),
        ).pack(side="left")
        tk.Button(
            nav_row, text="ペアリング・比較実行へ →",
            bg="#4a9eff", fg="white", font=("", 10, "bold"),
            command=self._go_to_pair_build,
        ).pack(side="right")

        self._on_mode()

    def _on_mode(self) -> None:
        state = "normal" if self._mode.get() == "key" else "disabled"
        self._entry_key.config(state=state)

    def _go_to_pair_build(self) -> None:
        if self._switch_to_pair_build:
            self._switch_to_pair_build()

    # ------------------------------------------------------------------ 現在のUI値

    def get_compare_options(self) -> dict:
        """現在の比較オプションを返す。比較ペア構築タブから参照される。"""
        return {
            "output_dir":    self._out_dir.get().strip(),
            "sheet":         self._sheet.get().strip(),
            "include_cols":  self._cols.get().strip(),
            "matchers":      self._matchers.get().strip(),
            "strikethrough": self._strike.get(),
            "open_browser":  self._open_br.get(),
            "diff_mode":     self._mode.get(),
            "key_cols":      self._key_cols.get().strip(),
        }

    def save_state(self) -> None:
        """現在のUI値を設定に書き戻す（ウィンドウを閉じる前に呼ばれる）。"""
        cfg.set_tab("dir_diff", self.get_compare_options())
