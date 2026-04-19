"""タブ① ファイル比較。"""
import os
import queue
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Callable

from . import settings as cfg
from .widgets import FileSelectRow
from .worker import get_worker


class TabFileDiff(tk.Frame):

    def __init__(self, parent, log: Callable[[str], None]) -> None:
        super().__init__(parent)
        self._log = log
        self._result_q: "queue.Queue | None" = None

        # 変数
        self._old     = tk.StringVar(value=cfg.get("file_diff", "old_file"))
        self._new     = tk.StringVar(value=cfg.get("file_diff", "new_file"))
        self._out     = tk.StringVar(value=cfg.get("file_diff", "output"))
        self._sheet   = tk.StringVar(value=cfg.get("file_diff", "sheet"))
        self._cols    = tk.StringVar(value=cfg.get("file_diff", "include_cols"))
        self._matchers= tk.StringVar(value=cfg.get("file_diff", "matchers"))
        self._key_cols= tk.StringVar(value=cfg.get("file_diff", "key_cols"))
        self._strike  = tk.BooleanVar(value=cfg.get("file_diff", "strikethrough"))
        self._open_br = tk.BooleanVar(value=cfg.get("file_diff", "open_browser"))
        self._mode    = tk.StringVar(value=cfg.get("file_diff", "diff_mode", "lcs"))

        self._opt_built = False
        self._opt_open  = False

        self._build()

    # ------------------------------------------------------------------ レイアウト

    def _build(self) -> None:
        pad = {"padx": 6, "pady": 3}

        # ファイル選択
        grp_files = tk.LabelFrame(self, text="ファイル")
        grp_files.pack(fill="x", **pad)
        FileSelectRow(
            grp_files, "旧ファイル", self._old,
            filetypes=[("Excel", "*.xlsx"), ("All", "*.*")],
        ).pack(fill="x", padx=6, pady=2)
        FileSelectRow(
            grp_files, "新ファイル", self._new,
            filetypes=[("Excel", "*.xlsx"), ("All", "*.*")],
        ).pack(fill="x", padx=6, pady=2)

        # 差分モード
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

        # オプション折りたたみ + 実行ボタン（同じ行）
        ctrl_row = tk.Frame(self)
        ctrl_row.pack(fill="x", padx=6, pady=(6, 2))

        self._lbl_opt = tk.Label(ctrl_row, text="▶ オプション", cursor="hand2", fg="blue")
        self._lbl_opt.pack(side="left")
        self._lbl_opt.bind("<Button-1>", self._toggle_opt)

        self._btn_run = tk.Button(
            ctrl_row, text="実行", width=16,
            bg="#4a9eff", fg="white", font=("", 10, "bold"),
            command=self._run,
        )
        self._btn_run.pack(side="right")

        # オプション本体（初期は非表示）
        self._grp_opt = tk.LabelFrame(self, text="オプション")

        self._on_mode()

    def _on_mode(self) -> None:
        state = "normal" if self._mode.get() == "key" else "disabled"
        self._entry_key.config(state=state)

    def _toggle_opt(self, _=None) -> None:
        if self._opt_open:
            self._grp_opt.pack_forget()
            self._opt_open = False
            self._lbl_opt.config(text="▶ オプション")
        else:
            if not self._opt_built:
                self._build_opt()
                self._opt_built = True
            self._grp_opt.pack(fill="x", padx=6, pady=3)
            self._opt_open = True
            self._lbl_opt.config(text="▼ オプション")

    def _build_opt(self) -> None:
        pad = {"padx": 6, "pady": 2}
        g = self._grp_opt

        FileSelectRow(
            g, "出力HTML", self._out,
            filetypes=[("HTML", "*.html"), ("All", "*.*")],
        ).pack(fill="x", **pad)

        fr = tk.Frame(g)
        fr.pack(fill="x", **pad)
        tk.Label(fr, text="比較シート", width=14, anchor="w").pack(side="left")
        tk.Entry(fr, textvariable=self._sheet).pack(side="left", fill="x", expand=True)
        tk.Label(fr, text="空=全シート", fg="gray").pack(side="left", padx=4)

        fr2 = tk.Frame(g)
        fr2.pack(fill="x", **pad)
        tk.Label(fr2, text="比較列", width=14, anchor="w").pack(side="left")
        tk.Entry(fr2, textvariable=self._cols).pack(side="left", fill="x", expand=True)
        tk.Label(fr2, text="例: A:C,E", fg="gray").pack(side="left", padx=4)

        FileSelectRow(
            g, "マッチャーJSON", self._matchers,
            filetypes=[("JSON", "*.json"), ("All", "*.*")],
        ).pack(fill="x", **pad)

        tk.Checkbutton(
            g, text="取り消し線も差分として扱う", variable=self._strike,
        ).pack(anchor="w", **pad)
        tk.Checkbutton(
            g, text="完了後ブラウザで開く", variable=self._open_br,
        ).pack(anchor="w", **pad)

    # ------------------------------------------------------------------ 状態保存

    def save_state(self) -> None:
        """現在のUI値を設定に書き戻す（ウィンドウを閉じる前に呼ばれる）。"""
        cfg.set_tab("file_diff", {
            "old_file":      self._old.get(),
            "new_file":      self._new.get(),
            "output":        self._out.get(),
            "sheet":         self._sheet.get(),
            "include_cols":  self._cols.get(),
            "matchers":      self._matchers.get(),
            "strikethrough": self._strike.get(),
            "open_browser":  self._open_br.get(),
            "diff_mode":     self._mode.get(),
            "key_cols":      self._key_cols.get(),
        })

    # ------------------------------------------------------------------ 実行

    def _run(self) -> None:
        old = self._old.get().strip()
        new = self._new.get().strip()
        if not old or not new:
            messagebox.showerror("エラー", "旧ファイルと新ファイルを指定してください")
            return
        if not os.path.isfile(old):
            messagebox.showerror("エラー", f"ファイルが見つかりません:\n{old}")
            return
        if not os.path.isfile(new):
            messagebox.showerror("エラー", f"ファイルが見つかりません:\n{new}")
            return
        if self._mode.get() == "key" and not self._key_cols.get().strip():
            messagebox.showerror("エラー", "キーJOINモード: キー列を指定してください")
            return

        cfg.set_tab("file_diff", {
            "old_file": old, "new_file": new,
            "output": self._out.get(),
            "sheet": self._sheet.get(),
            "include_cols": self._cols.get(),
            "matchers": self._matchers.get(),
            "strikethrough": self._strike.get(),
            "open_browser": self._open_br.get(),
            "diff_mode": self._mode.get(),
            "key_cols": self._key_cols.get(),
        })
        cfg.save()

        self._btn_run.config(state="disabled", text="実行中...")
        self._log("ファイル比較を開始します...")

        self._result_q = get_worker().submit(
            self._do_diff,
            old, new,
            self._out.get().strip(), self._sheet.get().strip(),
            self._cols.get().strip(), self._matchers.get().strip(),
            self._strike.get(), self._mode.get(),
            self._key_cols.get().strip(), self._open_br.get(),
        )
        self.after(100, self._poll)

    def _poll(self) -> None:
        if self._result_q is None:
            return
        try:
            status, val = self._result_q.get_nowait()
            self._btn_run.config(state="normal", text="実行")
            if status == "err":
                self._log(f"エラー: {val}")
        except queue.Empty:
            self.after(100, self._poll)

    def _do_diff(
        self, old_file, new_file, output, sheet, include_cols,
        matchers_file, strikethrough, diff_mode, key_cols_str, open_browser,
    ) -> None:
        from excel_diff.reader import read_workbook
        from excel_diff.diff_engine import diff_files, RowTag
        from excel_diff.html_renderer import render
        from excel_diff.matcher import DiffConfig, parse_col_spec, parse_col_list, load_config
        from openpyxl.utils import get_column_letter

        self._log(f"読み込み中: {old_file}")
        old_sheets = read_workbook(old_file, strikethrough, sheet or None)
        self._log(f"読み込み中: {new_file}")
        new_sheets = read_workbook(new_file, strikethrough, sheet or None)

        if matchers_file and os.path.isfile(matchers_file):
            config = load_config(matchers_file)
            self._log(f"マッチャー: {len(config.matchers)} 件ロード")
        else:
            config = DiffConfig()

        if include_cols:
            try:
                config.global_col_filter = parse_col_spec(include_cols)
            except Exception as e:
                self._log(f"警告: 比較列の解析エラー: {e}")

        if diff_mode == "key":
            config.key_cols = parse_col_list(key_cols_str)
            config.diff_mode = "key"
            disp = ", ".join(get_column_letter(c + 1) for c in config.key_cols)
            self._log(f"差分モード: キーJOIN  キー列: {disp}")
        else:
            config.diff_mode = "lcs"
            self._log("差分モード: LCS（行の出現順）")

        self._log("差分計算中...")
        file_diff = diff_files(
            old_sheets, new_sheets, old_file, new_file,
            include_strike=strikethrough, config=config,
        )

        out_path = output or str(Path(new_file).parent / f"{Path(new_file).stem}_diff.html")
        Path(out_path).write_text(render(file_diff), encoding="utf-8")

        if file_diff.has_differences:
            delete = sum(1 for sd in file_diff.sheet_diffs
                         for rd in sd.row_diffs if rd.tag == RowTag.DELETE)
            insert = sum(1 for sd in file_diff.sheet_diffs
                         for rd in sd.row_diffs if rd.tag == RowTag.INSERT)
            modify = sum(1 for sd in file_diff.sheet_diffs
                         for rd in sd.row_diffs if rd.tag == RowTag.MODIFY)
            self._log(f"差分あり (削除 {delete} 追加 {insert} 変更 {modify}) → {out_path}")
        else:
            self._log("差分なし")

        if open_browser:
            webbrowser.open(Path(out_path).resolve().as_uri())
