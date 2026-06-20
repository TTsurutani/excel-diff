"""タブ② フォルダ比較（ペアリング・比較実行）。"""
import os
import queue
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Callable, Optional

from excel_diff.utils import generate_output_dir

from . import settings as cfg
from .widgets import FileSelectRow
from .worker import get_worker


class TabPatterns(tk.Frame):

    def __init__(
        self,
        parent,
        log: Callable[[str], None],
        get_compare_options: Optional[Callable] = None,
        # 旧引数との後方互換（無視）
        switch_to_dir_diff: Optional[Callable] = None,
        get_dir_diff_options: Optional[Callable] = None,
    ) -> None:
        super().__init__(parent)
        self._log = log
        self._get_compare_options = get_compare_options or get_dir_diff_options
        self._result_q: "queue.Queue | None" = None
        self._pairs: list = []
        self._compare_open_browser = True
        self._rebuild_after_id = None
        self._patterns: list = []
        self._sort_col: str = ""
        self._sort_rev: bool = False

        self._old_dir   = tk.StringVar(value=cfg.get("pair_build", "old_dir"))
        self._new_dir   = tk.StringVar(value=cfg.get("pair_build", "new_dir"))
        self._pairing   = tk.StringVar(value=cfg.get("pair_build", "pairing", "exact"))
        self._pairs_f   = tk.StringVar(value=cfg.get("pair_build", "pairs_file"))
        self._pat_regex = tk.StringVar(value=cfg.get("pair_build", "pattern_regex", ""))

        self._build()

        self._old_dir.trace_add(  "write", lambda *_: self._schedule_rebuild())
        self._new_dir.trace_add(  "write", lambda *_: self._schedule_rebuild())
        self._pairing.trace_add(  "write", lambda *_: self._on_pairing_change())
        self._pairs_f.trace_add(  "write", lambda *_: self._schedule_rebuild())
        self._pat_regex.trace_add("write", lambda *_: self._schedule_rebuild())

        self._reload_patterns()
        self._on_pairing_change()
        self._rebuild_pairs_now()
        self._refresh_list()

    # ================================================================== レイアウト

    def _build(self) -> None:
        pad = {"padx": 6, "pady": 3}

        # ── フォルダ選択 ───────────────────────────────────────────────
        grp_folders = tk.LabelFrame(self, text="フォルダ")
        grp_folders.pack(fill="x", **pad)
        FileSelectRow(grp_folders, "旧フォルダ", self._old_dir, mode="dir").pack(
            fill="x", padx=6, pady=2)
        FileSelectRow(grp_folders, "新フォルダ", self._new_dir, mode="dir").pack(
            fill="x", padx=6, pady=2)

        # ── ペアリング方法 ─────────────────────────────────────────────
        grp_pairing = tk.LabelFrame(self, text="ペアリング方法")
        grp_pairing.pack(fill="x", **pad)

        tk.Radiobutton(
            grp_pairing,
            text="完全一致（同名ファイルを対応付ける・デフォルト）",
            variable=self._pairing, value="exact",
        ).pack(anchor="w", padx=8, pady=(4, 0))

        # ペアJSON 行
        fr_pairs = tk.Frame(grp_pairing)
        fr_pairs.pack(anchor="w", padx=8, pady=(2, 0))
        tk.Radiobutton(
            fr_pairs, text="ペアJSON",
            variable=self._pairing, value="pairs",
        ).pack(side="left")
        tk.Label(fr_pairs, text="ファイル").pack(side="left", padx=(6, 2))
        self._entry_pairs = tk.Entry(fr_pairs, textvariable=self._pairs_f, width=28)
        self._entry_pairs.pack(side="left")
        self._btn_pairs = tk.Button(
            fr_pairs, text="参照", width=5, command=self._browse_pairs,
        )
        self._btn_pairs.pack(side="left", padx=2)

        # パターン 行（正規表現直接入力 + 既存読込コンボ）
        fr_pat = tk.Frame(grp_pairing)
        fr_pat.pack(fill="x", padx=8, pady=(2, 0))
        tk.Radiobutton(
            fr_pat, text="パターン",
            variable=self._pairing, value="pattern",
        ).pack(side="left")
        tk.Label(fr_pat, text="正規表現").pack(side="left", padx=(6, 2))
        self._entry_pat_regex = tk.Entry(fr_pat, textvariable=self._pat_regex, width=32)
        self._entry_pat_regex.pack(side="left", padx=(0, 8))
        tk.Label(fr_pat, text="既存から読込:").pack(side="left")
        self._cmb_pat = ttk.Combobox(fr_pat, width=18, state="disabled")
        self._cmb_pat.pack(side="left", padx=2)
        self._cmb_pat.bind("<<ComboboxSelected>>", self._on_pattern_select)

        # ウィザード 行
        fr_wiz_radio = tk.Frame(grp_pairing)
        fr_wiz_radio.pack(anchor="w", padx=8, pady=(2, 0))
        tk.Radiobutton(
            fr_wiz_radio, text="ウィザード（ファイル名の類似度で自動探索）",
            variable=self._pairing, value="wizard",
        ).pack(side="left")

        fr_wiz_ctrl = tk.Frame(grp_pairing)
        fr_wiz_ctrl.pack(fill="x", padx=24, pady=(2, 4))
        tk.Label(fr_wiz_ctrl, text="しきい値", width=8, anchor="w").pack(side="left")
        self._s1_thr = tk.DoubleVar(value=0.30)
        self._s1_thr_lbl = tk.Label(fr_wiz_ctrl, text="0.30", width=5)
        self._s1_thr_lbl.pack(side="right")
        tk.Label(fr_wiz_ctrl, text="1.0", fg="gray").pack(side="right")
        tk.Scale(
            fr_wiz_ctrl, variable=self._s1_thr, from_=0.0, to=1.0,
            resolution=0.05, orient="horizontal", showvalue=False,
            command=lambda v: self._s1_thr_lbl.config(text=f"{float(v):.2f}"),
        ).pack(side="left", fill="x", expand=True)
        tk.Label(fr_wiz_ctrl, text="0.0", fg="gray").pack(side="left")
        self._btn_discover = tk.Button(
            fr_wiz_ctrl, text="探索実行", width=10,
            bg="#4a9eff", fg="white", font=("", 9, "bold"),
            command=self._run_discover,
        )
        self._btn_discover.pack(side="right", padx=(8, 0))

        # ── ボタン行・注記（先に bottom pack して展開エリアに押しつぶされないようにする）
        btn_row = tk.Frame(self)
        btn_row.pack(side="bottom", fill="x", padx=8, pady=(2, 6))
        tk.Button(btn_row, text="JSON保存", command=self._save_pairs_json).pack(side="left")
        tk.Button(
            btn_row, text="パターン登録", command=self._open_register_dialog,
        ).pack(side="left", padx=8)
        self._btn_compare = tk.Button(
            btn_row, text="比較実行", width=14,
            bg="#4a9eff", fg="white", font=("", 10, "bold"),
            command=self._run_compare_pairs,
        )
        self._btn_compare.pack(side="right")

        tk.Label(
            self,
            text="※「旧のみ」「新のみ」の行は比較対象外として扱われます",
            fg="gray", font=("", 8),
        ).pack(side="bottom", anchor="w", padx=8)

        # ── PanedWindow: ペアリスト（上）/ 保存済みパターン一覧（下）─────
        paned = ttk.PanedWindow(self, orient="vertical")
        paned.pack(fill="both", expand=True, padx=4, pady=4)

        # 上段: ペアリスト
        grp_pairs = tk.LabelFrame(paned, text="ペアリスト")
        paned.add(grp_pairs, weight=2)

        self._summary_var = tk.StringVar(value="")
        tk.Label(grp_pairs, textvariable=self._summary_var,
                 font=("", 8), fg="#444444", anchor="w").pack(
            fill="x", padx=6, pady=(2, 0))

        cols = ("old", "new", "score", "kind")
        self._tree_pairs = ttk.Treeview(
            grp_pairs, columns=cols, show="headings", height=6, selectmode="browse",
        )
        for col, w in zip(cols, (190, 190, 60, 80)):
            self._tree_pairs.heading(col, text=self._PAIR_HEADS[col],
                                     command=lambda c=col: self._on_pair_heading_click(c))
            self._tree_pairs.column(col, width=w, anchor="w")
        self._tree_pairs.tag_configure("unmatched", foreground="#888888")
        sb2 = ttk.Scrollbar(grp_pairs, orient="vertical", command=self._tree_pairs.yview)
        self._tree_pairs.configure(yscrollcommand=sb2.set)
        self._tree_pairs.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=4)
        sb2.pack(side="left", fill="y", pady=4)

        # 下段: 保存済みパターン一覧
        grp_list = tk.LabelFrame(paned, text="保存済みパターン一覧")
        paned.add(grp_list, weight=1)

        cols2 = ("id", "name", "regex", "created_at")
        self._tree_list = ttk.Treeview(
            grp_list, columns=cols2, show="headings", height=4, selectmode="browse",
        )
        for col, head, w in zip(cols2, ("ID", "名前", "正規表現", "作成日"),
                                 (100, 120, 260, 90)):
            self._tree_list.heading(col, text=head)
            self._tree_list.column(col, width=w, anchor="w")
        sb = ttk.Scrollbar(grp_list, orient="vertical", command=self._tree_list.yview)
        self._tree_list.configure(yscrollcommand=sb.set)
        self._tree_list.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=4)
        sb.pack(side="left", fill="y", pady=4)

        btn_fr = tk.Frame(grp_list)
        btn_fr.pack(side="left", padx=6, pady=4, anchor="n")
        tk.Button(btn_fr, text="編集", width=8, command=self._edit_pattern).pack(pady=2)
        tk.Button(btn_fr, text="削除", width=8, command=self._delete_pattern).pack(pady=2)

    # ================================================================== ペアリング

    def _browse_pairs(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("JSON", "*.json"), ("All", "*.*")],
        )
        if path:
            self._pairs_f.set(path)

    def _reload_all(self) -> None:
        self._reload_patterns()
        self._refresh_list()

    def _reload_patterns(self) -> None:
        try:
            from excel_diff.patterns import PatternStore
            store = PatternStore(cfg.patterns_file())
            self._patterns = store.list_all()
        except Exception as e:
            self._log(f"パターン読み込みエラー: {e}")
            self._patterns = []
        self._cmb_pat["values"] = [f"{p.id}  {p.name}" for p in self._patterns]

    def _on_pattern_select(self, event=None) -> None:
        """コンボから既存パターンを選択したとき正規表現フィールドへ展開する。"""
        sel = self._cmb_pat.get().strip()
        if not sel:
            return
        pat_id = sel.split()[0]
        pat = next((p for p in self._patterns if p.id == pat_id), None)
        if pat:
            self._pat_regex.set(pat.key_regex)

    def _on_pairing_change(self) -> None:
        method = self._pairing.get()
        self._entry_pairs.config(    state="normal"   if method == "pairs"   else "disabled")
        self._btn_pairs.config(      state="normal"   if method == "pairs"   else "disabled")
        self._entry_pat_regex.config(state="normal"   if method == "pattern" else "disabled")
        self._cmb_pat.config(        state="readonly" if method == "pattern" else "disabled")
        self._btn_discover.config(   state="normal"   if method == "wizard"  else "disabled")
        if method != "wizard":
            self._schedule_rebuild()

    def _schedule_rebuild(self) -> None:
        """200ms デバウンスでペアリスト再構築。"""
        if self._rebuild_after_id is not None:
            try:
                self.after_cancel(self._rebuild_after_id)
            except Exception:
                pass
        self._rebuild_after_id = self.after(200, self._rebuild_pairs_now)

    def _rebuild_pairs_now(self) -> None:
        self._rebuild_after_id = None
        method = self._pairing.get()
        if method == "wizard":
            return

        old = self._old_dir.get().strip()
        new = self._new_dir.get().strip()

        if not old or not new or not os.path.isdir(old) or not os.path.isdir(new):
            self._pairs = []
            self._populate_pairs()
            return

        try:
            if method == "exact":
                from excel_diff.file_pairing import FilePair
                old_files = {
                    f for f in os.listdir(old)
                    if f.lower().endswith(".xlsx") and not f.startswith("~$")
                }
                new_files = {
                    f for f in os.listdir(new)
                    if f.lower().endswith(".xlsx") and not f.startswith("~$")
                }
                all_names = sorted(old_files | new_files)
                self._pairs = [
                    FilePair(
                        old_name=name if name in old_files else None,
                        new_name=name if name in new_files else None,
                        score=1.0, matched_by="exact",
                    )
                    for name in all_names
                ]

            elif method == "pairs":
                pf = self._pairs_f.get().strip()
                if not pf or not os.path.isfile(pf):
                    self._pairs = []
                else:
                    from excel_diff.file_pairing import load_pairs
                    self._pairs = load_pairs(pf)

            elif method == "pattern":
                regex = self._pat_regex.get().strip()
                if not regex:
                    self._pairs = []
                else:
                    from excel_diff.file_pairing import apply_pattern
                    self._pairs = apply_pattern(old, new, regex)

        except Exception as e:
            self._log(f"ペアリスト構築エラー: {e}")
            self._pairs = []

        self._populate_pairs()

    _KIND_MAP = {
        "exact": "完全一致", "auto": "自動", "pattern": "パターン",
        "unmatched_old": "旧のみ", "unmatched_new": "新のみ",
    }
    _PAIR_HEADS = {"old": "旧ファイル", "new": "新ファイル", "score": "スコア", "kind": "種別"}

    def _on_pair_heading_click(self, col: str) -> None:
        if self._sort_col == col:
            self._sort_rev = not self._sort_rev
        else:
            self._sort_col = col
            self._sort_rev = False
        self._populate_pairs()

    def _sorted_pairs(self) -> list:
        if not self._sort_col:
            return self._pairs
        col, rev = self._sort_col, self._sort_rev
        matched   = [p for p in self._pairs if p.old_name and p.new_name]
        unmatched = [p for p in self._pairs if not p.old_name or not p.new_name]
        if col == "old":
            key = lambda p: (p.old_name or "").lower()
        elif col == "new":
            key = lambda p: (p.new_name or "").lower()
        elif col == "score":
            key = lambda p: p.score
        else:
            key = lambda p: self._KIND_MAP.get(p.matched_by, p.matched_by)
        um_key = (lambda p: (p.old_name or p.new_name or "").lower()) if col in ("old", "new") else key
        return sorted(matched, key=key, reverse=rev) + sorted(unmatched, key=um_key, reverse=rev)

    def _populate_pairs(self) -> None:
        for row in self._tree_pairs.get_children():
            self._tree_pairs.delete(row)
        for col, head in self._PAIR_HEADS.items():
            ind = (" ▼" if self._sort_rev else " ▲") if col == self._sort_col else ""
            self._tree_pairs.heading(col, text=head + ind)
        paired   = [p for p in self._pairs if p.old_name and p.new_name]
        only_old = sum(1 for p in self._pairs if p.old_name and not p.new_name)
        only_new = sum(1 for p in self._pairs if p.new_name and not p.old_name)
        sc1      = sum(1 for p in paired if p.score >= 1.0)
        self._summary_var.set(
            f"旧 {len(paired) + only_old}件  新 {len(paired) + only_new}件  "
            f"ペア {len(paired)}件（スコア=1.0: {sc1}件・<1.0: {len(paired) - sc1}件）  "
            f"旧のみ {only_old}件  新のみ {only_new}件"
        )
        for i, p in enumerate(self._sorted_pairs()):
            old_disp   = p.old_name or "（なし）"
            new_disp   = p.new_name or "（なし）"
            score_disp = f"{p.score:.2f}" if p.score > 0 else "-"
            kind_disp  = self._KIND_MAP.get(p.matched_by, p.matched_by)
            tags = ("unmatched",) if not p.old_name or not p.new_name else ()
            self._tree_pairs.insert(
                "", "end", iid=str(i),
                values=(old_disp, new_disp, score_disp, kind_disp),
                tags=tags,
            )

    # ================================================================== ウィザード探索

    def _run_discover(self) -> None:
        old = self._old_dir.get().strip()
        new = self._new_dir.get().strip()
        if not old or not new:
            messagebox.showerror("エラー", "旧フォルダと新フォルダを指定してください")
            return
        if not os.path.isdir(old):
            messagebox.showerror("エラー", f"フォルダが見つかりません:\n{old}")
            return
        if not os.path.isdir(new):
            messagebox.showerror("エラー", f"フォルダが見つかりません:\n{new}")
            return
        self._btn_discover.config(state="disabled", text="探索中...")
        self._log(f"ペア候補を探索中: {old} / {new}")
        self._result_q = get_worker().submit(
            self._do_discover, old, new, self._s1_thr.get(),
        )
        self.after(100, self._poll_discover)

    def _do_discover(self, old_dir: str, new_dir: str, threshold: float) -> list:
        from excel_diff.file_pairing import discover_pairs
        return discover_pairs(old_dir, new_dir, threshold)

    def _poll_discover(self) -> None:
        if self._result_q is None:
            return
        try:
            status, val = self._result_q.get_nowait()
            self._btn_discover.config(state="normal", text="探索実行")
            if status == "err":
                self._log(f"探索エラー: {val}")
            else:
                self._pairs = val
                self._log(f"探索完了: {len(val)} ペア候補")
                self._populate_pairs()
        except queue.Empty:
            self.after(100, self._poll_discover)

    # ================================================================== JSON保存

    def _save_pairs_json(self) -> None:
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json"), ("All", "*.*")],
        )
        if not path:
            return
        try:
            from excel_diff.file_pairing import save_pairs
            save_pairs(self._pairs, path)
            self._log(f"ペアJSON保存: {path}（ペアリング方法「ペアJSON」で再利用可）")
        except Exception as e:
            self._log(f"保存エラー: {e}")

    # ================================================================== パターン登録

    def _open_register_dialog(self) -> None:
        """現在の正規表現にID・名前を付けて patterns.json に保存するダイアログを開く。"""
        regex = self._pat_regex.get().strip()
        if not regex:
            messagebox.showerror("エラー", "パターン登録するには正規表現を入力してください")
            return

        dlg = tk.Toplevel(self)
        dlg.title("パターン登録")
        dlg.resizable(False, False)
        dlg.transient(self.winfo_toplevel())
        dlg.grab_set()

        frm = tk.Frame(dlg, padx=14, pady=10)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)

        # 正規表現（読み取り専用表示）
        tk.Label(frm, text="正規表現", anchor="w").grid(
            row=0, column=0, sticky="w", pady=4)
        tk.Label(frm, text=regex, relief="groove", anchor="w",
                 font=("Courier", 9), width=44).grid(
            row=0, column=1, sticky="ew", padx=4, pady=4)

        id_var   = tk.StringVar()
        name_var = tk.StringVar()
        desc_var = tk.StringVar()

        for row, (label, var, w) in enumerate(
            (("ID",   id_var,   16),
             ("名前", name_var, 28),
             ("説明", desc_var, 40)),
            start=1,
        ):
            tk.Label(frm, text=label, anchor="w").grid(
                row=row, column=0, sticky="w", pady=4)
            tk.Entry(frm, textvariable=var, width=w).grid(
                row=row, column=1, sticky="ew", padx=4, pady=4)

        btn_frm = tk.Frame(frm)
        btn_frm.grid(row=4, column=0, columnspan=2, pady=(10, 0))

        def do_save() -> None:
            pat_id   = id_var.get().strip()
            pat_name = name_var.get().strip()
            if not pat_id:
                messagebox.showerror("エラー", "IDを入力してください", parent=dlg)
                return
            if not pat_name:
                messagebox.showerror("エラー", "名前を入力してください", parent=dlg)
                return
            try:
                from excel_diff.patterns import PatternStore, PatternDef
                from datetime import date
                store = PatternStore(cfg.patterns_file())
                if store.get(pat_id):
                    if not messagebox.askyesno(
                        "確認", f"パターン「{pat_id}」は既に存在します。上書きしますか？",
                        parent=dlg,
                    ):
                        return
                store.add_or_update(PatternDef(
                    id=pat_id,
                    name=pat_name,
                    key_regex=regex,
                    description=desc_var.get().strip(),
                    example_old_dir=self._old_dir.get(),
                    example_new_dir=self._new_dir.get(),
                    created_at=date.today().isoformat(),
                ))
                store.save()
                self._log(f"パターン保存: [{pat_id}] {pat_name}  regex={regex}")
                self._reload_all()
                dlg.destroy()
            except Exception as e:
                self._log(f"保存エラー: {e}")

        tk.Button(
            btn_frm, text="保存", command=do_save,
            bg="#4a9eff", fg="white", font=("", 9, "bold"), width=10,
        ).pack(side="left", padx=4)
        tk.Button(
            btn_frm, text="キャンセル", command=dlg.destroy, width=10,
        ).pack(side="left", padx=4)

        dlg.wait_window()

    # ================================================================== 比較実行

    def _run_compare_pairs(self) -> None:
        matched = [p for p in self._pairs if p.old_name and p.new_name]
        if not matched:
            messagebox.showinfo("情報", "比較可能なペアがありません")
            return

        old = self._old_dir.get().strip()
        new = self._new_dir.get().strip()

        options = (
            self._get_compare_options()
            if self._get_compare_options is not None
            else cfg.data("dir_diff")
        )
        self._compare_open_browser = options.get("open_browser", True)
        unmatched = [p for p in self._pairs if not p.old_name or not p.new_name]

        self._log(f"比較実行: {len(matched)} 件")
        self._btn_compare.config(state="disabled", text="実行中...")
        self._result_q = get_worker().submit(
            self._do_compare_pairs, matched, unmatched, old, new, options,
        )
        self.after(100, self._poll_compare)

    def _do_compare_pairs(self, matched, unmatched, old_dir, new_dir, options: dict):
        from excel_diff.reader import read_workbook, filter_sheets_by_pattern
        from excel_diff.diff_engine import diff_files
        from excel_diff.html_renderer import render
        from excel_diff.matcher import DiffConfig, parse_col_spec, parse_col_list, load_config
        from excel_diff.__main__ import _render_index_html, _write_index_xlsx

        sheet_old_pat  = options.get("sheet_old") or None
        sheet_new_pat  = options.get("sheet_new") or None
        strikethrough  = options.get("strikethrough", False)
        include_cols   = options.get("include_cols", "")
        matchers_file  = options.get("matchers", "")
        diff_mode      = options.get("diff_mode", "lcs")
        key_cols_str   = options.get("key_cols", "")
        output_dir_opt = options.get("output_dir", "")

        if matchers_file and os.path.isfile(matchers_file):
            config = load_config(matchers_file)
        else:
            config = DiffConfig()
        if include_cols:
            try:
                config.global_col_filter = parse_col_spec(include_cols)
            except Exception:
                pass
        if diff_mode == "key" and key_cols_str:
            config.key_cols = parse_col_list(key_cols_str)
            config.diff_mode = "key"
        else:
            config.diff_mode = "lcs"

        from openpyxl.utils import get_column_letter as _gcl
        log_lines = ["─" * 36]
        log_lines.append(f"[実行条件] 旧: {old_dir}")
        log_lines.append(f"[実行条件] 新: {new_dir}")
        if config.global_col_filter is not None:
            col_letters = sorted(_gcl(i + 1) for i in config.global_col_filter)
            log_lines.append(f"[実行条件] 比較列: {', '.join(col_letters)}  (raw='{include_cols}')")
        else:
            log_lines.append(f"[実行条件] 比較列: 全列  (raw='{include_cols}')")
        if config.diff_mode == "key" and config.key_cols:
            key_letters = ", ".join(_gcl(c + 1) for c in config.key_cols)
            log_lines.append(f"[実行条件] キー列: {key_letters}")
        log_lines.append(f"[実行条件] 差分モード: {config.diff_mode}")
        sheet_log = []
        if sheet_old_pat:
            sheet_log.append(f"旧={sheet_old_pat}")
        if sheet_new_pat:
            sheet_log.append(f"新={sheet_new_pat}")
        log_lines.append(f"[実行条件] シート: {', '.join(sheet_log) if sheet_log else '全シート'}")
        log_lines.append("─" * 36)
        for line in log_lines:
            self._log(line)

        out_dir = output_dir_opt or generate_output_dir(
            old_dir, new_dir, base_dir=str(Path(new_dir).parent / "diff")
        )
        Path(out_dir).mkdir(parents=True, exist_ok=True)

        results = []
        skipped = []
        for pair in matched:
            old_path = os.path.join(old_dir, pair.old_name)
            new_path = os.path.join(new_dir, pair.new_name)
            try:
                old_sheets = read_workbook(old_path, strikethrough)
                if sheet_old_pat:
                    old_sheets = filter_sheets_by_pattern(old_sheets, sheet_old_pat)
                    if not old_sheets:
                        self._log(f"  ⚠ スキップ ({pair.old_name}): 旧シートパターン '{sheet_old_pat}' にマッチなし")
                        skipped.append(pair.old_name)
                        continue
                new_sheets = read_workbook(new_path, strikethrough)
                if sheet_new_pat:
                    new_sheets = filter_sheets_by_pattern(new_sheets, sheet_new_pat)
                    if not new_sheets:
                        self._log(f"  ⚠ スキップ ({pair.new_name}): 新シートパターン '{sheet_new_pat}' にマッチなし")
                        skipped.append(pair.old_name)
                        continue
                fd = diff_files(
                    old_sheets, new_sheets, old_path, new_path,
                    include_strike=strikethrough, config=config,
                )
                out_path = os.path.join(out_dir, f"{Path(pair.new_name).stem}_diff.html")
                Path(out_path).write_text(render(fd), encoding="utf-8")
                results.append((pair, fd, out_path))
            except Exception as e:
                self._log(f"  ⚠ スキップ ({pair.old_name}): {e}")
                skipped.append(pair.old_name)

        # index_path = os.path.join(out_dir, "★index.html")
        # Path(index_path).write_text(
        #     _render_index_html(results, unmatched, old_dir, new_dir), encoding="utf-8",
        # )
        index_xlsx_path = os.path.join(out_dir, "★index.xlsx")
        _write_index_xlsx(results, unmatched, old_dir, new_dir, index_xlsx_path)
        return index_xlsx_path, skipped

    def _poll_compare(self) -> None:
        if self._result_q is None:
            return
        try:
            status, val = self._result_q.get_nowait()
            self._btn_compare.config(state="normal", text="比較実行")
            if status == "err":
                self._log(f"比較エラー: {val}")
            else:
                index_path, skipped = val
                if skipped:
                    for name in skipped:
                        self._log(f"  ⚠ スキップ: {name}（無効な xlsx）")
                self._log(f"比較完了 → {index_path}")
                if self._compare_open_browser:
                    os.startfile(index_path)
        except queue.Empty:
            self.after(100, self._poll_compare)

    # ================================================================== パターン一覧

    def _refresh_list(self) -> None:
        for row in self._tree_list.get_children():
            self._tree_list.delete(row)
        try:
            from excel_diff.patterns import PatternStore
            for p in PatternStore(cfg.patterns_file()).list_all():
                self._tree_list.insert("", "end", iid=p.id,
                                       values=(p.id, p.name, p.key_regex, p.created_at))
        except Exception as e:
            self._log(f"パターン一覧の読み込みエラー: {e}")

    def _edit_pattern(self) -> None:
        sel = self._tree_list.selection()
        if not sel:
            messagebox.showinfo("情報", "編集するパターンを選択してください")
            return
        pat_id = sel[0]
        pat = next((p for p in self._patterns if p.id == pat_id), None)
        if pat is None:
            messagebox.showerror("エラー", f"パターン「{pat_id}」が見つかりません")
            return

        dlg = tk.Toplevel(self)
        dlg.title("パターン編集")
        dlg.resizable(False, False)
        dlg.transient(self.winfo_toplevel())
        dlg.grab_set()

        frm = tk.Frame(dlg, padx=14, pady=10)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)

        tk.Label(frm, text="ID", anchor="w").grid(row=0, column=0, sticky="w", pady=4)
        tk.Label(frm, text=pat.id, relief="groove", anchor="w",
                 font=("Courier", 9), width=16).grid(
            row=0, column=1, sticky="ew", padx=4, pady=4)

        name_var  = tk.StringVar(value=pat.name)
        regex_var = tk.StringVar(value=pat.key_regex)
        desc_var  = tk.StringVar(value=pat.description)

        for row, (label, var, w) in enumerate(
            (("名前",     name_var,  28),
             ("正規表現", regex_var, 40),
             ("説明",     desc_var,  40)),
            start=1,
        ):
            tk.Label(frm, text=label, anchor="w").grid(
                row=row, column=0, sticky="w", pady=4)
            tk.Entry(frm, textvariable=var, width=w).grid(
                row=row, column=1, sticky="ew", padx=4, pady=4)

        btn_frm = tk.Frame(frm)
        btn_frm.grid(row=4, column=0, columnspan=2, pady=(10, 0))

        def do_save() -> None:
            new_name  = name_var.get().strip()
            new_regex = regex_var.get().strip()
            if not new_name:
                messagebox.showerror("エラー", "名前を入力してください", parent=dlg)
                return
            if not new_regex:
                messagebox.showerror("エラー", "正規表現を入力してください", parent=dlg)
                return
            try:
                from excel_diff.patterns import PatternStore, PatternDef
                store = PatternStore(cfg.patterns_file())
                store.add_or_update(PatternDef(
                    id=pat.id,
                    name=new_name,
                    key_regex=new_regex,
                    description=desc_var.get().strip(),
                    example_old_dir=pat.example_old_dir,
                    example_new_dir=pat.example_new_dir,
                    created_at=pat.created_at,
                ))
                store.save()
                self._log(f"パターン更新: [{pat.id}] {new_name}  regex={new_regex}")
                self._reload_all()
                dlg.destroy()
            except Exception as e:
                self._log(f"更新エラー: {e}")

        tk.Button(
            btn_frm, text="保存", command=do_save,
            bg="#4a9eff", fg="white", font=("", 9, "bold"), width=10,
        ).pack(side="left", padx=4)
        tk.Button(
            btn_frm, text="キャンセル", command=dlg.destroy, width=10,
        ).pack(side="left", padx=4)

        dlg.wait_window()

    def _delete_pattern(self) -> None:
        sel = self._tree_list.selection()
        if not sel:
            messagebox.showinfo("情報", "削除するパターンを選択してください")
            return
        pat_id = sel[0]
        if not messagebox.askyesno("確認", f"パターン「{pat_id}」を削除しますか？"):
            return
        try:
            from excel_diff.patterns import PatternStore
            store = PatternStore(cfg.patterns_file())
            store._patterns = [p for p in store._patterns if p.id != pat_id]
            store.save()
            self._log(f"パターン削除: {pat_id}")
            self._reload_all()
        except Exception as e:
            self._log(f"削除エラー: {e}")

    # ================================================================== 状態保存

    def get_snapshot(self) -> dict:
        """現在の UI 値をすべて dict で返す（プロファイル保存用）。"""
        return {
            "old_dir":       self._old_dir.get(),
            "new_dir":       self._new_dir.get(),
            "pairing":       self._pairing.get(),
            "pairs_file":    self._pairs_f.get(),
            "pattern_regex": self._pat_regex.get(),
        }

    def load_from_snapshot(self, snap: dict) -> None:
        """スナップショットの値を各変数に反映する（プロファイル読み込み用）。"""
        mapping = {
            "old_dir":       self._old_dir,
            "new_dir":       self._new_dir,
            "pairing":       self._pairing,
            "pairs_file":    self._pairs_f,
            "pattern_regex": self._pat_regex,
        }
        for key, var in mapping.items():
            if key in snap:
                var.set(snap[key])
        # _pairing の trace が _on_pairing_change を自動発火する

    def save_state(self) -> None:
        """現在のUI値を設定に書き戻す（ウィンドウを閉じる前に呼ばれる）。"""
        cfg.set_tab("pair_build", self.get_snapshot())
