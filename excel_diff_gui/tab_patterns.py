"""タブ② フォルダ比較（ペアリング・比較実行）。"""
import os
import queue
import re
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Callable, Optional

from . import settings as cfg
from .widgets import FileSelectRow
from .worker import get_worker


_TEMPLATES = [
    ("日付8桁（例: report_20240101.xlsx）", r"^(.+?)_\d{8}\.xlsx$"),
    ("日付6桁（例: report_202401.xlsx）",   r"^(.+?)_\d{6}\.xlsx$"),
    ("バージョン番号（例: report_v2.xlsx）", r"^(.+?)_v\d+\.xlsx$"),
    ("連番（例: report_001.xlsx）",         r"^(.+?)_\d+\.xlsx$"),
    ("手動入力",                            ""),
]

_ERR_MSG = {
    "invalid_regex": "正規表現の文法エラー。括弧の対応や記号のエスケープを確認してください。",
    "no_match":      "マッチしないファイルがあります。ファイル名のパターンと正規表現が合っていません。テンプレートを変えて試してください。",
    "key_mismatch":  "旧と新でキーが一致しないペアがあります。キャプチャ範囲が変動部分（日付等）を含んでいる可能性があります。",
    "key_collision": "同一キーに複数ファイルがマッチしています。正規表現が広すぎます。",
}


# ────────────────────────────────────────────── パターン生成ヘルパー

def _lcp_len(a: str, b: str) -> int:
    n = 0
    for x, y in zip(a, b):
        if x == y:
            n += 1
        else:
            break
    return n


def _lcs_len(a: str, b: str) -> int:
    n = 0
    for x, y in zip(reversed(a), reversed(b)):
        if x == y:
            n += 1
        else:
            break
    return n


def _classify_var(v: str) -> str:
    if re.fullmatch(r"\d{8}", v):  return r"\d{8}"
    if re.fullmatch(r"\d{6}", v):  return r"\d{6}"
    if re.fullmatch(r"v\d+",  v):  return r"v\d+"
    if re.fullmatch(r"\d+",   v):  return r"\d+"
    return r".+"


# ──────────────────────────────────────────────────────────────────────

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
        self._validated_ok = False
        self._compare_open_browser = True
        self._rebuild_after_id = None
        self._patterns: list = []

        # フォルダ選択
        self._old_dir = tk.StringVar(value=cfg.get("pair_build", "old_dir"))
        self._new_dir = tk.StringVar(value=cfg.get("pair_build", "new_dir"))

        # ペアリング方法
        self._pairing = tk.StringVar(value=cfg.get("pair_build", "pairing", "exact"))
        self._pairs_f = tk.StringVar(value=cfg.get("pair_build", "pairs_file"))
        self._pat_id  = tk.StringVar(value=cfg.get("pair_build", "pattern_id"))

        self._build()

        # 変更トレース（ウィザード以外は自動でペアリスト再構築）
        self._old_dir.trace_add("write", lambda *_: self._schedule_rebuild())
        self._new_dir.trace_add("write", lambda *_: self._schedule_rebuild())
        self._pairing.trace_add("write", lambda *_: self._on_pairing_change())
        self._pairs_f.trace_add("write", lambda *_: self._schedule_rebuild())
        self._pat_id.trace_add("write",  lambda *_: self._schedule_rebuild())

        # 初期化
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

        # ペアJSON + パターン（同一行に2列並置）
        fr_2col = tk.Frame(grp_pairing)
        fr_2col.pack(fill="x", padx=8, pady=(2, 0))

        # 左列: ペアJSON
        fr_pairs = tk.Frame(fr_2col)
        fr_pairs.pack(side="left")
        tk.Radiobutton(
            fr_pairs, text="ペアJSON",
            variable=self._pairing, value="pairs",
        ).pack(side="left")
        tk.Label(fr_pairs, text="ファイル").pack(side="left", padx=(6, 2))
        self._entry_pairs = tk.Entry(fr_pairs, textvariable=self._pairs_f, width=22)
        self._entry_pairs.pack(side="left")
        self._btn_pairs = tk.Button(
            fr_pairs, text="参照", width=5, command=self._browse_pairs,
        )
        self._btn_pairs.pack(side="left", padx=2)

        tk.Label(fr_2col, text="  ").pack(side="left")  # 列間スペーサ

        # 右列: パターン
        fr_pat = tk.Frame(fr_2col)
        fr_pat.pack(side="left")
        tk.Radiobutton(
            fr_pat, text="パターン",
            variable=self._pairing, value="pattern",
        ).pack(side="left")
        tk.Label(fr_pat, text="名前").pack(side="left", padx=(6, 2))
        self._cmb_pat = ttk.Combobox(
            fr_pat, textvariable=self._pat_id, width=18, state="disabled",
        )
        self._cmb_pat.pack(side="left")
        tk.Button(
            fr_pat, text="更新", width=5, command=self._reload_patterns,
        ).pack(side="left", padx=2)

        # ウィザード行（ラジオ）
        fr_wiz_radio = tk.Frame(grp_pairing)
        fr_wiz_radio.pack(anchor="w", padx=8, pady=(2, 0))
        tk.Radiobutton(
            fr_wiz_radio, text="ウィザード（ファイル名の類似度で自動探索）",
            variable=self._pairing, value="wizard",
        ).pack(side="left")

        # ウィザード行（しきい値 + 探索実行）
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

        # ── PanedWindow: 上段＝ペアリスト/ウィザード、下段＝パターン一覧 ─
        paned = ttk.PanedWindow(self, orient="vertical")
        paned.pack(fill="both", expand=True, padx=4, pady=4)

        # 上段: 切り替えフレーム（ペアリスト ↔ パターン生成ウィザード）
        self._fr_switchable = tk.Frame(paned)
        paned.add(self._fr_switchable, weight=2)

        self._fr_main_view = tk.Frame(self._fr_switchable)
        self._build_main_view(self._fr_main_view)

        self._fr_step3_view = tk.Frame(self._fr_switchable)
        self._build_step3(self._fr_step3_view)

        # 下段: 保存済みパターン一覧
        grp_list = tk.LabelFrame(paned, text="保存済みパターン一覧")
        paned.add(grp_list, weight=1)

        cols = ("id", "name", "regex", "created_at")
        self._tree_list = ttk.Treeview(
            grp_list, columns=cols, show="headings", height=4, selectmode="browse",
        )
        for col, head, w in zip(cols, ("ID", "名前", "正規表現", "作成日"),
                                 (100, 120, 260, 90)):
            self._tree_list.heading(col, text=head)
            self._tree_list.column(col, width=w, anchor="w")
        sb = ttk.Scrollbar(grp_list, orient="vertical", command=self._tree_list.yview)
        self._tree_list.configure(yscrollcommand=sb.set)
        self._tree_list.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=4)
        sb.pack(side="left", fill="y", pady=4)

        btn_fr = tk.Frame(grp_list)
        btn_fr.pack(side="left", padx=6, pady=4, anchor="n")
        tk.Button(btn_fr, text="削除", width=8, command=self._delete_pattern).pack(pady=2)
        tk.Button(btn_fr, text="更新", width=8, command=self._refresh_list).pack(pady=2)

        # 初期表示
        self._show_main_view()

    # ------------------------------------------------------------------ メインビュー

    def _build_main_view(self, parent: tk.Frame) -> None:
        # ボタン行を先に side="bottom" で pack する。
        # こうしないと grp_pairs の expand=True が縦スペースを全部取り、
        # ボタンが不可視になる。
        btn_row = tk.Frame(parent)
        btn_row.pack(side="bottom", fill="x", padx=8, pady=(4, 8))
        tk.Button(btn_row, text="JSON保存", command=self._save_pairs_json).pack(side="left")
        tk.Button(
            btn_row, text="パターン生成 →", command=self._goto_step3,
        ).pack(side="left", padx=8)
        self._btn_compare = tk.Button(
            btn_row, text="比較実行", width=14,
            bg="#4a9eff", fg="white", font=("", 10, "bold"),
            command=self._run_compare_pairs,
        )
        self._btn_compare.pack(side="right")

        tk.Label(
            parent,
            text="※「旧のみ」「新のみ」の行は比較対象外として扱われます",
            fg="gray", font=("", 8),
        ).pack(side="bottom", anchor="w", padx=8)

        # ペアリストは最後に pack（expand=True で残りを埋める）
        grp_pairs = tk.LabelFrame(parent, text="ペアリスト")
        grp_pairs.pack(fill="both", expand=True, padx=4, pady=(4, 2))

        cols = ("old", "new", "score", "kind")
        self._tree_pairs = ttk.Treeview(
            grp_pairs, columns=cols, show="headings", height=6, selectmode="browse",
        )
        for col, head, w in zip(cols, ("旧ファイル", "新ファイル", "スコア", "種別"),
                                 (190, 190, 60, 80)):
            self._tree_pairs.heading(col, text=head)
            self._tree_pairs.column(col, width=w, anchor="w")
        self._tree_pairs.tag_configure("unmatched", foreground="#888888")

        sb2 = ttk.Scrollbar(grp_pairs, orient="vertical", command=self._tree_pairs.yview)
        self._tree_pairs.configure(yscrollcommand=sb2.set)
        self._tree_pairs.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=4)
        sb2.pack(side="left", fill="y", pady=4)

    def _show_main_view(self) -> None:
        self._fr_step3_view.pack_forget()
        self._fr_main_view.pack(fill="both", expand=True)

    def _show_step3_view(self) -> None:
        self._fr_main_view.pack_forget()
        self._fr_step3_view.pack(fill="both", expand=True)

    # ================================================================== ペアリング

    def _browse_pairs(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("JSON", "*.json"), ("All", "*.*")],
        )
        if path:
            self._pairs_f.set(path)

    def _reload_patterns(self) -> None:
        try:
            from excel_diff.patterns import PatternStore
            store = PatternStore(cfg.patterns_file())
            self._patterns = store.list_all()
        except Exception:
            self._patterns = []
        values = [f"{p.id}  {p.name}" for p in self._patterns]
        self._cmb_pat["values"] = values
        # 保存済みIDにマッチするエントリを復元
        saved_id = cfg.get("pair_build", "pattern_id", "")
        if saved_id:
            matching = [v for v in values if v.split()[0] == saved_id]
            if matching and not self._pat_id.get().strip():
                self._pat_id.set(matching[0])
        if not self._pat_id.get().strip() and values:
            self._pat_id.set(values[0])

    def _on_pairing_change(self) -> None:
        method = self._pairing.get()
        self._entry_pairs.config(state="normal"   if method == "pairs"   else "disabled")
        self._btn_pairs.config(  state="normal"   if method == "pairs"   else "disabled")
        self._cmb_pat.config(    state="readonly"  if method == "pattern" else "disabled")
        self._btn_discover.config(state="normal"  if method == "wizard"  else "disabled")
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
            return  # 探索実行ボタンでのみ更新

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
                pat_sel = self._pat_id.get().strip()
                pat_id = pat_sel.split()[0] if pat_sel else ""
                if not pat_id:
                    self._pairs = []
                else:
                    from excel_diff.patterns import PatternStore
                    from excel_diff.file_pairing import apply_pattern
                    store = PatternStore(cfg.patterns_file())
                    pat = store.get(pat_id)
                    self._pairs = apply_pattern(old, new, pat.key_regex) if pat else []

        except Exception as e:
            self._log(f"ペアリスト構築エラー: {e}")
            self._pairs = []

        self._populate_pairs()

    def _populate_pairs(self) -> None:
        for row in self._tree_pairs.get_children():
            self._tree_pairs.delete(row)
        kind_map = {
            "exact":         "完全一致",
            "auto":          "自動",
            "pattern":       "パターン",
            "unmatched_old": "旧のみ",
            "unmatched_new": "新のみ",
        }
        for i, p in enumerate(self._pairs):
            old_disp   = p.old_name or "（なし）"
            new_disp   = p.new_name or "（なし）"
            score_disp = f"{p.score:.2f}" if p.score > 0 else "-"
            kind_disp  = kind_map.get(p.matched_by, p.matched_by)
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
        from excel_diff.reader import read_workbook
        from excel_diff.diff_engine import diff_files
        from excel_diff.html_renderer import render
        from excel_diff.matcher import DiffConfig, parse_col_spec, parse_col_list, load_config
        from excel_diff.__main__ import _render_index_html

        sheet          = options.get("sheet") or None
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

        # 実行条件サマリ
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
        log_lines.append(f"[実行条件] シート: {sheet or '全シート'}")
        log_lines.append("─" * 36)
        for line in log_lines:
            self._log(line)

        old_name = Path(old_dir).name
        new_name = Path(new_dir).name
        out_dir = output_dir_opt or str(Path(new_dir).parent / f"{old_name}_vs_{new_name}")
        Path(out_dir).mkdir(parents=True, exist_ok=True)

        results = []
        skipped = []
        for pair in matched:
            old_path = os.path.join(old_dir, pair.old_name)
            new_path = os.path.join(new_dir, pair.new_name)
            try:
                old_sheets = read_workbook(old_path, strikethrough, sheet)
                new_sheets = read_workbook(new_path, strikethrough, sheet)
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

        index_path = os.path.join(out_dir, "★index.html")
        Path(index_path).write_text(
            _render_index_html(results, unmatched, old_dir, new_dir), encoding="utf-8",
        )
        return index_path, skipped

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
                    webbrowser.open(Path(index_path).resolve().as_uri())
        except queue.Empty:
            self.after(100, self._poll_compare)

    # ================================================================== パターン生成（案A: 画面切替）

    def _goto_step3(self) -> None:
        matched = [p for p in self._pairs if p.old_name and p.new_name]
        if not matched:
            messagebox.showinfo("情報", "パターン生成には比較可能なペアが必要です")
            return
        suggested = self._smart_suggest_regex(self._pairs)
        self._s3_regex.set(suggested)
        from_template = self._s3_tpl.get() != _TEMPLATES[-1][0]
        self._s3_manual.set(not from_template)
        self._on_manual_toggle()
        self._update_preview()
        self._validated_ok = False
        self._btn_save_pat.config(state="disabled")
        self._lbl_validate.config(text="")
        self._show_step3_view()

    def _build_step3(self, parent: tk.Frame) -> None:
        tk.Label(
            parent, text="パターン生成・検証・保存", font=("", 9, "bold"),
        ).pack(anchor="w", padx=8, pady=(6, 2))

        meta_fr = tk.Frame(parent)
        meta_fr.pack(fill="x", padx=8, pady=2)
        self._s3_id   = tk.StringVar()
        self._s3_name = tk.StringVar()
        self._s3_desc = tk.StringVar()
        for label, var, w in (
            ("ID", self._s3_id, 12), ("名前", self._s3_name, 20), ("説明", self._s3_desc, 28)
        ):
            tk.Label(meta_fr, text=label).pack(side="left")
            tk.Entry(meta_fr, textvariable=var, width=w).pack(side="left", padx=(2, 10))

        tpl_fr = tk.Frame(parent)
        tpl_fr.pack(fill="x", padx=8, pady=2)
        tk.Label(tpl_fr, text="テンプレート", width=12, anchor="w").pack(side="left")
        self._s3_tpl = tk.StringVar(value=_TEMPLATES[0][0])
        self._cmb_tpl = ttk.Combobox(
            tpl_fr, textvariable=self._s3_tpl,
            values=[t[0] for t in _TEMPLATES], state="readonly", width=36,
        )
        self._cmb_tpl.pack(side="left")
        self._cmb_tpl.bind("<<ComboboxSelected>>", lambda _: self._on_template_change())

        re_fr = tk.Frame(parent)
        re_fr.pack(fill="x", padx=8, pady=2)
        tk.Label(re_fr, text="正規表現", width=12, anchor="w").pack(side="left")
        self._s3_regex = tk.StringVar()
        self._entry_regex = tk.Entry(re_fr, textvariable=self._s3_regex, width=46)
        self._entry_regex.pack(side="left", padx=(0, 8))
        self._s3_manual = tk.BooleanVar(value=False)
        tk.Checkbutton(
            re_fr, text="手動入力モード", variable=self._s3_manual,
            command=self._on_manual_toggle,
        ).pack(side="left")
        self._s3_regex.trace_add("write", lambda *_: self._update_preview())

        # ボタン行を side="bottom" で先に pack（expand=True に隠されないよう）
        btn_row = tk.Frame(parent)
        btn_row.pack(side="bottom", fill="x", padx=8, pady=(4, 8))
        tk.Button(
            btn_row, text="← ペアリストへ戻る", command=self._show_main_view,
        ).pack(side="left")
        tk.Button(
            btn_row, text="検証実行", command=self._validate_pattern,
        ).pack(side="left", padx=8)
        self._btn_save_pat = tk.Button(
            btn_row, text="保存", state="disabled",
            bg="#4a9eff", fg="white", font=("", 9, "bold"),
            command=self._save_pattern,
        )
        self._btn_save_pat.pack(side="right")

        # 検証ラベルも side="bottom" で先に pack（折り返し有効）
        self._lbl_validate = tk.Label(
            parent, text="", wraplength=700, justify="left", fg="red", font=("", 9),
        )
        self._lbl_validate.pack(side="bottom", fill="x", anchor="w", padx=8, pady=2)

        # プレビューラベル＋ツリーを最後に pack（expand=True で残りを埋める）
        tk.Label(parent, text="キー抽出プレビュー", font=("", 8, "bold")).pack(
            anchor="w", padx=8, pady=(6, 0))

        fr_prev = tk.Frame(parent)
        fr_prev.pack(fill="both", expand=True, padx=8, pady=(0, 4))

        cols = ("old", "new", "key", "status")
        self._tree_prev = ttk.Treeview(
            fr_prev, columns=cols, show="headings", height=5, selectmode="none",
        )
        for col, head, w in zip(cols, ("旧ファイル", "新ファイル", "抽出キー", "状態"),
                                 (180, 180, 120, 60)):
            self._tree_prev.heading(col, text=head)
            self._tree_prev.column(col, width=w, anchor="w")
        self._tree_prev.tag_configure("ok",        foreground="#1a7f37")
        self._tree_prev.tag_configure("mismatch",  background="#fff3b0")
        self._tree_prev.tag_configure("error",     background="#ffe0e0")
        self._tree_prev.tag_configure("unmatched", foreground="#888888")

        sb3 = ttk.Scrollbar(fr_prev, orient="vertical", command=self._tree_prev.yview)
        self._tree_prev.configure(yscrollcommand=sb3.set)
        self._tree_prev.pack(side="left", fill="both", expand=True)
        sb3.pack(side="left", fill="y")

        self._on_manual_toggle()

    def _on_template_change(self) -> None:
        if self._s3_manual.get():
            return
        tpl_name = self._s3_tpl.get()
        for name, regex in _TEMPLATES:
            if name == tpl_name:
                if regex:
                    self._s3_regex.set(regex)
                break

    def _on_manual_toggle(self) -> None:
        state = "normal" if self._s3_manual.get() else "readonly"
        self._entry_regex.config(state=state)
        self._cmb_tpl.config(state="disabled" if self._s3_manual.get() else "readonly")

    def _update_preview(self) -> None:
        for row in self._tree_prev.get_children():
            self._tree_prev.delete(row)
        regex_str = self._s3_regex.get().strip()
        if not regex_str:
            return
        try:
            pattern = re.compile(regex_str)
        except re.error:
            return

        matched   = [p for p in self._pairs if p.old_name and p.new_name]
        unmatched = [p for p in self._pairs if not p.old_name or not p.new_name]

        for p in matched:
            m_old = pattern.fullmatch(p.old_name) if p.old_name else None
            m_new = pattern.fullmatch(p.new_name) if p.new_name else None
            if m_old and m_new:
                key_old, key_new = m_old.group(1), m_new.group(1)
                if key_old == key_new:
                    tag, key_disp, status = "ok", key_old, "✓"
                else:
                    tag, key_disp, status = "mismatch", f"{key_old} ≠ {key_new}", "不一致"
            else:
                tag, key_disp, status = "error", "(マッチなし)", "✗"
            self._tree_prev.insert("", "end",
                                   values=(p.old_name, p.new_name, key_disp, status),
                                   tags=(tag,))
        for p in unmatched:
            self._tree_prev.insert("", "end",
                                   values=(p.old_name or "（なし）", p.new_name or "（なし）", "-", "-"),
                                   tags=("unmatched",))

    def _validate_pattern(self) -> None:
        regex_str = self._s3_regex.get().strip()
        if not regex_str:
            self._lbl_validate.config(text="正規表現を入力してください", fg="red")
            return
        matched = [p for p in self._pairs if p.old_name and p.new_name]
        if not matched:
            self._lbl_validate.config(text="検証対象のペアがありません", fg="red")
            return
        try:
            from excel_diff.file_pairing import validate_regex
            errors = validate_regex(matched, regex_str)
        except Exception as e:
            self._lbl_validate.config(text=f"検証エラー: {e}", fg="red")
            return

        if errors:
            msgs = [f"・{_ERR_MSG.get(e.kind, e.kind)}\n  詳細: {e.details}"
                    for e in errors[:3]]
            self._lbl_validate.config(text="\n".join(msgs), fg="red")
            self._validated_ok = False
            self._btn_save_pat.config(state="disabled")
        else:
            self._lbl_validate.config(text="検証OK。保存できます。", fg="#1a7f37")
            self._validated_ok = True
            self._btn_save_pat.config(state="normal")

    def _save_pattern(self) -> None:
        pat_id   = self._s3_id.get().strip()
        pat_name = self._s3_name.get().strip()
        regex    = self._s3_regex.get().strip()
        if not pat_id:
            messagebox.showerror("エラー", "IDを入力してください")
            return
        if not pat_name:
            messagebox.showerror("エラー", "名前を入力してください")
            return
        try:
            from excel_diff.patterns import PatternStore, PatternDef
            from datetime import date
            store = PatternStore(cfg.patterns_file())
            if store.get(pat_id):
                if not messagebox.askyesno("確認", f"パターン「{pat_id}」は既に存在します。上書きしますか？"):
                    return
            store.add_or_update(PatternDef(
                id=pat_id, name=pat_name, key_regex=regex,
                description=self._s3_desc.get().strip(),
                example_old_dir=self._old_dir.get(),
                example_new_dir=self._new_dir.get(),
                created_at=date.today().isoformat(),
            ))
            store.save()
            self._log(f"パターン保存: [{pat_id}] {pat_name}  regex={regex}")
            self._refresh_list()
            self._show_main_view()
        except Exception as e:
            self._log(f"保存エラー: {e}")

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
            self._refresh_list()
        except Exception as e:
            self._log(f"削除エラー: {e}")

    # ================================================================== インテリジェント正規表現生成

    def _smart_suggest_regex(self, pairs: list) -> str:
        all_files = [f for p in pairs for f in (p.old_name, p.new_name) if f]
        if not all_files:
            return ""

        best_tpl_name, best_tpl_regex, best_score = _TEMPLATES[-1][0], "", 0.0
        for tpl_name, tpl_regex in _TEMPLATES[:-1]:
            try:
                pat = re.compile(tpl_regex)
            except re.error:
                continue
            score = sum(1 for f in all_files if pat.fullmatch(f)) / len(all_files)
            if score > best_score:
                best_score, best_tpl_name, best_tpl_regex = score, tpl_name, tpl_regex

        if best_score >= 0.5:
            self._s3_tpl.set(best_tpl_name)
            return best_tpl_regex

        matched = [p for p in pairs if p.old_name and p.new_name]
        try:
            from excel_diff.file_pairing import generate_regex
            result = generate_regex(matched)
            if result:
                self._s3_tpl.set(_TEMPLATES[-1][0])
                return result
        except Exception:
            pass

        differing = [
            (p.old_name, p.new_name)
            for p in matched if p.old_name != p.new_name
        ]
        if not differing:
            return ""

        analyses = []
        ext_set: set = set()
        for old_f, new_f in differing:
            old_base, old_ext = os.path.splitext(old_f)
            new_base          = os.path.splitext(new_f)[0]
            ext_set.add(old_ext.lower())
            lcp = _lcp_len(old_base, new_base)
            lcs = _lcs_len(old_base, new_base)
            old_var = old_base[lcp : len(old_base) - lcs] if lcs else old_base[lcp:]
            new_var = new_base[lcp : len(new_base) - lcs] if lcs else new_base[lcp:]
            if not old_var or not new_var:
                continue
            prefix_raw = old_base[:lcp]
            sep = prefix_raw[-1] if prefix_raw and prefix_raw[-1] in "_-" else ""
            analyses.append((sep, _classify_var(old_var), _classify_var(new_var)))

        if not analyses:
            return ""
        seps = {a[0] for a in analyses}
        if len(seps) > 1:
            return ""
        sep    = list(seps)[0]
        sep_re = re.escape(sep) if sep else ""
        var_pats = {a[1] for a in analyses} | {a[2] for a in analyses}
        if r".+" in var_pats and len(var_pats) > 1:
            var_pats.discard(r".+")
        var_re = (
            list(var_pats)[0]
            if len(var_pats) == 1
            else f"(?:{'|'.join(sorted(var_pats))})"
        )
        ext    = list(ext_set)[0] if len(ext_set) == 1 else ".xlsx"
        ext_re = re.escape(ext)
        self._s3_tpl.set(_TEMPLATES[-1][0])
        return f"^(.+?){sep_re}{var_re}{ext_re}$"

    # ================================================================== 状態保存

    def save_state(self) -> None:
        """現在のUI値を設定に書き戻す（ウィンドウを閉じる前に呼ばれる）。"""
        pat_sel = self._pat_id.get().strip()
        pat_id  = pat_sel.split()[0] if pat_sel else ""
        cfg.set_tab("pair_build", {
            "old_dir":    self._old_dir.get(),
            "new_dir":    self._new_dir.get(),
            "pairing":    self._pairing.get(),
            "pairs_file": self._pairs_f.get(),
            "pattern_id": pat_id,
        })
