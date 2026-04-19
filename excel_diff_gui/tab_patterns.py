"""タブ④ パターン管理。"""
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
        self, parent,
        log: Callable[[str], None],
        switch_to_dir_diff: Optional[Callable] = None,
    ) -> None:
        super().__init__(parent)
        self._log = log
        self._switch_to_dir_diff = switch_to_dir_diff
        self._result_q: "queue.Queue | None" = None
        self._pairs: list = []
        self._old_dir = ""
        self._new_dir = ""
        self._validated_ok = False
        self._compare_open_browser = True

        self._build()

    # ================================================================== レイアウト

    def _build(self) -> None:
        # 上段（パターン一覧）と下段（ウィザード）を縦分割ペインで配置
        paned = ttk.PanedWindow(self, orient="vertical")
        paned.pack(fill="both", expand=True, padx=4, pady=4)

        # ── 上段: パターン一覧 ──────────────────────────────────────────
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

        # ── 下段: ウィザード ────────────────────────────────────────────
        self._wiz_outer = tk.LabelFrame(paned, text="新規パターン作成ウィザード")
        paned.add(self._wiz_outer, weight=2)

        self._fr_step1 = tk.Frame(self._wiz_outer)
        self._fr_step2 = tk.Frame(self._wiz_outer)
        self._fr_step3 = tk.Frame(self._wiz_outer)

        self._build_step1(self._fr_step1)
        self._build_step2(self._fr_step2)
        self._build_step3(self._fr_step3)

        self._show_step(1)
        self._refresh_list()

    # ------------------------------------------------------------------ パターン一覧

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

    # ================================================================== ステップ①

    def _build_step1(self, parent: tk.Frame) -> None:
        tk.Label(parent, text="ステップ①: ペア候補探索", font=("", 9, "bold")).pack(
            anchor="w", padx=8, pady=(6, 2))

        self._s1_old = tk.StringVar()
        self._s1_new = tk.StringVar()
        FileSelectRow(parent, "旧フォルダ", self._s1_old, mode="dir").pack(
            fill="x", padx=8, pady=2)
        FileSelectRow(parent, "新フォルダ", self._s1_new, mode="dir").pack(
            fill="x", padx=8, pady=2)

        fr_thr = tk.Frame(parent)
        fr_thr.pack(fill="x", padx=8, pady=4)
        tk.Label(fr_thr, text="しきい値", width=10, anchor="w").pack(side="left")
        self._s1_thr = tk.DoubleVar(value=0.30)
        self._s1_thr_lbl = tk.Label(fr_thr, text="0.30", width=5)
        self._s1_thr_lbl.pack(side="right")
        tk.Label(fr_thr, text="1.0", fg="gray").pack(side="right")
        tk.Scale(
            fr_thr, variable=self._s1_thr, from_=0.0, to=1.0,
            resolution=0.05, orient="horizontal", showvalue=False,
            command=lambda v: self._s1_thr_lbl.config(text=f"{float(v):.2f}"),
        ).pack(side="left", fill="x", expand=True)
        tk.Label(fr_thr, text="0.0", fg="gray").pack(side="left")

        btn_row = tk.Frame(parent)
        btn_row.pack(fill="x", padx=8, pady=(4, 8))
        self._s1_btn = tk.Button(
            btn_row, text="探索実行", width=14,
            bg="#4a9eff", fg="white", font=("", 10, "bold"),
            command=self._run_discover,
        )
        self._s1_btn.pack(side="right")

    def _run_discover(self) -> None:
        old = self._s1_old.get().strip()
        new = self._s1_new.get().strip()
        if not old or not new:
            messagebox.showerror("エラー", "旧フォルダと新フォルダを指定してください")
            return
        if not os.path.isdir(old):
            messagebox.showerror("エラー", f"フォルダが見つかりません:\n{old}")
            return
        if not os.path.isdir(new):
            messagebox.showerror("エラー", f"フォルダが見つかりません:\n{new}")
            return
        self._old_dir = old
        self._new_dir = new
        self._s1_btn.config(state="disabled", text="探索中...")
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
            self._s1_btn.config(state="normal", text="探索実行")
            if status == "err":
                self._log(f"探索エラー: {val}")
            else:
                self._pairs = val
                self._log(f"探索完了: {len(val)} ペア候補")
                self._populate_step2()
                self._show_step(2)
        except queue.Empty:
            self.after(100, self._poll_discover)

    # ================================================================== ステップ②

    def _build_step2(self, parent: tk.Frame) -> None:
        tk.Label(parent, text="ステップ②: ペア確認・調整", font=("", 9, "bold")).pack(
            anchor="w", padx=8, pady=(6, 2))

        # 探索対象フォルダを表示
        self._lbl_dirs = tk.Label(
            parent, text="", fg="#444444", font=("", 8),
            anchor="w", justify="left",
        )
        self._lbl_dirs.pack(anchor="w", padx=8, pady=(0, 4))

        cols = ("old", "new", "score", "kind")
        self._tree_pairs = ttk.Treeview(
            parent, columns=cols, show="headings", height=6, selectmode="browse",
        )
        for col, head, w in zip(cols, ("旧ファイル", "新ファイル", "スコア", "種別"),
                                 (190, 190, 60, 80)):
            self._tree_pairs.heading(col, text=head)
            self._tree_pairs.column(col, width=w, anchor="w")
        self._tree_pairs.tag_configure("unmatched", foreground="#888888")

        sb2 = ttk.Scrollbar(parent, orient="vertical", command=self._tree_pairs.yview)
        self._tree_pairs.configure(yscrollcommand=sb2.set)
        fr_tree = tk.Frame(parent)
        fr_tree.pack(fill="both", expand=True, padx=8, pady=(0, 2))
        self._tree_pairs.pack(side="left", fill="both", expand=True)
        sb2.pack(side="left", fill="y")

        tk.Label(parent,
                 text="※「旧のみ」「新のみ」の行は比較対象外として扱われます",
                 fg="gray", font=("", 8)).pack(anchor="w", padx=8)

        # ボタン 2 行
        btn_row1 = tk.Frame(parent)
        btn_row1.pack(fill="x", padx=8, pady=(6, 2))
        tk.Button(btn_row1, text="← やり直す",
                  command=self._back_to_step1).pack(side="left")
        tk.Button(btn_row1, text="パターン生成 →",
                  bg="#4a9eff", fg="white", font=("", 9, "bold"),
                  command=self._goto_step3).pack(side="right")

        btn_row2 = tk.Frame(parent)
        btn_row2.pack(fill="x", padx=8, pady=(0, 8))
        tk.Button(btn_row2, text="JSON 保存",
                  command=self._save_pairs_json).pack(side="left")
        tk.Button(btn_row2, text="そのまま比較",      # ← 変更
                  command=self._run_compare_pairs).pack(side="left", padx=8)

    def _populate_step2(self) -> None:
        self._lbl_dirs.config(text=f"旧: {self._old_dir}\n新: {self._new_dir}")
        for row in self._tree_pairs.get_children():
            self._tree_pairs.delete(row)
        kind_map = {
            "exact": "完全一致", "auto": "自動", "pattern": "パターン",
            "unmatched_old": "旧のみ", "unmatched_new": "新のみ",
        }
        for i, p in enumerate(self._pairs):
            old_disp = p.old_name or "（なし）"
            new_disp = p.new_name or "（なし）"
            score_disp = f"{p.score:.2f}" if p.score > 0 else "-"
            kind_disp = kind_map.get(p.matched_by, p.matched_by)
            tags = ("unmatched",) if not p.old_name or not p.new_name else ()
            self._tree_pairs.insert("", "end", iid=str(i),
                                    values=(old_disp, new_disp, score_disp, kind_disp),
                                    tags=tags)

    def _back_to_step1(self) -> None:
        self._pairs = []
        self._show_step(1)

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
            self._log(f"ペアJSON保存: {path}（フォルダ比較タブの「ペアJSON」で再利用可）")
        except Exception as e:
            self._log(f"保存エラー: {e}")

    # ── 比較確認ポップアップ ─────────────────────────────────────────────

    def _ask_compare_or_settings(self) -> Optional[str]:
        """'ok' / 'settings' / None(閉じた) を返す"""
        result: dict = {"v": None}
        dlg = tk.Toplevel(self)
        dlg.title("比較実行の確認")
        dlg.resizable(False, False)
        dlg.transient(self.winfo_toplevel())
        dlg.grab_set()

        tk.Label(
            dlg,
            text="フォルダ比較に設定されたオプションで\n実行されますが良いですか？",
            pady=10, padx=20,
        ).pack()

        btn_fr = tk.Frame(dlg)
        btn_fr.pack(pady=(0, 12), padx=20)
        tk.Button(btn_fr, text="OK", width=10,
                  command=lambda: (result.__setitem__("v", "ok"), dlg.destroy())
                  ).pack(side="left", padx=6)
        tk.Button(btn_fr, text="フォルダ比較の設定に戻る",
                  command=lambda: (result.__setitem__("v", "settings"), dlg.destroy())
                  ).pack(side="left", padx=6)

        self.winfo_toplevel().update_idletasks()
        rx = self.winfo_toplevel().winfo_x()
        ry = self.winfo_toplevel().winfo_y()
        rw = self.winfo_toplevel().winfo_width()
        rh = self.winfo_toplevel().winfo_height()
        dlg.update_idletasks()
        dw, dh = dlg.winfo_width(), dlg.winfo_height()
        dlg.geometry(f"+{rx + (rw - dw) // 2}+{ry + (rh - dh) // 2}")

        dlg.wait_window()
        return result["v"]

    def _run_compare_pairs(self) -> None:
        matched = [p for p in self._pairs if p.old_name and p.new_name]
        if not matched:
            messagebox.showinfo("情報", "比較可能なペアがありません")
            return

        choice = self._ask_compare_or_settings()
        if choice == "settings":
            if self._switch_to_dir_diff:
                self._switch_to_dir_diff()
            return
        if choice != "ok":
            return

        options = cfg.data("dir_diff")
        self._compare_open_browser = options.get("open_browser", True)
        self._log(f"そのまま比較: {len(matched)} 件（フォルダ比較タブの設定を使用）")
        self._result_q = get_worker().submit(
            self._do_compare_pairs, matched, self._old_dir, self._new_dir, options,
        )
        self.after(100, self._poll_compare)

    def _do_compare_pairs(self, matched, old_dir, new_dir, options: dict):
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
                fd = diff_files(old_sheets, new_sheets, old_path, new_path,
                                include_strike=strikethrough, config=config)
                out_path = os.path.join(out_dir, f"{Path(pair.new_name).stem}_diff.html")
                Path(out_path).write_text(render(fd), encoding="utf-8")
                results.append((pair, fd, out_path))
            except Exception as e:
                skipped.append(pair.old_name)
                # skipped リストは return 後に index に含めないためここでは記録のみ

        unmatched = [p for p in self._pairs if not p.old_name or not p.new_name]
        index_path = os.path.join(out_dir, "index.html")
        Path(index_path).write_text(
            _render_index_html(results, unmatched, old_dir, new_dir), encoding="utf-8")
        return index_path, skipped

    def _poll_compare(self) -> None:
        if self._result_q is None:
            return
        try:
            status, val = self._result_q.get_nowait()
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
                self._back_to_step1()
        except queue.Empty:
            self.after(100, self._poll_compare)

    def _goto_step3(self) -> None:
        matched = [p for p in self._pairs if p.old_name and p.new_name]
        if not matched:
            messagebox.showinfo("情報", "パターン生成には比較可能なペアが必要です")
            return
        suggested = self._smart_suggest_regex(self._pairs)
        self._s3_regex.set(suggested)
        # テンプレートから生成された場合は readonly、それ以外は手動入力モード
        from_template = self._s3_tpl.get() != _TEMPLATES[-1][0]
        self._s3_manual.set(not from_template)
        self._on_manual_toggle()
        self._update_preview()
        self._validated_ok = False
        self._btn_save_pat.config(state="disabled")
        self._lbl_validate.config(text="")
        self._show_step(3)

    # ================================================================== ステップ③

    def _build_step3(self, parent: tk.Frame) -> None:
        tk.Label(parent, text="ステップ③: パターン生成・検証・保存", font=("", 9, "bold")).pack(
            anchor="w", padx=8, pady=(6, 2))

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

        tk.Label(parent, text="キー抽出プレビュー", font=("", 8, "bold")).pack(
            anchor="w", padx=8, pady=(6, 0))

        cols = ("old", "new", "key", "status")
        self._tree_prev = ttk.Treeview(
            parent, columns=cols, show="headings", height=5, selectmode="none",
        )
        for col, head, w in zip(cols, ("旧ファイル", "新ファイル", "抽出キー", "状態"),
                                 (180, 180, 120, 60)):
            self._tree_prev.heading(col, text=head)
            self._tree_prev.column(col, width=w, anchor="w")
        self._tree_prev.tag_configure("ok",        foreground="#1a7f37")
        self._tree_prev.tag_configure("mismatch",  background="#fff3b0")
        self._tree_prev.tag_configure("error",     background="#ffe0e0")
        self._tree_prev.tag_configure("unmatched", foreground="#888888")

        sb3 = ttk.Scrollbar(parent, orient="vertical", command=self._tree_prev.yview)
        self._tree_prev.configure(yscrollcommand=sb3.set)
        fr_prev = tk.Frame(parent)
        fr_prev.pack(fill="both", expand=True, padx=8, pady=(0, 4))
        self._tree_prev.pack(side="left", fill="both", expand=True)
        sb3.pack(side="left", fill="y")

        self._lbl_validate = tk.Label(parent, text="", wraplength=500, justify="left",
                                      fg="red", font=("", 9))
        self._lbl_validate.pack(anchor="w", padx=8, pady=2)

        btn_row = tk.Frame(parent)
        btn_row.pack(fill="x", padx=8, pady=(4, 8))
        tk.Button(btn_row, text="← 戻る", command=self._back_to_step2).pack(side="left")
        tk.Button(btn_row, text="検証実行", command=self._validate_pattern).pack(side="left", padx=8)
        self._btn_save_pat = tk.Button(
            btn_row, text="保存", state="disabled",
            bg="#4a9eff", fg="white", font=("", 9, "bold"),
            command=self._save_pattern,
        )
        self._btn_save_pat.pack(side="right")

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
                example_old_dir=self._old_dir,
                example_new_dir=self._new_dir,
                created_at=date.today().isoformat(),
            ))
            store.save()
            self._log(f"パターン保存: [{pat_id}] {pat_name}  regex={regex}")
            self._refresh_list()
            self._pairs = []
            self._show_step(1)
        except Exception as e:
            self._log(f"保存エラー: {e}")

    def _back_to_step2(self) -> None:
        self._show_step(2)

    # ================================================================== インテリジェント正規表現生成

    def _smart_suggest_regex(self, pairs: list) -> str:
        """ファイル名パターンを分析し、3段階で最適な正規表現を自動提案する。"""
        all_files = [f for p in pairs for f in (p.old_name, p.new_name) if f]
        if not all_files:
            return ""

        # ── Step 1: 定義済みテンプレートで最もマッチするものを採用 ──────
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

        # ── Step 2: file_pairing.generate_regex() を試す ────────────────
        matched = [p for p in pairs if p.old_name and p.new_name]
        try:
            from excel_diff.file_pairing import generate_regex
            result = generate_regex(matched)
            if result:
                self._s3_tpl.set(_TEMPLATES[-1][0])
                return result
        except Exception:
            pass

        # ── Step 3: LCP/LCS 方式で変動部分を特定してパターン生成 ─────────
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
            return ""  # セパレータが不統一

        sep    = list(seps)[0]
        sep_re = re.escape(sep) if sep else ""

        var_pats = {a[1] for a in analyses} | {a[2] for a in analyses}
        if r".+" in var_pats and len(var_pats) > 1:
            var_pats.discard(r".+")   # 具体パターンを優先
        var_re = (
            list(var_pats)[0]
            if len(var_pats) == 1
            else f"(?:{'|'.join(sorted(var_pats))})"
        )

        ext    = list(ext_set)[0] if len(ext_set) == 1 else ".xlsx"
        ext_re = re.escape(ext)

        self._s3_tpl.set(_TEMPLATES[-1][0])
        return f"^(.+?){sep_re}{var_re}{ext_re}$"

    # ================================================================== ステップ切り替え

    def _show_step(self, n: int) -> None:
        for fr in (self._fr_step1, self._fr_step2, self._fr_step3):
            fr.pack_forget()
        {1: self._fr_step1, 2: self._fr_step2, 3: self._fr_step3}[n].pack(
            fill="both", expand=True)
