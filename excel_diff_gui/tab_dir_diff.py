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
        load_profile_cb: Optional[Callable] = None,
        get_snapshot_cb: Optional[Callable] = None,
    ) -> None:
        super().__init__(parent)
        self._log = log
        self._switch_to_pair_build = switch_to_pair_build
        self._load_profile_cb = load_profile_cb
        self._get_snapshot_cb = get_snapshot_cb
        self._profile_ids: list = []
        self._profile_cmb: Optional[ttk.Combobox] = None

        self._out_dir   = tk.StringVar(value=cfg.get("dir_diff", "output_dir"))
        self._sheet_old = tk.StringVar(value=cfg.get("dir_diff", "sheet_old"))
        self._sheet_new = tk.StringVar(value=cfg.get("dir_diff", "sheet_new"))
        self._cols      = tk.StringVar(value=cfg.get("dir_diff", "include_cols"))
        self._matchers = tk.StringVar(value=cfg.get("dir_diff", "matchers"))
        self._key_cols = tk.StringVar(value=cfg.get("dir_diff", "key_cols"))
        self._strike   = tk.BooleanVar(value=cfg.get("dir_diff", "strikethrough"))
        self._open_br  = tk.BooleanVar(value=cfg.get("dir_diff", "open_browser", True))
        self._mode     = tk.StringVar(value=cfg.get("dir_diff", "diff_mode", "lcs"))

        self._build()

    # ------------------------------------------------------------------ レイアウト

    def _build(self) -> None:
        pad = {"padx": 6, "pady": 3}

        # ── 設定セット（プロファイル）セレクタ ────────────────────────────
        grp_prof = tk.LabelFrame(self, text="設定セット")
        grp_prof.pack(fill="x", **pad)

        prof_row = tk.Frame(grp_prof)
        prof_row.pack(fill="x", padx=6, pady=4)

        self._profile_cmb = ttk.Combobox(
            prof_row, state="readonly", width=28,
        )
        self._profile_cmb.pack(side="left")

        tk.Button(
            prof_row, text="読み込み", width=8,
            command=self._load_profile,
        ).pack(side="left", padx=(4, 0))

        tk.Button(
            prof_row, text="現在の設定を保存...", width=16,
            command=self._save_profile,
        ).pack(side="left", padx=(8, 0))

        tk.Button(
            prof_row, text="管理...", width=8,
            command=self._manage_profiles,
        ).pack(side="left", padx=(4, 0))

        self.refresh_profile_combobox()

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
        tk.Label(fr, text="旧シートパターン", width=14, anchor="w").pack(side="left")
        tk.Entry(fr, textvariable=self._sheet_old).pack(side="left", fill="x", expand=True)
        tk.Label(fr, text="空=全シート / 正規表現", fg="gray").pack(side="left", padx=4)

        fr_sn = tk.Frame(grp_opt)
        fr_sn.pack(fill="x", padx=6, pady=2)
        tk.Label(fr_sn, text="新シートパターン", width=14, anchor="w").pack(side="left")
        tk.Entry(fr_sn, textvariable=self._sheet_new).pack(side="left", fill="x", expand=True)
        tk.Label(fr_sn, text="空=全シート / 正規表現", fg="gray").pack(side="left", padx=4)

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

    # ------------------------------------------------------------------ プロファイル操作

    def refresh_profile_combobox(self) -> None:
        """保存済みプロファイル一覧を Combobox に反映する。"""
        if self._profile_cmb is None:
            return
        profiles = cfg.get_profiles()
        self._profile_ids = [p["id"] for p in profiles]
        self._profile_cmb["values"] = [p["name"] for p in profiles]
        if not profiles:
            self._profile_cmb.set("")

    def _load_profile(self) -> None:
        idx = self._profile_cmb.current()
        if idx < 0:
            from tkinter import messagebox
            messagebox.showinfo("情報", "読み込む設定セットを選択してください")
            return
        profile_id = self._profile_ids[idx]
        if self._load_profile_cb:
            self._load_profile_cb(profile_id)

    def _save_profile(self) -> None:
        from .profile_dialog import ProfileSaveDialog
        result = ProfileSaveDialog(self).show()
        if result and result[0] == "save":
            _, name, note = result
            if self._get_snapshot_cb is None:
                return
            snap = self._get_snapshot_cb()
            cfg.save_profile(name, note, snap)
            cfg.save()
            self.refresh_profile_combobox()
            self._log(f"設定セット保存: {name}")

    def _manage_profiles(self) -> None:
        from .profile_dialog import ProfileManageDialog
        ProfileManageDialog(
            self, on_change=self.refresh_profile_combobox,
        ).show()

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
            "sheet_old":     self._sheet_old.get().strip(),
            "sheet_new":     self._sheet_new.get().strip(),
            "include_cols":  self._cols.get().strip(),
            "matchers":      self._matchers.get().strip(),
            "strikethrough": self._strike.get(),
            "open_browser":  self._open_br.get(),
            "diff_mode":     self._mode.get(),
            "key_cols":      self._key_cols.get().strip(),
        }

    def get_snapshot(self) -> dict:
        """現在の UI 値をすべて dict で返す（プロファイル保存用）。"""
        return self.get_compare_options()

    def load_from_snapshot(self, snap: dict) -> None:
        """スナップショットの値を各変数に反映する（プロファイル読み込み用）。"""
        mapping = {
            "output_dir":    self._out_dir,
            "sheet_old":     self._sheet_old,
            "sheet_new":     self._sheet_new,
            "include_cols":  self._cols,
            "matchers":      self._matchers,
            "key_cols":      self._key_cols,
            "diff_mode":     self._mode,
        }
        for key, var in mapping.items():
            if key in snap:
                var.set(snap[key])
        if "strikethrough" in snap:
            self._strike.set(bool(snap["strikethrough"]))
        if "open_browser" in snap:
            self._open_br.set(bool(snap["open_browser"]))
        self._on_mode()

    def save_state(self) -> None:
        """現在のUI値を設定に書き戻す（ウィンドウを閉じる前に呼ばれる）。"""
        cfg.set_tab("dir_diff", self.get_compare_options())
