"""メインウィンドウ。"""
import tkinter as tk
from tkinter import ttk

# 強制シャットダウン等で正常終了できない場合に備えた定期autosaveの間隔
_AUTOSAVE_INTERVAL_MS = 30_000

from . import settings as cfg
from .tab_dir_diff import TabDirDiff
from .tab_file_diff import TabFileDiff
from .tab_patterns import TabPatterns
from .tab_sheet_diff import TabSheetDiff
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
        paned.add(nb, weight=4)

        self._log_area = LogArea(paned, height=10)
        paned.add(self._log_area, weight=2)

        # タブ①: フォルダ比較（条件設定）
        tab_dir = TabDirDiff(
            nb, self._log,
            load_profile_cb=self.load_profile,
            get_snapshot_cb=self._get_current_snapshot,
        )
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

        # タブ⑤: シート比較
        self._tab_sheet_cmp = TabSheetDiff(nb, self._log)

        nb.add(tab_dir,              text="フォルダ比較（条件設定）")
        nb.add(tab_patterns,         text="フォルダ比較（ペアリング・比較実行）")
        nb.add(self._tab_file,       text="ファイル比較")
        nb.add(self._tab_split,      text="シート分解")
        nb.add(self._tab_sheet_cmp,  text="シート比較")

        self.protocol("WM_DELETE_WINDOW", self._quit)
        self.after(_AUTOSAVE_INTERVAL_MS, self._autosave)

    def _log(self, msg: str) -> None:
        self._log_area.log(msg)

    def _get_current_snapshot(self) -> dict:
        """全タブの現在値を収集してスナップショット dict を返す。"""
        return {
            "dir_diff":      self._tab_dir.get_snapshot(),
            "pair_build":    self._tab_patterns.get_snapshot(),
            "file_diff":     self._tab_file.get_snapshot(),
            "split":         self._tab_split.get_snapshot(),
            "sheet_compare": self._tab_sheet_cmp.get_snapshot(),
        }

    def load_profile(self, profile_id: str) -> None:
        """指定IDのプロファイルを全タブに読み込む。"""
        profile = next(
            (p for p in cfg.get_profiles() if p["id"] == profile_id), None
        )
        if not profile:
            return
        snap = profile.get("snapshot", {})
        self._tab_dir.load_from_snapshot(snap.get("dir_diff", {}))
        self._tab_patterns.load_from_snapshot(snap.get("pair_build", {}))
        self._tab_file.load_from_snapshot(snap.get("file_diff", {}))
        self._tab_split.load_from_snapshot(snap.get("split", {}))
        self._tab_sheet_cmp.load_from_snapshot(snap.get("sheet_compare", {}))
        self._tab_dir.refresh_profile_combobox()
        self._log(f"設定セット読み込み: {profile['name']}")

    def _save_all_tab_states(self) -> None:
        """各タブの現在UI値を _data に書き戻す。"""
        for tab in (self._tab_file, self._tab_dir, self._tab_patterns,
                    self._tab_split, self._tab_sheet_cmp):
            tab.save_state()

    def _autosave(self) -> None:
        """強制シャットダウン等で _quit() を経由せず終了した場合の保険として
        定期的に gui_settings.json へ保存する（プロファイルの問い合わせは行わない）。"""
        try:
            self._save_all_tab_states()
            cfg.save()
        except Exception:
            pass
        finally:
            self.after(_AUTOSAVE_INTERVAL_MS, self._autosave)

    def _quit(self) -> None:
        # 各タブの現在UI値を _data に書き戻してから保存
        self._save_all_tab_states()

        # 保存済みプロファイルが1件以上あり、かつどれとも一致しなければ問い合わせ
        if cfg.get_profiles():
            snap = self._get_current_snapshot()
            if cfg.find_matching_profile(snap) is None:
                from .profile_dialog import ProfileSaveDialog
                result = ProfileSaveDialog(self, on_exit=True).show()
                if result is None:  # キャンセル → 終了しない
                    return
                if result[0] == "save":
                    _, name, note = result
                    cfg.save_profile(name, note, snap)

        cfg.save()
        self.destroy()
