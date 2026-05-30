"""プロファイル（設定セット）の保存・管理ダイアログ。"""
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Callable, Optional, Tuple

from . import settings as cfg


class ProfileSaveDialog:
    """プロファイルの名前・メモを入力して保存するダイアログ。

    show() の戻り値:
        ("save", name, note) — 保存ボタン押下
        ("skip",)            — 保存しないボタン押下（on_exit=True の場合のみ）
        None                 — キャンセル（終了しない）
    """

    def __init__(
        self,
        parent: tk.Misc,
        name: str = "",
        note: str = "",
        on_exit: bool = False,
    ) -> None:
        self._parent = parent
        self._initial_name = name
        self._initial_note = note
        self._on_exit = on_exit
        self._result: Optional[Tuple] = None

    def show(self) -> Optional[Tuple]:
        dlg = tk.Toplevel(self._parent)
        dlg.title("設定セットを保存")
        dlg.resizable(False, False)
        dlg.transient(self._parent.winfo_toplevel())
        dlg.grab_set()

        frm = tk.Frame(dlg, padx=16, pady=12)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)

        if self._on_exit:
            tk.Label(
                frm,
                text="現在の設定はどの設定セットとも一致しません。\n名前をつけて保存しますか？",
                justify="left",
            ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
            start_row = 1
        else:
            start_row = 0

        name_var = tk.StringVar(value=self._initial_name)
        note_var = tk.StringVar(value=self._initial_note)

        tk.Label(frm, text="名前", anchor="w").grid(
            row=start_row, column=0, sticky="w", pady=4, padx=(0, 8))
        name_entry = tk.Entry(frm, textvariable=name_var, width=32)
        name_entry.grid(row=start_row, column=1, sticky="ew", pady=4)

        tk.Label(frm, text="メモ", anchor="w").grid(
            row=start_row + 1, column=0, sticky="w", pady=4, padx=(0, 8))
        tk.Entry(frm, textvariable=note_var, width=32).grid(
            row=start_row + 1, column=1, sticky="ew", pady=4)

        btn_frm = tk.Frame(frm)
        btn_frm.grid(
            row=start_row + 2, column=0, columnspan=2, pady=(12, 0))

        def do_save() -> None:
            n = name_var.get().strip()
            if not n:
                messagebox.showerror("エラー", "名前を入力してください", parent=dlg)
                return
            self._result = ("save", n, note_var.get().strip())
            dlg.destroy()

        def do_skip() -> None:
            self._result = ("skip",)
            dlg.destroy()

        def do_cancel() -> None:
            self._result = None
            dlg.destroy()

        if self._on_exit:
            tk.Button(
                btn_frm, text="保存して終了",
                bg="#4a9eff", fg="white", font=("", 9, "bold"), width=14,
                command=do_save,
            ).pack(side="left", padx=4)
            tk.Button(
                btn_frm, text="保存せず終了", width=12,
                command=do_skip,
            ).pack(side="left", padx=4)
            tk.Button(
                btn_frm, text="キャンセル", width=10,
                command=do_cancel,
            ).pack(side="left", padx=4)
        else:
            tk.Button(
                btn_frm, text="保存",
                bg="#4a9eff", fg="white", font=("", 9, "bold"), width=10,
                command=do_save,
            ).pack(side="left", padx=4)
            tk.Button(
                btn_frm, text="キャンセル", width=10,
                command=do_cancel,
            ).pack(side="left", padx=4)

        name_entry.focus_set()
        dlg.wait_window()
        return self._result


class ProfileManageDialog:
    """保存済みプロファイルの一覧・編集・削除ダイアログ。"""

    def __init__(
        self,
        parent: tk.Misc,
        on_change: Optional[Callable] = None,
    ) -> None:
        self._parent = parent
        self._on_change = on_change

    def show(self) -> None:
        dlg = tk.Toplevel(self._parent)
        dlg.title("設定セットの管理")
        dlg.resizable(True, True)
        dlg.transient(self._parent.winfo_toplevel())
        dlg.grab_set()
        dlg.minsize(520, 280)

        frm = tk.Frame(dlg, padx=10, pady=8)
        frm.pack(fill="both", expand=True)
        frm.rowconfigure(0, weight=1)
        frm.columnconfigure(0, weight=1)

        cols = ("name", "note", "created_at")
        tree = ttk.Treeview(
            frm, columns=cols, show="headings",
            height=8, selectmode="browse",
        )
        for col, head, w in zip(
            cols, ("名前", "メモ", "作成日時"), (180, 220, 130)
        ):
            tree.heading(col, text=head)
            tree.column(col, width=w, anchor="w")

        sb = ttk.Scrollbar(frm, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        btn_frm = tk.Frame(frm)
        btn_frm.grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))

        def _refresh() -> None:
            for item in tree.get_children():
                tree.delete(item)
            for p in cfg.get_profiles():
                tree.insert(
                    "", "end", iid=p["id"],
                    values=(p["name"], p.get("note", ""), p.get("created_at", "")),
                )

        def _edit() -> None:
            sel = tree.selection()
            if not sel:
                messagebox.showinfo("情報", "編集するセットを選択してください",
                                    parent=dlg)
                return
            profile_id = sel[0]
            profile = next(
                (p for p in cfg.get_profiles() if p["id"] == profile_id), None
            )
            if not profile:
                return
            _open_edit_dialog(dlg, profile, lambda: (_refresh(), _notify()))

        def _delete() -> None:
            sel = tree.selection()
            if not sel:
                messagebox.showinfo("情報", "削除するセットを選択してください",
                                    parent=dlg)
                return
            profile_id = sel[0]
            name = tree.item(profile_id, "values")[0]
            if not messagebox.askyesno(
                "確認", f"設定セット「{name}」を削除しますか？", parent=dlg
            ):
                return
            cfg.delete_profile(profile_id)
            cfg.save()
            _refresh()
            _notify()

        def _notify() -> None:
            if self._on_change:
                self._on_change()

        tk.Button(btn_frm, text="編集", width=8, command=_edit).pack(
            side="left", padx=(0, 4))
        tk.Button(btn_frm, text="削除", width=8, command=_delete).pack(
            side="left")

        _refresh()
        dlg.wait_window()


def _open_edit_dialog(
    parent: tk.Misc,
    profile: dict,
    on_save: Callable,
) -> None:
    """プロファイルの名前・メモを編集する小ダイアログ。"""
    dlg = tk.Toplevel(parent)
    dlg.title("設定セットを編集")
    dlg.resizable(False, False)
    dlg.transient(parent.winfo_toplevel())
    dlg.grab_set()

    frm = tk.Frame(dlg, padx=14, pady=10)
    frm.pack(fill="both", expand=True)
    frm.columnconfigure(1, weight=1)

    name_var = tk.StringVar(value=profile.get("name", ""))
    note_var = tk.StringVar(value=profile.get("note", ""))

    tk.Label(frm, text="名前", anchor="w").grid(
        row=0, column=0, sticky="w", pady=4, padx=(0, 8))
    name_entry = tk.Entry(frm, textvariable=name_var, width=30)
    name_entry.grid(row=0, column=1, sticky="ew", pady=4)

    tk.Label(frm, text="メモ", anchor="w").grid(
        row=1, column=0, sticky="w", pady=4, padx=(0, 8))
    tk.Entry(frm, textvariable=note_var, width=30).grid(
        row=1, column=1, sticky="ew", pady=4)

    btn_frm = tk.Frame(frm)
    btn_frm.grid(row=2, column=0, columnspan=2, pady=(10, 0))

    def do_save() -> None:
        n = name_var.get().strip()
        if not n:
            messagebox.showerror("エラー", "名前を入力してください", parent=dlg)
            return
        cfg.update_profile(profile["id"], name=n, note=note_var.get().strip())
        cfg.save()
        on_save()
        dlg.destroy()

    tk.Button(
        btn_frm, text="保存",
        bg="#4a9eff", fg="white", font=("", 9, "bold"), width=10,
        command=do_save,
    ).pack(side="left", padx=4)
    tk.Button(btn_frm, text="キャンセル", width=10,
              command=dlg.destroy).pack(side="left", padx=4)

    name_entry.focus_set()
    dlg.wait_window()
