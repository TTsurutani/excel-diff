"""共通ウィジェット。"""
import tkinter as tk
from tkinter import filedialog, ttk
from datetime import datetime
from typing import Optional


def _parse_dnd_path(data: str) -> str:
    """DnD イベントの data から最初のパスを取り出す。
    Windows では空白を含むパスが {} で囲まれる。
    """
    data = data.strip()
    if data.startswith("{"):
        end = data.find("}")
        if end > 0:
            return data[1:end]
    return data.split()[0] if data else ""


class FileSelectRow(tk.Frame):
    """Label + Entry + 参照ボタンの1行。pack() で使う。
    tkinterdnd2 が利用可能なとき Entry へのファイルドロップも受け付ける。
    """

    def __init__(
        self,
        parent,
        label: str,
        var: tk.StringVar,
        mode: str = "file",           # "file" or "dir"
        filetypes: Optional[list] = None,
    ) -> None:
        super().__init__(parent)
        self._var = var
        self._mode = mode
        self._filetypes = filetypes or [("All files", "*.*")]

        tk.Label(self, text=label, width=14, anchor="w").pack(side="left")
        self._entry = tk.Entry(self, textvariable=var)
        self._entry.pack(side="left", fill="x", expand=True, padx=(0, 4))
        tk.Button(self, text="参照", width=6, command=self._browse).pack(side="left")

        self._setup_dnd(self._entry)

    def _setup_dnd(self, entry: tk.Entry) -> None:
        try:
            from tkinterdnd2 import DND_FILES
            entry.drop_target_register(DND_FILES)
            entry.dnd_bind("<<DragEnter>>", self._on_drag_enter)
            entry.dnd_bind("<<DragLeave>>", self._on_drag_leave)
            entry.dnd_bind("<<Drop>>",      self._on_drop)
        except Exception:
            pass

    def _on_drag_enter(self, event) -> None:
        event.widget.configure(background="#cce5ff")

    def _on_drag_leave(self, event) -> None:
        event.widget.configure(background="white")

    def _on_drop(self, event) -> None:
        event.widget.configure(background="white")
        path = _parse_dnd_path(event.data)
        if path:
            self._var.set(path)

    def _browse(self) -> None:
        if self._mode == "dir":
            path = filedialog.askdirectory()
        else:
            path = filedialog.askopenfilename(filetypes=self._filetypes)
        if path:
            self._var.set(path)


class LogArea(tk.Frame):
    """タイムスタンプ付きログ表示エリア。"""

    def __init__(self, parent, height: int = 7) -> None:
        super().__init__(parent)

        bar = tk.Frame(self)
        bar.pack(fill="x")
        tk.Label(bar, text="ログ", font=("", 9, "bold")).pack(side="left")
        tk.Button(bar, text="クリア", command=self.clear, width=7).pack(side="right")
        tk.Button(bar, text="コピー", command=self._copy, width=7).pack(side="right", padx=2)

        frame = tk.Frame(self)
        frame.pack(fill="both", expand=True)
        sb = ttk.Scrollbar(frame)
        sb.pack(side="right", fill="y")
        self._text = tk.Text(
            frame, height=height, state="disabled",
            font=("Courier New", 9), yscrollcommand=sb.set,
        )
        self._text.pack(side="left", fill="both", expand=True)
        sb.config(command=self._text.yview)

    def log(self, msg: str) -> None:
        ts = datetime.now().strftime("%H:%M:%S")
        self._text.config(state="normal")
        self._text.insert("end", f"[{ts}] {msg}\n")
        self._text.see("end")
        self._text.config(state="disabled")

    def clear(self) -> None:
        self._text.config(state="normal")
        self._text.delete("1.0", "end")
        self._text.config(state="disabled")

    def _copy(self) -> None:
        self.clipboard_clear()
        self.clipboard_append(self._text.get("1.0", "end"))
