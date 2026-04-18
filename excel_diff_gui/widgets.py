"""共通ウィジェット。"""
import tkinter as tk
from tkinter import filedialog, ttk
from datetime import datetime
from typing import Optional


class FileSelectRow(tk.Frame):
    """Label + Entry + 参照ボタンの1行。pack() で使う。"""

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
        tk.Entry(self, textvariable=var).pack(side="left", fill="x", expand=True, padx=(0, 4))
        tk.Button(self, text="参照", width=6, command=self._browse).pack(side="left")

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
