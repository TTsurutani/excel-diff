"""タブ③ シート分解（スケルトン）。"""
import tkinter as tk
from typing import Callable


class TabSplit(tk.Frame):
    def __init__(self, parent, log: Callable[[str], None]) -> None:
        super().__init__(parent)
        tk.Label(self, text="タブ③ シート分解（未実装）", font=("", 12)).pack(padx=20, pady=30)
