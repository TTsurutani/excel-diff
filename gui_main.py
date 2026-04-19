"""PyInstaller EXE ビルド用エントリポイント。
python -m excel_diff_gui での起動は excel_diff_gui/__main__.py を使う。
"""
from excel_diff_gui.app import App

App().mainloop()
