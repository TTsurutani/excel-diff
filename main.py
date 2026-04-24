"""PyInstaller EXE 用エントリポイント。

excel_diff/__main__.py を直接ビルドすると相対インポートが失敗するため、
パッケージ外からモジュールとして呼び出す。
"""
from excel_diff.__main__ import main

if __name__ == "__main__":
    main()
