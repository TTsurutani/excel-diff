"""汎用ユーティリティ関数。"""
from __future__ import annotations

import datetime
from pathlib import Path


def generate_output_dir(old_dir: str, new_dir: str, base_dir: str | None = None) -> str:
    """デフォルト出力フォルダ名を生成する。

    フォルダ名は ``{旧フォルダ名}_vs_{新フォルダ名}_{タイムスタンプ}`` の形式。
    base_dir を指定した場合はその配下のパスを返す。省略時はフォルダ名のみ返す。
    """
    old_name = Path(old_dir).name
    new_name = Path(new_dir).name
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    dirname = f"{old_name}_vs_{new_name}_{ts}"
    if base_dir is not None:
        return str(Path(base_dir) / dirname)
    return dirname
