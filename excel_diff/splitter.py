"""
Excelブックをシート単位のファイルに分解するモジュール。

split_workbook(path, prefix, suffix, output_dir) → list[str]
  各シートを <output_dir>/<prefix><シート名><suffix>.xlsx として保存し、
  出力ファイルパスのリストを返す。
"""
from __future__ import annotations

import re
from pathlib import Path


# ファイル名として使えない文字（Windows / macOS / Linux 共通の危険文字）
_INVALID_CHARS = re.compile(r'[\\/:*?"<>|]')


def _safe_filename(sheet_name: str) -> str:
    """シート名をファイル名として安全な文字列に変換する。"""
    return _INVALID_CHARS.sub("_", sheet_name)


def split_workbook(
    path: str,
    prefix: str = "",
    suffix: str = "",
    output_dir: str | None = None,
) -> list[str]:
    """
    ブックを1シート1ファイルに分解して保存する。

    Parameters
    ----------
    path       : 入力Excelファイルパス (.xlsx)
    prefix     : 出力ファイル名の前置文字列
    suffix     : 出力ファイル名の後置文字列（拡張子の前）
    output_dir : 出力先ディレクトリ（Noneの場合はブックと同じフォルダ）

    Returns
    -------
    出力ファイルパスのリスト（シート順）
    """
    try:
        from openpyxl import load_workbook
    except ImportError as e:
        raise ImportError("openpyxl が必要です: pip install openpyxl") from e

    src_path = Path(path)
    out_dir = Path(output_dir) if output_dir else src_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    # シート名を事前確認（ファイルを1回だけ開く）
    wb_meta = load_workbook(path, read_only=True, data_only=True)
    sheet_names = wb_meta.sheetnames
    wb_meta.close()

    output_paths: list[str] = []

    for sheet_name in sheet_names:
        # 元ブックを毎回フルロードしてシート以外を削除
        wb = load_workbook(path)
        for name in wb.sheetnames:
            if name != sheet_name:
                del wb[name]

        safe_name = _safe_filename(sheet_name)
        out_filename = f"{prefix}{safe_name}{suffix}.xlsx"
        out_path = out_dir / out_filename
        wb.save(str(out_path))
        wb.close()
        output_paths.append(str(out_path))

    return output_paths
