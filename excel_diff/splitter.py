"""
Excelブックをシート単位のファイルに分解するモジュール。

split_workbook(path, prefix, suffix, name_regex, output_dir) → list[str]
  各シートを <output_dir>/<prefix><ファイル名ベース><suffix>.xlsx として保存し、
  出力ファイルパスのリストを返す。
"""
from __future__ import annotations

import re
import warnings
from pathlib import Path


# ファイル名として使えない文字（Windows / macOS / Linux 共通の危険文字）
_INVALID_CHARS = re.compile(r'[\\/:*?"<>|]')


def _safe_filename(sheet_name: str) -> str:
    """シート名をファイル名として安全な文字列に変換する。"""
    return _INVALID_CHARS.sub("_", sheet_name)


def _apply_name_regex(sheet_name: str, pattern: re.Pattern[str]) -> str:
    """
    正規表現の第1キャプチャグループにマッチした部分を返す。
    マッチしない場合はシート名全体にフォールバックして警告を出す。
    """
    m = pattern.search(sheet_name)
    if m and m.lastindex and m.lastindex >= 1:
        return m.group(1)
    warnings.warn(
        f"--name-regex がシート '{sheet_name}' にマッチしませんでした。シート名をそのまま使用します。",
        stacklevel=3,
    )
    return sheet_name


def split_workbook(
    path: str,
    prefix: str = "",
    suffix: str = "",
    name_regex: str | None = None,
    output_dir: str | None = None,
) -> list[str]:
    """
    ブックを1シート1ファイルに分解して保存する。

    Parameters
    ----------
    path       : 入力Excelファイルパス (.xlsx)
    prefix     : 出力ファイル名の前置文字列
    suffix     : 出力ファイル名の後置文字列（拡張子の前）
    name_regex : ファイル名ベース抽出用正規表現（第1キャプチャグループを使用）。
                 Noneの場合はシート名をそのまま使用。
    output_dir : 出力先ディレクトリ（Noneの場合はブックと同じフォルダ）

    Returns
    -------
    出力ファイルパスのリスト（シート順）
    """
    try:
        from openpyxl import load_workbook
    except ImportError as e:
        raise ImportError("openpyxl が必要です: pip install openpyxl") from e

    compiled_regex: re.Pattern[str] | None = None
    if name_regex:
        compiled_regex = re.compile(name_regex)
        if compiled_regex.groups < 1:
            raise ValueError(
                f"--name-regex にはキャプチャグループ () が1つ以上必要です: {name_regex!r}"
            )

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

        # 名前定義を全削除（他シート参照による #REF! エラー防止）
        for dn in list(wb.defined_names):
            del wb.defined_names[dn]

        # ファイル名ベースを決定
        if compiled_regex is not None:
            name_base = _apply_name_regex(sheet_name, compiled_regex)
        else:
            name_base = sheet_name

        safe_name = _safe_filename(name_base)
        out_filename = f"{prefix}{safe_name}{suffix}.xlsx"
        out_path = out_dir / out_filename
        wb.save(str(out_path))
        wb.close()
        output_paths.append(str(out_path))

    return output_paths
