"""
カスタムマッチャーモジュール。

特定の列に対して「旧値→新値」が意図的な変換である場合に
差分なしとして扱うための仕組みを提供する。

設定ファイル（JSON）の例:
[
  {
    "type": "mapping",
    "column": "B",
    "sheet": null,
    "pairs": [["旧コード001", "新コード001"], ["旧コード002", "新コード002"]]
  },
  {
    "type": "mapping_file",
    "column": "C",
    "sheet": "売上",
    "file": "code_mapping.csv",
    "old_col": 0,
    "new_col": 1,
    "has_header": true
  }
]
"""
from __future__ import annotations

import csv
import json
import os
from abc import ABC, abstractmethod
from typing import Any, Optional

from openpyxl.utils import column_index_from_string


# ---------------------------------------------------------------------------
# 正規化キー用センチネル
# ---------------------------------------------------------------------------
_MAPPED_SENTINEL = "__excel_diff_mapped__"


class ColumnMatcher(ABC):
    """特定列のカスタム等値判定の基底クラス。"""

    def __init__(self, column_idx: int, sheet: Optional[str]):
        self.column_idx = column_idx      # 0始まり
        self.sheet = sheet                # None = 全シートに適用

    def applies_to(self, sheet_name: str, col_idx: int) -> bool:
        if col_idx != self.column_idx:
            return False
        if self.sheet is not None and self.sheet != sheet_name:
            return False
        return True

    @abstractmethod
    def matches(self, old_val: Any, new_val: Any) -> bool:
        """旧値と新値が「等値」とみなせる場合 True を返す。"""

    @abstractmethod
    def normalize_old(self, val: Any) -> Any:
        """
        旧ファイル側のセル値をLCS用正規化キーに変換する。
        マッピングのキーに該当する場合は (_MAPPED_SENTINEL, old_val) を返す。
        """

    @abstractmethod
    def normalize_new(self, val: Any) -> Any:
        """
        新ファイル側のセル値をLCS用正規化キーに変換する。
        マッピングの値（変換後）に該当する場合は (_MAPPED_SENTINEL, old_val) を返す。
        """


class MappingMatcher(ColumnMatcher):
    """
    対比表（旧値 → 新値）によるマッチャー。
    旧値が forward のキーに存在し、新値が期待値と一致すれば差分なし。
    """

    def __init__(
        self,
        column_idx: int,
        sheet: Optional[str],
        pairs: list[tuple[Any, Any]],
    ):
        super().__init__(column_idx, sheet)
        self.forward: dict[Any, Any] = {old: new for old, new in pairs}
        self.reverse: dict[Any, Any] = {new: old for old, new in pairs}

    def matches(self, old_val: Any, new_val: Any) -> bool:
        if old_val in self.forward:
            return self.forward[old_val] == new_val
        # 旧値がマッピングのキーにない場合は通常等値比較
        return old_val == new_val

    def normalize_old(self, val: Any) -> Any:
        if val in self.forward:
            return (_MAPPED_SENTINEL, val)
        return val

    def normalize_new(self, val: Any) -> Any:
        if val in self.reverse:
            return (_MAPPED_SENTINEL, self.reverse[val])
        return val


# ---------------------------------------------------------------------------
# ファクトリ関数
# ---------------------------------------------------------------------------

def _parse_column(col_spec: Any) -> int:
    """列指定を 0始まりインデックスに変換する。A=0, B=1 など。"""
    if isinstance(col_spec, int):
        return col_spec
    if isinstance(col_spec, str):
        # 数字文字列なら整数として扱う
        if col_spec.isdigit():
            return int(col_spec)
        # 列記号 (A, B, AA, ...) → 0始まりに変換
        return column_index_from_string(col_spec.upper()) - 1
    raise ValueError(f"列指定が不正です: {col_spec!r}")


def _load_pairs_from_csv(
    file_path: str,
    old_col: Any,
    new_col: Any,
    has_header: bool,
    base_dir: str,
) -> list[tuple[Any, Any]]:
    """CSVファイルから (旧値, 新値) のペアリストを読み込む。"""
    full_path = os.path.join(base_dir, file_path) if not os.path.isabs(file_path) else file_path
    pairs: list[tuple[Any, Any]] = []

    with open(full_path, encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        rows = list(reader)

    if has_header and rows:
        header = rows[0]
        rows = rows[1:]
        # 列名指定のサポート
        if isinstance(old_col, str) and not old_col.isdigit():
            old_col = header.index(old_col)
        if isinstance(new_col, str) and not new_col.isdigit():
            new_col = header.index(new_col)

    old_idx = int(old_col)
    new_idx = int(new_col)

    for row in rows:
        if len(row) > max(old_idx, new_idx):
            pairs.append((row[old_idx], row[new_idx]))

    return pairs


def _load_pairs_from_xlsx(
    file_path: str,
    old_col: Any,
    new_col: Any,
    has_header: bool,
    base_dir: str,
) -> list[tuple[Any, Any]]:
    """Excelファイルから (旧値, 新値) のペアリストを読み込む。"""
    import openpyxl
    full_path = os.path.join(base_dir, file_path) if not os.path.isabs(file_path) else file_path
    wb = openpyxl.load_workbook(full_path, data_only=True, read_only=True)
    ws = wb.active
    pairs: list[tuple[Any, Any]] = []

    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if has_header and rows:
        header = list(rows[0])
        rows = rows[1:]
        if isinstance(old_col, str) and not str(old_col).isdigit():
            old_col = header.index(old_col)
        if isinstance(new_col, str) and not str(new_col).isdigit():
            new_col = header.index(new_col)

    old_idx = int(old_col)
    new_idx = int(new_col)

    for row in rows:
        if len(row) > max(old_idx, new_idx):
            pairs.append((row[old_idx], row[new_idx]))

    return pairs


def parse_col_spec(spec: str) -> set[int]:
    """
    列範囲指定文字列を 0始まりインデックスの集合に変換する。

    Examples
    --------
    "A"       → {0}
    "A:C"     → {0, 1, 2}
    "A:C,E"   → {0, 1, 2, 4}
    "A,C:E,G" → {0, 2, 3, 4, 6}
    "1,3:5"   → {0, 2, 3, 4}   # 1始まり整数も受け付ける
    """
    result: set[int] = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if ":" in part:
            left, right = part.split(":", 1)
            result.update(range(_parse_column(left.strip()), _parse_column(right.strip()) + 1))
        else:
            result.add(_parse_column(part))
    return result


# ---------------------------------------------------------------------------
# DiffConfig: マッチャー + 列フィルタをまとめる設定オブジェクト
# ---------------------------------------------------------------------------

class DiffConfig:
    """
    diff_files() に渡す設定をまとめたクラス。

    Attributes
    ----------
    matchers:
        カスタムマッチャーのリスト
    global_col_filter:
        全シートに適用する列フィルタ（0始まりインデックスの集合）。
        None の場合は全列比較。
    sheet_col_filters:
        シート名ごとの列フィルタ。global_col_filter より優先される。
    """

    def __init__(
        self,
        matchers: Optional[list[ColumnMatcher]] = None,
        global_col_filter: Optional[set[int]] = None,
        sheet_col_filters: Optional[dict[str, set[int]]] = None,
    ):
        self.matchers: list[ColumnMatcher] = matchers or []
        self.global_col_filter: Optional[set[int]] = global_col_filter
        self.sheet_col_filters: dict[str, set[int]] = sheet_col_filters or {}

    def get_col_filter(self, sheet_name: str) -> Optional[set[int]]:
        """シート名に対応する列フィルタを返す（なければ全列）。"""
        if sheet_name in self.sheet_col_filters:
            return self.sheet_col_filters[sheet_name]
        return self.global_col_filter

    @property
    def matcher_count(self) -> int:
        return len(self.matchers)


def load_config(config_path: str) -> DiffConfig:
    """
    JSONファイルから DiffConfig を生成して返す。

    対応フォーマット:
    1. 旧来の配列形式（マッチャーのみ）:
       [ {"type": "mapping", ...}, ... ]

    2. 新形式（列フィルタ + マッチャー）:
       {
         "include_cols": "A:C,E",
         "sheets": {
           "売上": { "include_cols": "A,C:F" }
         },
         "matchers": [ {"type": "mapping", ...} ]
       }
    """
    base_dir = os.path.dirname(os.path.abspath(config_path))

    with open(config_path, encoding="utf-8") as f:
        raw = json.load(f)

    # --- 旧来フォーマット（配列）への後方互換 ---
    if isinstance(raw, list):
        matchers = _parse_matchers(raw, base_dir)
        return DiffConfig(matchers=matchers)

    # --- 新形式（辞書）---
    matchers = _parse_matchers(raw.get("matchers", []), base_dir)

    global_filter: Optional[set[int]] = None
    if "include_cols" in raw:
        global_filter = parse_col_spec(str(raw["include_cols"]))

    sheet_filters: dict[str, set[int]] = {}
    for sheet_name, sheet_cfg in raw.get("sheets", {}).items():
        if "include_cols" in sheet_cfg:
            sheet_filters[sheet_name] = parse_col_spec(str(sheet_cfg["include_cols"]))

    return DiffConfig(
        matchers=matchers,
        global_col_filter=global_filter,
        sheet_col_filters=sheet_filters,
    )


def _parse_matchers(entries: list, base_dir: str) -> list[ColumnMatcher]:
    """マッチャーエントリのリストを ColumnMatcher リストに変換する。"""
    matchers: list[ColumnMatcher] = []

    for entry in entries:
        col_idx = _parse_column(entry["column"])
        sheet = entry.get("sheet")
        matcher_type = entry.get("type", "mapping")

        if matcher_type == "mapping":
            pairs = [(p[0], p[1]) for p in entry["pairs"]]
            matchers.append(MappingMatcher(col_idx, sheet, pairs))

        elif matcher_type == "mapping_file":
            file_path = entry["file"]
            old_col = entry.get("old_col", 0)
            new_col = entry.get("new_col", 1)
            has_header = entry.get("has_header", False)

            ext = os.path.splitext(file_path)[1].lower()
            if ext in (".xlsx", ".xlsm"):
                pairs = _load_pairs_from_xlsx(file_path, old_col, new_col, has_header, base_dir)
            else:
                pairs = _load_pairs_from_csv(file_path, old_col, new_col, has_header, base_dir)

            matchers.append(MappingMatcher(col_idx, sheet, pairs))

        else:
            raise ValueError(f"未知のマッチャータイプ: {matcher_type!r}")

    return matchers


# 後方互換エイリアス
def load_matchers(config_path: str) -> list[ColumnMatcher]:
    """後方互換用。新規コードは load_config() を使用すること。"""
    return load_config(config_path).matchers
