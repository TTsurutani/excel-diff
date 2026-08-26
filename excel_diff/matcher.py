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
  },
  {
    "type": "numeric",
    "column": "I"
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
_EQUIV_SENTINEL = "__excel_diff_equiv__"
_NUMERIC_SENTINEL = "__excel_diff_numeric__"

# column に "*" を指定した場合、シート内の全列に一括適用する
ALL_COLUMNS = "*"


class ColumnMatcher(ABC):
    """特定列のカスタム等値判定の基底クラス。"""

    def __init__(self, column_idx: Any, sheet: Optional[str]):
        self.column_idx = column_idx      # 0始まりの整数、または ALL_COLUMNS ("*")
        self.sheet = sheet                # None = 全シートに適用

    def applies_to(self, sheet_name: str, col_idx: int) -> bool:
        if self.column_idx != ALL_COLUMNS and col_idx != self.column_idx:
            return False
        if self.sheet is not None and self.sheet != sheet_name:
            return False
        return True

    def can_handle(self, old_val: Any, new_val: Any) -> bool:
        """
        このマッチャー本来のロジックで判定できる値の組か（issue #23）。

        ChainMatcher がサブマッチャーを順に試す際に使う。デフォルトは常に
        True（＝単体使用時は常にこのマッチャーの判定を採用する、従来通り
        の挙動）。「両方が数値の場合のみ扱える」のように適用条件が限定的な
        マッチャー（NumericMatcher 等）はオーバーライドする。
        """
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


def _blank(val: Any) -> Any:
    """空文字列を None に丸める（未入力セルとの表記ゆれを吸収）。"""
    return None if val == "" else val


class EquivalenceMatcher(ColumnMatcher):
    """
    指定した値の集合を常に同一とみなす、対称な同一視マッチャー。

    MappingMatcher と異なり「旧値→新値への変換が完了しているか」を
    検証するものではない。集合に含まれる値同士は、変化の有無や方向に
    関わらず常に等値として扱う（例: "-" と "" を同一視する）。
    """

    def __init__(
        self,
        column_idx: Any,
        sheet: Optional[str],
        values: list[Any],
    ):
        super().__init__(column_idx, sheet)
        group = {_blank(v) for v in values}
        if None in group:
            group.add("")  # 空文字/未入力(None)双方を同一グループとして扱う
        self.group: set[Any] = group

    def _canon(self, val: Any) -> Any:
        return _EQUIV_SENTINEL if _blank(val) in self.group else _blank(val)

    def matches(self, old_val: Any, new_val: Any) -> bool:
        return self._canon(old_val) == self._canon(new_val)

    def normalize_old(self, val: Any) -> Any:
        canon = self._canon(val)
        return (_EQUIV_SENTINEL, canon) if canon is _EQUIV_SENTINEL else val

    def normalize_new(self, val: Any) -> Any:
        canon = self._canon(val)
        return (_EQUIV_SENTINEL, canon) if canon is _EQUIV_SENTINEL else val


class NumericMatcher(ColumnMatcher):
    """
    数値として等価であれば型（int / float / 数字文字列）を問わず同一視する
    マッチャー（issue #14）。

    old_val・new_val の両方が数値に変換できる場合のみ can_handle() が True
    を返し、float() 変換後の値同士を比較する（例: 128 と "128" は同一視さ
    れる）。どちらか一方でも数値に変換できない場合（"-" など）は
    can_handle() が False を返す — 単体で使う場合は通常の等値比較に
    フォールバックするが、ChainMatcher 経由で使う場合は次のサブマッチャー
    （例: equivalence による "-"/空欄 の同一視、issue #20 / #23）に判定を
    委ねられる。

    呼び出し元（_cell_equal / _normalize_row_key）で渡される値は既に
    _normalize_val() を経ている（_x000D_ 除去・改行コード統一・空文字列→None
    済み）前提のため、このクラス自身では文字列の表記ゆれ吸収は行わない。
    """

    @staticmethod
    def _to_float(val: Any) -> Optional[float]:
        if val is None:
            return None
        # bool は int のサブクラスで True==1 / False==0 と評価されるため、
        # 意図せず数値と一致してしまわないよう明示的に対象外とする。
        if isinstance(val, bool):
            return None
        if isinstance(val, (int, float)):
            return float(val)
        if isinstance(val, str):
            try:
                return float(val)
            except ValueError:
                return None
        return None

    def can_handle(self, old_val: Any, new_val: Any) -> bool:
        return (
            self._to_float(old_val) is not None
            and self._to_float(new_val) is not None
        )

    def matches(self, old_val: Any, new_val: Any) -> bool:
        old_num = self._to_float(old_val)
        new_num = self._to_float(new_val)
        if old_num is not None and new_num is not None:
            return old_num == new_num
        # 単体使用時（ChainMatcher を介さない場合）は通常比較にフォールバック
        return old_val == new_val

    def _normalize(self, val: Any) -> Any:
        num = self._to_float(val)
        return (_NUMERIC_SENTINEL, num) if num is not None else val

    def normalize_old(self, val: Any) -> Any:
        return self._normalize(val)

    def normalize_new(self, val: Any) -> Any:
        return self._normalize(val)


class ChainMatcher(ColumnMatcher):
    """
    複数のマッチャーを順に試し、最初に can_handle() が True を返した
    サブマッチャーの matches() / normalize_old() / normalize_new() を
    採用する合成マッチャー（issue #23）。

    「numeric で扱えない値（"-" など）は equivalence にフォールバックする」
    といった、独立した複数の関心事を1列に重ねて適用したい場合に使う。
    NumericMatcher.blank_values（issue #20）のような専用オプションを個別の
    マッチャー実装に追加していく代わりに、既存のマッチャーをそのまま組み
    合わせられる。

    どのサブマッチャーも can_handle() が False を返した場合は、最後の
    サブマッチャーの判定・正規化をそのまま採用する（フォールバック段は
    常に can_handle=True となる equivalence 等を置く運用を想定）。
    """

    def __init__(
        self,
        column_idx: Any,
        sheet: Optional[str],
        sub_matchers: list[ColumnMatcher],
    ):
        super().__init__(column_idx, sheet)
        if not sub_matchers:
            raise ValueError("ChainMatcher には1つ以上のサブマッチャーが必要です")
        self.sub_matchers = sub_matchers

    def _select(self, old_val: Any, new_val: Any) -> ColumnMatcher:
        for m in self.sub_matchers:
            if m.can_handle(old_val, new_val):
                return m
        return self.sub_matchers[-1]

    def can_handle(self, old_val: Any, new_val: Any) -> bool:
        # chain自体は常に何らかのサブマッチャーへフォールバックできるため True
        return True

    def matches(self, old_val: Any, new_val: Any) -> bool:
        return self._select(old_val, new_val).matches(old_val, new_val)

    def normalize_old(self, val: Any) -> Any:
        # normalize時点では相手側の値が分からないため、can_handle は
        # 自分自身の値だけで判定できるマッチャー（数値変換可否など）を
        # 前提とする。old側・new側で選ばれるサブマッチャーが食い違わない
        # よう、can_handle(val, val) で自己判定する。
        return self._select(val, val).normalize_old(val)

    def normalize_new(self, val: Any) -> Any:
        return self._select(val, val).normalize_new(val)


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


def parse_col_list(spec: str) -> list[int]:
    """
    列指定文字列を 0始まりインデックスの **順序付きリスト** に変換する。
    key_cols のように列の順序が複合キーの構成順に影響する場合に使う。

    Examples
    --------
    "B,C"   → [1, 2]
    "C,B"   → [2, 1]   # 指定順を保持
    "B"     → [1]
    """
    result: list[int] = []
    for part in spec.split(","):
        part = part.strip()
        if part:
            result.append(_parse_column(part))
    return result


# ---------------------------------------------------------------------------
# DiffConfig: マッチャー + 列フィルタ + 差分モードをまとめる設定オブジェクト
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
    diff_mode:
        差分計算モード。"lcs"（デフォルト）または "key"。
        "key" の場合は key_cols で指定した列をキーとして行を JOIN する。
    key_cols:
        キー JOIN モード時の複合キー列（0始まりインデックスのリスト）。
        指定順がキーの構成順に対応する（例: [1, 2] → B列・C列の複合キー）。
        diff_mode が "lcs" のときは無視される。
    sub_key_cols:
        主キーで一意に対応付けられなかった行を救済するための、2段目の
        照合キー（0始まりインデックスのリスト）。key_cols と重複する列は
        指定できない。
    """

    def __init__(
        self,
        matchers: Optional[list[ColumnMatcher]] = None,
        global_col_filter: Optional[set[int]] = None,
        sheet_col_filters: Optional[dict[str, set[int]]] = None,
        diff_mode: str = "lcs",
        key_cols: Optional[list[int]] = None,
        sub_key_cols: Optional[list[int]] = None,
    ):
        self.matchers: list[ColumnMatcher] = matchers or []
        self.global_col_filter: Optional[set[int]] = global_col_filter
        self.sheet_col_filters: dict[str, set[int]] = sheet_col_filters or {}
        self.diff_mode: str = diff_mode          # "lcs" or "key"
        self.key_cols: list[int] = key_cols or []
        self.sub_key_cols: list[int] = sub_key_cols or []
        self.validate_subkey_config()

    def validate_subkey_config(self) -> None:
        """
        sub_key_cols の設定妥当性を検証する。__init__ から呼ばれるほか、
        CLI（_build_config）のように属性を後から書き換える経路でも、
        全属性を確定させた後に明示的に呼び出して検証する。
        """
        if self.sub_key_cols:
            if self.diff_mode != "key":
                raise ValueError(
                    "sub_key_cols は diff_mode が 'key' のときのみ指定できます"
                )
            if not self.key_cols:
                raise ValueError(
                    "sub_key_cols を指定する場合は key_cols も指定してください"
                )
            overlap = set(self.key_cols) & set(self.sub_key_cols)
            if overlap:
                raise ValueError(
                    f"key_cols と sub_key_cols に重複する列があります: {sorted(overlap)}"
                )

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

    # diff_mode / key_cols / sub_key_cols
    diff_mode: str = raw.get("diff_mode", "lcs")
    key_cols: list[int] = []
    if "key_cols" in raw:
        raw_keys = raw["key_cols"]
        if isinstance(raw_keys, str):
            key_cols = parse_col_list(raw_keys)
        elif isinstance(raw_keys, list):
            key_cols = [_parse_column(c) for c in raw_keys]

    sub_key_cols: list[int] = []
    if "sub_key_cols" in raw:
        raw_sub_keys = raw["sub_key_cols"]
        if isinstance(raw_sub_keys, str):
            sub_key_cols = parse_col_list(raw_sub_keys)
        elif isinstance(raw_sub_keys, list):
            sub_key_cols = [_parse_column(c) for c in raw_sub_keys]

    return DiffConfig(
        matchers=matchers,
        global_col_filter=global_filter,
        sheet_col_filters=sheet_filters,
        diff_mode=diff_mode,
        key_cols=key_cols,
        sub_key_cols=sub_key_cols,
    )


def _parse_matcher_column(col_spec: Any) -> Any:
    """
    マッチャーの column 指定を解釈する。
    "*" の場合はシート内の全列に一括適用する ALL_COLUMNS を返す。
    """
    if isinstance(col_spec, str) and col_spec.strip() == ALL_COLUMNS:
        return ALL_COLUMNS
    return _parse_column(col_spec)


def _build_matcher(
    entry: dict, col_idx: Any, sheet: Optional[str], base_dir: str
) -> ColumnMatcher:
    """
    1つのマッチャーエントリを ColumnMatcher に変換する。
    "chain" タイプの "of" 配下エントリの変換にも再帰的に使う（issue #23）。
    サブエントリに "column"/"sheet" を書く必要はない（chain全体の値を継承）。
    """
    matcher_type = entry.get("type", "mapping")

    if matcher_type == "mapping":
        pairs = [(p[0], p[1]) for p in entry["pairs"]]
        return MappingMatcher(col_idx, sheet, pairs)

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

        return MappingMatcher(col_idx, sheet, pairs)

    elif matcher_type == "equivalence":
        values = entry["values"]
        return EquivalenceMatcher(col_idx, sheet, values)

    elif matcher_type == "numeric":
        return NumericMatcher(col_idx, sheet)

    elif matcher_type == "chain":
        sub_entries = entry["of"]
        sub_matchers = [
            _build_matcher(sub, col_idx, sheet, base_dir) for sub in sub_entries
        ]
        return ChainMatcher(col_idx, sheet, sub_matchers)

    else:
        raise ValueError(f"未知のマッチャータイプ: {matcher_type!r}")


def _parse_matchers(entries: list, base_dir: str) -> list[ColumnMatcher]:
    """マッチャーエントリのリストを ColumnMatcher リストに変換する。"""
    matchers: list[ColumnMatcher] = []

    for entry in entries:
        col_idx = _parse_matcher_column(entry["column"])
        sheet = entry.get("sheet")
        matchers.append(_build_matcher(entry, col_idx, sheet, base_dir))

    return matchers


# 後方互換エイリアス
def load_matchers(config_path: str) -> list[ColumnMatcher]:
    """後方互換用。新規コードは load_config() を使用すること。"""
    return load_config(config_path).matchers
