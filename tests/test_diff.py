"""
差分エンジンのユニットテスト。
openpyxl に依存せず、データモデルを直接組み立てて検証する。

実行:
  python tests/test_diff.py
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from excel_diff.reader import CellData, RowData, SheetData
from excel_diff.diff_engine import diff_files, RowTag
from excel_diff.matcher import MappingMatcher


# ---------------------------------------------------------------------------
# ヘルパー
# ---------------------------------------------------------------------------

def make_sheet(name: str, data: list[list]) -> SheetData:
    rows = [
        RowData(row_idx=i + 1, cells=[CellData(v) for v in row])
        for i, row in enumerate(data)
    ]
    max_col = max((len(r) for r in data), default=0)
    return SheetData(name=name, rows=rows, max_col=max_col)


def run_diff(old_data, new_data, matchers=None):
    old = {"Sheet1": make_sheet("Sheet1", old_data)}
    new = {"Sheet1": make_sheet("Sheet1", new_data)}
    return diff_files(old, new, "old.xlsx", "new.xlsx", matchers=matchers)


# ---------------------------------------------------------------------------
# テストケース
# ---------------------------------------------------------------------------

PASS = []
FAIL = []


def test(name: str, fn):
    try:
        fn()
        print(f"  PASS  {name}")
        PASS.append(name)
    except AssertionError as e:
        print(f"  FAIL  {name}: {e}")
        FAIL.append(name)
    except Exception as e:
        print(f"  ERROR {name}: {type(e).__name__}: {e}")
        FAIL.append(name)


# --- 差分なし ---

def t_no_diff():
    result = run_diff([["A", "B"], [1, 2]], [["A", "B"], [1, 2]])
    assert not result.has_differences
    sd = result.sheet_diffs[0]
    assert all(rd.tag == RowTag.EQUAL for rd in sd.row_diffs)


# --- セル変更 ---

def t_cell_change():
    result = run_diff([["A", "B"], [1, 2]], [["A", "B"], [1, 99]])
    assert result.has_differences
    sd = result.sheet_diffs[0]
    modified = [rd for rd in sd.row_diffs if rd.tag == RowTag.MODIFY]
    assert len(modified) == 1
    assert len(modified[0].cell_diffs) == 1
    assert modified[0].cell_diffs[0].col_idx == 1
    assert modified[0].cell_diffs[0].old_cell.value == 2
    assert modified[0].cell_diffs[0].new_cell.value == 99


# --- 行挿入（中間） ---

def t_row_insert():
    result = run_diff(
        [["A"], ["B"], ["C"]],
        [["A"], ["X"], ["B"], ["C"]],
    )
    assert result.has_differences
    sd = result.sheet_diffs[0]
    tags = [rd.tag for rd in sd.row_diffs]
    assert RowTag.INSERT in tags
    # A, B, C はすべて EQUAL のまま残る
    equal_vals = [rd.new_row.cells[0].value for rd in sd.row_diffs if rd.tag == RowTag.EQUAL]
    assert "A" in equal_vals
    assert "B" in equal_vals
    assert "C" in equal_vals


# --- 行削除（中間） ---

def t_row_delete():
    result = run_diff(
        [["A"], ["B"], ["C"]],
        [["A"], ["C"]],
    )
    assert result.has_differences
    sd = result.sheet_diffs[0]
    tags = [rd.tag for rd in sd.row_diffs]
    assert RowTag.DELETE in tags
    deleted_vals = [rd.old_row.cells[0].value for rd in sd.row_diffs if rd.tag == RowTag.DELETE]
    assert "B" in deleted_vals


# --- 挿入と削除が混在 ---

def t_insert_and_delete():
    # B を削除して末尾に Y を追加 → DELETE と INSERT が両立する
    # Old: [H, A, B, C, D]
    # New: [H, A, C, D, Y]  ← B削除、Y追加
    # LCS = [H, A, C, D] → B:DELETE, Y:INSERT
    result = run_diff(
        [["ヘッダ"], ["A"], ["B"], ["C"], ["D"]],
        [["ヘッダ"], ["A"], ["C"], ["D"], ["Y"]],
    )
    assert result.has_differences
    sd = result.sheet_diffs[0]
    tags = [rd.tag for rd in sd.row_diffs]
    assert RowTag.INSERT in tags
    assert RowTag.DELETE in tags
    assert RowTag.EQUAL in tags


# --- シート追加 ---

def t_sheet_added():
    old = {"Sheet1": make_sheet("Sheet1", [["A"]])}
    new = {
        "Sheet1": make_sheet("Sheet1", [["A"]]),
        "Sheet2": make_sheet("Sheet2", [["B"]]),
    }
    result = diff_files(old, new, "old.xlsx", "new.xlsx")
    assert result.has_differences
    statuses = {sd.name: sd.status for sd in result.sheet_diffs}
    assert statuses["Sheet1"] == "equal"
    assert statuses["Sheet2"] == "added"


# --- シート削除 ---

def t_sheet_deleted():
    old = {
        "Sheet1": make_sheet("Sheet1", [["A"]]),
        "Sheet2": make_sheet("Sheet2", [["B"]]),
    }
    new = {"Sheet1": make_sheet("Sheet1", [["A"]])}
    result = diff_files(old, new, "old.xlsx", "new.xlsx")
    assert result.has_differences
    statuses = {sd.name: sd.status for sd in result.sheet_diffs}
    assert statuses["Sheet2"] == "deleted"


# --- カスタムマッチャー: 一致（差分なし）---

def t_matcher_equal():
    matcher = MappingMatcher(
        column_idx=1,
        sheet=None,
        pairs=[("旧コード", "新コード")],
    )
    result = run_diff(
        [["名前", "コード"], ["A商事", "旧コード"]],
        [["名前", "コード"], ["A商事", "新コード"]],
        matchers=[matcher],
    )
    assert not result.has_differences, "マッチャー一致なのに差分ありと判定された"


# --- カスタムマッチャー: マッピング外（差分あり）---

def t_matcher_not_in_mapping():
    matcher = MappingMatcher(
        column_idx=1,
        sheet=None,
        pairs=[("旧コード", "新コード")],
    )
    result = run_diff(
        [["名前", "コード"], ["A商事", "旧コード"]],
        [["名前", "コード"], ["A商事", "未知コード"]],  # マッピング外
        matchers=[matcher],
    )
    assert result.has_differences, "マッピング外なのに差分なしと判定された"


# --- カスタムマッチャー: 一部列のみ一致、別列に差分 ---

def t_matcher_partial():
    matcher = MappingMatcher(
        column_idx=1,
        sheet=None,
        pairs=[("旧コード", "新コード")],
    )
    result = run_diff(
        [["名前", "コード", "金額"], ["A商事", "旧コード", 100]],
        [["名前", "コード", "金額"], ["A商事", "新コード", 999]],  # 金額変更
        matchers=[matcher],
    )
    assert result.has_differences
    sd = result.sheet_diffs[0]
    modified = [rd for rd in sd.row_diffs if rd.tag == RowTag.MODIFY]
    assert len(modified) == 1
    # col_idx=1 (コード列) は差分なし、col_idx=2 (金額列) だけ差分
    changed_cols = {cd.col_idx for cd in modified[0].cell_diffs}
    assert 2 in changed_cols, "金額の差分が検出されなかった"
    assert 1 not in changed_cols, "コード列がマッチャーで除外されなかった"


# --- カスタムマッチャー: シート指定（別シートには適用しない）---

def t_matcher_sheet_scope():
    matcher = MappingMatcher(
        column_idx=1,
        sheet="対象シート",  # このシートのみ適用
        pairs=[("旧コード", "新コード")],
    )
    old = {
        "対象シート":   make_sheet("対象シート",   [["X", "旧コード"]]),
        "対象外シート": make_sheet("対象外シート", [["X", "旧コード"]]),
    }
    new = {
        "対象シート":   make_sheet("対象シート",   [["X", "新コード"]]),  # マッチャー適用 → EQUAL
        "対象外シート": make_sheet("対象外シート", [["X", "新コード"]]),  # 適用外 → MODIFY
    }
    result = diff_files(old, new, "old.xlsx", "new.xlsx", matchers=[matcher])
    statuses = {sd.name: sd.status for sd in result.sheet_diffs}
    assert statuses["対象シート"]   == "equal",    "対象シートがEQUALにならなかった"
    assert statuses["対象外シート"] == "modified", "対象外シートがMODIFIEDにならなかった"


# --- 行番号の正確性 ---

def t_row_numbers():
    result = run_diff(
        [["A"], ["B"], ["C"], ["D"]],
        [["A"], ["C"], ["D"]],  # B 削除
    )
    sd = result.sheet_diffs[0]
    deleted = [rd for rd in sd.row_diffs if rd.tag == RowTag.DELETE]
    assert len(deleted) == 1
    assert deleted[0].old_row.row_idx == 2, f"削除行の行番号が {deleted[0].old_row.row_idx} になった（期待値: 2）"

    # C, D の新側行番号が正しいか
    equal_rows = [rd for rd in sd.row_diffs if rd.tag == RowTag.EQUAL and rd.old_row is not None]
    old_idxs = [rd.old_row.row_idx for rd in equal_rows]
    new_idxs = [rd.new_row.row_idx for rd in equal_rows]
    assert old_idxs == [1, 3, 4], f"old行番号が {old_idxs} になった"
    assert new_idxs == [1, 2, 3], f"new行番号が {new_idxs} になった"


# ---------------------------------------------------------------------------
# メイン
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    print("=" * 50)
    print("excel-diff ユニットテスト")
    print("=" * 50)

    test("差分なし",                            t_no_diff)
    test("セル変更",                            t_cell_change)
    test("行挿入（中間）",                      t_row_insert)
    test("行削除（中間）",                      t_row_delete)
    test("挿入と削除が混在",                    t_insert_and_delete)
    test("シート追加",                          t_sheet_added)
    test("シート削除",                          t_sheet_deleted)
    test("マッチャー: 一致→差分なし",           t_matcher_equal)
    test("マッチャー: マッピング外→差分あり",   t_matcher_not_in_mapping)
    test("マッチャー: 一部列のみ一致",           t_matcher_partial)
    test("マッチャー: シートスコープ",           t_matcher_sheet_scope)
    test("行番号の正確性",                      t_row_numbers)

    print("=" * 50)
    print(f"結果: {len(PASS)} PASS / {len(FAIL)} FAIL")
    if FAIL:
        sys.exit(1)
