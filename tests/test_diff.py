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
from excel_diff.matcher import (
    MappingMatcher,
    EquivalenceMatcher,
    ALL_COLUMNS,
    DiffConfig,
    parse_col_spec,
    load_config,
)
from excel_diff.html_renderer import _render_cell_pair_diff, _render_row_pair
from excel_diff.diff_engine import RowDiff, CellDiff


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


def run_diff(old_data, new_data, matchers=None, config=None):
    old = {"Sheet1": make_sheet("Sheet1", old_data)}
    new = {"Sheet1": make_sheet("Sheet1", new_data)}
    return diff_files(old, new, "old.xlsx", "new.xlsx", matchers=matchers, config=config)


# ---------------------------------------------------------------------------
# テストケース
# ---------------------------------------------------------------------------

PASS = []
FAIL = []


def _run_test(name: str, fn):
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


# --- カスタムマッチャー: equivalence（"-" と "" を同一視）---

def t_equivalence_hyphen_blank():
    matcher = EquivalenceMatcher(
        column_idx=1,
        sheet=None,
        values=["-", ""],
    )
    result = run_diff(
        [["名前", "備考"], ["A商事", "-"], ["B商事", ""], ["C商事", "変化なし"]],
        [["名前", "備考"], ["A商事", ""], ["B商事", "-"], ["C商事", "変化なし"]],
        matchers=[matcher],
    )
    assert not result.has_differences, "「-」と空白の同一視で差分が検出された"


def t_equivalence_unchanged_not_flagged():
    """マッピングと違い、変化していない値まで誤検出しないこと。"""
    matcher = EquivalenceMatcher(
        column_idx=1,
        sheet=None,
        values=["-", ""],
    )
    result = run_diff(
        [["名前", "備考"], ["A商事", "-"], ["B商事", ""]],
        [["名前", "備考"], ["A商事", "-"], ["B商事", ""]],  # 完全に無変更
        matchers=[matcher],
    )
    assert not result.has_differences, "無変更の行が誤って差分ありと判定された"


def t_equivalence_outside_group_still_diffs():
    matcher = EquivalenceMatcher(
        column_idx=1,
        sheet=None,
        values=["-", ""],
    )
    result = run_diff(
        [["名前", "備考"], ["A商事", "-"]],
        [["名前", "備考"], ["A商事", "別の値"]],  # 同一視グループ外への変化
        matchers=[matcher],
    )
    assert result.has_differences, "グループ外への変化が差分なしと判定された"


def t_equivalence_all_columns_wildcard():
    matcher = EquivalenceMatcher(
        column_idx=ALL_COLUMNS,
        sheet=None,
        values=["-", ""],
    )
    result = run_diff(
        [["-", "値A"], ["", "値B"]],
        [["", "値A"], ["-", "値B"]],  # 全列で "-"/"" が入れ替わっているだけ
        matchers=[matcher],
    )
    assert not result.has_differences, "全列ワイルドカードで差分が検出された"


# --- カスタムマッチャー適用列でも _x000D_ / 改行コード差異を正規化する（issue #19）---

def t_equivalence_normalizes_x000d():
    """
    全列適用(ALL_COLUMNS)の equivalence マッチャーが指定されていても、
    Excel由来の _x000D_ (セル内改行の内部テキスト表現) や CRLF/CR の
    表記ゆれは差分として検出されないこと。
    """
    matcher = EquivalenceMatcher(
        column_idx=ALL_COLUMNS,
        sheet=None,
        values=["-", ""],
    )
    result = run_diff(
        [["名前", "選択肢"], ["A商事", "0:いいえ_x000D_\n1:はい"]],
        [["名前", "選択肢"], ["A商事", "0:いいえ\n1:はい"]],
        matchers=[matcher],
    )
    assert not result.has_differences, "_x000D_ の表記ゆれがマッチャー適用列で誤検出された"


def t_mapping_matcher_normalizes_x000d():
    """MappingMatcher が適用される列でも同様に正規化されること。"""
    matcher = MappingMatcher(
        column_idx=1,
        sheet=None,
        pairs=[("旧コード", "新コード")],
    )
    result = run_diff(
        [["名前", "コード"], ["A商事", "説明\r\n続き"]],
        [["名前", "コード"], ["A商事", "説明\n続き"]],  # CRLF → LF のみの差
        matchers=[matcher],
    )
    assert not result.has_differences, "CRLF表記ゆれがMappingMatcher適用列で誤検出された"


def t_lcs_row_key_normalizes_x000d_unmatched_column():
    """
    マッチャー未適用列でも、LCSの行マッチング用キー生成(_normalize_row_key)が
    _x000D_ / 改行コードを正規化すること（他列は完全一致の行がEQUALになる）。
    """
    result = run_diff(
        [["No", "選択肢"], [1, "0:いいえ_x000D_\n1:はい"]],
        [["No", "選択肢"], [1, "0:いいえ\n1:はい"]],
    )
    assert not result.has_differences, "_x000D_ の表記ゆれがLCS行マッチングで誤検出された"


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
# 列フィルタ テスト
# ---------------------------------------------------------------------------

def t_parse_col_spec_single():
    """parse_col_spec: 単一列 "B" → {1}"""
    assert parse_col_spec("B") == {1}


def t_parse_col_spec_range():
    """parse_col_spec: 連続範囲 "A:C" → {0,1,2}"""
    assert parse_col_spec("A:C") == {0, 1, 2}


def t_parse_col_spec_noncontiguous():
    """parse_col_spec: 飛び地 "A,C,E" → {0,2,4}"""
    assert parse_col_spec("A,C,E") == {0, 2, 4}


def t_parse_col_spec_mixed():
    """parse_col_spec: 混在 "A:C,E" → {0,1,2,4}"""
    assert parse_col_spec("A:C,E") == {0, 1, 2, 4}


def t_col_filter_excludes_diff():
    """列フィルタ: 除外列の変更は差分として検出しない"""
    # 列A(idx=0)のみ比較、列B(idx=1)は除外
    # 列Bが変わっても差分なし
    config = DiffConfig(global_col_filter={0})
    result = run_diff(
        [["同じ", "旧値"]],
        [["同じ", "新値"]],
        config=config,
    )
    assert not result.has_differences, "除外列の変更が差分として検出された"


def t_col_filter_detects_included_diff():
    """列フィルタ: 対象列の変更は差分として検出し、除外列は cell_diffs に含まない"""
    # 列A・B(idx=0,1)を比較対象、列C(idx=2)は除外
    # 列B が変わる → MODIFY として検出され、cell_diffs は col_idx=1 のみ
    config = DiffConfig(global_col_filter={0, 1})
    result = run_diff(
        [["同じID", "旧値B", "除外列値"]],
        [["同じID", "新値B", "除外列値変更"]],
        config=config,
    )
    assert result.has_differences, "対象列の変更が検出されなかった"
    sd = result.sheet_diffs[0]
    modified = [rd for rd in sd.row_diffs if rd.tag == RowTag.MODIFY]
    assert len(modified) == 1, f"MODIFY行が {len(modified)} 件（期待値: 1）"
    changed_cols = {cd.col_idx for cd in modified[0].cell_diffs}
    assert 1 in changed_cols, "B列(idx=1)の差分が検出されなかった"
    assert 0 not in changed_cols, "A列(idx=0)が誤って差分に含まれた"
    assert 2 not in changed_cols, "除外列(idx=2)が cell_diffs に含まれた"


def t_col_filter_row_lcs():
    """列フィルタ: 除外列の相違がLCSに影響しない（行マッチが正しく行われる）"""
    # 列B(idx=1)は除外（列A=idx=0のみ比較）
    # 旧行: ["A", "X"]  新行: ["A", "Y"] → 除外列Bが違っても EQUAL
    config = DiffConfig(global_col_filter={0})
    result = run_diff(
        [["A", "X"]],
        [["A", "Y"]],
        config=config,
    )
    assert not result.has_differences


def t_col_filter_stored_in_sheet_diff():
    """SheetDiff に col_filter が格納されている"""
    config = DiffConfig(global_col_filter={0, 2})
    result = run_diff(
        [["A", "B", "C"]],
        [["A", "B", "C"]],
        config=config,
    )
    sd = result.sheet_diffs[0]
    assert sd.col_filter == {0, 2}


# ---------------------------------------------------------------------------
# 文字レベルdiff テスト (_render_cell_pair_diff)
# ---------------------------------------------------------------------------

def char_diff(old_val, new_val) -> tuple[str, str]:
    """CellData を作って _render_cell_pair_diff を呼ぶ薄いラッパー。"""
    return _render_cell_pair_diff(CellData(old_val), CellData(new_val))


def assert_marks(html_str: str, expected_texts: list[str], mark_class: str):
    """expected_texts のそれぞれが <mark class="..."> で囲まれているか検証する。"""
    for text in expected_texts:
        tag = f'<mark class="{mark_class}">{text}</mark>'
        assert tag in html_str, f'"{text}" が {mark_class} でマークされていない\n  actual: {html_str}'


def assert_raises(exc_type, fn):
    """fn() を実行し、exc_type が送出されることを検証する。"""
    try:
        fn()
    except exc_type:
        return
    except Exception as e:
        raise AssertionError(
            f"{exc_type.__name__} ではなく {type(e).__name__} が発生した: {e}"
        )
    raise AssertionError(f"{exc_type.__name__} が発生しなかった")


def assert_no_mark(html_str: str, texts: list[str]):
    """texts のいずれも <mark> で囲まれていないことを検証する。"""
    for text in texts:
        assert f'<mark' not in html_str.split(text)[0].rsplit('<mark', 1)[-1] \
               or text not in html_str, \
            f'"{text}" が意図せずマークされている\n  actual: {html_str}'


# --- 全角テキスト ---

def t_char_zenkaku_append():
    """全角: 末尾追加 "みかん" → "みかんジュース" """
    old_h, new_h = char_diff("みかん", "みかんジュース")
    assert "みかん" in old_h, "旧: 共通部分が消えた"
    assert_marks(new_h, ["ジュース"], "char-new")
    assert "char-del" not in old_h, "旧: 削除マークが出るはずない"


def t_char_zenkaku_prefix():
    """全角: 先頭追加 "東京" → "新東京" """
    old_h, new_h = char_diff("東京", "新東京")
    assert_marks(new_h, ["新"], "char-new")
    assert "東京" in new_h, "新: 共通部分が消えた"


def t_char_zenkaku_middle():
    """全角: 中間1文字変更 "東京都渋谷区" → "東京都新宿区" """
    old_h, new_h = char_diff("東京都渋谷区", "東京都新宿区")
    assert_marks(old_h, ["渋谷"], "char-del")
    assert_marks(new_h, ["新宿"], "char-new")
    assert "東京都" in old_h and "東京都" in new_h, "共通接頭辞が消えた"
    assert "区" in old_h and "区" in new_h, "共通接尾辞が消えた"


def t_char_zenkaku_full_replace():
    """全角: 完全置換 "りんご" → "バナナ" """
    old_h, new_h = char_diff("りんご", "バナナ")
    assert_marks(old_h, ["りんご"], "char-del")
    assert_marks(new_h, ["バナナ"], "char-new")


def t_char_zenkaku_trim():
    """全角: 末尾削除 "みかんジュース" → "みかん" """
    old_h, new_h = char_diff("みかんジュース", "みかん")
    assert_marks(old_h, ["ジュース"], "char-del")
    assert "char-new" not in new_h


# --- 半角英数 ---

def t_char_ascii_last():
    """半角: 末尾1文字変更 "ABC" → "ABD" """
    old_h, new_h = char_diff("ABC", "ABD")
    assert_marks(old_h, ["C"], "char-del")
    assert_marks(new_h, ["D"], "char-new")
    assert "AB" in old_h and "AB" in new_h


def t_char_ascii_middle():
    """半角: 中間変更 "商品コードA-001" → "商品コードB-001" """
    old_h, new_h = char_diff("商品コードA-001", "商品コードB-001")
    assert_marks(old_h, ["A"], "char-del")
    assert_marks(new_h, ["B"], "char-new")
    assert "商品コード" in old_h
    assert "-001" in old_h and "-001" in new_h


def t_char_version():
    """半角: バージョン番号 "v1.0.0" → "v1.2.0" """
    old_h, new_h = char_diff("v1.0.0", "v1.2.0")
    assert_marks(old_h, ["0"], "char-del")   # 2番目の0が変わる
    assert_marks(new_h, ["2"], "char-new")


# --- 数値 ---

def t_char_number_single():
    """数値: 1桁変更 60 → 70 """
    old_h, new_h = char_diff(60, 70)
    assert_marks(old_h, ["6"], "char-del")
    assert_marks(new_h, ["7"], "char-new")
    assert "0" in old_h and "0" in new_h


def t_char_number_middle():
    """数値: 4桁→4桁 1200 → 1400 """
    old_h, new_h = char_diff(1200, 1400)
    assert_marks(old_h, ["2"], "char-del")
    assert_marks(new_h, ["4"], "char-new")
    assert "1" in old_h and "00" in old_h


def t_char_number_digit_increase():
    """数値: 桁数増加 999 → 1000 """
    old_h, new_h = char_diff(999, 1000)
    # 全文字が変わる or 部分的にマークされる（どちらでも可）
    # 重要なのは old に char-del、new に char-new が含まれること
    assert "char-del" in old_h
    assert "char-new" in new_h


def t_char_float():
    """数値: 小数 3.14 → 3.15 """
    old_h, new_h = char_diff(3.14, 3.15)
    assert_marks(old_h, ["4"], "char-del")
    assert_marks(new_h, ["5"], "char-new")
    assert "3.1" in old_h and "3.1" in new_h


# --- 特殊ケース ---

def t_char_none_to_value():
    """空→値: None → "新規追加" """
    old_h, new_h = char_diff(None, "新規追加")
    assert old_h == "", f"旧: None は空文字列になるはず。actual: {old_h}"
    assert_marks(new_h, ["新規追加"], "char-new")


def t_char_value_to_none():
    """値→空: "削除予定" → None """
    old_h, new_h = char_diff("削除予定", None)
    assert_marks(old_h, ["削除予定"], "char-del")
    assert new_h == "", f"新: None は空文字列になるはず。actual: {new_h}"


def t_char_same_value():
    """同一値: 差分なし → mark タグなし """
    old_h, new_h = char_diff("変更なし", "変更なし")
    assert "char-del" not in old_h, "同一値なのに char-del が出た"
    assert "char-new" not in new_h, "同一値なのに char-new が出た"
    assert "変更なし" in old_h and "変更なし" in new_h


def t_char_zenkaku_numbers():
    """全角数字: "１２３" → "１２４" """
    old_h, new_h = char_diff("１２３", "１２４")
    assert_marks(old_h, ["３"], "char-del")
    assert_marks(new_h, ["４"], "char-new")
    assert "１２" in old_h and "１２" in new_h


def t_char_mixed_date():
    """混在: "2024年1月期" → "2024年3月期" """
    old_h, new_h = char_diff("2024年1月期", "2024年3月期")
    assert_marks(old_h, ["1"], "char-del")
    assert_marks(new_h, ["3"], "char-new")
    assert "2024年" in old_h and "月期" in old_h


# ---------------------------------------------------------------------------
# サブキー照合: DiffConfig バリデーション
# ---------------------------------------------------------------------------

def t_subkey_overlap_with_key_cols_raises():
    """主キー列とサブキー列に同じ列を指定するとエラー"""
    assert_raises(
        ValueError,
        lambda: DiffConfig(diff_mode="key", key_cols=[1, 2], sub_key_cols=[2, 3]),
    )


def t_subkey_without_key_mode_raises():
    """diff_mode が key でないのに sub_key_cols を指定するとエラー"""
    assert_raises(
        ValueError,
        lambda: DiffConfig(diff_mode="lcs", key_cols=[1], sub_key_cols=[2]),
    )


def t_subkey_without_key_cols_raises():
    """key_cols が空なのに sub_key_cols だけ指定するとエラー"""
    assert_raises(
        ValueError,
        lambda: DiffConfig(diff_mode="key", key_cols=[], sub_key_cols=[1]),
    )


def t_load_config_reads_sub_key_cols():
    """JSON設定ファイル（--matchers）から sub_key_cols を読み込める（key_colsと同様）。"""
    import json
    import tempfile
    data = {
        "diff_mode": "key",
        "key_cols": "A",
        "sub_key_cols": "B",
        "matchers": [],
    }
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".json", delete=False, encoding="utf-8"
    ) as f:
        json.dump(data, f)
        path = f.name
    try:
        config = load_config(path)
        assert config.key_cols == [0], f"key_cols が {config.key_cols}"
        assert config.sub_key_cols == [1], f"sub_key_cols が {config.sub_key_cols}"
    finally:
        os.remove(path)


# ---------------------------------------------------------------------------
# サブキー照合: コアロジック（diff_files 経由）
# ---------------------------------------------------------------------------

def t_subkey_rescues_unmatched_pair():
    """主キーで一致しなかった行同士が、一意なサブキーで救済される。
    救済されたペアでは主キー列も差分比較対象になる。"""
    config = DiffConfig(diff_mode="key", key_cols=[0], sub_key_cols=[1])
    result = run_diff(
        [["K1", "S1", "X"], ["K2", "S2", "Y"]],
        [["K1", "S1", "X"], ["K3", "S2", "Z"]],
        config=config,
    )
    sd = result.sheet_diffs[0]

    # K1行は主キー一致でそのままEQUAL
    equal_rows = [rd for rd in sd.row_diffs if rd.tag == RowTag.EQUAL]
    assert len(equal_rows) == 1, f"EQUAL行が{len(equal_rows)}件（期待値: 1）"

    # K2/K3行はサブキー(S2)で救済されMODIFY。主キー列(col0)も差分に含まれる
    modified = [rd for rd in sd.row_diffs if rd.tag == RowTag.MODIFY]
    assert len(modified) == 1, f"MODIFY行が{len(modified)}件（期待値: 1）"
    rd = modified[0]
    assert rd.matched_by == "subkey", f"matched_by が {rd.matched_by}（期待値: 'subkey'）"
    changed_cols = {cd.col_idx for cd in rd.cell_diffs}
    assert 0 in changed_cols, "サブキー救済ペアで主キー列(col0)の差分が検出されなかった"
    assert 2 in changed_cols, "通常列(col2)の差分が検出されなかった"

    assert not any(rd.tag in (RowTag.DELETE, RowTag.INSERT) for rd in sd.row_diffs), \
        "サブキーで救済できるはずの行が削除/追加になった"


def t_subkey_ambiguous_match_falls_back():
    """サブキー値が旧側2件・新側1件など曖昧な場合は救済せず、
    そのままDELETE/INSERTになる（安全側フォールバック）。"""
    config = DiffConfig(diff_mode="key", key_cols=[0], sub_key_cols=[1])
    result = run_diff(
        [["K1", "S1", "A"], ["K2", "SX", "B"], ["K3", "SX", "C"]],
        [["K1", "S1", "A"], ["K9", "SX", "D"]],
        config=config,
    )
    sd = result.sheet_diffs[0]

    assert not any(rd.matched_by == "subkey" for rd in sd.row_diffs), \
        "曖昧一致なのにサブキー救済されてしまった"
    deleted_keys = {rd.old_row.cells[0].value for rd in sd.row_diffs if rd.tag == RowTag.DELETE}
    inserted_keys = {rd.new_row.cells[0].value for rd in sd.row_diffs if rd.tag == RowTag.INSERT}
    assert deleted_keys == {"K2", "K3"}, f"DELETE行のキーが {deleted_keys}"
    assert inserted_keys == {"K9"}, f"INSERT行のキーが {inserted_keys}"


def t_subkey_blank_never_matches():
    """サブキー値が旧・新とも空欄の行同士はサブキー一致とみなさない。"""
    config = DiffConfig(diff_mode="key", key_cols=[0], sub_key_cols=[1])
    result = run_diff(
        [["K1", "S1", "A"], ["K2", None, "B"]],
        [["K1", "S1", "A"], ["K9", None, "C"]],
        config=config,
    )
    sd = result.sheet_diffs[0]

    assert not any(rd.matched_by == "subkey" for rd in sd.row_diffs), \
        "空欄サブキー同士なのに救済されてしまった"
    tags = {rd.tag for rd in sd.row_diffs}
    assert RowTag.DELETE in tags and RowTag.INSERT in tags


def t_subkey_rescues_duplicate_key_overflow_row():
    """主キーが重複した行（2件目以降）も、サブキー救済の対象プールに含まれる。"""
    config = DiffConfig(diff_mode="key", key_cols=[0], sub_key_cols=[1])
    result = run_diff(
        [["K1", "SA", "first"], ["K1", "SB", "second"]],   # K1が重複、2件目はold_nullへ
        [["K1", "SA", "first"], ["K2", "SB", "newsecond"]],  # K2は主キーで対応先なし
        config=config,
    )
    sd = result.sheet_diffs[0]

    rescued = [rd for rd in sd.row_diffs if rd.matched_by == "subkey"]
    assert len(rescued) == 1, f"サブキー救済された行が{len(rescued)}件（期待値: 1）"
    assert rescued[0].old_row.cells[2].value == "second"
    assert rescued[0].new_row.cells[2].value == "newsecond"


# ---------------------------------------------------------------------------
# サブキー照合: HTML出力の色分け (_render_row_pair)
# ---------------------------------------------------------------------------

def t_subkey_preserves_old_row_order_when_rescued():
    """サブキー救済されたペアも、旧ファイルの行順の位置に表示される
    （SPEC.md 5-6: 表示順は旧ファイルの行順を基準とする、という既存の不変条件）。"""
    config = DiffConfig(diff_mode="key", key_cols=[0], sub_key_cols=[1])
    result = run_diff(
        [["K1", "SA", "a"], ["K2", "SB", "b"], ["K3", "SC", "c"]],
        [["K1", "SA", "a"], ["K3", "SC", "c"], ["K9", "SB", "d"]],
        config=config,
    )
    sd = result.sheet_diffs[0]
    keys = [
        (rd.old_row.cells[0].value if rd.old_row else None,
         rd.new_row.cells[0].value if rd.new_row else None)
        for rd in sd.row_diffs
    ]
    assert keys == [("K1", "K1"), ("K2", "K9"), ("K3", "K3")], (
        f"旧ファイルの行順（K1, K2/K9, K3）を保持していない: {keys}"
    )


def t_subkey_preserves_old_row_order_when_unrescued():
    """サブキーで救済されない（曖昧一致にもならない、単に候補がない）DELETE行も、
    旧ファイルの行順の位置に表示される。"""
    config = DiffConfig(diff_mode="key", key_cols=[0], sub_key_cols=[1])
    result = run_diff(
        [["K1", "SA"], ["K2", "SX"], ["K3", "SY"]],
        [["K1", "SA"], ["K3", "SY"]],
        config=config,
    )
    sd = result.sheet_diffs[0]
    keys = [
        (rd.old_row.cells[0].value if rd.old_row else None,
         rd.new_row.cells[0].value if rd.new_row else None)
        for rd in sd.row_diffs
    ]
    assert keys == [("K1", "K1"), ("K2", None), ("K3", "K3")], (
        f"旧ファイルの行順（K1, K2(削除), K3）を保持していない: {keys}"
    )


def t_subkey_old_row_index_stays_ascending():
    """不変条件: old_row を持つ行は、シート内のどこであっても
    old_row_idx が単調増加でなければならない（表示順が旧ファイルの
    行順から外れていないことの直接的な検証）。
    サブキーで救済されるかどうかに関わらず、ランダムなキー配置で検証する。"""
    config = DiffConfig(diff_mode="key", key_cols=[0], sub_key_cols=[1])
    result = run_diff(
        [
            ["K1", "SA", "a"], ["K2", "SB", "b"], ["K3", "SC", "c"],
            ["K4", "SD", "d"], ["K5", None, "e"], ["K5", "SF", "f"],  # K5重複
        ],
        [
            ["K1", "SA", "a"], ["K3", "SX", "c2"],  # K3は主キー一致(MODIFY)
            ["K9", "SB", "b2"],                      # K2をサブキー救済
            ["K10", "SD", "d2"],                     # K4をサブキー救済
            ["K11", "SF", "f2"],                     # 重複キー行2件目を救済
        ],
        config=config,
    )
    sd = result.sheet_diffs[0]
    old_indices = [rd.old_row.row_idx for rd in sd.row_diffs if rd.old_row is not None]
    assert old_indices == sorted(old_indices), (
        f"old_row_idx が単調増加でない（表示順が旧ファイルの行順から外れている）: {old_indices}"
    )


def t_subkey_matched_row_highlights_key_col():
    """サブキー救済ペアでは、旧主キー列のセルに専用クラス cell-modified-subkey が付く。
    サブキーと無関係な通常の変更列は従来通り cell-modified のまま。"""
    old_row = make_sheet("S", [["K2", "B"]]).rows[0]
    new_row = make_sheet("S", [["K9", "C"]]).rows[0]
    rd = RowDiff(
        RowTag.MODIFY,
        old_row,
        new_row,
        cell_diffs=[
            CellDiff(col_idx=0, old_cell=old_row.cells[0], new_cell=new_row.cells[0]),
            CellDiff(col_idx=1, old_cell=old_row.cells[1], new_cell=new_row.cells[1]),
        ],
        matched_by="subkey",
    )
    left_tr, right_tr = _render_row_pair(rd, max_cols=2, col_filter=None, row_idx=1, key_cols=[0])

    assert 'class="cell-modified-subkey"' in left_tr, f"旧側に専用クラスが付かなかった: {left_tr}"
    assert 'class="cell-modified-subkey"' in right_tr, f"新側に専用クラスが付かなかった: {right_tr}"
    # col1（通常の変更列）は通常のcell-modifiedのまま
    assert left_tr.count('class="cell-modified"') == 1, f"通常列の扱いが変わった: {left_tr}"
    assert right_tr.count('class="cell-modified"') == 1, f"通常列の扱いが変わった: {right_tr}"


def t_normal_modify_row_has_no_subkey_class():
    """主キー一致による通常のMODIFY行には cell-modified-subkey は付かない。"""
    old_row = make_sheet("S", [["A", "old"]]).rows[0]
    new_row = make_sheet("S", [["A", "new"]]).rows[0]
    rd = RowDiff(
        RowTag.MODIFY,
        old_row,
        new_row,
        cell_diffs=[CellDiff(col_idx=1, old_cell=old_row.cells[1], new_cell=new_row.cells[1])],
    )
    left_tr, right_tr = _render_row_pair(rd, max_cols=2, col_filter=None, row_idx=1, key_cols=[0])

    assert "cell-modified-subkey" not in left_tr
    assert "cell-modified-subkey" not in right_tr


# ---------------------------------------------------------------------------
# メイン
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    print("=" * 50)
    print("excel-diff ユニットテスト")
    print("=" * 50)

    print("--- 列フィルタ ---")
    _run_test("parse_col_spec: 単一列",              t_parse_col_spec_single)
    _run_test("parse_col_spec: 連続範囲",            t_parse_col_spec_range)
    _run_test("parse_col_spec: 飛び地",              t_parse_col_spec_noncontiguous)
    _run_test("parse_col_spec: 混在",                t_parse_col_spec_mixed)
    _run_test("列フィルタ: 除外列は差分なし",         t_col_filter_excludes_diff)
    _run_test("列フィルタ: 対象列の差分を検出",       t_col_filter_detects_included_diff)
    _run_test("列フィルタ: LCSに影響しない",          t_col_filter_row_lcs)
    _run_test("列フィルタ: SheetDiffに格納",          t_col_filter_stored_in_sheet_diff)

    print()
    print("--- 差分エンジン ---")
    _run_test("差分なし",                            t_no_diff)
    _run_test("セル変更",                            t_cell_change)
    _run_test("行挿入（中間）",                      t_row_insert)
    _run_test("行削除（中間）",                      t_row_delete)
    _run_test("挿入と削除が混在",                    t_insert_and_delete)
    _run_test("シート追加",                          t_sheet_added)
    _run_test("シート削除",                          t_sheet_deleted)
    _run_test("マッチャー: 一致→差分なし",           t_matcher_equal)
    _run_test("マッチャー: マッピング外→差分あり",   t_matcher_not_in_mapping)
    _run_test("マッチャー: 一部列のみ一致",           t_matcher_partial)
    _run_test("マッチャー: シートスコープ",           t_matcher_sheet_scope)
    _run_test("equiv: 「-」と空白を同一視",   t_equivalence_hyphen_blank)
    _run_test("equiv: 無変更を誤検出しない", t_equivalence_unchanged_not_flagged)
    _run_test("equiv: グループ外は差分あり", t_equivalence_outside_group_still_diffs)
    _run_test("equiv: 全列ワイルドカード",   t_equivalence_all_columns_wildcard)
    _run_test("equiv: _x000D_の表記ゆれを正規化",     t_equivalence_normalizes_x000d)
    _run_test("mapping: CRLF表記ゆれを正規化",         t_mapping_matcher_normalizes_x000d)
    _run_test("LCS行キー: 未適用列の_x000D_を正規化",  t_lcs_row_key_normalizes_x000d_unmatched_column)
    _run_test("行番号の正確性",                      t_row_numbers)

    print()
    print("--- サブキー照合 ---")
    _run_test("サブキー: 主キーと重複でエラー",       t_subkey_overlap_with_key_cols_raises)
    _run_test("サブキー: LCSモードで指定するとエラー", t_subkey_without_key_mode_raises)
    _run_test("サブキー: 主キー未指定でエラー",         t_subkey_without_key_cols_raises)
    _run_test("サブキー: JSON設定から読み込める",       t_load_config_reads_sub_key_cols)
    _run_test("サブキー: 一意な一致で救済",             t_subkey_rescues_unmatched_pair)
    _run_test("サブキー: 曖昧一致は安全側フォールバック", t_subkey_ambiguous_match_falls_back)
    _run_test("サブキー: 空欄同士は一致させない",         t_subkey_blank_never_matches)
    _run_test("サブキー: 重複キー行の残りも救済対象",     t_subkey_rescues_duplicate_key_overflow_row)
    _run_test("サブキー: 救済ペアも旧行順を保持",         t_subkey_preserves_old_row_order_when_rescued)
    _run_test("サブキー: 未救済DELETEも旧行順を保持",     t_subkey_preserves_old_row_order_when_unrescued)
    _run_test("サブキー: old_row_idxが常に単調増加",      t_subkey_old_row_index_stays_ascending)
    _run_test("サブキー: 主キー列に専用色クラス",         t_subkey_matched_row_highlights_key_col)
    _run_test("サブキー: 通常MODIFYには専用色が付かない", t_normal_modify_row_has_no_subkey_class)

    print()
    print("--- 文字レベルdiff ---")
    _run_test("全角: 末尾追加",                      t_char_zenkaku_append)
    _run_test("全角: 先頭追加",                      t_char_zenkaku_prefix)
    _run_test("全角: 中間変更",                      t_char_zenkaku_middle)
    _run_test("全角: 完全置換",                      t_char_zenkaku_full_replace)
    _run_test("全角: 末尾削除",                      t_char_zenkaku_trim)
    _run_test("半角: 末尾1文字変更",                 t_char_ascii_last)
    _run_test("半角: 中間変更",                      t_char_ascii_middle)
    _run_test("半角: バージョン番号",                 t_char_version)
    _run_test("数値: 1桁変更",                       t_char_number_single)
    _run_test("数値: 4桁変更",                       t_char_number_middle)
    _run_test("数値: 桁数増加",                      t_char_number_digit_increase)
    _run_test("数値: 小数",                          t_char_float)
    _run_test("特殊: None → 値",                    t_char_none_to_value)
    _run_test("特殊: 値 → None",                    t_char_value_to_none)
    _run_test("特殊: 同一値（markなし）",             t_char_same_value)
    _run_test("全角数字変更",                        t_char_zenkaku_numbers)
    _run_test("混在: 日付文字列",                     t_char_mixed_date)

    print("=" * 50)
    print(f"結果: {len(PASS)} PASS / {len(FAIL)} FAIL")
    if FAIL:
        sys.exit(1)
