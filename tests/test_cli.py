"""
CLI引数 → DiffConfig 変換のユニットテスト。

実行:
  python tests/test_cli.py
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import argparse

from excel_diff.__main__ import _build_parser, _build_config, _apply_profile, _profiles_dir


# ---------------------------------------------------------------------------
# ヘルパー
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


def build_config(argv: list[str]):
    args = _build_parser().parse_args(argv)
    return _build_config(args)


def assert_raises(exc_type, fn):
    try:
        fn()
    except exc_type:
        return
    except Exception as e:
        raise AssertionError(
            f"{exc_type.__name__} ではなく {type(e).__name__} が発生した: {e}"
        )
    raise AssertionError(f"{exc_type.__name__} が発生しなかった")


# ---------------------------------------------------------------------------
# テストケース
# ---------------------------------------------------------------------------

def t_sub_key_cols_parsed_with_key_cols():
    """--key-cols と --sub-key-cols を併用すると、両方が DiffConfig に反映される"""
    config = build_config(["--key-cols", "A", "--sub-key-cols", "B"])
    assert config.diff_mode == "key"
    assert config.key_cols == [0]
    assert config.sub_key_cols == [1]


def t_sub_key_cols_without_key_cols_exits():
    """--key-cols なしで --sub-key-cols だけ指定するとエラー終了する"""
    assert_raises(SystemExit, lambda: build_config(["--sub-key-cols", "B"]))


def t_sub_key_cols_overlap_with_key_cols_exits():
    """--key-cols と --sub-key-cols に同じ列を指定するとエラー終了する"""
    assert_raises(
        SystemExit,
        lambda: build_config(["--key-cols", "A", "--sub-key-cols", "A"]),
    )


def t_profile_dir_diff_applies_sub_key_cols():
    """--profile（dir_diffタブ相当）の sub_key_cols が args に反映される
    （p-pipelineが --dir ... --profile <name> で使う経路）。"""
    profile_name = "__test_subkey_profile__"
    profile_path = _profiles_dir() / f"{profile_name}.toml"
    profile_path.write_text(
        '[dir_diff]\ndiff_mode = "key"\nkey_cols = "A"\nsub_key_cols = "B"\n',
        encoding="utf-8",
    )
    try:
        args = argparse.Namespace(
            profile=profile_name,
            dir=["old_dir", "new_dir"],
            split=None,
            old_file=None,
            new_file=None,
            output_dir=None, sheet_old=None, sheet_new=None,
            include_cols=None, matchers=None, strikethrough=False,
            open=True, diff_mode=None, key_cols=None, sub_key_cols=None,
        )
        _apply_profile(args)
        assert args.key_cols == "A", f"key_cols が {args.key_cols!r}"
        assert args.sub_key_cols == "B", f"sub_key_cols が {args.sub_key_cols!r}（反映されていない）"
    finally:
        profile_path.unlink(missing_ok=True)


def t_matchers_json_with_invalid_subkey_config_exits_cleanly():
    """--matchers のJSONに不正な sub_key_cols 設定があった場合、生の
    ValueError で落ちず、他のバリデーションと同様にエラー終了する。"""
    import json
    import tempfile
    data = {"diff_mode": "lcs", "sub_key_cols": "B", "matchers": []}
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".json", delete=False, encoding="utf-8"
    ) as f:
        json.dump(data, f)
        path = f.name
    try:
        assert_raises(SystemExit, lambda: build_config(["--matchers", path]))
    finally:
        os.remove(path)


def t_sub_key_cols_with_lcs_mode_exits():
    """--diff-mode lcs と --sub-key-cols を併用するとエラー終了する"""
    assert_raises(
        SystemExit,
        lambda: build_config(
            ["--key-cols", "A", "--diff-mode", "lcs", "--sub-key-cols", "B"]
        ),
    )


# ---------------------------------------------------------------------------
# --excel-summary / --header-row
# ---------------------------------------------------------------------------

def t_excel_summary_without_key_mode_exits():
    """--excel-summary は --diff-mode key（--key-cols指定）が前提。lcsモードだとエラー終了する"""
    assert_raises(SystemExit, lambda: build_config(["--excel-summary"]))


def t_excel_summary_with_key_mode_succeeds():
    """--key-cols併用（keyモード）なら --excel-summary はエラーにならない"""
    config = build_config(["--key-cols", "A", "--excel-summary"])
    assert config.diff_mode == "key"


def t_excel_summary_default_is_none():
    """--excel-summary 未指定時は args.excel_summary が None のまま"""
    args = _build_parser().parse_args(["--key-cols", "A"])
    assert args.excel_summary is None


def t_excel_summary_bare_flag_yields_empty_string():
    """--excel-summary をパス省略で指定すると const の空文字列になる"""
    args = _build_parser().parse_args(["--key-cols", "A", "--excel-summary"])
    assert args.excel_summary == ""


def t_excel_summary_with_path():
    """--excel-summary にパスを指定するとそのまま args に反映される"""
    args = _build_parser().parse_args(["--key-cols", "A", "--excel-summary", "out.xlsx"])
    assert args.excel_summary == "out.xlsx"


def t_header_row_default_is_one():
    """--header-row 未指定時のデフォルトは1"""
    args = _build_parser().parse_args([])
    assert args.header_row == 1


def t_header_row_explicit_value():
    """--header-row を明示指定すると反映される（0=ヘッダーなし扱いも許可）"""
    args = _build_parser().parse_args(["--header-row", "0"])
    assert args.header_row == 0
    args = _build_parser().parse_args(["--header-row", "5"])
    assert args.header_row == 5


def t_profile_dir_diff_applies_excel_summary_and_header_row():
    """--profile（dir_diffタブ相当）の excel_summary/header_row が args に反映される"""
    profile_name = "__test_excel_summary_profile__"
    profile_path = _profiles_dir() / f"{profile_name}.toml"
    profile_path.write_text(
        '[dir_diff]\ndiff_mode = "key"\nkey_cols = "A"\n'
        'excel_summary = true\nheader_row = 5\n',
        encoding="utf-8",
    )
    try:
        args = argparse.Namespace(
            profile=profile_name,
            dir=["old_dir", "new_dir"],
            split=None,
            old_file=None,
            new_file=None,
            output_dir=None, sheet_old=None, sheet_new=None,
            include_cols=None, matchers=None, strikethrough=False,
            open=True, diff_mode=None, key_cols=None, sub_key_cols=None,
            excel_summary=None, header_row=1,
        )
        _apply_profile(args)
        assert args.excel_summary is True, f"excel_summary が {args.excel_summary!r}"
        assert args.header_row == 5, f"header_row が {args.header_row!r}"
    finally:
        profile_path.unlink(missing_ok=True)


# ---------------------------------------------------------------------------
# メイン
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    print("=" * 50)
    print("excel-diff CLIユニットテスト")
    print("=" * 50)

    _run_test("サブキー: --key-colsと併用で反映",       t_sub_key_cols_parsed_with_key_cols)
    _run_test("サブキー: --key-cols無しでエラー終了",   t_sub_key_cols_without_key_cols_exits)
    _run_test("サブキー: --key-colsと重複でエラー終了", t_sub_key_cols_overlap_with_key_cols_exits)
    _run_test("サブキー: lcsモードでエラー終了",        t_sub_key_cols_with_lcs_mode_exits)
    _run_test("サブキー: --profile(dir_diff)で反映",    t_profile_dir_diff_applies_sub_key_cols)
    _run_test("サブキー: matchers JSON不正設定でエラー終了", t_matchers_json_with_invalid_subkey_config_exits_cleanly)

    print()
    _run_test("集約Excel: keyモード無しでエラー終了",     t_excel_summary_without_key_mode_exits)
    _run_test("集約Excel: keyモード併用で成功",           t_excel_summary_with_key_mode_succeeds)
    _run_test("集約Excel: 未指定はNone",                 t_excel_summary_default_is_none)
    _run_test("集約Excel: パス省略でconst空文字列",       t_excel_summary_bare_flag_yields_empty_string)
    _run_test("集約Excel: パス指定が反映される",           t_excel_summary_with_path)
    _run_test("ヘッダー行: デフォルトは1",                t_header_row_default_is_one)
    _run_test("ヘッダー行: 明示指定が反映される",          t_header_row_explicit_value)
    _run_test(
        "集約Excel: --profile(dir_diff)で反映",
        t_profile_dir_diff_applies_excel_summary_and_header_row,
    )

    print("=" * 50)
    print(f"結果: {len(PASS)} PASS / {len(FAIL)} FAIL")
    if FAIL:
        sys.exit(1)
