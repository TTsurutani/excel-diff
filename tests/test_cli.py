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

    print("=" * 50)
    print(f"結果: {len(PASS)} PASS / {len(FAIL)} FAIL")
    if FAIL:
        sys.exit(1)
