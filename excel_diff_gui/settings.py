"""gui_settings.json の読み書き。"""
import json
from pathlib import Path
from typing import Any


_DEFAULT: dict[str, Any] = {
    "file_diff": {
        "old_file": "",
        "new_file": "",
        "output": "",
        "sheet": "",
        "include_cols": "",
        "matchers": "",
        "strikethrough": False,
        "open_browser": True,
        "diff_mode": "lcs",
        "key_cols": "",
    },
    "dir_diff": {
        "output_dir": "",
        "sheet": "",
        "include_cols": "",
        "matchers": "",
        "strikethrough": False,
        "open_browser": True,
        "diff_mode": "lcs",
        "key_cols": "",
    },
    "pair_build": {
        "old_dir": "",
        "new_dir": "",
        "pairing": "exact",
        "pairs_file": "",
        "pattern_id": "",
    },
    "split": {
        "book_file": "",
        "prefix": "",
        "suffix": "",
        "name_regex": "",
        "output_dir": "",
    },
    "split_presets": [
        {"name": "括弧前の名前（例: 売上（Sales）→売上）", "regex": "^([^（]+)"},
        {"name": "番号プレフィックス除去（例: 01_概要→概要）", "regex": r"^\d+_(.+)"},
        {"name": "日付サフィックス除去（例: report_20240101→report）", "regex": r"^(.+?)_\d{8}$"},
        {"name": "バージョン番号除去（例: 報告書_v2→報告書）", "regex": r"^(.+?)_v\d+$"},
    ],
}

def _data_dir() -> Path:
    """設定・パターンファイルの保存先ディレクトリ。
    EXE実行時は EXE と同じフォルダ、スクリプト実行時はプロジェクトルート。
    """
    import sys
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent.parent


def patterns_file() -> str:
    """patterns.json の絶対パスを返す。"""
    return str(_data_dir() / "patterns.json")


_settings_path = _data_dir() / "gui_settings.json"
_data: dict[str, Any] = {}


def _ensure_loaded() -> None:
    global _data
    if _data:
        return
    import copy
    _data = copy.deepcopy(_DEFAULT)
    if _settings_path.exists():
        try:
            loaded = json.loads(_settings_path.read_text(encoding="utf-8"))
            for tab, vals in loaded.items():
                if tab in _data:
                    if isinstance(vals, dict) and isinstance(_data[tab], dict):
                        _data[tab].update(vals)
                    else:
                        _data[tab] = vals
        except Exception:
            pass


def get(tab: str, key: str, default: Any = None) -> Any:
    _ensure_loaded()
    return _data.get(tab, {}).get(key, default)


def set_tab(tab: str, values: dict[str, Any]) -> None:
    _ensure_loaded()
    _data[tab] = values


def save() -> None:
    _ensure_loaded()
    try:
        _settings_path.write_text(
            json.dumps(_data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception:
        pass


def data(tab: str) -> dict[str, Any]:
    _ensure_loaded()
    return _data.get(tab, {})


def get_split_presets() -> list:
    _ensure_loaded()
    return _data.get("split_presets", [])


def set_split_presets(presets: list) -> None:
    _ensure_loaded()
    _data["split_presets"] = presets
