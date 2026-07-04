"""gui_settings.json の読み書き。"""
import json
from datetime import datetime
from pathlib import Path
from typing import Any, Optional


_DEFAULT: dict[str, Any] = {
    "profiles": [],
    "file_diff": {
        "old_file": "",
        "new_file": "",
        "output": "",
        "sheet_old": "",
        "sheet_new": "",
        "include_cols": "",
        "matchers": "",
        "strikethrough": False,
        "open_browser": True,
        "diff_mode": "lcs",
        "key_cols": "",
    },
    "dir_diff": {
        "output_dir": "",
        "sheet_old": "",
        "sheet_new": "",
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
    "sheet_compare": {
        "old_file":      "",
        "new_file":      "",
        "old_sheet":     "",
        "new_sheet":     "",
        "output_dir":    "",
        "include_cols":  "",
        "strikethrough": False,
        "open_browser":  True,
        "diff_mode":     "lcs",
        "key_cols":      "",
    },
    "split_presets": [
        {"name": "括弧前の名前（例: 売上（Sales）→売上）", "regex": "^([^(（]+)"},
        {"name": "番号プレフィックス除去（例: 01_概要→概要）", "regex": r"^\d+_(.+)"},
        {"name": "日付サフィックス除去（例: report_20240101→report）", "regex": r"^(.+?)_\d{8}$"},
        {"name": "バージョン番号除去（例: 報告書_v2→報告書）", "regex": r"^(.+?)_v\d+$"},
    ],
}


def _data_dir() -> Path:
    """設定・パターンファイルの保存先ディレクトリ。

    EXE実行時（frozen）は「exeの一つ上のフォルダ」（dist/の親＝プロジェクト
    ルート）に既存のソースツリー（excel_diff パッケージ）があればそちらを
    使う。dist/ を直接の保存先にすると dist/ の再ビルドでデータが失われる
    ため、Preprocessing-Tools と同様プロジェクトルート側を優先する。
    該当しない場合（exe単体を持ち出した場合等）は exe と同じフォルダを使う。
    スクリプト実行時はプロジェクトルート。
    """
    import sys
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).parent
        project_root = exe_dir.parent
        if (project_root / "excel_diff").is_dir():
            return project_root
        return exe_dir
    return Path(__file__).parent.parent


def patterns_file() -> str:
    """patterns.json の絶対パスを返す。"""
    return str(_data_dir() / "patterns.json")


_settings_path = _data_dir() / "gui_settings.json"
_data: dict[str, Any] = {}


def _migrate(data: dict) -> None:
    """既知の不具合データを自動修正するマイグレーション処理。"""
    # 旧プリセット: 全角（のみ → 半角・全角両対応に修正
    for p in data.get("split_presets", []):
        if p.get("regex") == "^([^（]+)":
            p["regex"] = "^([^(（]+)"
    # sheet → sheet_old / sheet_new への移行
    for tab in ("file_diff", "dir_diff"):
        if tab in data and "sheet" in data[tab]:
            old_val = data[tab].pop("sheet")
            data[tab].setdefault("sheet_old", old_val)
            data[tab].setdefault("sheet_new", "")
    # profiles キーが無い旧バージョンのデータに補完
    data.setdefault("profiles", [])


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
    _migrate(_data)


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


# ── プロファイル（設定セット）管理 ─────────────────────────────────────────

# 同一性チェック時に除外するパス系フィールド（FileSelectRow を使うフィールド）。
# 新しいパスフィールドをタブに追加したら、ここにも追記すること。
_PATH_KEYS: dict[str, set[str]] = {
    "dir_diff":      {"output_dir", "matchers"},
    "pair_build":    {"old_dir", "new_dir", "pairs_file"},
    "file_diff":     {"old_file", "new_file", "output", "matchers"},
    "split":         {"book_file", "output_dir"},
    "sheet_compare": {"old_file", "new_file", "output_dir"},
}


def get_profiles() -> list[dict]:
    _ensure_loaded()
    return _data.get("profiles", [])


def save_profile(name: str, note: str, snapshot: dict) -> str:
    """名前・メモ・スナップショットを持つプロファイルを保存し id を返す。"""
    _ensure_loaded()
    now = datetime.now()
    profile_id = now.strftime("%Y%m%d_%H%M%S")
    _data.setdefault("profiles", []).append({
        "id":         profile_id,
        "name":       name,
        "note":       note,
        "created_at": now.isoformat(timespec="seconds"),
        "snapshot":   snapshot,
    })
    return profile_id


def update_profile(
    profile_id: str,
    name: Optional[str] = None,
    note: Optional[str] = None,
) -> None:
    """既存プロファイルの名前・メモを更新する。"""
    _ensure_loaded()
    for p in _data.get("profiles", []):
        if p["id"] == profile_id:
            if name is not None:
                p["name"] = name
            if note is not None:
                p["note"] = note
            return


def delete_profile(profile_id: str) -> None:
    """指定 id のプロファイルを削除する。"""
    _ensure_loaded()
    _data["profiles"] = [
        p for p in _data.get("profiles", []) if p["id"] != profile_id
    ]


def build_identity_snapshot(snapshot: dict) -> dict:
    """パスフィールドを除いた同一性チェック用スナップショットを返す。"""
    result = {}
    for tab, vals in snapshot.items():
        exclude = _PATH_KEYS.get(tab, set())
        result[tab] = {k: v for k, v in vals.items() if k not in exclude}
    return result


def find_matching_profile(snapshot: dict) -> Optional[dict]:
    """現在のスナップショット（パス除外後）と一致するプロファイルを返す。"""
    _ensure_loaded()
    identity = build_identity_snapshot(snapshot)
    for p in _data.get("profiles", []):
        if build_identity_snapshot(p.get("snapshot", {})) == identity:
            return p
    return None
