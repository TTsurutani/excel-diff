"""gui_settings.json の読み書き。"""
import json
import tomllib
from datetime import datetime
from pathlib import Path
from typing import Any, Optional


_DEFAULT: dict[str, Any] = {
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
        "sub_key_cols": "",
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
        "sub_key_cols": "",
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
        "sub_key_cols":  "",
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


def _migrate_legacy_profiles(loaded: dict) -> None:
    """gui_settings.json > profiles[]（旧形式）を profiles/<name>.toml に変換する。

    変換後、gui_settings.json 側の profiles キーは空にする（呼び出し元で
    書き戻し・保存される）。
    """
    legacy = loaded.pop("profiles", None)
    if not legacy:
        return
    for p in legacy:
        name = p.get("name") or p.get("id")
        if not name:
            continue
        path = _profile_path(name)
        if path.exists():
            continue
        _write_profile_file(
            path,
            note=p.get("note", ""),
            created_at=p.get("created_at", ""),
            snapshot=p.get("snapshot", {}),
        )


def _ensure_loaded() -> None:
    global _data
    if _data:
        return
    import copy
    _data = copy.deepcopy(_DEFAULT)
    if _settings_path.exists():
        try:
            loaded = json.loads(_settings_path.read_text(encoding="utf-8"))
            _migrate_legacy_profiles(loaded)
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


# プロファイル名に使えない文字（Windowsのファイル名制約）。
_INVALID_NAME_CHARS = '\\/:*?"<>|'


def _validate_profile_name(name: str) -> None:
    if not name or not name.strip():
        raise ValueError("プロファイル名を入力してください")
    if any(c in name for c in _INVALID_NAME_CHARS):
        raise ValueError(
            f"プロファイル名に次の文字は使えません: {_INVALID_NAME_CHARS}"
        )


def _profiles_dir() -> Path:
    d = _data_dir() / "profiles"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _profile_path(name: str) -> Path:
    return _profiles_dir() / f"{name}.toml"


def _toml_quote(v: str) -> str:
    return '"' + v.replace("\\", "\\\\").replace('"', '\\"') + '"'


def _toml_scalar(v: Any) -> str:
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, (int, float)):
        return str(v)
    return _toml_quote(str(v))


def _write_profile_file(
    path: Path, note: str, created_at: str, snapshot: dict
) -> None:
    """プロファイルを name.toml として書き出す（トップレベル: note/created_at、
    タブごとに [tab_name] テーブル）。"""
    lines = [
        f"note       = {_toml_quote(note)}",
        f"created_at = {_toml_quote(created_at)}",
    ]
    for tab, vals in snapshot.items():
        lines.append("")
        lines.append(f"[{tab}]")
        for k, v in vals.items():
            lines.append(f"{k} = {_toml_scalar(v)}")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def _read_profile_file(path: Path) -> dict:
    with open(path, "rb") as f:
        raw = tomllib.load(f)
    snapshot = {k: v for k, v in raw.items() if isinstance(v, dict)}
    return {
        "id":         path.stem,
        "name":       path.stem,
        "note":       raw.get("note", ""),
        "created_at": raw.get("created_at", ""),
        "snapshot":   snapshot,
    }


def get_profiles() -> list[dict]:
    """profiles/*.toml を名前順に読み込んで返す。"""
    _ensure_loaded()  # 旧形式（gui_settings.json > profiles[]）からの移行を保証する
    profiles = []
    for path in sorted(_profiles_dir().glob("*.toml")):
        try:
            profiles.append(_read_profile_file(path))
        except Exception:
            continue
    return profiles


def save_profile(
    name: str, note: str, snapshot: dict, *, overwrite: bool = False
) -> str:
    """名前・メモ・スナップショットを持つプロファイルを profiles/<name>.toml
    として保存し、プロファイル名（id を兼ねる）を返す。
    同名のプロファイルが既に存在する場合、overwrite=False なら
    FileExistsError を送出する（呼び出し側で上書き確認すること）。
    """
    _validate_profile_name(name)
    path = _profile_path(name)
    if path.exists() and not overwrite:
        raise FileExistsError(name)
    now = datetime.now()
    _write_profile_file(path, note, now.isoformat(timespec="seconds"), snapshot)
    return name


def update_profile(
    profile_id: str,
    name: Optional[str] = None,
    note: Optional[str] = None,
) -> None:
    """既存プロファイルの名前・メモを更新する（名前変更時はファイル名も変更）。"""
    old_path = _profile_path(profile_id)
    if not old_path.exists():
        return
    existing = _read_profile_file(old_path)
    new_name = name if name else profile_id
    new_note = note if note is not None else existing["note"]
    if new_name != profile_id:
        _validate_profile_name(new_name)
        new_path = _profile_path(new_name)
        if new_path.exists():
            raise FileExistsError(new_name)
    else:
        new_path = old_path
    _write_profile_file(
        new_path, new_note, existing["created_at"], existing["snapshot"]
    )
    if new_path != old_path:
        old_path.unlink()


def delete_profile(profile_id: str) -> None:
    """指定名（id）のプロファイルを削除する。"""
    path = _profile_path(profile_id)
    if path.exists():
        path.unlink()


def build_identity_snapshot(snapshot: dict) -> dict:
    """パスフィールドを除いた同一性チェック用スナップショットを返す。"""
    result = {}
    for tab, vals in snapshot.items():
        exclude = _PATH_KEYS.get(tab, set())
        result[tab] = {k: v for k, v in vals.items() if k not in exclude}
    return result


def find_matching_profile(snapshot: dict) -> Optional[dict]:
    """現在のスナップショット（パス除外後）と一致するプロファイルを返す。"""
    identity = build_identity_snapshot(snapshot)
    for p in get_profiles():
        if build_identity_snapshot(p.get("snapshot", {})) == identity:
            return p
    return None
