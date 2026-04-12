"""
ファイルペアリングモジュール。

--discover : 類似度ベースでペア候補を探索
generate_regex : 確認済みペアから正規表現を自動生成
validate_regex : 正規表現がペアを正しく再現できるか検証
apply_pattern  : 正規表現を使ってフォルダ内ファイルをペアリング
"""
from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Optional


# ---------------------------------------------------------------------------
# データクラス
# ---------------------------------------------------------------------------

@dataclass
class FilePair:
    old_name: Optional[str]
    new_name: Optional[str]
    score: float
    matched_by: str  # 'exact' | 'pattern' | 'auto' | 'unmatched_old' | 'unmatched_new'


@dataclass
class ValidationError:
    kind: str    # 'no_match' | 'key_mismatch' | 'key_collision'
    details: str


# ---------------------------------------------------------------------------
# ユーティリティ
# ---------------------------------------------------------------------------

def _xlsx_files(directory: str) -> list[str]:
    return sorted(
        f for f in os.listdir(directory)
        if f.lower().endswith(".xlsx") and not f.startswith("~$")
    )


# ---------------------------------------------------------------------------
# discover: 類似度ベースのペア候補探索
# ---------------------------------------------------------------------------

def discover_pairs(
    old_dir: str,
    new_dir: str,
    threshold: float = 0.6,
) -> list[FilePair]:
    """
    旧・新フォルダのファイルを類似度スコアでペアリングする。
    完全一致を優先し、残りをSequenceMatcherスコアで貪欲マッチ。
    """
    old_files = _xlsx_files(old_dir)
    new_files = _xlsx_files(new_dir)
    new_set = set(new_files)

    pairs: list[FilePair] = []
    used_old: set[str] = set()
    used_new: set[str] = set()

    # 完全一致を優先
    for f in old_files:
        if f in new_set:
            pairs.append(FilePair(old_name=f, new_name=f, score=1.0, matched_by="exact"))
            used_old.add(f)
            used_new.add(f)

    remaining_old = [f for f in old_files if f not in used_old]
    remaining_new = [f for f in new_files if f not in used_new]

    # 類似度スコアを全候補ペアで計算
    candidates: list[tuple[float, str, str]] = []
    for old_f in remaining_old:
        for new_f in remaining_new:
            score = SequenceMatcher(None, old_f, new_f, autojunk=False).ratio()
            candidates.append((score, old_f, new_f))
    candidates.sort(key=lambda x: -x[0])

    matched_old: set[str] = set()
    matched_new: set[str] = set()

    for score, old_f, new_f in candidates:
        if score < threshold:
            break
        if old_f not in matched_old and new_f not in matched_new:
            pairs.append(FilePair(
                old_name=old_f, new_name=new_f,
                score=round(score, 3), matched_by="auto",
            ))
            matched_old.add(old_f)
            matched_new.add(new_f)

    # 未マッチ
    for f in remaining_old:
        if f not in matched_old:
            pairs.append(FilePair(old_name=f, new_name=None, score=0.0, matched_by="unmatched_old"))
    for f in remaining_new:
        if f not in matched_new:
            pairs.append(FilePair(old_name=None, new_name=f, score=0.0, matched_by="unmatched_new"))

    return pairs


def save_pairs(pairs: list[FilePair], path: str) -> None:
    data = [
        {"old_name": p.old_name, "new_name": p.new_name,
         "score": p.score, "matched_by": p.matched_by}
        for p in pairs
    ]
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_pairs(path: str) -> list[FilePair]:
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    return [FilePair(**d) for d in data]


# ---------------------------------------------------------------------------
# generate_regex: 確認済みペアから正規表現を自動生成
# ---------------------------------------------------------------------------

# 可変部分の分類パターン（長い順に試す）
_VAR_PATTERNS: list[tuple[str, re.Pattern]] = [
    (r"\d{8}", re.compile(r"^\d{8}$")),
    (r"\d{6}", re.compile(r"^\d{6}$")),
    (r"v\d+",  re.compile(r"^v\d+$")),
    (r"\d+",   re.compile(r"^\d+$")),
]


def _classify_var(old_var: str, new_var: str) -> Optional[str]:
    """2つの可変部分に共通する正規表現パターンを返す。"""
    for pattern_str, pattern_re in _VAR_PATTERNS:
        if pattern_re.fullmatch(old_var) and pattern_re.fullmatch(new_var):
            return pattern_str
    return None


def _split_stem(stem: str) -> tuple[str, str, str]:
    """
    ファイルステムを (prefix, separator, variable) に分割する。
    末尾の _XXX または -XXX を可変部分として扱う。
    """
    for sep in ("_", "-"):
        idx = stem.rfind(sep)
        if idx > 0:
            return stem[:idx], sep, stem[idx + 1:]
    return stem, "", ""


def generate_regex(pairs: list[FilePair]) -> Optional[str]:
    """
    確認済みペアから key_regex の候補を生成する。
    capture group 1 がキーになる正規表現を返す。
    全ペアが自動分類できない場合は None を返す。
    """
    differing = [
        (p.old_name, p.new_name)
        for p in pairs
        if p.old_name and p.new_name and p.old_name != p.new_name
    ]

    if not differing:
        return None  # 差分ペアなし → パターン不要

    var_patterns: list[str] = []
    sep_chars: set[str] = set()
    ext_set: set[str] = set()

    for old_f, new_f in differing:
        old_stem, old_ext = os.path.splitext(old_f)
        new_stem, new_ext = os.path.splitext(new_f)
        ext_set.add(re.escape(old_ext))
        ext_set.add(re.escape(new_ext))

        old_prefix, old_sep, old_var = _split_stem(old_stem)
        new_prefix, new_sep, new_var = _split_stem(new_stem)

        # プレフィックスまたはセパレータが違う → 自動生成不可
        if old_sep != new_sep or old_prefix != new_prefix:
            return None

        classified = _classify_var(old_var, new_var)
        if classified is None:
            return None

        sep_chars.add(re.escape(old_sep))
        var_patterns.append(classified)

    if not var_patterns:
        return None

    sep = list(sep_chars)[0] if len(sep_chars) == 1 else f"[{''.join(sep_chars)}]"
    ext = list(ext_set)[0] if len(ext_set) == 1 else f"(?:{'|'.join(ext_set)})"

    # 重複排除・順序保持
    unique_vars = list(dict.fromkeys(var_patterns))
    var_re = unique_vars[0] if len(unique_vars) == 1 else f"(?:{'|'.join(unique_vars)})"

    return f"^(.+?){sep}{var_re}{ext}$"


# ---------------------------------------------------------------------------
# validate_regex: 正規表現の検証
# ---------------------------------------------------------------------------

def validate_regex(pairs: list[FilePair], key_regex: str) -> list[ValidationError]:
    """
    key_regex が確認済みペアと同じペアリングを100%再現できるか検証する。
    エラーリストを返す（空なら OK）。
    """
    try:
        pattern = re.compile(key_regex)
    except re.error as e:
        return [ValidationError("invalid_regex", f"正規表現エラー: {e}")]

    errors: list[ValidationError] = []
    key_to_old: dict[str, str] = {}
    key_to_new: dict[str, str] = {}

    matched_pairs = [p for p in pairs if p.old_name and p.new_name]

    for p in matched_pairs:
        m_old = pattern.fullmatch(p.old_name)
        m_new = pattern.fullmatch(p.new_name)

        if not m_old:
            errors.append(ValidationError("no_match", f"{p.old_name} が正規表現にマッチしない"))
            continue
        if not m_new:
            errors.append(ValidationError("no_match", f"{p.new_name} が正規表現にマッチしない"))
            continue

        key_old = m_old.group(1)
        key_new = m_new.group(1)

        if key_old != key_new:
            errors.append(ValidationError(
                "key_mismatch",
                f'{p.old_name} のキー "{key_old}" ≠ {p.new_name} のキー "{key_new}"',
            ))
            continue

        key = key_old
        if key in key_to_old and key_to_old[key] != p.old_name:
            errors.append(ValidationError(
                "key_collision",
                f'キー "{key}" が {key_to_old[key]} と {p.old_name} の両方にマッチ',
            ))
        else:
            key_to_old[key] = p.old_name

        if key in key_to_new and key_to_new[key] != p.new_name:
            errors.append(ValidationError(
                "key_collision",
                f'キー "{key}" が {key_to_new[key]} と {p.new_name} の両方にマッチ',
            ))
        else:
            key_to_new[key] = p.new_name

    return errors


# ---------------------------------------------------------------------------
# apply_pattern: パターンを使ってフォルダをペアリング
# ---------------------------------------------------------------------------

def apply_pattern(old_dir: str, new_dir: str, key_regex: str) -> list[FilePair]:
    """
    key_regex のキャプチャグループ1をキーとして、旧・新フォルダをペアリングする。
    正規表現にマッチしないファイルは unmatched として返す。
    """
    pattern = re.compile(key_regex)
    old_files = _xlsx_files(old_dir)
    new_files = _xlsx_files(new_dir)

    def extract_key(fname: str) -> Optional[str]:
        m = pattern.fullmatch(fname)
        return m.group(1) if m else None

    old_keyed: dict[str, str] = {}
    old_unkeyed: list[str] = []
    for f in old_files:
        k = extract_key(f)
        if k is not None:
            old_keyed[k] = f
        else:
            old_unkeyed.append(f)

    new_keyed: dict[str, str] = {}
    new_unkeyed: list[str] = []
    for f in new_files:
        k = extract_key(f)
        if k is not None:
            new_keyed[k] = f
        else:
            new_unkeyed.append(f)

    pairs: list[FilePair] = []
    used_new_keys: set[str] = set()

    for key, old_f in sorted(old_keyed.items()):
        if key in new_keyed:
            pairs.append(FilePair(
                old_name=old_f, new_name=new_keyed[key],
                score=1.0, matched_by="pattern",
            ))
            used_new_keys.add(key)
        else:
            pairs.append(FilePair(old_name=old_f, new_name=None, score=0.0, matched_by="unmatched_old"))

    for key, new_f in sorted(new_keyed.items()):
        if key not in used_new_keys:
            pairs.append(FilePair(old_name=None, new_name=new_f, score=0.0, matched_by="unmatched_new"))

    for f in old_unkeyed:
        pairs.append(FilePair(old_name=f, new_name=None, score=0.0, matched_by="unmatched_old"))
    for f in new_unkeyed:
        pairs.append(FilePair(old_name=None, new_name=f, score=0.0, matched_by="unmatched_new"))

    return pairs
