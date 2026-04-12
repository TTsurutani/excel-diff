"""
ファイルペアリングパターンの管理モジュール。

patterns.json にパターンを保存・読込・一覧・取得する。
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Optional


DEFAULT_PATTERNS_FILE = "patterns.json"


@dataclass
class PatternDef:
    id: str
    name: str
    key_regex: str
    description: str = ""
    example_old_dir: str = ""
    example_new_dir: str = ""
    created_at: str = ""


class PatternStore:
    def __init__(self, path: str = DEFAULT_PATTERNS_FILE):
        self.path = path
        self._patterns: list[PatternDef] = []
        if Path(path).exists():
            self._load()

    def _load(self) -> None:
        with open(self.path, encoding="utf-8") as f:
            data = json.load(f)
        self._patterns = [PatternDef(**p) for p in data.get("patterns", [])]

    def save(self) -> None:
        data = {
            "patterns": [
                {
                    "id": p.id,
                    "name": p.name,
                    "key_regex": p.key_regex,
                    "description": p.description,
                    "example_old_dir": p.example_old_dir,
                    "example_new_dir": p.example_new_dir,
                    "created_at": p.created_at,
                }
                for p in self._patterns
            ]
        }
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def get(self, pattern_id: str) -> Optional[PatternDef]:
        for p in self._patterns:
            if p.id == pattern_id:
                return p
        return None

    def add_or_update(self, pattern: PatternDef) -> None:
        for i, p in enumerate(self._patterns):
            if p.id == pattern.id:
                self._patterns[i] = pattern
                return
        self._patterns.append(pattern)

    def list_all(self) -> list[PatternDef]:
        return list(self._patterns)
