"""
テスト用 Excel ファイルを生成するスクリプト。

使い方:
  python tests/make_fixtures.py
"""
import os
import sys

# パッケージを参照できるようにパスを通す
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import openpyxl
from pathlib import Path

FIXTURES_DIR = Path(__file__).parent / "fixtures"


def make_basic():
    """基本テスト: 行の追加・削除・セル変更が混在する例（日本語テキスト含む）"""
    # --- 旧ファイル ---
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    for row in [
        ["商品名",   "数量", "単価",  "合計"],
        ["りんご",   10,    100,    1000],
        ["バナナ",    5,     80,     400],   # ← 新ファイルで削除
        ["みかん",   20,    60,     1200],  # ← 単価・合計が変わる
        ["ぶどう",    3,    300,     900],
        ["いちご",    8,    200,    1600],
    ]:
        ws.append(row)
    wb.save(FIXTURES_DIR / "basic_old.xlsx")

    # --- 新ファイル ---
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    for row in [
        ["商品名",   "数量", "単価",  "合計"],
        ["りんご",   10,    100,    1000],
        # バナナ 削除
        ["みかん",   20,    70,     1400],  # 単価・合計 変更
        ["メロン",    2,    500,    1000],  # 新規 追加
        ["ぶどう",    3,    300,     900],
        ["いちご",    8,    200,    1600],
    ]:
        ws.append(row)
    wb.save(FIXTURES_DIR / "basic_new.xlsx")
    print("生成: basic_old.xlsx / basic_new.xlsx")


def make_multisheet():
    """複数シートテスト: シート追加・削除を含む"""
    # --- 旧ファイル ---
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "1月"
    for row in [["品目", "金額"], ["A", 100], ["B", 200], ["C", 300]]:
        ws1.append(row)
    ws2 = wb.create_sheet("2月")
    for row in [["品目", "金額"], ["A", 150], ["B", 250]]:
        ws2.append(row)
    ws3 = wb.create_sheet("削除シート")
    ws3.append(["このシートは削除されます"])
    wb.save(FIXTURES_DIR / "multi_old.xlsx")

    # --- 新ファイル ---
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "1月"
    for row in [["品目", "金額"], ["A", 100], ["B", 220], ["C", 300]]:  # B 変更
        ws1.append(row)
    ws2 = wb.create_sheet("2月")
    for row in [["品目", "金額"], ["A", 150], ["B", 250], ["D", 400]]:  # D 追加
        ws2.append(row)
    ws4 = wb.create_sheet("追加シート")  # 削除シート → 追加シート
    ws4.append(["このシートは追加されました"])
    wb.save(FIXTURES_DIR / "multi_new.xlsx")
    print("生成: multi_old.xlsx / multi_new.xlsx")


def make_no_diff():
    """差分なしテスト: 完全に同一のファイル"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for row in [["名前", "値"], ["A", 1], ["B", 2], ["C", 3]]:
        ws.append(row)
    wb.save(FIXTURES_DIR / "nodiff_old.xlsx")
    wb.save(FIXTURES_DIR / "nodiff_new.xlsx")
    print("生成: nodiff_old.xlsx / nodiff_new.xlsx")


def make_matcher_fixture():
    """カスタムマッチャーテスト用: コード改番が含まれる例"""
    # --- 旧ファイル ---
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "取引"
    for row in [
        ["取引先名",    "旧コード", "金額"],
        ["A商事",      "CODE001", 10000],
        ["B物産",      "CODE002", 20000],
        ["C株式会社",  "CODE003", 30000],  # ← コードは変わるが金額も変わる（差分あり）
        ["D商会",      "CODE999", 50000],  # ← マッピング外（差分として検出）
    ]:
        ws.append(row)
    wb.save(FIXTURES_DIR / "matcher_old.xlsx")

    # --- 新ファイル（コード改番後）---
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "取引"
    for row in [
        ["取引先名",    "新コード", "金額"],
        ["A商事",      "NEW001",  10000],  # CODE001→NEW001: マッチャーでEQUAL
        ["B物産",      "NEW002",  20000],  # CODE002→NEW002: マッチャーでEQUAL
        ["C株式会社",  "NEW003",  99999],  # CODE003→NEW003はEQUAL、金額変更は差分
        ["D商会",      "UNKNOWN", 50000],  # マッピング外の変換 → 差分あり
    ]:
        ws.append(row)
    wb.save(FIXTURES_DIR / "matcher_new.xlsx")

    # --- マッチャー設定 JSON ---
    import json
    config = [
        {
            "type": "mapping",
            "column": "B",
            "sheet": None,
            "pairs": [
                ["CODE001", "NEW001"],
                ["CODE002", "NEW002"],
                ["CODE003", "NEW003"],
            ]
        }
    ]
    with open(FIXTURES_DIR / "matchers.json", "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)

    print("生成: matcher_old.xlsx / matcher_new.xlsx / matchers.json")


if __name__ == "__main__":
    FIXTURES_DIR.mkdir(parents=True, exist_ok=True)
    make_basic()
    make_multisheet()
    make_no_diff()
    make_matcher_fixture()
    print(f"\nすべてのテストデータを生成しました: {FIXTURES_DIR}/")
