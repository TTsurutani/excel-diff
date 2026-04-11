# excel-diff 仕様書

## 1. 概要

Excelファイルの差分をHTMLで出力するCLIツール。
行の挿入・削除があっても位置ズレを起こさず、GitHub風のdiff表示を実現する。

---

## 2. 動作環境

- OS: Windows 10/11
- 実行形式: `excel-diff.exe`（PyInstallerでビルド）
- Pythonランタイム不要（exeに同梱）

---

## 3. 実行ファイル構成

```
excel_diff/
├── __init__.py
├── __main__.py          CLIエントリポイント
├── reader.py            Excelファイル読込
├── matcher.py           カスタムマッチャー定義・読込
├── diff_engine.py       差分アルゴリズム
└── html_renderer.py     HTML出力

tests/
├── make_fixtures.py     テスト用Excelファイル生成
└── test_diff.py         ユニットテスト

requirements.txt         依存ライブラリ（openpyxlのみ）
build.bat                PyInstallerビルドスクリプト
```

---

## 4. CLI仕様

### 4-1. ファイル比較

```
excel-diff.exe <旧ファイル.xlsx> <新ファイル.xlsx> [オプション]
```

### 4-2. フォルダ一括比較

```
excel-diff.exe --dir <旧フォルダ> <新フォルダ> [オプション]
```

両フォルダ内の同名xlsxファイルを検索し、それぞれ比較する。
片方にしか存在しないファイルも「追加」「削除」として報告する。

### 4-3. オプション

| オプション | 説明 | デフォルト |
|---|---|---|
| `-o`, `--output PATH` | 出力HTMLファイルパス | `<新ファイル名>_diff.html` |
| `--output-dir DIR` | フォルダ比較時の出力先フォルダ | `<旧フォルダ>_vs_<新フォルダ>/` |
| `--sheet NAME` | 比較対象シートを絞り込み | 全シート |
| `--strikethrough` | 取り消し線の有無も差分として扱う | 無効（値のみ比較） |
| `--matchers FILE` | カスタムマッチャー設定JSONファイル | なし |
| `--open` | 生成後にブラウザで自動オープン | 無効 |

### 4-4. VBAからの呼び出し例

```vba
Dim wsh As Object
Set wsh = CreateObject("WScript.Shell")
Dim cmd As String
cmd = "excel-diff.exe """ & oldPath & """ """ & newPath & """ -o """ & outPath & """"
wsh.Run cmd, 0, True   ' 第3引数True = 完了まで待機
' 結果HTMLを開く
Shell "explorer """ & outPath & """"
```

---

## 5. 差分アルゴリズム

### 5-1. 全体フロー

```
[Excelファイル読込]
    ↓
[シートのマッチング]  ← 同名シートを対応付け
    ↓
[行レベルLCS diff]    ← SequenceMatcherで行挿入・削除を追跡
    ↓
[セルレベルdiff]      ← replaceブロックのみ、列単位で比較
    ↓
[HTML出力]
```

### 5-2. Excelファイル読込

- `openpyxl` を使用、`data_only=True` で計算済み値を取得
- マージセル: slave cell（値がNoneのセル）はそのままNoneとして扱う
- 末尾の空行（全セルNone）はトリミングして除去

### 5-3. シートマッチング

| 状況 | 扱い |
|---|---|
| 両方に存在するシート | 同名で対応付け、差分比較 |
| 旧ファイルにのみ存在 | 「シート削除」として全行DELETE表示 |
| 新ファイルにのみ存在 | 「シート追加」として全行INSERT表示 |

### 5-4. 行レベルLCS diff

1. 各行を「行ハッシュ（tuple）」に変換
2. `difflib.SequenceMatcher(autojunk=False)` でLCS計算
3. 結果のopcodeに基づき行操作を分類

| opcode | 行操作 |
|---|---|
| `equal` | EQUAL（変更なし） |
| `delete` | DELETE（旧ファイルの行が削除） |
| `insert` | INSERT（新ファイルの行が追加） |
| `replace` | replaceブロック内でペアリング → MODIFY or DELETE/INSERT |

`replace` ブロックのペアリング：
- `min(旧行数, 新行数)` の行を先頭から1対1でMODIFYとして対応付け
- 余った行はDELETE/INSERTとして扱う

### 5-5. セルレベルdiff（MODIFYの行のみ）

1. 旧行・新行の列数を `max(旧列数, 新列数)` に揃えてNoneでパディング
2. 列ごとに比較し、異なるセルを `CellDiff` としてリスト化
3. `--strikethrough` 指定時は値に加えて取り消し線の有無も比較

---

## 6. カスタムマッチャー

### 6-1. 用途

特定列の新旧値が「意図的な変換」である場合に、差分なしとして扱う。

例：コードマスタの改番があり、変換前→変換後のペア一覧が存在する場合。
対比表通りの変換は差分なし、対比表外の変換は差分ありとして報告する。

### 6-2. 設定ファイル（JSON）

```json
[
  {
    "type": "mapping",
    "column": "B",
    "sheet": null,
    "pairs": [
      ["旧コード001", "新コード001"],
      ["旧コード002", "新コード002"]
    ]
  },
  {
    "type": "mapping_file",
    "column": "C",
    "sheet": "売上",
    "file": "code_mapping.csv",
    "old_col": 0,
    "new_col": 1,
    "has_header": true
  }
]
```

| フィールド | 説明 |
|---|---|
| `type` | `"mapping"`（インライン）または `"mapping_file"`（外部ファイル） |
| `column` | 対象列（Excelの列記号: `"A"`, `"B"` ... または 0始まり整数） |
| `sheet` | 対象シート名（`null` の場合は全シートに適用） |
| `pairs` | `[旧値, 新値]` のリスト（typeが `mapping` の場合） |
| `file` | CSVまたはxlsxファイルパス（typeが `mapping_file` の場合） |
| `old_col` | 対比表の「変換前」列インデックスまたは列名 |
| `new_col` | 対比表の「変換後」列インデックスまたは列名 |
| `has_header` | 対比ファイルの1行目がヘッダか（デフォルト: `false`） |

### 6-3. マッチャーの動作

**行ハッシュへの組み込み（LCS用正規化）**

カスタムマッチャーが設定された列は、ハッシュ計算時に値を「正規化キー」に変換する。

```
マッピング: {"旧コード001": "新コード001"}
列Bに適用

旧行の列B: "旧コード001" → ハッシュキー: ("__mapped__", "旧コード001")
新行の列B: "新コード001" → 逆引きして → ハッシュキー: ("__mapped__", "旧コード001")

→ 同一ハッシュ → LCSがEQUALと判定
```

これによりカスタムマッチャーが行整列（LCS）とセル比較の両方に一貫して効く。

**セルdiffへの組み込み**

MODIFY行のセル比較時、対象列にマッチャーが適用され `True` を返した場合は差分として報告しない。

### 6-4. 対応外の値の扱い

- 旧値がマッピングのキーに存在しない → 通常比較（差分として扱う）
- 旧値がマッピングのキーに存在するが、新値が期待値と異なる → 差分として扱う

---

## 7. HTML出力仕様

### 7-1. 構造

- 1ファイルに全シートを収める
- ヘッダ: 比較ファイルパス・比較日時
- ナビゲーション: シート名リンク（変更ステータス付き）
- サマリ: 変更シート数・変更行数の統計
- 各シートのdiffテーブル（サイドバイサイド表示）

### 7-2. 表示形式

```
行番号 | 旧ファイルのセル列... | （区切り） | 行番号 | 新ファイルのセル列...
```

列ヘッダ（A, B, C...）を1行目に表示。

### 7-3. 色分け

| 状態 | 背景色 |
|---|---|
| 変更なし行（EQUAL） | 白 |
| 削除行（DELETE） | 薄赤（`#ffeef0`） |
| 追加行（INSERT） | 薄緑（`#e6ffed`） |
| 変更行（MODIFY） | 薄黄（`#fffbdd`） |
| 変更セル（旧値側） | 濃赤（`#fdb8c0`） |
| 変更セル（新値側） | 濃緑（`#acf2bd`） |

### 7-4. 機能

- 「変更行のみ表示」トグルボタン（JavaScriptで EQUAL行の表示/非表示）
- カスタムマッチャー使用時はヘッダに「マッチャー設定: ○件適用」と表示
- 文字エンコーディング: UTF-8

---

## 8. データモデル

```
CellData
  value: Any               セル値（計算済み）
  strikethrough: bool      取り消し線あり/なし

RowData
  row_idx: int             元ファイルでの行番号（1始まり）
  cells: list[CellData]

SheetData
  name: str                シート名
  rows: list[RowData]
  max_col: int             最大列数

ColumnMatcher（抽象）
  column_idx: int          対象列（0始まり）
  sheet: Optional[str]     対象シート名（Noneは全シート）
  matches(old, new) → bool

MappingMatcher（ColumnMatcherの実装）
  forward: dict            {旧値: 新値}
  reverse: dict            {新値: 旧値}（逆引き用）

CellDiff
  col_idx: int             変更列（0始まり）
  old_cell: CellData
  new_cell: CellData

RowDiff
  tag: RowTag              EQUAL / DELETE / INSERT / MODIFY
  old_row: Optional[RowData]
  new_row: Optional[RowData]
  cell_diffs: list[CellDiff]   MODIFYの場合のみ

SheetDiff
  name: str
  status: str              "equal" / "modified" / "added" / "deleted"
  row_diffs: list[RowDiff]
  max_cols: int
  col_letters: list[str]   ["A", "B", "C", ...]

FileDiff
  old_path: str
  new_path: str
  sheet_diffs: list[SheetDiff]
  has_differences: bool
```

---

## 9. 制約・スコープ外

| 項目 | 扱い |
|---|---|
| 列の挿入・削除 | スコープ外（将来の拡張課題）。列数は固定として扱う |
| セルの書式（色・フォント等） | スコープ外。`--strikethrough` のみ対応 |
| マージセル範囲の変更 | スコープ外。値比較のみ（slave cellはNone扱い） |
| 数式テキストの比較 | スコープ外（`data_only=True` で計算済み値のみ） |
| グラフ・画像 | 無視 |
| .xls 形式（旧形式） | 非対応（.xlsxのみ） |

---

## 10. テスト方針

- `tests/make_fixtures.py` でテスト用Excelファイルを自動生成（外部ファイル依存なし）
- `tests/test_diff.py` でdiffエンジンのユニットテスト（openpyxlを使わずデータモデル直接）

### テストケース一覧

| ケース | 内容 |
|---|---|
| 差分なし | 完全一致のシート |
| セル変更 | 1セルだけ変更 |
| 行挿入 | 中間に行追加 |
| 行削除 | 中間の行削除 |
| 行挿入＋削除混在 | 複数箇所の挿入・削除が同時に発生 |
| シート追加 | 新ファイルにシートが増えた |
| シート削除 | 新ファイルからシートが消えた |
| カスタムマッチャー（一致） | 対比表通りの変換 → 差分なし |
| カスタムマッチャー（不一致） | 対比表外の変換 → 差分あり |
| 日本語テキスト | 日本語値の比較 |
| 取り消し線 | `--strikethrough` 有効時の比較 |

---

## 11. ビルド手順

```bat
:: 初回セットアップ
pip install openpyxl pyinstaller

:: テスト実行
python tests/make_fixtures.py
python tests/test_diff.py

:: exeビルド
pyinstaller --onefile --name excel-diff --clean excel_diff/__main__.py

:: 成果物
:: dist/excel-diff.exe（約25MB）
```

---

## 12. 将来の拡張候補（スコープ外）

- 列レベルのLCS diff（列挿入・削除への対応）
- GUIモード（tkinterでファイル選択ダイアログ）
- 変更行のみ表示モード（デフォルト）/ 全行表示モード
- .xls 形式のサポート（xlrd経由）
- マッチャータイプの追加（正規表現、数値許容誤差など）
