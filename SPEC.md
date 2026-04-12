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

## 3. ファイル構成

```
excel_diff/
├── __init__.py
├── __main__.py          CLIエントリポイント
├── reader.py            Excelファイル読込
├── matcher.py           カスタムマッチャー・列フィルタ定義
├── diff_engine.py       差分アルゴリズム
├── html_renderer.py     HTML出力
├── file_pairing.py      ファイルペアリング（discover / 正規表現生成・検証 / パターン適用）
├── patterns.py          パターン定義の永続化（PatternStore / patterns.json）
└── splitter.py          ブックをシート単位ファイルに分解

tests/
├── make_fixtures.py     テスト用Excelファイル生成
└── test_diff.py         ユニットテスト（37件）

SPEC.md
README.md
requirements.txt         依存ライブラリ（openpyxlのみ）
build.bat                PyInstallerビルドスクリプト
patterns.json            保存済みパターン定義（ユーザー作成）
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

両フォルダ内のxlsxファイルをペアリングして比較する。
デフォルトはファイル名の完全一致でペアリング。`--pattern` 指定時は保存済みパターンを使用。
片方にしか存在しないファイル（またはパターン未マッチファイル）は「比較対象外」としてコンソールに報告する。

### 4-3. ペアリング候補の探索

```
excel-diff.exe --discover <旧フォルダ> <新フォルダ> [-o pairs.json] [--threshold 0.6]
```

類似度スコアでペア候補を探索し、JSONに保存する（人手で確認・修正するためのたたき台）。

### 4-4. パターンの生成・保存

```
excel-diff.exe --gen-pattern pairs.json --id <ID> --name <名前> [--regex <正規表現>]
```

確認済みペアJSONから `key_regex` を自動生成し `patterns.json` に保存する。
自動生成に失敗した場合は `--regex` で正規表現を手動指定できる（その場合も検証を実行）。

### 4-5. パターン一覧表示

```
excel-diff.exe --list-patterns [--patterns-file FILE]
```

### 4-6. ブックのシート分解

```
excel-diff.exe --split <ブック.xlsx> [--prefix TEXT] [--suffix TEXT] [--output-dir DIR]
```

1つのブックを「1シート = 1ファイル」に分解する。
出力ファイル名は `<prefix><シート名><suffix>.xlsx`。
シート名にファイル名不正文字（`\ / : * ? " < > |`）が含まれる場合は `_` に置換する。

### 4-7. オプション一覧

**ファイル比較・フォルダ比較**

| オプション | 説明 | デフォルト |
|---|---|---|
| `-o`, `--output PATH` | 出力HTMLファイルパス | `<新ファイル名>_diff.html` |
| `--output-dir DIR` | フォルダ比較時の出力先フォルダ | `<旧フォルダ>_vs_<新フォルダ>/` |
| `--pattern ID` | フォルダ比較時のペアリングパターンID | なし（完全一致） |
| `--patterns-file FILE` | パターン定義ファイルのパス | `patterns.json` |
| `--sheet NAME` | 比較対象シートを絞り込み | 全シート |
| `--strikethrough` | 取り消し線の有無も差分として扱う | 無効（値のみ比較） |
| `--matchers FILE` | カスタムマッチャー／列フィルタ設定JSONファイル | なし |
| `--include-cols SPEC` | 比較対象列の指定（例: `A:C,E`） | 全列比較 |
| `--open` | 生成後にブラウザで自動オープン | 無効 |

**ペアリングパターン管理**

| オプション | 説明 |
|---|---|
| `--discover OLD NEW` | ファイルペア候補を探索してJSONに保存 |
| `--threshold SCORE` | `--discover` の類似度しきい値（0〜1） デフォルト: 0.6 |
| `--gen-pattern FILE` | 確認済みペアJSONからパターンを生成・保存 |
| `--id ID` | `--gen-pattern`: パターンID |
| `--name NAME` | `--gen-pattern`: パターン名 |
| `--regex REGEX` | `--gen-pattern`: 正規表現を手動指定 |
| `--list-patterns` | 保存済みパターンを一覧表示 |

**ブック分解**

| オプション | 説明 | デフォルト |
|---|---|---|
| `--split FILE` | 分解するブックのパス | — |
| `--prefix TEXT` | 出力ファイル名の前置文字列 | なし |
| `--suffix TEXT` | 出力ファイル名の後置文字列（拡張子の前） | なし |
| `--output-dir DIR` | 出力先フォルダ | ブックと同フォルダ |

### 4-8. コンソール出力形式

**ファイル比較**

```
差分あり: sales.xlsx  (削除 1行、追加 2行、変更 3行)  → result.html
```

**フォルダ比較**

```
差分なし: 1 ファイル
差分あり: 2 ファイル
  sales.xlsx      (削除 1行、追加 2行、変更 3行)  → diffs\sales_diff.html
  inventory.xlsx  (削除 0行、追加 1行、変更 0行)  → diffs\inventory_diff.html
比較対象外: 1 ファイル
  [旧のみ] archive.xlsx
```

### 4-9. VBAからの呼び出し例

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
    ↓                    ※比較対象列のみハッシュに含める
[replaceブロックの行ペアリング]  ← 類似度スコアで最適ペアを選択
    ↓
[セルレベルdiff]      ← MODIFYの行のみ、比較対象列で比較
    ↓
[文字レベルdiff]      ← 変更セルの値を文字単位でdiff
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

1. 各行を「行ハッシュ（tuple）」に変換（比較対象列のみ使用）
2. `difflib.SequenceMatcher(autojunk=False)` でLCS計算
3. 結果のopcodeに基づき行操作を分類

| opcode | 行操作 |
|---|---|
| `equal` | EQUAL（変更なし） |
| `delete` | DELETE（旧ファイルの行が削除） |
| `insert` | INSERT（新ファイルの行が追加） |
| `replace` | replaceブロック内で類似度ペアリング → MODIFY or DELETE/INSERT |

### 5-5. replaceブロックのペアリング

`replace` ブロック内の旧行と新行を類似度スコアで最適にペアリングする。

- 比較対象列における一致セル数 ÷ 比較対象列数 をスコアとする
- スコアの高い順にグリーディーにペア確定（各行は1回のみ使用）
- スコアが 0.0（比較対象列で1セルも一致しない）場合はペア化せず DELETE/INSERT として扱う
- ペア確定した行はMODIFYとして処理

### 5-6. セルレベルdiff（MODIFYの行のみ）

1. 旧行・新行の列数を `max(旧列数, 新列数)` に揃えてNoneでパディング
2. 比較対象列のみ列ごとに比較し、異なるセルを `CellDiff` としてリスト化
3. `--strikethrough` 指定時は値に加えて取り消し線の有無も比較

### 5-7. 文字レベルdiff

変更セル（`CellDiff`）の旧値・新値を文字列化し、`difflib.SequenceMatcher` で文字単位のdiffを計算。
削除文字は `<mark class="char-del">` 、追加文字は `<mark class="char-new">` でHTMLマーク。

---

## 6. 比較列フィルタ

### 6-1. 用途

ファイル単位またはシート単位で「比較する列」を限定する。
指定外の列は差分計算から除外されるが、HTML上では薄いグレーで値を表示して文脈を維持する。

### 6-2. 列範囲の指定形式

| 指定例 | 対象列 |
|---|---|
| `A` | A列のみ |
| `A:C` | A〜C列（連続） |
| `A,C,E` | A・C・E列（飛び地） |
| `A:C,E` | A〜C列とE列（混在） |
| `1,3:5` | 1列目・3〜5列目（1始まり整数も使用可） |

### 6-3. 指定方法

**CLIオプション（全シートに適用）**

```
excel-diff.exe old.xlsx new.xlsx --include-cols "A:C,E"
```

**設定ファイル（シート別指定も可能）**

```json
{
  "include_cols": "A:C,E",
  "sheets": {
    "売上": { "include_cols": "A,C:F" }
  },
  "matchers": []
}
```

- `include_cols`（ルート）: 全シートに適用するグローバルフィルタ
- `sheets.<シート名>.include_cols`: シート個別フィルタ（グローバル設定より優先）
- `--include-cols` CLIオプションはJSON設定のグローバルフィルタを上書きする

### 6-4. 除外列の扱い

- 行ハッシュ計算（LCS用）から除外
- 行類似度スコア計算から除外
- セルdiff対象から除外（除外列のセル変化は差分として報告しない）
- HTML表示: グレー文字 + 薄いグレー背景（`.cell-excluded` クラス）でそのまま表示

---

## 7. カスタムマッチャー

### 7-1. 用途

特定列の新旧値が「意図的な変換」である場合に、差分なしとして扱う。

例：コードマスタの改番があり、変換前→変換後のペア一覧が存在する場合。
対比表通りの変換は差分なし、対比表外の変換は差分ありとして報告する。

### 7-2. 設定ファイル（JSON）

**新形式（列フィルタとマッチャーを同一ファイルで指定）**

```json
{
  "include_cols": "A:C",
  "matchers": [
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
}
```

**旧形式（後方互換: 配列形式）**

```json
[
  { "type": "mapping", "column": "B", "pairs": [["旧コード001", "新コード001"]] }
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

### 7-3. マッチャーの動作

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

### 7-4. 対応外の値の扱い

- 旧値がマッピングのキーに存在しない → 通常比較（差分として扱う）
- 旧値がマッピングのキーに存在するが、新値が期待値と異なる → 差分として扱う

---

## 8. HTML出力仕様

### 8-1. 構造

- 1ファイルに全シートを収める（CSS/JS内包の自己完結型）
- ヘッダ: 比較ファイルパス・比較日時
- ナビゲーション: シート名リンク（変更ステータス付き）
- サマリ: 変更シート数・変更行数の統計・マッチャー適用件数
- 各シートのdiffテーブル（サイドバイサイド表示または上下表示）

### 8-2. 表示形式

**サイドバイサイド（デフォルト）**

```
行番号 | 旧ファイルのセル列... | （区切り） | 行番号 | 新ファイルのセル列...
```

**上下表示（列数が多い場合向け）**

MODIFY行は「旧行（薄赤）→ 新行（薄緑）」の2行で表示する。
DELETE/INSERT は1行で表示。EQUAL は旧値1行で表示。

列ヘッダ（A, B, C...）を1行目に表示。

### 8-3. 色分け

| 状態 | 表示 |
|---|---|
| 変更なし行（EQUAL） | 白背景 |
| 削除行（DELETE） | 全セル薄赤（`#ffeef0`） |
| 追加行（INSERT） | 全セル薄緑（`#e6ffed`） |
| 変更行（MODIFY）/ SBS | 行番号のみ黄（`#fff5b1`）、セルは白背景 |
| 変更行（MODIFY）/ 上下・旧行 | 全セル薄赤（`#ffeef0`） |
| 変更行（MODIFY）/ 上下・新行 | 全セル薄緑（`#e6ffed`） |
| 変更セルの削除文字 | `<mark class="char-del">` 赤マーク（`#ffc0c0`） |
| 変更セルの追加文字 | `<mark class="char-new">` 緑マーク（`#2da44e`） |
| 比較除外列 | グレー文字・グレー背景（`.cell-excluded`） |

### 8-4. 機能

- 「変更行のみ表示」トグルボタン（JavaScriptで EQUAL行の表示/非表示）
- 「上下表示に切り替え」ボタン（`.view-sbs` と `.view-stacked` を `display:none` で切り替え）
  - 両ビューはHTML生成時に同時にレンダリングされ、JSで表示/非表示を切り替える
  - ボタンラベルはビュー切り替えに合わせて「サイドバイサイドに戻す」↔「上下表示に切り替え」と変化
- カスタムマッチャー使用時はサマリに「カスタムマッチャー: ○件適用」と表示
- 文字エンコーディング: UTF-8

---

## 9. データモデル

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
  normalize_old(val) → Any    LCS用正規化キー（旧ファイル側）
  normalize_new(val) → Any    LCS用正規化キー（新ファイル側）

MappingMatcher（ColumnMatcherの実装）
  forward: dict            {旧値: 新値}
  reverse: dict            {新値: 旧値}（逆引き用）

DiffConfig
  matchers: list[ColumnMatcher]
  global_col_filter: Optional[set[int]]   全シート共通フィルタ（0始まり列インデックス集合）
  sheet_col_filters: dict[str, set[int]]  シート別フィルタ（グローバルより優先）
  get_col_filter(sheet_name) → Optional[set[int]]

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
  col_filter: Optional[set[int]]   比較対象列（Noneは全列）

FileDiff
  old_path: str
  new_path: str
  sheet_diffs: list[SheetDiff]
  has_differences: bool
  matcher_count: int        適用されたマッチャー数

FilePair                    （file_pairing.py）
  old_name: Optional[str]   旧フォルダ内ファイル名（Noneは新規追加ファイル）
  new_name: Optional[str]   新フォルダ内ファイル名（Noneは削除ファイル）
  score: float              類似度スコア（0.0〜1.0）
  matched_by: str           'exact' | 'pattern' | 'auto' | 'unmatched_old' | 'unmatched_new'

ValidationError             （file_pairing.py）
  kind: str                 'no_match' | 'key_mismatch' | 'key_collision' | 'invalid_regex'
  details: str              エラーの詳細説明

PatternDef                  （patterns.py）
  id: str                   パターンID（ユーザー定義）
  name: str                 パターン名（表示用）
  key_regex: str            キャプチャグループ1をキーとする正規表現
  description: str          説明（省略可）
  example_old_dir: str      生成時に使用した旧フォルダパス（省略可）
  example_new_dir: str      生成時に使用した新フォルダパス（省略可）
  created_at: str           作成日（ISO 8601）

PatternStore                （patterns.py）
  path: str                 patterns.jsonのパス
  get(id) → Optional[PatternDef]
  add_or_update(pattern)
  list_all() → list[PatternDef]
  save()
```

---

## 10. ファイルペアリングパターン管理

### 10-1. 概要

月次レポートなど、定期的にファイル名の一部（日付・バージョン番号）が変わるフォルダ比較に対応する。
「正規表現のキャプチャグループ1をキーとして旧ファイルと新ファイルをペアリングする」というルールをパターンとして保存・再利用できる。

### 10-2. ワークフロー

```
① discover  → ② 手動確認・編集 → ③ gen-pattern → ④ --pattern で繰り返し使用
```

1. `--discover` で類似度ベースのペア候補を探索し `pairs.json` に出力
2. `pairs.json` を確認・修正してペアを確定させる
3. `--gen-pattern` で `key_regex` を自動生成（または手動指定）し `patterns.json` に保存
4. 以降は `--dir ... --pattern <ID>` で同じルールを適用

### 10-3. key_regex の仕様

- 正規表現のキャプチャグループ1（`(...)` 部分）がペアリングキーになる
- 旧ファイル名と新ファイル名で group(1) が一致したもの同士をペアとする
- パターンにマッチしないファイルは「比較対象外（パターン未マッチ）」として扱う

**自動生成ロジック**

`_split_stem()` でファイルステムを `(prefix, sep, variable)` に分割し、可変部分を正規表現に変換する。

| 可変部分の例 | 変換後 |
|---|---|
| `20240101` / `20240201` | `\d{8}` |
| `202401` / `202402` | `\d{6}` |
| `v1` / `v2` | `v\d+` |
| `001` / `002` | `\d+` |

生成例: `^(.+?)_(?:\d{8}|v\d+)\.xlsx$`

### 10-4. 検証ルール

`--gen-pattern` 実行時、保存前に以下を検証する（1件でもエラーがあれば保存しない）。

| エラー種別 | 内容 |
|---|---|
| `invalid_regex` | 正規表現の文法エラー |
| `no_match` | 確認済みペアのいずれかのファイル名が正規表現にマッチしない |
| `key_mismatch` | 旧ファイルのキー ≠ 新ファイルのキー |
| `key_collision` | 同一キーに複数ファイルがマッチ |

### 10-5. patterns.json フォーマット

```json
{
  "patterns": [
    {
      "id": "monthly",
      "name": "月次レポート",
      "key_regex": "^(.+?)_(?:\\d{8}|v\\d+)\\.xlsx$",
      "description": "",
      "example_old_dir": "old_202401",
      "example_new_dir": "new_202402",
      "created_at": "2024-02-01"
    }
  ]
}
```

---

## 11. 制約・スコープ外

| 項目 | 扱い |
|---|---|
| 列の挿入・削除 | スコープ外（将来の拡張課題）。列数は固定として扱う |
| セルの書式（色・フォント等） | スコープ外。`--strikethrough` のみ対応 |
| マージセル範囲の変更 | スコープ外。値比較のみ（slave cellはNone扱い） |
| 数式テキストの比較 | スコープ外（`data_only=True` で計算済み値のみ） |
| グラフ・画像 | 無視 |
| .xls 形式（旧形式） | 非対応（.xlsxのみ） |

---

## 12. テスト方針

- `tests/make_fixtures.py` でテスト用Excelファイルを自動生成（外部ファイル依存なし）
- `tests/test_diff.py` でdiffエンジンのユニットテスト（openpyxlを使わずデータモデル直接）

### テストケース一覧

**列フィルタ（8件）**

| ケース | 内容 |
|---|---|
| parse_col_spec: 単一列 | `"B"` → `{1}` |
| parse_col_spec: 連続範囲 | `"A:C"` → `{0,1,2}` |
| parse_col_spec: 飛び地 | `"A,C,E"` → `{0,2,4}` |
| parse_col_spec: 混在 | `"A:C,E"` → `{0,1,2,4}` |
| 除外列は差分なし | 除外列のセル変化を差分として検出しない |
| 対象列の差分を検出 | 比較対象列の変更は検出され、除外列はcell_diffsに含まない |
| LCSに影響しない | 除外列の相違が行マッチ（LCS）に影響しない |
| SheetDiffに格納 | col_filterがSheetDiffフィールドに正しく格納される |

**差分エンジン（12件）**

| ケース | 内容 |
|---|---|
| 差分なし | 完全一致のシート |
| セル変更 | 1セルだけ変更 |
| 行挿入（中間） | 中間に行追加 |
| 行削除（中間） | 中間の行削除 |
| 挿入と削除が混在 | 複数箇所の挿入・削除が同時に発生 |
| シート追加 | 新ファイルにシートが増えた |
| シート削除 | 新ファイルからシートが消えた |
| マッチャー: 一致→差分なし | 対比表通りの変換 → 差分なし |
| マッチャー: マッピング外→差分あり | 対比表外の変換 → 差分あり |
| マッチャー: 一部列のみ一致 | マッチャー列は差分なし、別列の差分は検出 |
| マッチャー: シートスコープ | シート指定マッチャーが対象シートのみに適用 |
| 行番号の正確性 | 挿入・削除後も行番号がずれない |

**文字レベルdiff（17件）**

| ケース | 内容 |
|---|---|
| 全角: 末尾追加 | "みかん" → "みかんジュース" |
| 全角: 先頭追加 | "東京" → "新東京" |
| 全角: 中間変更 | "東京都渋谷区" → "東京都新宿区" |
| 全角: 完全置換 | "りんご" → "バナナ" |
| 全角: 末尾削除 | "みかんジュース" → "みかん" |
| 半角: 末尾1文字変更 | "ABC" → "ABD" |
| 半角: 中間変更 | "商品コードA-001" → "商品コードB-001" |
| 半角: バージョン番号 | "v1.0.0" → "v1.2.0" |
| 数値: 1桁変更 | 60 → 70 |
| 数値: 4桁変更 | 1200 → 1400 |
| 数値: 桁数増加 | 999 → 1000 |
| 数値: 小数 | 3.14 → 3.15 |
| 特殊: None → 値 | None → "新規追加" |
| 特殊: 値 → None | "削除予定" → None |
| 特殊: 同一値（markなし） | "変更なし" → "変更なし" |
| 全角数字変更 | "１２３" → "１２４" |
| 混在: 日付文字列 | "2024年1月期" → "2024年3月期" |

---

## 13. ビルド手順

```bat
:: 初回セットアップ
python -m venv .venv
.venv\Scripts\activate
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

## 14. 将来の拡張候補（スコープ外）

- 列レベルのLCS diff（列挿入・削除への対応）
- GUIモード（tkinterでファイル選択ダイアログ）
- .xls 形式のサポート（xlrd経由）
- マッチャータイプの追加（正規表現、数値許容誤差など）
- シートフィルタ（`--sheet` の複数指定）
