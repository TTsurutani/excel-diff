# excel-diff

ExcelファイルをWinMerge風にdiffするCLIツール。
行の挿入・削除があっても位置ズレせず、変更セルを文字単位でハイライトしたHTMLを出力する。

---

## インストール

```bat
git clone https://github.com/TTsurutani/excel-diff.git
cd excel-diff
setup.bat
```

`setup.bat` が仮想環境の作成・依存インストール・テストデータ生成を一括で行う。

---

## 基本的な使い方

### ファイルを1対1で比較する

```bat
python -m excel_diff old.xlsx new.xlsx
```

新ファイルと同じフォルダに `<新ファイル名>_diff.html` が生成される。

```bat
:: 出力先を指定する
python -m excel_diff old.xlsx new.xlsx -o result.html

:: 生成後ブラウザで開く
python -m excel_diff old.xlsx new.xlsx --open
```

### フォルダ内の全xlsxをまとめて比較する

```bat
python -m excel_diff --dir old_dir new_dir
```

同名ファイルをそれぞれ比較し、`old_dir_vs_new_dir/` フォルダに各ファイルの差分HTMLを出力する。

```bat
:: 出力先フォルダを指定する
python -m excel_diff --dir old_dir new_dir --output-dir diffs/
```

---

## 差分モード

### LCSモード（デフォルト）

行の出現順を基準に `SequenceMatcher`（LCS）で差分を計算する。

```bat
python -m excel_diff old.xlsx new.xlsx
```

### キーJOINモード（`--key-cols`）

指定した列の値をキーとして行を結合し、キーの一致で差分を判定する。
行の並び替えがあっても、同じキーの行同士が正確に対応付けられる。

```bat
:: C列をキーとして使用
python -m excel_diff old.xlsx new.xlsx --key-cols C

:: B列・C列の複合キー
python -m excel_diff old.xlsx new.xlsx --key-cols B,C
```

**動作仕様：**

| 状態 | 判定 |
|---|---|
| 旧・新 両方にキーが存在 | EQUAL または MODIFY（セル単位比較） |
| 旧にのみキーが存在 | DELETE |
| 新にのみキーが存在 | INSERT |
| キー列が空値の行 | 末尾でLCSにフォールバック |

表示順は旧ファイルの行順を基準とし、新にしかないキーは末尾に追加。  
キーは重複なし前提（重複があった場合は先頭行のみキー扱いとし、残りはLCSフォールバック）。

---

## コンソール出力例

```
差分なし: 1 ファイル
差分あり: 2 ファイル
  sales.xlsx      (削除 1行、追加 2行、変更 3行)  → diffs\sales_diff.html
  inventory.xlsx  (削除 0行、追加 1行、変更 0行)  → diffs\inventory_diff.html
```

---

## HTML出力の機能

### 表示レイアウト

- **サイドバイサイド表示**（デフォルト）: 旧ファイル左・新ファイル右の2パネル並列
- **上下表示**: 「上下表示」ボタンで切替。列数が多い場合も縦スクロールで読める

### 行の色分け

| 状態 | 表示 |
|---|---|
| 変更なし行（EQUAL） | 白背景 |
| 削除行（DELETE） | 全セル薄赤 |
| 追加行（INSERT） | 全セル薄緑 |
| 変更行（MODIFY） | 行番号のみ黄・**変更セルのみ**黄背景 |
| 対応なし側の空き行 | グレー背景 |
| 比較除外列 | グレー文字・薄グレー背景 |

変更セルは文字単位でマーク（削除文字=赤、追加文字=緑）。

### ツールバー

**表示グループ**

| ボタン | 機能 |
|---|---|
| 変更行のみ（デフォルトON） | EQUAL行・対応なし行を折りたたみ/展開 |
| 上下表示 | 左右パネル ⇄ 上下パネル切替 |

**スクロールグループ**

| ボタン | 機能 |
|---|---|
| 3列固定 | A/B/C列を横スクロール中も固定表示 |
| 垂直同期（デフォルトON） | 左右パネルの垂直スクロールを連動 |
| 水平同期（デフォルトON） | 左右パネルの水平スクロールを連動 |

ON/OFF状態は青背景+太字で視覚的に区別。各ボタンにホバーでツールチップ表示。

### その他

- ページ読み込み時にデフォルトで「変更行のみ」表示
- 列ヘッダ（A/B/C…）をスティッキー表示（スクロール中も固定）
- 行番号列を横スクロール時もスティッキー表示
- シートナビゲーションリンク（変更ステータスバッジ付き）
- パネル高さをJSで動的計算し、水平スクロールバーを常時ビューポート内に表示

---

## オプション一覧

### ファイル比較・フォルダ比較

| オプション | 説明 |
|---|---|
| `-o PATH` | 出力HTMLのパスを指定 |
| `--output-dir DIR` | フォルダ比較時の出力先フォルダ |
| `--pairs FILE` | フォルダ比較時に使用するペアJSONファイル（`--discover` で生成したもの） |
| `--pattern ID` | フォルダ比較時のファイルペアリングパターンID（`--pairs` より低優先） |
| `--patterns-file FILE` | パターン定義ファイルのパス（デフォルト: `patterns.json`） |
| `--sheet NAME` | 指定シートのみ比較 |
| `--strikethrough` | 取り消し線の有無も差分として扱う |
| `--matchers FILE` | カスタムマッチャー／列フィルタ設定JSON |
| `--include-cols SPEC` | 比較対象列を指定（例: `A:C,E`） |
| `--key-cols SPEC` | キーJOIN差分モードのキー列（例: `C` / `B,C`）。指定するとキーモードが有効になる |
| `--diff-mode MODE` | 差分モード: `lcs`（デフォルト）または `key` |
| `--open` | 生成後にブラウザで自動オープン |

### ペアリングパターン管理

| オプション | 説明 |
|---|---|
| `--discover OLD NEW` | ファイルペア候補を探索してJSONに保存 |
| `--threshold SCORE` | `--discover` の類似度しきい値（0〜1、デフォルト: 0.6） |
| `--gen-pattern FILE` | 確認済みペアJSONからパターンを生成・保存 |
| `--id ID` | `--gen-pattern`: パターンID |
| `--name NAME` | `--gen-pattern`: パターン名 |
| `--regex REGEX` | `--gen-pattern`: 正規表現を手動指定 |
| `--list-patterns` | 保存済みパターンを一覧表示 |

### ブック分解

| オプション | 説明 |
|---|---|
| `--split FILE` | ブックをシート単位のファイルに分解する |
| `--prefix TEXT` | 出力ファイル名の前置文字列 |
| `--suffix TEXT` | 出力ファイル名の後置文字列（拡張子の前） |
| `--output-dir DIR` | 出力先フォルダ（省略時はブックと同じフォルダ） |

---

## 比較対象列の絞り込み（`--include-cols`）

特定の列だけを比較対象にすることができる。除外列は差分計算に含まれないが、HTMLにはグレーで表示される。

```bat
:: B〜U列のみ比較
python -m excel_diff old.xlsx new.xlsx --include-cols "B:U"

:: A〜C列とE列のみ比較
python -m excel_diff old.xlsx new.xlsx --include-cols "A:C,E"
```

設定ファイルでシートごとに指定することも可能（詳細は `SPEC.md` 参照）。

---

## ファイル名が一致しないフォルダ比較

月次レポートなど、ファイル名の一部が変わる定期比較に対応するパターン管理機能。

### ① ペア候補を探索する

```bat
python -m excel_diff --discover old_202401 new_202402 -o pairs.json
```

```
ペア候補:
  [自動 94%] 売上_20240101.xlsx  →  売上_20240201.xlsx
  [自動 91%] 報告書_v1.xlsx      →  報告書_v2.xlsx
  [未対応-旧] archive.xlsx
→ pairs.json に保存しました
```

`pairs.json` を確認・編集してペアを確定させる。

### ② ペアJSONを使って直接比較する

確認済みのペアJSONをそのまま比較に使う。1回限りの比較や、正規表現パターンに落とし込めない命名規則にも対応できる。

```bat
python -m excel_diff --dir old_202401 new_202402 --pairs pairs.json
```

### ③ パターンを生成・保存する（繰り返し比較用）

同じ命名規則のフォルダを定期的に比較する場合は、ペアから正規表現パターンを生成・保存しておく。

```bat
python -m excel_diff --gen-pattern pairs.json --id monthly --name "月次レポート"
```

```
提案パターン: ^(.+?)_(?:\d{8}|v\d+)\.xlsx$
[OK] 検証OK: 2 ペアすべてが正しく再現されます
→ patterns.json に保存しました
```

### ④ パターンを使って比較する

```bat
python -m excel_diff --dir old_202402 new_202403 --pattern monthly
```

---

## カスタムマッチャー

コードの改番など「旧値→新値が意図的な変換」である列を差分なしとして扱いたい場合に使う。

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
    }
  ]
}
```

```bat
python -m excel_diff old.xlsx new.xlsx --matchers matchers.json
```

外部CSVやExcelから対比表を読み込む `mapping_file` タイプもある（詳細は `SPEC.md` 参照）。

---

## ブックをシート単位に分解する（`--split`）

1つのExcelブックを「シート1枚 = ファイル1つ」に分解する。

```bat
:: シート名のみ
python -m excel_diff --split book.xlsx

:: 前置文字列を付ける → 2024_シート名.xlsx
python -m excel_diff --split book.xlsx --prefix "2024_"

:: 後置文字列を付ける → シート名_確定.xlsx
python -m excel_diff --split book.xlsx --suffix "_確定"

:: 前置・後置と出力先フォルダを指定
python -m excel_diff --split book.xlsx --prefix "2024_" --suffix "_v2" --output-dir out/
```

---

## exeビルド（Python不要な配布用）

```bat
build.bat
```

`dist/excel-diff.exe` が生成される。

---

## 開発・テスト

```bat
:: テストデータ生成
python tests/make_fixtures.py

:: ユニットテスト実行
python tests/test_diff.py
```

詳細な仕様・アルゴリズム・データモデルは [SPEC.md](SPEC.md) を参照。
未対応課題・今後の実装予定は [TODO.md](TODO.md) を参照。
