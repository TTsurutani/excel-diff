# excel-diff

ExcelファイルをGitHub風にdiffするCLIツール。
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

## コンソール出力例

```
差分なし: 1 ファイル
差分あり: 2 ファイル
  sales.xlsx      (削除 1行、追加 2行、変更 3行)  → diffs\sales_diff.html
  inventory.xlsx  (削除 0行、追加 1行、変更 0行)  → diffs\inventory_diff.html
```

---

## HTML出力のイメージ

- **サイドバイサイド表示**（旧ファイル左、新ファイル右）
- 行の色分け: 削除行=赤、追加行=緑、変更行=行番号のみ黄
- 変更セルは文字単位でマーク（削除文字=赤、追加文字=緑）
- 「変更行のみ表示」ボタンで変更なし行を折りたたみ
- **「上下表示に切り替え」ボタン**で列数が多い場合でも縦スクロールで読める表示に切替可能

---

## オプション一覧

### ファイル比較・フォルダ比較

| オプション | 説明 |
|---|---|
| `-o PATH` | 出力HTMLのパスを指定 |
| `--output-dir DIR` | フォルダ比較時の出力先フォルダ |
| `--pattern ID` | フォルダ比較時のファイルペアリングパターンID |
| `--patterns-file FILE` | パターン定義ファイルのパス（デフォルト: `patterns.json`） |
| `--sheet NAME` | 指定シートのみ比較 |
| `--strikethrough` | 取り消し線の有無も差分として扱う |
| `--matchers FILE` | カスタムマッチャー／列フィルタ設定JSON |
| `--include-cols SPEC` | 比較対象列を指定（例: `A:C,E`） |
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

### ② パターンを生成・保存する

```bat
python -m excel_diff --gen-pattern pairs.json --id monthly --name "月次レポート"
```

```
提案パターン: ^(.+?)_(?:\d{8}|v\d+)\.xlsx$
[OK] 検証OK: 2 ペアすべてが正しく再現されます
→ patterns.json に保存しました
```

自動生成に失敗した場合は `--regex` で正規表現を手動指定できる（その場合も検証を実行）。

### ③ パターンを使って比較する

```bat
python -m excel_diff --dir old_202402 new_202403 --pattern monthly
```

```
パターン「月次レポート」を使用
  正規表現: ^(.+?)_(?:\d{8}|v\d+)\.xlsx$
ペアリング:
  売上_20240201.xlsx  →  売上_20240301.xlsx

差分あり: 1 ファイル
  売上_20240201.xlsx → 売上_20240301.xlsx  (削除 1行、追加 1行、変更 1行)

比較対象外: 1 ファイル
  [旧のみ] archive.xlsx
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
出力ファイル名は `<prefix><シート名><suffix>.xlsx` の形式。

```bat
:: シート名のみ（プレフィックス・サフィックスなし）
python -m excel_diff --split book.xlsx

:: シート名に前置文字列を付ける → 2024_シート名.xlsx
python -m excel_diff --split book.xlsx --prefix "2024_"

:: シート名に後置文字列を付ける → シート名_確定.xlsx
python -m excel_diff --split book.xlsx --suffix "_確定"

:: 前置・後置の両方を指定し、出力先フォルダも指定する
python -m excel_diff --split book.xlsx --prefix "2024_" --suffix "_v2" --output-dir out/
```

コンソール出力例:

```
分解中: book.xlsx
  前置: 2024_

3 シートを分解しました:
  → out\2024_売上.xlsx
  → out\2024_在庫.xlsx
  → out\2024_マスタ.xlsx
```

シート名にファイル名として使えない文字（`\ / : * ? " < > |`）が含まれる場合は自動的に `_` に置換される。

---

## exeビルド（Python不要な配布用）

```bat
build.bat
```

`dist/excel-diff.exe` が生成される。Python未インストールの環境でも動作する。

```bat
:: VBAからの呼び出し例
excel-diff.exe old.xlsx new.xlsx -o result.html
```

---

## 開発・テスト

```bat
:: テストデータ生成
python tests/make_fixtures.py

:: ユニットテスト実行（37件）
python tests/test_diff.py
```

詳細な仕様・アルゴリズム・データモデルは [SPEC.md](SPEC.md) を参照。
