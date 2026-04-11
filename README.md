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
  sales.xlsx    (削除 1行、追加 2行、変更 3行)  → diffs\sales_diff.html
  inventory.xlsx  (削除 0行、追加 1行、変更 0行)  → diffs\inventory_diff.html
```

---

## HTML出力のイメージ

- **サイドバイサイド表示**（旧ファイル左、新ファイル右）
- 行の色分け: 削除行=赤、追加行=緑、変更行=行番号のみ黄
- 変更セルは文字単位でマーク（削除文字=赤、追加文字=緑）
- 「変更行のみ表示」トグルボタンで変更なし行を折りたたみ可能

---

## オプション一覧

| オプション | 説明 |
|---|---|
| `-o PATH` | 出力HTMLのパスを指定 |
| `--output-dir DIR` | フォルダ比較時の出力先フォルダ |
| `--sheet NAME` | 指定シートのみ比較 |
| `--strikethrough` | 取り消し線の有無も差分として扱う |
| `--matchers FILE` | カスタムマッチャー／列フィルタ設定JSON |
| `--include-cols SPEC` | 比較対象列を指定（例: `A:C,E`） |
| `--open` | 生成後にブラウザで自動オープン |

---

## 比較対象列の絞り込み（`--include-cols`）

特定の列だけを比較対象にすることができる。除外列は差分計算に含まれないが、HTMLにはグレーで表示される。

```bat
:: A〜C列とE列のみ比較
python -m excel_diff old.xlsx new.xlsx --include-cols "A:C,E"
```

設定ファイルでシートごとに指定することも可能（詳細は後述）。

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
