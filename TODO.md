# TODO / 未対応課題

## 優先度: 高

### ~~[TODO-001] キーJOIN方式の差分モード追加~~ ✅ 完了

**概要**  
現在のLCSベースの差分に加え、指定した列をキーとして行をJOINする差分モードを追加する。
行の出現順に依存せず、キーの一致/不一致で DELETE・INSERT・MODIFY を判定できる。

**背景**  
LCSは行の並び順を前提とするため、行が並び替わった場合に意図しない差分が発生する。
テーブル定義のような「各行にユニークなキーが存在するデータ」ではJOIN方式の方が適切。

**設計方針**

```python
# run_diff.py / CLI での指定イメージ
config.key_cols = [1, 2]    # B列・C列を複合キーとして使用（0始まり）
config.diff_mode = "key"    # "lcs"（デフォルト）or "key"
```

**内部処理**

```
old_map = { (B値, C値): row }  for row in old_rows
new_map = { (B値, C値): row }  for row in new_rows

all_keys = old_keysの出現順 ∪ (new_keysのみにあるものを末尾に追加)

for key in all_keys:
    old有・new有 → セル単位比較 → EQUAL or MODIFY
    old有・new無 → DELETE
    old無・new有 → INSERT
```

**前提・制約**
- キーは複合可（単独列も可）、キー重複なし前提
- キー列自体が変更された行は「DELETE + INSERT」として表示（キーが変わったら別行扱い）
- キー列に空値がある行は通常行として末尾にまとめる（またはLCSフォールバック）
- 表示順はoldの行番号順を基準とする

**実装対象ファイル**
- `excel_diff/matcher.py`: `DiffConfig` に `key_cols: list[int]` と `diff_mode: str` を追加
- `excel_diff/diff_engine.py`: `_diff_sheet_rows_by_key()` を新規追加、`diff_files()` でモード分岐
- `excel_diff/__main__.py`: `--key-cols` / `--diff-mode` オプション追加
- `excel_diff/html_renderer.py`: 変更なし（RowDiff構造は同じ）

**工数見込み**: 小（diff_engine.pyに50行程度追加）

---

## 優先度: 中

### ~~[TODO-002] 水平スクロールバーの表示問題~~ ✅ 完了

**概要**  
行数が多くパネルが一画面に収まらない場合、水平スクロールバーがパネル内に表示されない。
データが少なく一画面に収まる場合は正常に表示される。

**現象**  
- `sheets-container` が `overflow-y: auto` のため、`.sheet-panels` の `max-height` が効かず
  パネルがコンテンツ全体に伸び、水平スクロールバーがデータ最下部に押し出される

**試みた対策（いずれも未解決）**  
- `overflow-x: hidden` → `visible` → 削除
- グリッド → フレックスボックスへの変更
- `min-width: 0` の追加
- `panel-wrapper` 構造の導入
- `-webkit-scrollbar` の削除
- `overflow: hidden` / `height` の各種組み合わせ

**根本的な解決案（未実装）**  
`sheets-container` 自体をスクロール要素にするのをやめ、
各 `.panel` が `height: calc(100vh - Xpx)` で固定高さを持つ構造にする。
複数シートがある場合の縦ナビゲーションは別途検討が必要。

---

---

### [TODO-003] split_workbook の高速化（zipfile 参照追跡方式）

**概要**  
現在の `splitter.py` は openpyxl で xlsx をシート数分フルロード＆保存するため、  
123シートで約6分かかる。zipfile を直接操作する方式で 50倍以上の高速化が見込まれる。

**失敗した試みと原因**  
1. **ET.tostring() 方式**: `xml.etree.ElementTree` で workbook.xml を書き換えると  
   名前空間プレフィックスが変わり（例: `r:id` → `ns0:id`）、Excel が読めなくなる。  
2. **正規表現バイト操作方式**: `workbook.xml` / `workbook.xml.rels` / `[Content_Types].xml` を  
   バイト列の正規表現で書き換えても、以下の孤立ファイルが ZIP 内に残存し Excel が破損と判定する。  
   - `xl/calcChain.xml` — 全シートのセル計算順序を参照
   - `xl/worksheets/_rels/sheetN.xml.rels` — 他シートの外部参照
   - `xl/printerSettings/printerSettingsN.bin` — 各シートのプリンター設定
   - `xl/drawings/*.vml`, `xl/comments*.xml` — 他シートの描画・コメント

**正しい実装方針（未実装）**  
「対象シートから到達できるファイルのみを ZIP に含める」ホワイトリスト方式が必要。

```
手順:
1. workbook.xml を解析して対象シートの rId → sheetN.xml を特定
2. xl/worksheets/_rels/sheetN.xml.rels を解析し参照ファイルを収集
3. [Content_Types].xml のエントリも収集ファイルだけ残す
4. calcChain.xml は常に除外（参照が全シートにまたがるため再生成不可）
5. 上記ファイルのみを新 ZIP に書き込む
6. workbook.xml / workbook.xml.rels / [Content_Types].xml はバイト操作で修正
   （ET.tostring() は使わず、raw bytes の正規表現のみ使用）
```

**工数見込み**: 中（splitter.py に 100〜150行の追加実装）  
**期待効果**: 123シートで 6分 → 数秒（57倍以上）

---

## 優先度: 低（将来の拡張候補）

- 列レベルのLCS diff（列挿入・削除への対応）
- .xls 形式のサポート
- マッチャータイプの追加（正規表現、数値許容誤差など）
- `--sheet` の複数指定
