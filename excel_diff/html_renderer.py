"""
HTML出力モジュール。
差分結果を自己完結型のHTMLファイルとして生成する（CSS/JS内包）。
"""
from __future__ import annotations

import html
from datetime import datetime
from difflib import SequenceMatcher
from typing import Optional

from .diff_engine import FileDiff, SheetDiff, RowDiff, RowTag
from .reader import CellData


# ---------------------------------------------------------------------------
# スタイルシート
# ---------------------------------------------------------------------------

_CSS = """
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
body {
  font-family: 'Consolas', 'Courier New', 'Meiryo', monospace;
  font-size: 13px;
  background: #f6f8fa;
  color: #24292e;
}
/* ヘッダ */
.site-header {
  background: #24292e;
  color: #fff;
  padding: 12px 24px;
  display: flex;
  align-items: center;
  gap: 24px;
}
.site-header h1 { font-size: 15px; font-weight: bold; }
.site-header .meta { font-size: 12px; color: #8b949e; }
/* ナビゲーション */
.nav {
  background: #fff;
  border-bottom: 1px solid #e1e4e8;
  padding: 8px 24px;
  display: flex;
  align-items: center;
  gap: 8px;
  flex-wrap: wrap;
}
.nav-label { font-size: 12px; color: #57606a; margin-right: 8px; }
.nav a {
  text-decoration: none;
  color: #0969da;
  font-size: 12px;
  padding: 3px 8px;
  border-radius: 4px;
  border: 1px solid #e1e4e8;
  display: inline-flex;
  align-items: center;
  gap: 4px;
}
.nav a:hover { background: #f0f6ff; }
/* バッジ */
.badge {
  display: inline-block;
  padding: 1px 6px;
  border-radius: 10px;
  font-size: 11px;
  font-weight: bold;
}
.badge-modified { background: #ddf4ff; color: #0550ae; }
.badge-added    { background: #e6ffed; color: #1a7f37; }
.badge-deleted  { background: #ffeef0; color: #cf222e; }
.badge-equal    { background: #eee;    color: #666; }
/* サマリ */
.summary {
  margin: 16px 24px 0;
  padding: 10px 16px;
  background: #fff;
  border: 1px solid #e1e4e8;
  border-radius: 6px;
  font-size: 13px;
  color: #57606a;
  display: flex;
  align-items: center;
  gap: 16px;
  flex-wrap: wrap;
}
.summary strong { color: #24292e; }
.summary .matcher-note { font-size: 12px; color: #7d5eaa; }
/* トグルボタン */
.toolbar {
  margin: 12px 24px 0;
  display: flex;
  gap: 8px;
}
.btn {
  padding: 4px 12px;
  border-radius: 6px;
  border: 1px solid #e1e4e8;
  background: #fff;
  color: #24292e;
  font-size: 12px;
  cursor: pointer;
}
.btn:hover { background: #f0f0f0; }
/* シートセクション */
.sheet-section { margin: 16px 24px; }
.sheet-header {
  font-size: 13px;
  font-weight: bold;
  padding: 8px 12px;
  background: #fff;
  border: 1px solid #e1e4e8;
  border-bottom: none;
  border-radius: 6px 6px 0 0;
  display: flex;
  align-items: center;
  gap: 8px;
}
.no-diff-msg {
  background: #fff;
  border: 1px solid #e1e4e8;
  border-radius: 0 0 6px 6px;
  padding: 10px 16px;
  color: #57606a;
  font-size: 12px;
}
/* diffテーブル */
.diff-wrap { overflow-x: auto; }
.diff-table {
  width: 100%;
  border-collapse: collapse;
  background: #fff;
  border: 1px solid #e1e4e8;
  border-radius: 0 0 6px 6px;
  table-layout: auto;
}
.diff-table th {
  background: #f6f8fa;
  padding: 3px 8px;
  border: 1px solid #e1e4e8;
  text-align: center;
  font-size: 11px;
  color: #57606a;
  white-space: nowrap;
}
.diff-table td {
  padding: 2px 8px;
  border-right: 1px solid #eaecef;
  vertical-align: top;
  white-space: pre-wrap;
  word-break: break-all;
  max-width: 280px;
  min-width: 32px;
}
.line-num {
  width: 36px !important;
  min-width: 36px !important;
  max-width: 36px !important;
  text-align: right !important;
  color: #8b949e;
  font-size: 11px;
  user-select: none;
  background: #f6f8fa !important;
  border-right: 2px solid #e1e4e8 !important;
  padding-right: 6px !important;
}
.sep-col {
  width: 6px !important;
  min-width: 6px !important;
  max-width: 6px !important;
  background: #d0d7de !important;
  border: none !important;
  padding: 0 !important;
}
/* 行の色 */
.row-equal   td                   { background: #fff; }
.row-equal   .line-num            { background: #f6f8fa !important; }
.row-deleted td                   { background: #ffeef0; }
.row-deleted .line-num            { background: #ffd7dc !important; }
.row-inserted td                  { background: #e6ffed; }
.row-inserted .line-num           { background: #ccffd8 !important; }
/* MODIFY行: 変更なしセルは白、行番号のみ黄で「変更行」を示す */
.row-modified td                  { background: #fff; }
.row-modified .line-num           { background: #fff5b1 !important; }
/* 変更セル: 背景なし（文字マークだけで差分を示す） */
.cell-old { background: #fff !important; }
.cell-new { background: #fff !important; }
/* 文字レベルのdiff */
.char-del { background: #ffc0c0; color: #8b0000; border-radius: 2px; padding: 0 1px; }
.char-new { background: #2da44e; color: #fff;    border-radius: 2px; padding: 0 1px; }
/* 空セル */
.empty-side td { background: #fafafa !important; }
/* 取り消し線 */
.strike { text-decoration: line-through; }
/* 除外列（比較対象外） */
.cell-excluded { color: #bbb !important; background: #f8f8f8 !important; }
"""

# ---------------------------------------------------------------------------
# JavaScript（EQUALトグル）
# ---------------------------------------------------------------------------

_JS = """
function toggleEqual() {
  var rows = document.querySelectorAll('.row-equal');
  var btn = document.getElementById('btnToggleEqual');
  var showing = btn.getAttribute('data-showing') !== 'false';
  rows.forEach(function(r) { r.style.display = showing ? 'none' : ''; });
  btn.setAttribute('data-showing', showing ? 'false' : 'true');
  btn.textContent = showing ? '全行を表示' : '変更行のみ表示';
}
"""

# ---------------------------------------------------------------------------
# ヘルパー
# ---------------------------------------------------------------------------

def _e(text) -> str:
    return html.escape(str(text) if text is not None else "")


def _render_cell_value(cell: Optional[CellData]) -> str:
    if cell is None or cell.value is None:
        return ""
    text = _e(cell.value)
    if cell.strikethrough:
        text = f'<span class="strike">{text}</span>'
    return text


def _render_cell_pair_diff(
    old_cell: Optional[CellData],
    new_cell: Optional[CellData],
) -> tuple[str, str]:
    """
    変更セルのペアに対して文字レベルの diff HTML を生成し
    (old_html, new_html) を返す。
    """
    old_str = str(old_cell.value) if (old_cell and old_cell.value is not None) else ""
    new_str = str(new_cell.value) if (new_cell and new_cell.value is not None) else ""

    old_parts: list[str] = []
    new_parts: list[str] = []

    for tag, i1, i2, j1, j2 in SequenceMatcher(None, old_str, new_str, autojunk=False).get_opcodes():
        oc = _e(old_str[i1:i2])
        nc = _e(new_str[j1:j2])
        if tag == "equal":
            old_parts.append(oc)
            new_parts.append(nc)
        elif tag == "delete":
            old_parts.append(f'<mark class="char-del">{oc}</mark>')
        elif tag == "insert":
            new_parts.append(f'<mark class="char-new">{nc}</mark>')
        elif tag == "replace":
            old_parts.append(f'<mark class="char-del">{oc}</mark>')
            new_parts.append(f'<mark class="char-new">{nc}</mark>')

    old_html = "".join(old_parts)
    new_html = "".join(new_parts)

    # 取り消し線を外側でラップ
    if old_cell and old_cell.strikethrough:
        old_html = f'<span class="strike">{old_html}</span>'
    if new_cell and new_cell.strikethrough:
        new_html = f'<span class="strike">{new_html}</span>'

    return old_html, new_html


def _render_row(
    row_diff: RowDiff,
    max_cols: int,
    col_filter: Optional[set[int]] = None,
) -> str:
    tag = row_diff.tag
    css_class = {
        RowTag.EQUAL:  "row-equal",
        RowTag.DELETE: "row-deleted",
        RowTag.INSERT: "row-inserted",
        RowTag.MODIFY: "row-modified",
    }[tag]

    # 変更セルを {col_idx: CellDiff} で引けるように
    changed: dict[int, object] = {cd.col_idx: cd for cd in row_diff.cell_diffs}

    old_ln = str(row_diff.old_row.row_idx) if row_diff.old_row else ""
    new_ln = str(row_diff.new_row.row_idx) if row_diff.new_row else ""

    # MODIFY行の変更セルは文字diffを事前計算
    char_diffs: dict[int, tuple[str, str]] = {}
    if tag == RowTag.MODIFY:
        for col_idx, cd in changed.items():
            char_diffs[col_idx] = _render_cell_pair_diff(cd.old_cell, cd.new_cell)

    def _is_excluded(i: int) -> bool:
        return col_filter is not None and i not in col_filter

    # 旧側セル
    if row_diff.old_row:
        cells = row_diff.old_row.cells
        old_tds = ""
        for i in range(max_cols):
            cell = cells[i] if i < len(cells) else None
            if _is_excluded(i):
                old_tds += f'<td class="cell-excluded">{_render_cell_value(cell)}</td>'
            elif i in changed:
                content = char_diffs[i][0] if char_diffs else _render_cell_value(cell)
                old_tds += f'<td class="cell-old">{content}</td>'
            else:
                old_tds += f"<td>{_render_cell_value(cell)}</td>"
    else:
        old_tds = f'<td class="empty-side" colspan="{max_cols}"></td>'

    # 新側セル
    if row_diff.new_row:
        cells = row_diff.new_row.cells
        new_tds = ""
        for i in range(max_cols):
            cell = cells[i] if i < len(cells) else None
            if _is_excluded(i):
                new_tds += f'<td class="cell-excluded">{_render_cell_value(cell)}</td>'
            elif i in changed:
                content = char_diffs[i][1] if char_diffs else _render_cell_value(cell)
                new_tds += f'<td class="cell-new">{content}</td>'
            else:
                new_tds += f"<td>{_render_cell_value(cell)}</td>"
    else:
        new_tds = f'<td class="empty-side" colspan="{max_cols}"></td>'

    return (
        f'<tr class="{css_class}">'
        f'<td class="line-num">{old_ln}</td>{old_tds}'
        f'<td class="sep-col"></td>'
        f'<td class="line-num">{new_ln}</td>{new_tds}'
        f'</tr>'
    )


def _render_sheet(sheet_diff: SheetDiff) -> str:
    status_text = {
        "modified": "変更あり",
        "equal":    "変更なし",
        "added":    "追加",
        "deleted":  "削除",
    }[sheet_diff.status]

    anchor = f'id="sheet-{_e(sheet_diff.name)}"'
    header = (
        f'<div class="sheet-header" {anchor}>'
        f'シート: {_e(sheet_diff.name)}&nbsp;'
        f'<span class="badge badge-{sheet_diff.status}">{status_text}</span>'
        f'</div>'
    )

    if sheet_diff.status == "equal":
        return header + '<div class="no-diff-msg">変更なし</div>'

    # 列ヘッダ行
    col_ths_old = "".join(f'<th>{_e(c)}</th>' for c in sheet_diff.col_letters[:sheet_diff.max_cols])
    col_ths_new = col_ths_old
    header_row = (
        f'<tr>'
        f'<th class="line-num">#</th>{col_ths_old}'
        f'<th class="sep-col"></th>'
        f'<th class="line-num">#</th>{col_ths_new}'
        f'</tr>'
    )

    rows_html = "\n".join(
        _render_row(rd, sheet_diff.max_cols, sheet_diff.col_filter)
        for rd in sheet_diff.row_diffs
    )

    table = (
        f'<div class="diff-wrap">'
        f'<table class="diff-table">{header_row}\n{rows_html}</table>'
        f'</div>'
    )
    return header + table


# ---------------------------------------------------------------------------
# 公開関数
# ---------------------------------------------------------------------------

def render(file_diff: FileDiff) -> str:
    """
    FileDiff を受け取り、自己完結型HTMLを文字列で返す。
    """
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 統計
    cnt = {s: 0 for s in ["modified", "added", "deleted", "equal"]}
    for sd in file_diff.sheet_diffs:
        cnt[sd.status] += 1

    total_modify = sum(1 for sd in file_diff.sheet_diffs for rd in sd.row_diffs if rd.tag == RowTag.MODIFY)
    total_insert = sum(1 for sd in file_diff.sheet_diffs for rd in sd.row_diffs if rd.tag == RowTag.INSERT)
    total_delete = sum(1 for sd in file_diff.sheet_diffs for rd in sd.row_diffs if rd.tag == RowTag.DELETE)

    # ナビゲーション
    nav_items = []
    for sd in file_diff.sheet_diffs:
        nav_items.append(
            f'<a href="#sheet-{_e(sd.name)}">'
            f'{_e(sd.name)}'
            f'<span class="badge badge-{sd.status}">'
            + {"modified": "変更", "equal": "変更なし", "added": "追加", "deleted": "削除"}[sd.status]
            + f'</span></a>'
        )

    # サマリ
    summary_parts = []
    if cnt["modified"]: summary_parts.append(f'<strong>{cnt["modified"]}</strong> シート変更')
    if cnt["added"]:    summary_parts.append(f'<strong>{cnt["added"]}</strong> シート追加')
    if cnt["deleted"]:  summary_parts.append(f'<strong>{cnt["deleted"]}</strong> シート削除')
    if cnt["equal"]:    summary_parts.append(f'<strong>{cnt["equal"]}</strong> シート変更なし')
    summary_parts.append(
        f'行: <strong style="color:#1a7f37">+{total_insert}</strong> '
        f'<strong style="color:#cf222e">−{total_delete}</strong> '
        f'<strong style="color:#9a6700">~{total_modify}</strong>'
    )
    summary_html = "　|　".join(summary_parts)

    matcher_note = ""
    if file_diff.matcher_count > 0:
        matcher_note = (
            f'<span class="matcher-note">'
            f'カスタムマッチャー: {file_diff.matcher_count} 件適用'
            f'</span>'
        )

    # シートコンテンツ
    sheets_html = "\n".join(
        f'<div class="sheet-section">{_render_sheet(sd)}</div>'
        for sd in file_diff.sheet_diffs
    )

    old_name = _e(file_diff.old_path)
    new_name = _e(file_diff.new_path)

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Excel Diff: {old_name} vs {new_name}</title>
<style>{_CSS}</style>
</head>
<body>
<div class="site-header">
  <h1>Excel Diff</h1>
  <span class="meta">{old_name} &rarr; {new_name}</span>
  <span class="meta">比較日時: {now}</span>
</div>
<div class="nav">
  <span class="nav-label">シート:</span>
  {"".join(nav_items)}
</div>
<div class="summary">
  {summary_html}
  {matcher_note}
</div>
<div class="toolbar">
  <button class="btn" id="btnToggleEqual" data-showing="true" onclick="toggleEqual()">変更行のみ表示</button>
</div>
{sheets_html}
<script>{_JS}</script>
</body>
</html>"""
