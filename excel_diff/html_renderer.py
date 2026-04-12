"""
HTML出力モジュール。
差分結果を自己完結型のHTMLファイルとして生成する（CSS/JS内包）。
WinMerge スタイル: 左右分割パネル + 同期スクロール
"""
from __future__ import annotations

import html
import re
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
  font-size: 12px;
  background: #f6f8fa;
  color: #24292e;
  height: 100vh;
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

/* ─── トップバー ─── */
.top-bar {
  background: #24292e;
  color: #fff;
  padding: 6px 16px;
  display: flex;
  align-items: center;
  gap: 16px;
  flex-shrink: 0;
}
.top-bar h1 { font-size: 13px; font-weight: bold; }
.top-bar .meta { font-size: 11px; color: #8b949e; }

/* ─── インフォバー（サマリ + ツールバー） ─── */
.info-bar {
  background: #fff;
  border-bottom: 1px solid #e1e4e8;
  padding: 4px 16px;
  display: flex;
  align-items: center;
  gap: 12px;
  flex-wrap: wrap;
  flex-shrink: 0;
  font-size: 12px;
}
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
.btn {
  padding: 3px 10px;
  border-radius: 5px;
  border: 1px solid #e1e4e8;
  background: #fff;
  color: #24292e;
  font-size: 11px;
  cursor: pointer;
}
.btn:hover { background: #f0f0f0; }

/* ─── シートコンテナ（残り全高さ） ─── */
.sheets-container {
  flex: 1;
  min-height: 0;
  overflow-y: auto;
  overflow-x: hidden;
}

/* ─── シートセクション ─── */
.sheet-section {
  display: flex;
  flex-direction: column;
  border-bottom: 2px solid #ccc;
}
.sheet-header {
  background: #f6f8fa;
  border-bottom: 1px solid #e1e4e8;
  padding: 4px 12px;
  font-size: 12px;
  font-weight: bold;
  display: flex;
  align-items: center;
  gap: 8px;
  flex-shrink: 0;
}
.no-diff-msg {
  padding: 8px 16px;
  color: #57606a;
  font-size: 12px;
  background: #fff;
}

/* ─── 左右分割パネル ─── */
.sheet-panels {
  display: grid;
  grid-template-columns: 50% 50%;
  grid-template-rows: max-content 1fr;
  min-height: 300px;
  max-height: calc(100vh - 80px);
}

/* ファイルパスタイトル行 */
.file-title {
  font-weight: bold;
  color: #fff;
  background: linear-gradient(mediumblue, darkblue);
  padding: 3px 10px;
  font-size: 11px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  border-right: 1px solid #003;
}
.file-title:last-of-type { border-right: none; }

/* スクロール可能なパネル本体 */
.panel {
  overflow: auto;
  border-right: 1px solid #d0d7de;
  min-height: 0;
}
.panel:last-child { border-right: none; }

/* ─── diff テーブル ─── */
.diff-table {
  width: max-content;
  min-width: 100%;
  border-collapse: collapse;
  background: #fff;
}

/* ヘッダ行を各パネル内でスティッキー */
.diff-table thead tr th {
  position: sticky;
  top: 0;
  z-index: 10;
  background: #f6f8fa;
  padding: 2px 8px;
  border: 1px solid #e1e4e8;
  text-align: center;
  font-size: 11px;
  color: #57606a;
  white-space: nowrap;
}

.diff-table td {
  padding: 2px 8px;
  border-right: 1px solid #eaecef;
  border-bottom: 1px solid #f0f0f0;
  vertical-align: top;
  white-space: pre;
  min-width: 40px;
}

/* 行番号列：横方向スティッキー */
.line-num {
  position: sticky;
  left: 0;
  z-index: 5;
  width: 36px;
  min-width: 36px;
  text-align: right;
  color: #8b949e;
  font-size: 11px;
  user-select: none;
  background: #f6f8fa;
  border-right: 2px solid #e1e4e8 !important;
  padding-right: 6px !important;
}
.diff-table thead tr th.line-num { z-index: 15; }

/* ─── 行の色 ─── */
.row-equal td         { background: #fff; }
.row-equal .line-num  { background: #f6f8fa; }

.row-deleted td         { background: #ffeef0; }
.row-deleted .line-num  { background: #ffd7dc; }

.row-inserted td         { background: #e6ffed; }
.row-inserted .line-num  { background: #ccffd8; }

/* MODIFY行：行全体は白、変更セルのみ黄色 */
.row-modified td         { background: #fff; }
.row-modified .line-num  { background: #fff5b1; }  /* 行番号だけ色をつけて変更行であることを示す */
.cell-modified           { background: #fff8c5 !important; }

/* 対応のない行（削除/追加の空き側） */
.row-phantom td         { background: #f0f0f0 !important; }
.row-phantom .line-num  { background: #e8e8e8 !important; }

/* ─── 上下レイアウト ─── */
.sheet-panels.layout-vertical {
  grid-template-columns: 100%;
  grid-template-rows: max-content 1fr max-content 1fr;
  max-height: none;
}
.sheet-panels.layout-vertical .panel {
  max-height: 45vh;
  border-right: none;
  border-bottom: 1px solid #d0d7de;
}

/* ─── 文字レベル diff ─── */
.char-del { background: #ffc0c0; color: #8b0000; border-radius: 2px; padding: 0 1px; }
.char-new { background: #2da44e; color: #fff;    border-radius: 2px; padding: 0 1px; }

/* 比較対象外列 */
.cell-excluded { color: #bbb; background: #f8f8f8 !important; }

/* 取り消し線 */
.strike { text-decoration: line-through; }

/* 列固定（JS で動的に付与） */
.col-frozen {
  position: sticky;
  z-index: 4;
}
/* 固定列の背景を確保（行色に合わせて継承） */
.row-equal    .col-frozen { background: #fff; }
.row-deleted  .col-frozen { background: #ffeef0; }
.row-inserted .col-frozen { background: #e6ffed; }
.row-modified .col-frozen { background: #fff; }
.row-phantom  .col-frozen { background: #f0f0f0; }
.col-frozen.cell-modified { background: #fff8c5 !important; }
.col-frozen.cell-excluded { background: #f8f8f8 !important; }
"""

# ---------------------------------------------------------------------------
# JavaScript
# ---------------------------------------------------------------------------

_JS = """
// ── スクロール同期（垂直は常時、水平はトグル） ──────────────────────────
var _hSyncEnabled = true;

function toggleHSync() {
  _hSyncEnabled = !_hSyncEnabled;
  var btn = document.getElementById('btnHSync');
  btn.textContent = _hSyncEnabled ? '水平同期 ON' : '水平同期 OFF';
  btn.style.background = _hSyncEnabled ? '#ddf4ff' : '';
}

document.querySelectorAll('.panel-pair').forEach(function(pair) {
  var panels = pair.querySelectorAll('.panel');
  if (panels.length < 2) return;
  var left = panels[0], right = panels[1];
  var syncing = false;
  left.addEventListener('scroll', function() {
    if (syncing) return; syncing = true;
    right.scrollTop = left.scrollTop;                          // 垂直は常時同期
    if (_hSyncEnabled) right.scrollLeft = left.scrollLeft;    // 水平はトグル制御
    syncing = false;
  });
  right.addEventListener('scroll', function() {
    if (syncing) return; syncing = true;
    left.scrollTop = right.scrollTop;                         // 垂直は常時同期
    if (_hSyncEnabled) left.scrollLeft = right.scrollLeft;   // 水平はトグル制御
    syncing = false;
  });
});

// ── 行高さの均一化（DELETE/INSERT の phantom 行と実行の高さを揃える） ────
function equalizeRowHeights() {
  document.querySelectorAll('.panel-pair').forEach(function(pair) {
    var panels = Array.from(pair.querySelectorAll('.panel'));
    if (panels.length < 2) return;
    var lMap = {}, rMap = {};
    panels[0].querySelectorAll('tbody tr[data-row]').forEach(function(tr) {
      lMap[tr.getAttribute('data-row')] = tr;
    });
    panels[1].querySelectorAll('tbody tr[data-row]').forEach(function(tr) {
      rMap[tr.getAttribute('data-row')] = tr;
    });
    // いったんリセット
    Object.values(lMap).concat(Object.values(rMap)).forEach(function(tr) {
      tr.style.height = '';
    });
    // 高さを揃える
    Object.keys(lMap).forEach(function(k) {
      var ltr = lMap[k], rtr = rMap[k];
      if (!ltr || !rtr) return;
      var lh = ltr.offsetHeight, rh = rtr.offsetHeight;
      if (lh !== rh) {
        var maxH = Math.max(lh, rh) + 'px';
        ltr.style.height = maxH;
        rtr.style.height = maxH;
      }
    });
  });
}

// ── EQUAL 行の表示トグル ─────────────────────────────────────────────────
function toggleEqual() {
  var rows = document.querySelectorAll('.row-equal, .row-phantom');
  var btn = document.getElementById('btnToggleEqual');
  var showing = btn.getAttribute('data-showing') !== 'false';
  rows.forEach(function(r) { r.style.display = showing ? 'none' : ''; });
  btn.setAttribute('data-showing', showing ? 'false' : 'true');
  btn.textContent = showing ? '全行を表示' : '変更行のみ表示';
  equalizeRowHeights();
}

// ── 左右 ⇄ 上下 レイアウト切替 ─────────────────────────────────────────
function toggleLayout() {
  var panels = document.querySelectorAll('.sheet-panels');
  var btn = document.getElementById('btnToggleLayout');
  var isVertical = btn.getAttribute('data-layout') === 'vertical';
  panels.forEach(function(p) {
    if (isVertical) {
      p.classList.remove('layout-vertical');
    } else {
      p.classList.add('layout-vertical');
    }
  });
  btn.setAttribute('data-layout', isVertical ? 'horizontal' : 'vertical');
  btn.textContent = isVertical ? '上下表示に切替' : '左右表示に切替';
  setTimeout(equalizeRowHeights, 50);
}

// ── 先頭N列を固定 ────────────────────────────────────────────────────────
var _colFreezeActive = false;
var FREEZE_COLS = 3;  // 固定するデータ列数（A, B, C）

function toggleFreezeColumns() {
  _colFreezeActive = !_colFreezeActive;
  var btn = document.getElementById('btnFreezeCol');
  if (_colFreezeActive) {
    applyFreezeColumns(FREEZE_COLS);
    btn.textContent = '列固定 解除';
    btn.style.background = '#ddf4ff';
  } else {
    removeFreezeColumns();
    btn.textContent = '先頭' + FREEZE_COLS + '列を固定';
    btn.style.background = '';
  }
}

function applyFreezeColumns(numCols) {
  document.querySelectorAll('.panel-pair .diff-table').forEach(function(table) {
    // 1行目のセル幅から各列の left オフセットを計算
    var firstRow = table.querySelector('thead tr') || table.querySelector('tr');
    if (!firstRow) return;
    var firstCells = Array.from(firstRow.querySelectorAll('td, th'));
    // col 0 = line-num（CSS で既に sticky、left:0）
    // col 1〜numCols をJSで sticky にする
    var lefts = [0];
    for (var i = 0; i < numCols && i < firstCells.length - 1; i++) {
      lefts.push(lefts[i] + firstCells[i].getBoundingClientRect().width);
    }
    table.querySelectorAll('tr').forEach(function(tr) {
      var cells = Array.from(tr.querySelectorAll('td, th'));
      for (var i = 1; i <= numCols && i < cells.length; i++) {
        cells[i].classList.add('col-frozen');
        cells[i].style.left = lefts[i] + 'px';
      }
    });
  });
}

function removeFreezeColumns() {
  document.querySelectorAll('.col-frozen').forEach(function(cell) {
    cell.classList.remove('col-frozen');
    cell.style.left = '';
  });
}

// ── ページ読み込み後に行高さを均一化 ────────────────────────────────────
window.addEventListener('load', function() {
  equalizeRowHeights();
});
"""

# ---------------------------------------------------------------------------
# ヘルパー
# ---------------------------------------------------------------------------

def _e(text) -> str:
    return html.escape(str(text) if text is not None else "")


# CR の実体文字（\r）と、テキストとして埋め込まれた表現（x000D / x000d など）を除去
_CTRL_RE = re.compile(r'\r|[xX]000[dD]')

def _strip_ctrl(v) -> str:
    """制御コード（CR 等）を除去して文字列化する。
    - 実際の CR 文字（\\r / \\x0D）
    - テキストとして埋め込まれた 'x000D' / 'x000d' 表現
    の両方を取り除く。
    """
    if v is None:
        return ""
    return _CTRL_RE.sub('', str(v))


def _render_cell_value(cell: Optional[CellData]) -> str:
    if cell is None or cell.value is None:
        return ""
    text = _e(_strip_ctrl(cell.value))
    if cell.strikethrough:
        text = f'<span class="strike">{text}</span>'
    return text


def _render_cell_pair_diff(
    old_cell: Optional[CellData],
    new_cell: Optional[CellData],
) -> tuple[str, str]:
    """変更セルのペアに対して文字レベル diff HTML を生成し (old_html, new_html) を返す。"""
    old_str = _strip_ctrl(old_cell.value) if old_cell else ""
    new_str = _strip_ctrl(new_cell.value) if new_cell else ""

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

    if old_cell and old_cell.strikethrough:
        old_html = f'<span class="strike">{old_html}</span>'
    if new_cell and new_cell.strikethrough:
        new_html = f'<span class="strike">{new_html}</span>'

    return old_html, new_html


# ---------------------------------------------------------------------------
# 行レンダリング（左パネル用 / 右パネル用を別々に生成）
# ---------------------------------------------------------------------------

def _phantom_tr(max_cols: int, row_idx: int = 0) -> str:
    """対応行のない側に挿入する空行"""
    empty_tds = "".join(f"<td></td>" for _ in range(max_cols))
    return f'<tr class="row-phantom" data-row="{row_idx}"><td class="line-num"></td>{empty_tds}</tr>'


def _build_tds(
    row_data,
    max_cols: int,
    html_map: dict[int, str],
    col_filter: Optional[set[int]],
) -> str:
    if row_data is None:
        return "".join(f"<td></td>" for _ in range(max_cols))
    cells = row_data.cells
    tds = f'<td class="line-num">{row_data.row_idx}</td>'
    for i in range(max_cols):
        cell = cells[i] if i < len(cells) else None
        if col_filter is not None and i not in col_filter:
            tds += f'<td class="cell-excluded">{_render_cell_value(cell)}</td>'
        elif i in html_map:
            tds += f'<td class="cell-modified">{html_map[i]}</td>'
        else:
            tds += f"<td>{_render_cell_value(cell)}</td>"
    return tds


def _render_row_pair(
    row_diff: RowDiff,
    max_cols: int,
    col_filter: Optional[set[int]] = None,
    row_idx: int = 0,
) -> tuple[str, str]:
    """1つの RowDiff から (左パネル用 <tr>, 右パネル用 <tr>) を返す。"""
    tag = row_diff.tag
    changed = {cd.col_idx: cd for cd in row_diff.cell_diffs}
    ri = f' data-row="{row_idx}"'

    # MODIFY: 文字 diff を事前計算
    char_diffs: dict[int, tuple[str, str]] = {}
    if tag == RowTag.MODIFY:
        for col_idx, cd in changed.items():
            char_diffs[col_idx] = _render_cell_pair_diff(cd.old_cell, cd.new_cell)

    if tag == RowTag.EQUAL:
        tds_l = _build_tds(row_diff.old_row, max_cols, {}, col_filter)
        tds_r = _build_tds(row_diff.new_row, max_cols, {}, col_filter)
        left_tr  = f'<tr class="row-equal"{ri}>{tds_l}</tr>'
        right_tr = f'<tr class="row-equal"{ri}>{tds_r}</tr>'

    elif tag == RowTag.DELETE:
        tds_l = _build_tds(row_diff.old_row, max_cols, {}, col_filter)
        left_tr  = f'<tr class="row-deleted"{ri}>{tds_l}</tr>'
        right_tr = _phantom_tr(max_cols, row_idx)

    elif tag == RowTag.INSERT:
        tds_r = _build_tds(row_diff.new_row, max_cols, {}, col_filter)
        left_tr  = _phantom_tr(max_cols, row_idx)
        right_tr = f'<tr class="row-inserted"{ri}>{tds_r}</tr>'

    else:  # MODIFY
        old_map = {i: v[0] for i, v in char_diffs.items()}
        new_map = {i: v[1] for i, v in char_diffs.items()}
        tds_l = _build_tds(row_diff.old_row, max_cols, old_map, col_filter)
        tds_r = _build_tds(row_diff.new_row, max_cols, new_map, col_filter)
        left_tr  = f'<tr class="row-modified"{ri}>{tds_l}</tr>'
        right_tr = f'<tr class="row-modified"{ri}>{tds_r}</tr>'

    return left_tr, right_tr


# ---------------------------------------------------------------------------
# シートレンダリング
# ---------------------------------------------------------------------------

def _render_sheet(sheet_diff: SheetDiff, old_path: str, new_path: str) -> str:
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

    max_cols = sheet_diff.max_cols
    col_filter = sheet_diff.col_filter
    col_ths = "".join(
        f'<th>{_e(c)}</th>'
        for c in sheet_diff.col_letters[:max_cols]
    )
    header_tr = f'<tr><th class="line-num">#</th>{col_ths}</tr>'

    left_rows: list[str] = []
    right_rows: list[str] = []
    for i, rd in enumerate(sheet_diff.row_diffs):
        lt, rt = _render_row_pair(rd, max_cols, col_filter, i)
        left_rows.append(lt)
        right_rows.append(rt)

    def make_table(rows: list[str]) -> str:
        tbody = "\n".join(rows)
        return (
            f'<table class="diff-table">'
            f'<thead>{header_tr}</thead>'
            f'<tbody>{tbody}</tbody>'
            f'</table>'
        )

    old_label = _e(old_path)
    new_label = _e(new_path)

    panels_html = (
        f'<div class="sheet-panels panel-pair">'
        f'<div class="file-title">{old_label}</div>'
        f'<div class="file-title">{new_label}</div>'
        f'<div class="panel">{make_table(left_rows)}</div>'
        f'<div class="panel">{make_table(right_rows)}</div>'
        f'</div>'
    )

    return f'<div class="sheet-section">{header}{panels_html}</div>'


# ---------------------------------------------------------------------------
# 公開関数
# ---------------------------------------------------------------------------

def render(file_diff: FileDiff) -> str:
    """FileDiff を受け取り、自己完結型HTMLを文字列で返す。"""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cnt = {s: 0 for s in ["modified", "added", "deleted", "equal"]}
    for sd in file_diff.sheet_diffs:
        cnt[sd.status] += 1

    total_modify = sum(1 for sd in file_diff.sheet_diffs for rd in sd.row_diffs if rd.tag == RowTag.MODIFY)
    total_insert = sum(1 for sd in file_diff.sheet_diffs for rd in sd.row_diffs if rd.tag == RowTag.INSERT)
    total_delete = sum(1 for sd in file_diff.sheet_diffs for rd in sd.row_diffs if rd.tag == RowTag.DELETE)

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
            f'<span style="color:#7d5eaa; font-size:11px">'
            f'カスタムマッチャー {file_diff.matcher_count} 件適用</span>'
        )

    # ナビゲーション（シート一覧）
    nav_items = "".join(
        f'<a href="#sheet-{_e(sd.name)}" style="text-decoration:none; color:#0969da; '
        f'font-size:11px; padding:2px 8px; border:1px solid #e1e4e8; border-radius:4px;">'
        f'{_e(sd.name)}'
        f'<span class="badge badge-{sd.status}" style="margin-left:4px;">'
        + {"modified": "変更", "equal": "変更なし", "added": "追加", "deleted": "削除"}[sd.status]
        + f'</span></a>'
        for sd in file_diff.sheet_diffs
    )

    old_path = file_diff.old_path
    new_path = file_diff.new_path

    sheets_html = "\n".join(
        _render_sheet(sd, old_path, new_path)
        for sd in file_diff.sheet_diffs
    )

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Excel Diff</title>
<style>{_CSS}</style>
</head>
<body>
<div class="top-bar">
  <h1>Excel Diff</h1>
  <span class="meta">{_e(old_path)}</span>
  <span class="meta">→</span>
  <span class="meta">{_e(new_path)}</span>
  <span class="meta" style="margin-left:auto">{now}</span>
</div>
<div class="info-bar">
  {summary_html}
  {matcher_note}
  <span style="margin-left:auto; display:flex; gap:6px;">
    {nav_items}
    <button class="btn" id="btnToggleEqual" data-showing="true" onclick="toggleEqual()">変更行のみ表示</button>
    <button class="btn" id="btnToggleLayout" data-layout="horizontal" onclick="toggleLayout()">上下表示に切替</button>
    <button class="btn" id="btnFreezeCol" onclick="toggleFreezeColumns()">先頭3列を固定</button>
    <button class="btn" id="btnHSync" style="background:#ddf4ff" onclick="toggleHSync()">水平同期 ON</button>
  </span>
</div>
<div class="sheets-container">
{sheets_html}
</div>
<script>{_JS}</script>
</body>
</html>"""
