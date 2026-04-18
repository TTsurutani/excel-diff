"""
Excelブックをシート単位のファイルに分解するモジュール。

split_workbook(path, prefix, suffix, name_regex, output_dir) → list[str]
  各シートを <output_dir>/<prefix><ファイル名ベース><suffix>.xlsx として保存し、
  出力ファイルパスのリストを返す。

実装方針:
  xlsx は ZIP 形式のため、zipfile モジュールで直接操作する。
  openpyxl で毎回フルロードする旧方式と異なり、ファイル全体を1回だけメモリに
  読み込んでから各シートを出力するため大幅に高速（実測 50倍以上）。

  XML の書き換えは ET.tostring() による再シリアライズを避け、正規表現による
  バイト列の直接編集で行う。これにより名前空間宣言が壊れる問題を防ぐ。
  （ET.tostring() は名前空間プレフィックスを変更するため Excel が読めなくなる）

  修正対象ファイル（3点のみ）:
    xl/workbook.xml        ← 対象シートのみ残す・hidden 解除・definedNames 削除
    xl/_rels/workbook.xml.rels ← 対象シートの Relationship のみ残す
    [Content_Types].xml    ← 対象シートの Override のみ残す
"""
from __future__ import annotations

import io
import re
import warnings
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

# ファイル名として使えない文字（Windows / macOS / Linux 共通の危険文字）
_INVALID_CHARS = re.compile(r'[\\/:*?"<>|]')

# xlsx 内で使用する名前空間（メタ情報取得用）
_WB_NS  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def _safe_filename(sheet_name: str) -> str:
    """シート名をファイル名として安全な文字列に変換する。"""
    return _INVALID_CHARS.sub("_", sheet_name)


def _apply_name_regex(sheet_name: str, pattern: re.Pattern[str]) -> str:
    """
    正規表現の第1キャプチャグループにマッチした部分を返す。
    マッチしない場合はシート名全体にフォールバックして警告を出す。
    """
    m = pattern.search(sheet_name)
    if m and m.lastindex and m.lastindex >= 1:
        return m.group(1)
    warnings.warn(
        f"--name-regex がシート '{sheet_name}' にマッチしませんでした。シート名をそのまま使用します。",
        stacklevel=3,
    )
    return sheet_name


# ── XML バイト列の直接編集（名前空間を壊さないための正規表現方式） ──────────

def _patch_workbook_xml(data: bytes, target_rid: str) -> bytes:
    """
    workbook.xml から対象シート以外の <sheet> 要素を削除し、
    対象シートの state 属性（hidden）を除去し、<definedNames> を全削除する。
    r:id 属性で判定するためシート名のXMLエスケープを気にしなくてよい。
    """
    text = data.decode('utf-8')

    def _replace_sheet(m: re.Match) -> str:
        elem = m.group(0)
        rid_m = re.search(r':id="([^"]*)"', elem)
        if not rid_m:
            return elem
        if rid_m.group(1) != target_rid:
            return ''                              # 他シートは削除
        # state="hidden" 等を除去（unhide）
        return re.sub(r'\s+state="[^"]*"', '', elem)

    # <sheet ... /> 要素を処理
    text = re.sub(r'<sheet\b[^>]*/>', _replace_sheet, text)
    # <definedNames> ブロックを全削除（#REF! 防止）
    text = re.sub(r'<definedNames\b[^>]*>.*?</definedNames>', '', text, flags=re.DOTALL)
    text = re.sub(r'<definedNames\s*/>', '', text)

    return text.encode('utf-8')


def _patch_workbook_rels(data: bytes, target_rid: str) -> bytes:
    """
    workbook.xml.rels から、worksheets/ を Target に持つ Relationship のうち
    対象シート以外のものを削除する。
    """
    text = data.decode('utf-8')

    def _replace_rel(m: re.Match) -> str:
        elem = m.group(0)
        id_m  = re.search(r'\bId="([^"]*)"',     elem)
        tgt_m = re.search(r'\bTarget="([^"]*)"', elem)
        if not id_m or not tgt_m:
            return elem
        if tgt_m.group(1).startswith('worksheets/') and id_m.group(1) != target_rid:
            return ''  # 他シートの Relationship を削除
        return elem

    text = re.sub(r'<Relationship\b[^>]*/>', _replace_rel, text)
    return text.encode('utf-8')


def _patch_content_types(data: bytes, target_file: str) -> bytes:
    """
    [Content_Types].xml から /xl/worksheets/ 以下の Override のうち
    対象シート以外のものを削除する。
    """
    text = data.decode('utf-8')
    target_part = '/' + target_file   # /xl/worksheets/sheetN.xml

    def _replace_override(m: re.Match) -> str:
        elem = m.group(0)
        pn_m = re.search(r'\bPartName="([^"]*)"', elem)
        if not pn_m:
            return elem
        pn = pn_m.group(1)
        if pn.startswith('/xl/worksheets/') and pn != target_part:
            return ''  # 他シートの Override を削除
        return elem

    text = re.sub(r'<Override\b[^>]*/>', _replace_override, text)
    return text.encode('utf-8')


# ── メイン処理 ────────────────────────────────────────────────────────────────

def split_workbook(
    path: str,
    prefix: str = "",
    suffix: str = "",
    name_regex: str | None = None,
    output_dir: str | None = None,
) -> list[str]:
    """
    ブックを1シート1ファイルに分解して保存する。

    Parameters
    ----------
    path       : 入力Excelファイルパス (.xlsx)
    prefix     : 出力ファイル名の前置文字列
    suffix     : 出力ファイル名の後置文字列（拡張子の前）
    name_regex : ファイル名ベース抽出用正規表現（第1キャプチャグループを使用）。
                 Noneの場合はシート名をそのまま使用。
    output_dir : 出力先ディレクトリ（Noneの場合はブックと同じフォルダ）

    Returns
    -------
    出力ファイルパスのリスト（シート順）
    """
    compiled_regex: re.Pattern[str] | None = None
    if name_regex:
        compiled_regex = re.compile(name_regex)
        if compiled_regex.groups < 1:
            raise ValueError(
                f"--name-regex にはキャプチャグループ () が1つ以上必要です: {name_regex!r}"
            )

    src_path = Path(path)
    out_dir = Path(output_dir) if output_dir else src_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    # ── xlsx 全ファイルを一括メモリ読み込み ──────────────────────────────
    file_cache: dict[str, tuple[zipfile.ZipInfo, bytes]] = {}
    with zipfile.ZipFile(path, 'r') as zf:
        for item in zf.infolist():
            file_cache[item.filename] = (item, zf.read(item.filename))

    # workbook.xml からシート情報を取得（ETは読み取りのみに使用）
    wb_data = file_cache['xl/workbook.xml'][1]
    wb_root = ET.fromstring(wb_data)
    sheets_elem = wb_root.find(f'{{{_WB_NS}}}sheets')
    if sheets_elem is None:
        raise ValueError("xl/workbook.xml にシート情報が見つかりません")

    sheet_info: list[tuple[str, str, str]] = []  # (name, rid, state)
    for s in sheets_elem:
        sheet_info.append((
            s.get('name', ''),
            s.get(f'{{{_REL_NS}}}id', ''),
            s.get('state', 'visible'),
        ))

    # workbook.xml.rels から rid → ファイルパス マッピング
    rels_data = file_cache['xl/_rels/workbook.xml.rels'][1]
    rels_root = ET.fromstring(rels_data)
    rid_to_target: dict[str, str] = {
        r.get('Id', ''): r.get('Target', '') for r in rels_root
    }

    # ── シートごとに新しい xlsx を生成 ──────────────────────────────────
    output_paths: list[str] = []

    for target_name, target_rid, _state in sheet_info:
        target_rel  = rid_to_target.get(target_rid, '')   # worksheets/sheetN.xml
        target_file = f'xl/{target_rel}'                  # xl/worksheets/sheetN.xml

        # 除外するシートファイル（対象以外）
        other_files: set[str] = {
            f'xl/{rid_to_target[rid]}'
            for _, rid, _ in sheet_info
            if rid != target_rid and rid in rid_to_target
        }

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as dst:
            for fname, (item, data) in file_cache.items():
                # 対象外シートの XML はスキップ
                if fname in other_files:
                    continue

                # 3ファイルのみバイト列を直接編集（再シリアライズしない）
                if fname == 'xl/workbook.xml':
                    data = _patch_workbook_xml(data, target_rid)
                elif fname == 'xl/_rels/workbook.xml.rels':
                    data = _patch_workbook_rels(data, target_rid)
                elif fname == '[Content_Types].xml':
                    data = _patch_content_types(data, target_file)

                dst.writestr(item, data)

        # ファイル名を決定して保存
        if compiled_regex is not None:
            name_base = _apply_name_regex(target_name, compiled_regex)
        else:
            name_base = target_name

        safe_name    = _safe_filename(name_base)
        out_filename = f"{prefix}{safe_name}{suffix}.xlsx"
        out_path     = out_dir / out_filename
        out_path.write_bytes(buf.getvalue())
        output_paths.append(str(out_path))

    return output_paths
