"""
Excelブックをシート単位のファイルに分解するモジュール。

split_workbook(path, prefix, suffix, name_regex, output_dir) → list[str]
  各シートを <output_dir>/<prefix><ファイル名ベース><suffix>.xlsx として保存し、
  出力ファイルパスのリストを返す。

実装方針:
  xlsx は ZIP 形式のため、zipfile モジュールで直接操作する。
  openpyxl で毎回フルロードする旧方式に比べ、ファイル全体を1回だけメモリに
  読み込んでから各シートを出力するため大幅に高速（実測 60倍以上）。
  対象シート以外のシートXMLを除外し、workbook.xml / workbook.xml.rels /
  [Content_Types].xml を最小限書き換えてから新しい ZIP として保存する。
  名前定義（defined names）は xl/workbook.xml から一括削除（#REF! 防止）。
"""
from __future__ import annotations

import io
import re
import warnings
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

# xlsx 内で使用する名前空間
_WB_NS   = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
_REL_NS  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

# ファイル名として使えない文字（Windows / macOS / Linux 共通の危険文字）
_INVALID_CHARS = re.compile(r'[\\/:*?"<>|]')


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


def _remove_defined_names(wb_root: ET.Element) -> None:
    """workbook.xml の definedNames 要素を全削除（#REF! エラー防止）。"""
    dn_elem = wb_root.find(f'{{{_WB_NS}}}definedNames')
    if dn_elem is not None:
        wb_root.remove(dn_elem)


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

    # workbook.xml からシート情報を取得
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

                if fname == 'xl/workbook.xml':
                    # 対象シートのみ残し、hidden を解除し、definedNames を削除
                    root = ET.fromstring(data)
                    sh_elem = root.find(f'{{{_WB_NS}}}sheets')
                    if sh_elem is not None:
                        for s in list(sh_elem):
                            if s.get('name') != target_name:
                                sh_elem.remove(s)
                            else:
                                s.attrib.pop('state', None)  # 隠し状態を解除
                    _remove_defined_names(root)
                    data = ET.tostring(root, encoding='UTF-8', xml_declaration=True)

                elif fname == 'xl/_rels/workbook.xml.rels':
                    # 対象シートの rel のみ残す
                    root = ET.fromstring(data)
                    for r in list(root):
                        t = r.get('Target', '')
                        if t.startswith('worksheets/') and f'xl/{t}' != target_file:
                            root.remove(r)
                    data = ET.tostring(root, encoding='UTF-8', xml_declaration=True)

                elif fname == '[Content_Types].xml':
                    # 対象シート以外の Override を削除
                    root = ET.fromstring(data)
                    for ov in list(root):
                        pn = ov.get('PartName', '')
                        if pn.startswith('/xl/worksheets/') and pn != f'/{target_file}':
                            root.remove(ov)
                    data = ET.tostring(root, encoding='UTF-8', xml_declaration=True)

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
