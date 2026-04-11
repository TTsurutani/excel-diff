"""
excel-diff CLI エントリポイント。

使い方:
  # ファイル比較
  python -m excel_diff old.xlsx new.xlsx
  python -m excel_diff old.xlsx new.xlsx -o diff.html --open

  # フォルダ一括比較
  python -m excel_diff --dir old_dir/ new_dir/
  python -m excel_diff --dir old_dir/ new_dir/ --output-dir diffs/

  # カスタムマッチャー適用
  python -m excel_diff old.xlsx new.xlsx --matchers matchers.json
"""
from __future__ import annotations

import argparse
import os
import sys
import webbrowser
from pathlib import Path


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="excel-diff",
        description="Excelファイルの差分をHTMLで出力します",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
例:
  excel-diff old.xlsx new.xlsx
  excel-diff old.xlsx new.xlsx -o diff.html --open
  excel-diff --dir old_dir new_dir --output-dir diffs/
  excel-diff old.xlsx new.xlsx --matchers matchers.json
""",
    )
    p.add_argument("old_file", nargs="?", help="比較元ファイル (.xlsx)")
    p.add_argument("new_file", nargs="?", help="比較先ファイル (.xlsx)")
    p.add_argument("-o", "--output", metavar="PATH",
                   help="出力HTMLパス（省略時: <新ファイル名>_diff.html）")
    p.add_argument("--dir", nargs=2, metavar=("OLD_DIR", "NEW_DIR"),
                   help="フォルダ一括比較")
    p.add_argument("--output-dir", metavar="DIR",
                   help="フォルダ比較時の出力先ディレクトリ")
    p.add_argument("--sheet", metavar="NAME",
                   help="指定シートのみ比較（省略時: 全シート）")
    p.add_argument("--strikethrough", action="store_true",
                   help="取り消し線の有無も差分として扱う")
    p.add_argument("--matchers", metavar="FILE",
                   help="カスタムマッチャー設定JSONファイル")
    p.add_argument("--open", action="store_true",
                   help="生成後にブラウザで自動オープン")
    return p


def _default_output_path(new_file: str) -> str:
    stem = Path(new_file).stem
    parent = Path(new_file).parent
    return str(parent / f"{stem}_diff.html")


def _run_file_diff(args: argparse.Namespace) -> None:
    from .reader import read_workbook
    from .diff_engine import diff_files
    from .html_renderer import render
    from .matcher import load_matchers

    old_path = args.old_file
    new_path = args.new_file

    if not os.path.isfile(old_path):
        print(f"エラー: ファイルが見つかりません: {old_path}", file=sys.stderr)
        sys.exit(1)
    if not os.path.isfile(new_path):
        print(f"エラー: ファイルが見つかりません: {new_path}", file=sys.stderr)
        sys.exit(1)

    matchers = []
    if args.matchers:
        if not os.path.isfile(args.matchers):
            print(f"エラー: マッチャー設定ファイルが見つかりません: {args.matchers}", file=sys.stderr)
            sys.exit(1)
        matchers = load_matchers(args.matchers)
        print(f"カスタムマッチャー: {len(matchers)} 件ロード")

    print(f"読み込み中: {old_path}")
    old_sheets = read_workbook(old_path, args.strikethrough, args.sheet)
    print(f"読み込み中: {new_path}")
    new_sheets = read_workbook(new_path, args.strikethrough, args.sheet)

    print("差分計算中...")
    file_diff = diff_files(
        old_sheets, new_sheets,
        old_path, new_path,
        matchers=matchers,
        include_strike=args.strikethrough,
    )

    output_path = args.output or _default_output_path(new_path)
    html_content = render(file_diff)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    if file_diff.has_differences:
        print(f"差分あり → {output_path}")
    else:
        print(f"差分なし → {output_path}")

    if args.open:
        webbrowser.open(Path(output_path).resolve().as_uri())


def _run_dir_diff(args: argparse.Namespace) -> None:
    from .reader import read_workbook
    from .diff_engine import diff_files
    from .html_renderer import render
    from .matcher import load_matchers

    old_dir, new_dir = args.dir

    if not os.path.isdir(old_dir):
        print(f"エラー: ディレクトリが見つかりません: {old_dir}", file=sys.stderr)
        sys.exit(1)
    if not os.path.isdir(new_dir):
        print(f"エラー: ディレクトリが見つかりません: {new_dir}", file=sys.stderr)
        sys.exit(1)

    matchers = []
    if args.matchers:
        if not os.path.isfile(args.matchers):
            print(f"エラー: マッチャー設定ファイルが見つかりません: {args.matchers}", file=sys.stderr)
            sys.exit(1)
        matchers = load_matchers(args.matchers)
        print(f"カスタムマッチャー: {len(matchers)} 件ロード")

    # 両ディレクトリの .xlsx ファイル名を収集
    old_files = {
        f for f in os.listdir(old_dir)
        if f.lower().endswith(".xlsx") and not f.startswith("~$")
    }
    new_files = {
        f for f in os.listdir(new_dir)
        if f.lower().endswith(".xlsx") and not f.startswith("~$")
    }
    all_files = sorted(old_files | new_files)

    if not all_files:
        print("比較対象の .xlsx ファイルが見つかりませんでした。")
        return

    # 出力ディレクトリ
    if args.output_dir:
        out_dir = args.output_dir
    else:
        old_name = Path(old_dir).name
        new_name = Path(new_dir).name
        out_dir = f"{old_name}_vs_{new_name}"

    os.makedirs(out_dir, exist_ok=True)
    print(f"出力先: {out_dir}/")

    results = []
    for fname in all_files:
        old_path = os.path.join(old_dir, fname)
        new_path = os.path.join(new_dir, fname)

        if not os.path.isfile(old_path):
            print(f"  [追加] {fname}")
            old_sheets = {}
        else:
            old_sheets = read_workbook(old_path, args.strikethrough, args.sheet)

        if not os.path.isfile(new_path):
            print(f"  [削除] {fname}")
            new_sheets = {}
        else:
            new_sheets = read_workbook(new_path, args.strikethrough, args.sheet)

        file_diff = diff_files(
            old_sheets, new_sheets,
            old_path if os.path.isfile(old_path) else f"(なし)/{fname}",
            new_path if os.path.isfile(new_path) else f"(なし)/{fname}",
            matchers=matchers,
            include_strike=args.strikethrough,
        )

        stem = Path(fname).stem
        out_path = os.path.join(out_dir, f"{stem}_diff.html")
        html_content = render(file_diff)
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(html_content)

        status = "差分あり" if file_diff.has_differences else "変更なし"
        print(f"  [{status}] {fname} → {out_path}")
        results.append((fname, file_diff.has_differences, out_path))

    diff_count = sum(1 for _, d, _ in results if d)
    print(f"\n完了: {len(results)} ファイル比較、うち {diff_count} ファイルに差分")

    if args.open and results:
        # 差分があった最初のファイルを開く
        for _, has_diff, out_path in results:
            if has_diff:
                webbrowser.open(Path(out_path).resolve().as_uri())
                break


def main() -> None:
    parser = _build_parser()
    args = parser.parse_args()

    if args.dir:
        _run_dir_diff(args)
    elif args.old_file and args.new_file:
        _run_file_diff(args)
    else:
        parser.print_help()
        sys.exit(1)


if __name__ == "__main__":
    main()
