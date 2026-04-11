@echo off
chcp 65001 > nul
echo ====================================
echo excel-diff セットアップ
echo ====================================

:: 仮想環境を作成
if not exist ".venv" (
    echo [1/3] 仮想環境を作成中...
    python -m venv .venv
) else (
    echo [1/3] 仮想環境は既に存在します
)

:: 仮想環境を有効化して依存をインストール
echo [2/3] 依存ライブラリをインストール中...
call .venv\Scripts\activate.bat
pip install --upgrade pip -q
pip install -r requirements.txt

:: テストデータ生成
echo [3/3] テストデータを生成中...
python tests/make_fixtures.py

echo.
echo セットアップ完了。
echo 実行例: python -m excel_diff tests/fixtures/basic_old.xlsx tests/fixtures/basic_new.xlsx --open
echo テスト: python tests/test_diff.py
