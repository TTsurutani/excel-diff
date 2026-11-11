@echo off
chcp 65001 > nul
echo ====================================
echo excel-diff セットアップ
echo ====================================

:: uv が pyproject.toml と uv.lock から .venv を作る（Python のバージョンも uv が選ぶ）
echo [1/2] 仮想環境と依存ライブラリ...
uv sync
if errorlevel 1 exit /b 1

:: テストデータ生成
echo [2/2] テストデータを生成中...
uv run python tests/make_fixtures.py

echo.
echo セットアップ完了。
echo 実行例: uv run python -m excel_diff tests/fixtures/basic_old.xlsx tests/fixtures/basic_new.xlsx --open
echo テスト: uv run python tests/test_diff.py
