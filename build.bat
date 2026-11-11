@echo off
chcp 65001 > nul
echo ====================================
echo excel-diff EXEビルド
echo ====================================

:: 仮想環境が存在しなければセットアップを案内
if not exist ".venv" (
    echo エラー: 仮想環境が見つかりません。先に setup.bat を実行してください。
    exit /b 1
)

:: ビルド
echo ビルド中...
uv run pyinstaller ^
    --onefile ^
    --name excel-diff ^
    --clean ^
    main.py

echo.
if exist "dist\excel-diff.exe" (
    echo ビルド成功: dist\excel-diff.exe
    echo.
    echo 動作確認:
    echo   dist\excel-diff.exe tests\fixtures\basic_old.xlsx tests\fixtures\basic_new.xlsx --open
) else (
    echo ビルド失敗
    exit /b 1
)
