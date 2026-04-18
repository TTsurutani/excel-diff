@echo off
chcp 65001 > nul
echo ====================================
echo excel-diff-gui EXEビルド
echo ====================================

if not exist ".venv" (
    echo エラー: 仮想環境が見つかりません。先に setup.bat を実行してください。
    exit /b 1
)

call .venv\Scripts\activate.bat

pip show pyinstaller > nul 2>&1
if errorlevel 1 (
    echo PyInstaller をインストール中...
    pip install pyinstaller
)

echo ビルド中...
pyinstaller ^
    --onefile ^
    --noconsole ^
    --name excel-diff-gui ^
    --clean ^
    excel_diff_gui\__main__.py

echo.
if exist "dist\excel-diff-gui.exe" (
    echo ====================================
    echo ビルド成功: dist\excel-diff-gui.exe
    echo ====================================
    echo.
    echo 配布方法:
    echo   dist\excel-diff-gui.exe を単体で配布できます。
    echo   Python のインストール不要。
    echo.
    echo 自動生成されるファイル（EXE と同じフォルダ）:
    echo   gui_settings.json  ... ウィンドウ設定・各タブの入力値
    echo   patterns.json      ... パターン管理タブで作成したパターン定義
) else (
    echo ビルド失敗
    exit /b 1
)
