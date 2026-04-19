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

echo 実行中の excel-diff-gui.exe を確認中...
taskkill /f /im excel-diff-gui.exe > nul 2>&1
if errorlevel 1 (
    echo  → 起動していません。そのまま続行します。
) else (
    echo  → 強制終了しました。
)
echo.

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
    --collect-all tkinterdnd2 ^
    --clean ^
    gui_main.py

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
