@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo === 仮想環境のセットアップ ===

if not exist .venv (
    echo 仮想環境を作成中...
    python -m venv .venv
)

echo 仮想環境を有効化...
call .venv\Scripts\activate.bat

pip show xlwings >nul 2>&1
if errorlevel 1 (
    echo xlwingsをインストール中...
    pip install xlwings --quiet --no-warn-script-location
) else (
    echo xlwings: インストール済み
)

echo.
echo === 画像挿入スクリプトを実行 ===
python insert_image.py

echo.
pause
