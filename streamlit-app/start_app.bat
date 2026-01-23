@echo off
chcp 65001 > nul
title Release-Creator Streamlit App

echo ========================================
echo   Release-Creator 起動中...
echo ========================================
echo.

cd /d "%~dp0"

echo [INFO] 作業ディレクトリ: %CD%
echo [INFO] Streamlit アプリを起動します...
echo.

streamlit run app.py --server.port 8501

if errorlevel 1 (
    echo.
    echo [ERROR] アプリの起動に失敗しました。
    echo   - Python/Streamlit がインストールされているか確認してください
    echo   - 仮想環境が有効化されているか確認してください
    pause
)
