@echo off
chcp 65001 > nul 2>&1
setlocal EnableDelayedExpansion

:: ====================================================================
:: Release Creator - デバッグモード起動
:: ====================================================================
:: コンソールに Streamlit のログを全て表示します。
:: エラー発生時の原因調査用。
:: ====================================================================

title Release Creator [DEBUG]

:: パス解決
set "BASE_DIR=%~dp0"
if "!BASE_DIR:~-1!"=="\" set "BASE_DIR=!BASE_DIR:~0,-1!"

set "PYTHON_EXE=!BASE_DIR!\python\python.exe"
set "APP_DIR=!BASE_DIR!\app"

:: 存在チェック
if not exist "!PYTHON_EXE!" (
    echo [エラー] Python が見つかりません: !PYTHON_EXE!
    pause
    exit /b 1
)

if not exist "!APP_DIR!\app.py" (
    echo [エラー] app.py が見つかりません: !APP_DIR!\app.py
    pause
    exit /b 1
)

:: 環境情報表示
echo ============================================================
echo   Release Creator - デバッグモード
echo ============================================================
echo.
echo   Python: !PYTHON_EXE!
echo   App:    !APP_DIR!\app.py
echo.

:: Python バージョン表示
"!PYTHON_EXE!" --version 2>&1
echo.

:: pip list（インストール済みパッケージ）
echo --- インストール済みパッケージ ---
"!PYTHON_EXE!" -m pip list --format=columns 2>&1
echo.

:: 環境変数設定
set "ENABLE_UPLOAD_FEATURE=true"

:: ポート競合チェック
set "DEBUG_PORT=8501"
netstat -an | findstr ":!DEBUG_PORT! " | findstr "LISTENING" > nul 2>&1
if not errorlevel 1 (
    echo [警告] ポート !DEBUG_PORT! は既に使用中です。
    echo ReleaseCreator.bat が起動中の可能性があります。
    echo 先に stop.bat で停止するか、別のポートで起動します。
    echo.
    :: 空きポートを探す
    for /L %%p in (8502,1,8510) do (
        if "!DEBUG_PORT!"=="8501" (
            netstat -an | findstr ":%%p " | findstr "LISTENING" > nul 2>&1
            if errorlevel 1 set "DEBUG_PORT=%%p"
        )
    )
    if "!DEBUG_PORT!"=="8501" (
        echo [エラー] 利用可能なポートが見つかりません。
        pause
        exit /b 1
    )
    echo ポート !DEBUG_PORT! を使用します。
    echo.
)

:: Streamlit 起動（フォアグラウンド、全ログ表示）
echo --- Streamlit 起動 (port: !DEBUG_PORT!) ---
echo.
"!PYTHON_EXE!" -m streamlit run "!APP_DIR!\app.py" ^
    --server.port !DEBUG_PORT! ^
    --server.headless true ^
    --server.fileWatcherType none ^
    --browser.gatherUsageStats false ^
    --logger.level debug

echo.
echo --- Streamlit 終了 (exit code: !errorlevel!) ---
pause
