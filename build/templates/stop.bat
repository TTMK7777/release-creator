@echo off
chcp 65001 > nul 2>&1
setlocal EnableDelayedExpansion

:: ====================================================================
:: Release Creator - 停止スクリプト
:: ====================================================================
:: PID ファイルベースで該当プロセスのみを停止します。
:: 他の Python プロセスには影響しません。
:: ====================================================================

title Release Creator - 停止

set "PID_FILE=%TEMP%\release-creator.pid"

echo.
echo Release Creator を停止しています...
echo.

:: PID ファイルの存在チェック
if not exist "!PID_FILE!" (
    echo PID ファイルが見つかりません。
    echo アプリが起動していないか、既に停止しています。
    echo.
    echo --- フォールバック: ポート 8501-8510 で動作中のプロセスを確認 ---
    echo.
    set "FOUND_PORT=0"
    set "FALLBACK_PID="
    for /L %%p in (8501,1,8510) do (
        netstat -ano | findstr ":%%p " | findstr "LISTENING" > nul 2>&1
        if not errorlevel 1 (
            echo   ポート %%p で動作中:
            netstat -ano | findstr ":%%p " | findstr "LISTENING"
            set "FOUND_PORT=1"
            :: PIDを取得（最初のマッチのみ）
            for /f "tokens=5" %%q in ('netstat -ano ^| findstr ":%%p " ^| findstr "LISTENING"') do (
                if not defined FALLBACK_PID set "FALLBACK_PID=%%q"
            )
        )
    )
    if "!FOUND_PORT!"=="0" (
        echo ポート 8501-8510 でリッスンしているプロセスはありません。
    ) else if defined FALLBACK_PID (
        echo.
        set /p CONFIRM="PID !FALLBACK_PID! を強制終了しますか？ (Y/N): "
        if /i "!CONFIRM!"=="Y" (
            taskkill /PID !FALLBACK_PID! /T /F > nul 2>&1
            if !errorlevel! equ 0 (
                echo 停止しました。
            ) else (
                echo [警告] 停止に失敗しました。タスクマネージャーから終了してください。
            )
        )
    )
    echo.
    pause
    exit /b 0
)

:: PID 読み取り
set /p PID=<"!PID_FILE!"
for /f "tokens=*" %%a in ("!PID!") do set "PID=%%a"

:: プロセスが存在するか確認（python.exe であることも検証）
tasklist /fi "PID eq !PID!" 2>nul | findstr /i "python.exe" > nul 2>&1
if errorlevel 1 (
    echo PID !PID! のプロセスは既に終了しています。
    del "!PID_FILE!" 2>nul
    echo PID ファイルを削除しました。
    echo.
    pause
    exit /b 0
)

:: プロセスツリーごと終了（/T: 子プロセス含む）
echo PID !PID! を終了しています...
taskkill /PID !PID! /T /F > nul 2>&1
if !errorlevel! equ 0 (
    echo 停止しました。
) else (
    echo [警告] プロセスの停止に失敗しました。
    echo 管理者権限で再試行するか、タスクマネージャーから終了してください。
)

:: PID ファイル削除
del "!PID_FILE!" 2>nul

echo.
echo Release Creator を停止しました。
echo.
pause
