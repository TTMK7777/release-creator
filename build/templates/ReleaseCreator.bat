@echo off
chcp 65001 > nul 2>&1
setlocal EnableDelayedExpansion

:: ====================================================================
:: Release Creator - 統合ランチャー
:: ====================================================================
:: 初回: 共有フォルダ → ローカルコピー + デスクトップショートカット
:: 2回目以降: VERSION比較 → 差分更新 → 起動
:: ====================================================================

title Release Creator

:: --------------- 設定 ---------------
set "LOCAL_DIR=%LOCALAPPDATA%\ReleaseCreator"
set "PID_FILE=%TEMP%\release-creator.pid"
set "START_PORT=8501"
set "END_PORT=8510"

:: --------------- パス解決 ---------------
:: このBAT自体の場所
set "BAT_DIR=%~dp0"
:: 末尾の \ を除去
if "!BAT_DIR:~-1!"=="\" set "BAT_DIR=!BAT_DIR:~0,-1!"

:: ローカルインストールかどうか判定
set "IS_LOCAL=0"
if /i "!BAT_DIR!"=="!LOCAL_DIR!" set "IS_LOCAL=1"

:: ソースパス（共有フォルダ）の解決
set "SOURCE_DIR="
if "!IS_LOCAL!"=="1" (
    :: ローカルから起動 → .source_path から共有フォルダパスを読み取り
    if exist "!LOCAL_DIR!\.source_path" (
        set /p SOURCE_DIR=<"!LOCAL_DIR!\.source_path"
        :: 末尾の改行・空白を除去
        for /f "tokens=*" %%a in ("!SOURCE_DIR!") do set "SOURCE_DIR=%%a"
        :: 空ファイル対策
        if "!SOURCE_DIR!"=="" (
            echo [注意] .source_path が空です。オフラインモードで起動します。
        )
        :: パス検証: VERSION.txt が存在するディレクトリのみ信頼する
        if defined SOURCE_DIR (
            if not exist "!SOURCE_DIR!\VERSION.txt" (
                echo [注意] 共有フォルダにアクセスできません: !SOURCE_DIR!
                echo オフラインモードで起動します。
                set "SOURCE_DIR="
            )
        )
    )
) else (
    :: 共有フォルダから起動
    set "SOURCE_DIR=!BAT_DIR!"
)

:: ====================================================================
:: メインフロー
:: ====================================================================

:: ローカルコピーが存在するか？
if not exist "!LOCAL_DIR!\python\python.exe" (
    echo.
    echo ============================================================
    echo   Release Creator - 初回セットアップ
    echo ============================================================
    echo.
    goto :install
)

:: 更新チェック（ソースが利用可能な場合のみ）
if defined SOURCE_DIR (
    if exist "!SOURCE_DIR!\VERSION.txt" (
        call :check_update
    )
)

goto :launch

:: ====================================================================
:: 初回インストール
:: ====================================================================
:install
echo [1/3] ファイルをコピー中...（少々お待ちください）
echo.

:: ローカルディレクトリ作成
if not exist "!LOCAL_DIR!" mkdir "!LOCAL_DIR!"

:: python/ をコピー
echo   python/ をコピー中...
robocopy "!BAT_DIR!\python" "!LOCAL_DIR!\python" /E /NJH /NJS /NDL /NFL /NC /NS > nul 2>&1
if !errorlevel! geq 8 (
    echo [エラー] python/ のコピーに失敗しました ^(code: !errorlevel!^)
    echo ネットワーク接続とディスク空き容量を確認してください。
    pause
    exit /b 1
)

:: app/ をコピー
echo   app/ をコピー中...
robocopy "!BAT_DIR!\app" "!LOCAL_DIR!\app" /E /NJH /NJS /NDL /NFL /NC /NS > nul 2>&1
if !errorlevel! geq 8 (
    echo [エラー] app/ のコピーに失敗しました ^(code: !errorlevel!^)
    echo ネットワーク接続とディスク空き容量を確認してください。
    pause
    exit /b 1
)

:: ルートファイルをコピー（BAT, VERSION, README）
echo   設定ファイルをコピー中...
for %%F in (ReleaseCreator.bat start-debug.bat stop.bat VERSION.txt README.txt) do (
    if exist "!BAT_DIR!\%%F" copy /Y "!BAT_DIR!\%%F" "!LOCAL_DIR!\%%F" > nul 2>&1
)

:: ソースパスを記録（更新チェック用）
>"!LOCAL_DIR!\.source_path" echo !BAT_DIR!

echo.
echo [2/3] コピー完了

:: デスクトップショートカット作成
echo [3/3] デスクトップショートカットを作成中...
set "SHORTCUT=%USERPROFILE%\Desktop\Release Creator.lnk"
powershell -NoProfile -Command ^
    "$ws = New-Object -ComObject WScript.Shell; ^
     $sc = $ws.CreateShortcut($args[0]); ^
     $sc.TargetPath = $args[1]; ^
     $sc.WorkingDirectory = $args[2]; ^
     $sc.Description = 'Release Creator を起動'; ^
     $sc.Save()" ^
     "!SHORTCUT!" "!LOCAL_DIR!\ReleaseCreator.bat" "!LOCAL_DIR!"
if exist "!SHORTCUT!" (
    echo   デスクトップに「Release Creator」ショートカットを作成しました
) else (
    echo   [注意] ショートカット作成に失敗しました
    echo   !LOCAL_DIR!\ReleaseCreator.bat を直接実行してください
)

echo.
echo ============================================================
echo   セットアップ完了！
echo ============================================================
echo.

goto :launch

:: ====================================================================
:: 更新チェック
:: ====================================================================
:check_update
:: ソースの VERSION を読み取り
set "SOURCE_VER="
if exist "!SOURCE_DIR!\VERSION.txt" (
    set /p SOURCE_VER=<"!SOURCE_DIR!\VERSION.txt"
    for /f "tokens=*" %%a in ("!SOURCE_VER!") do set "SOURCE_VER=%%a"
)

:: ローカルの VERSION を読み取り
set "LOCAL_VER="
if exist "!LOCAL_DIR!\VERSION.txt" (
    set /p LOCAL_VER=<"!LOCAL_DIR!\VERSION.txt"
    for /f "tokens=*" %%a in ("!LOCAL_VER!") do set "LOCAL_VER=%%a"
)

:: バージョン比較
if "!SOURCE_VER!"=="!LOCAL_VER!" (
    goto :eof
)

:: ダウングレード防止（文字列比較、通常の x.y.z 形式で安全）
if "!SOURCE_VER!" LSS "!LOCAL_VER!" (
    echo [注意] ソースバージョン ^(!SOURCE_VER!^) がローカル ^(!LOCAL_VER!^) より古いため更新をスキップ
    goto :eof
)

echo.
echo ============================================================
echo   アップデートを検出: !LOCAL_VER! → !SOURCE_VER!
echo ============================================================
echo.

:: メジャーバージョン比較（python/ 更新が必要か判定）
for /f "tokens=1 delims=." %%a in ("!SOURCE_VER!") do set "SRC_MAJOR=%%a"
for /f "tokens=1 delims=." %%a in ("!LOCAL_VER!") do set "LOC_MAJOR=%%a"

if not "!SRC_MAJOR!"=="!LOC_MAJOR!" (
    echo [1/2] python/ を更新中（メジャーバージョン変更）...
    robocopy "!SOURCE_DIR!\python" "!LOCAL_DIR!\python" /MIR /NJH /NJS /NDL /NFL /NC /NS > nul 2>&1
    if !errorlevel! geq 8 (
        echo [警告] python/ の更新でエラーが発生しました ^(code: !errorlevel!^)
    ) else (
        echo   → python/ 更新完了
    )
) else (
    echo [1/2] python/ は変更なし（スキップ）
)

:: app/ は常に更新（/E = 追加コピー + 上書き）
echo [2/2] app/ を更新中...
robocopy "!SOURCE_DIR!\app" "!LOCAL_DIR!\app" /E /NJH /NJS /NDL /NFL /NC /NS > nul 2>&1
if !errorlevel! geq 8 (
    echo [警告] app/ の更新でエラーが発生しました ^(code: !errorlevel!^)
)

:: ルートファイル更新
for %%F in (ReleaseCreator.bat start-debug.bat stop.bat VERSION.txt README.txt) do (
    if exist "!SOURCE_DIR!\%%F" copy /Y "!SOURCE_DIR!\%%F" "!LOCAL_DIR!\%%F" > nul 2>&1
)

echo   → 更新完了 (v!SOURCE_VER!)
echo.
goto :eof

:: ====================================================================
:: アプリ起動
:: ====================================================================
:launch

:: 作業ディレクトリをローカルに設定
set "WORK_DIR=!LOCAL_DIR!"
if not exist "!WORK_DIR!\python\python.exe" (
    :: フォールバック: 現在のディレクトリから起動
    set "WORK_DIR=!BAT_DIR!"
)

set "PYTHON_EXE=!WORK_DIR!\python\python.exe"
set "APP_DIR=!WORK_DIR!\app"

:: python.exe の存在確認
if not exist "!PYTHON_EXE!" (
    echo.
    echo [エラー] Python が見つかりません: !PYTHON_EXE!
    echo セットアップが正しく完了していない可能性があります。
    echo.
    pause
    exit /b 1
)

:: app.py の存在確認
if not exist "!APP_DIR!\app.py" (
    echo.
    echo [エラー] アプリケーションが見つかりません: !APP_DIR!\app.py
    echo.
    pause
    exit /b 1
)

:: 環境変数設定
set "ENABLE_UPLOAD_FEATURE=true"
:: v8.2: 未公表ローカルランキングデータの共有フォルダパス
:: 共有フォルダを使用する場合は以下をコメント解除してパスを設定
:: set "LOCAL_DATA_PATH=\\server\oricon\release-creator\local_rankings"

:: --------------- 同時起動ガード ---------------
if exist "!PID_FILE!" (
    set /p EXISTING_PID=<"!PID_FILE!"
    for /f "tokens=*" %%a in ("!EXISTING_PID!") do set "EXISTING_PID=%%a"
    if defined EXISTING_PID (
        tasklist /fi "PID eq !EXISTING_PID!" 2>nul | findstr /i "python.exe" > nul 2>&1
        if not errorlevel 1 (
            echo.
            echo Release Creator は既に起動しています ^(PID: !EXISTING_PID!^)
            echo 既存のインスタンスのブラウザを開きます。
            echo.
            :: PIDからポートを逆引き
            set "EXISTING_PORT="
            for /L %%p in (!START_PORT!,1,!END_PORT!) do (
                if not defined EXISTING_PORT (
                    netstat -ano 2>nul | findstr ":%%p " | findstr "LISTENING" | findstr "!EXISTING_PID!" > nul 2>&1
                    if not errorlevel 1 set "EXISTING_PORT=%%p"
                )
            )
            if defined EXISTING_PORT (
                start "" "http://localhost:!EXISTING_PORT!"
                echo URL: http://localhost:!EXISTING_PORT!
            ) else (
                echo [注意] ポートを特定できませんでした。stop.bat で停止してから再起動してください。
            )
            echo.
            pause
            exit /b 0
        )
    )
    :: PIDプロセスが存在しない → 古いPIDファイルを削除
    del "!PID_FILE!" 2>nul
)

:: --------------- 空きポート検索 ---------------
set "PORT=0"
for /L %%p in (!START_PORT!,1,!END_PORT!) do (
    if "!PORT!"=="0" (
        netstat -an | findstr ":%%p " | findstr "LISTENING" > nul 2>&1
        if errorlevel 1 (
            set "PORT=%%p"
        )
    )
)

if "!PORT!"=="0" (
    echo.
    echo [エラー] 利用可能なポートが見つかりません（!START_PORT!-!END_PORT!）
    echo 他のアプリケーションがポートを使用している可能性があります。
    echo stop.bat を実行してから再試行してください。
    echo.
    pause
    exit /b 1
)

:: --------------- Streamlit 起動 ---------------
echo.
echo ============================================================
echo   Release Creator を起動しています...
echo   ポート: !PORT!
echo ============================================================
echo.

:: Streamlit をバックグラウンドで起動（python -m streamlit を使用）
start /b "" "!PYTHON_EXE!" -m streamlit run "!APP_DIR!\app.py" ^
    --server.port !PORT! ^
    --server.headless true ^
    --server.fileWatcherType none ^
    --browser.gatherUsageStats false

:: --------------- ポートチェックループ ---------------
echo アプリケーションの起動を待機中...
set "RETRIES=0"
set "MAX_RETRIES=30"

:wait_loop
if !RETRIES! geq !MAX_RETRIES! (
    echo.
    echo [警告] 起動タイムアウト（!MAX_RETRIES!秒）
    echo ブラウザを手動で開いてください: http://localhost:!PORT!
    goto :open_browser
)

:: ポート応答チェック（PowerShell で TCP 接続テスト）
powershell -NoProfile -Command ^
    "try { $c = New-Object Net.Sockets.TcpClient; $c.Connect('localhost', !PORT!); $c.Close(); exit 0 } catch { exit 1 }" > nul 2>&1

if !errorlevel! equ 0 (
    echo 起動完了！
    :: ポートからPIDを特定して記録（最初のマッチのみ: IPv4/IPv6 二重書込み防止）
    set "PORT_PID="
    for /f "tokens=5" %%p in ('netstat -ano ^| findstr ":!PORT! " ^| findstr "LISTENING"') do (
        if not defined PORT_PID set "PORT_PID=%%p"
    )
    if defined PORT_PID (
        >"!PID_FILE!" echo !PORT_PID!
    )
    goto :open_browser
)

set /a RETRIES+=1
timeout /t 1 /nobreak > nul
goto :wait_loop

:: --------------- ブラウザオープン ---------------
:open_browser
echo.
echo ブラウザを開いています... http://localhost:!PORT!
start "" "http://localhost:!PORT!"

echo.
echo ============================================================
echo   Release Creator が起動しました
echo   URL: http://localhost:!PORT!
echo.
echo   終了するには:
echo     - このウィンドウを閉じる
echo     - または stop.bat を実行
echo ============================================================
echo.
echo ウィンドウを閉じるとアプリも終了します。
pause > nul
