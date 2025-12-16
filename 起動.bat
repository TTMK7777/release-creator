@echo off
chcp 65001 >nul
cls
echo.
echo ╔════════════════════════════════════════════════════════════════╗
echo ║     オリコン顧客満足度 プレスリリース作成ツール                ║
echo ║                                                                ║
echo ║     TOPICS出し + 正誤チェック + 表/文章/Word/画像出力          ║
echo ╚════════════════════════════════════════════════════════════════╝
echo.

cd /d "%~dp0streamlit-app"

echo [起動中] Streamlitアプリを開始しています...
echo.
echo ■ ブラウザが自動的に開きます
echo ■ 開かない場合: http://localhost:8501 にアクセス
echo ■ 終了する場合: このウィンドウを閉じる または Ctrl+C
echo.
echo ────────────────────────────────────────────────────────────────
echo.

streamlit run app.py --server.headless true

pause
