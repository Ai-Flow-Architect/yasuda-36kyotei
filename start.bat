@echo off
chcp 65001 > /dev/null
echo ============================================
echo  36協定自動化ツール 起動中...
echo ============================================
echo.
echo ブラウザが自動で開きます。
echo 開かない場合は http://localhost:8501 を開いてください。
echo.
echo ★ このウィンドウは閉じないでください ★
echo   （閉じるとツールが停止します）
echo.

cd /d %~dp0
streamlit run app.py --server.headless false

pause
