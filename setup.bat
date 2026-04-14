@echo off
chcp 65001 > nul
echo ============================================
echo  36協定自動化ツール セットアップ
echo ============================================
echo.

:: Python確認
python --version > nul 2>&1
if errorlevel 1 (
    echo ❌ Pythonがインストールされていません。
    echo.
    echo 以下のURLからPythonをインストールしてください：
    echo https://www.python.org/downloads/
    echo.
    echo ★ インストール時に「Add Python to PATH」に必ずチェックを入れてください ★
    echo.
    echo インストール完了後、このファイルを再度ダブルクリックしてください。
    pause
    start https://www.python.org/downloads/
    exit /b 1
)

python --version
echo ✅ Python確認OK
echo.
echo 必要なライブラリをインストールしています...
echo （初回のみ数分かかります。そのままお待ちください）
echo.

pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo ❌ インストール中にエラーが発生しました。
    echo この画面をスクリーンショットして担当者に送ってください。
    pause
    exit /b 1
)

echo.
echo ============================================
echo  ✅ セットアップ完了！
echo.
echo  次回からは「start.bat」をダブルクリックするだけで
echo  ツールが起動します。
echo ============================================
pause
