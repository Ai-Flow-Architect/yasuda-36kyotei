@echo off
chcp 65001 > nul
echo ========================================
echo  36協定自動化ツール
echo ========================================
echo.

REM === 設定 ===
set EXCEL_FILE=input.xlsx
set OUTPUT_DIR=output

REM Excelファイル確認
if not exist "%EXCEL_FILE%" (
    echo エラー: %EXCEL_FILE% が見つかりません。
    echo 同じフォルダに input.xlsx を配置してください。
    echo.
    pause
    exit /b 1
)

echo 入力ファイル: %EXCEL_FILE%
echo 出力先: %OUTPUT_DIR%\
echo.
echo 処理を開始します...
echo.

python main.py "%EXCEL_FILE%" --output-dir "%OUTPUT_DIR%" --from-name "安田" --from-org "朝日事務所"

echo.
echo ========================================
echo  処理完了！ %OUTPUT_DIR% フォルダを確認してください。
echo ========================================
pause
