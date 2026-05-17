@echo off
echo ===================================
echo  XLS変換ツール .exe ビルドスクリプト
echo ===================================
echo.

echo [1/2] 必要なライブラリをインストール中...
pip install xlrd openpyxl pyinstaller
if %errorlevel% neq 0 (
    echo エラー: ライブラリのインストールに失敗しました
    pause
    exit /b 1
)

echo.
echo [2/2] .exeファイルを作成中...
pyinstaller --onefile --windowed --name "XLS変換ツール" xls_to_xlsx_gui.py
if %errorlevel% neq 0 (
    echo エラー: ビルドに失敗しました
    pause
    exit /b 1
)

echo.
echo ===================================
echo  完了！
echo  dist\XLS変換ツール.exe に作成されました
echo ===================================
pause
