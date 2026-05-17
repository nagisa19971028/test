@echo off
echo ===================================
echo  XLS to XLSX Converter - Build
echo ===================================
echo.

echo [1/2] Installing libraries...
pip install xlrd openpyxl pyinstaller
if %errorlevel% neq 0 (
    echo ERROR: Failed to install libraries.
    pause
    exit /b 1
)

echo.
echo [2/2] Building .exe file...
pyinstaller --onefile --windowed --name "XLS-to-XLSX" xls_to_xlsx_gui.py
if %errorlevel% neq 0 (
    echo ERROR: Build failed.
    pause
    exit /b 1
)

echo.
echo ===================================
echo  Done! Check dist\XLS-to-XLSX.exe
echo ===================================
pause
