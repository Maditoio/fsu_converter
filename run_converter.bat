@echo off
echo FSU CSV to Excel Converter
echo ========================
echo.

REM Change to the directory where this batch file is located
cd /d "%~dp0"

echo Installing dependencies...
pip install -r requirements.txt
echo.
echo Starting application...
python csv_to_excel_converter.py
pause