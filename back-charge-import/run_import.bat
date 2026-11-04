@echo off
REM Double-click after the accounting PDFs are in the loading folder.
REM Writes BC_Import_Ready.xlsx there; then click "Import Back Charges" in the tracking workbook.
cd /d "%~dp0"
python bc_import.py
echo.
pause
