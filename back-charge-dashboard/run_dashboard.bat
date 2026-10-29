@echo off
REM Double-click to rebuild the back charge dashboard from the P: drive workbook.
REM Outputs are written next to this file, then the dashboard opens in your browser.
cd /d "%~dp0"
python backcharge_dashboard.py --out "%~dp0"
if errorlevel 1 (
  echo.
  echo The dashboard was not built. See the message above.
  pause
  exit /b 1
)
start "" "%~dp0back_charge_dashboard.html"
