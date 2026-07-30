@echo off
REM Launch the Projection Model dashboard + Excel bridge.
setlocal

set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if not exist "%PY%" set "PY=python"

echo Starting Projection Model dashboard...
"%PY%" "%~dp0excel_bridge.py"

if errorlevel 1 (
  echo.
  echo The bridge exited with an error.
  echo If openpyxl is missing, run:  "%PY%" -m pip install openpyxl pywin32
  echo.
  pause
)
endlocal
