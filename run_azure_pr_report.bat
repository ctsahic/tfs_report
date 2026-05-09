@echo off
REM Batch wrapper to run azure_pr_report.py from repository root.
REM - Activates .venv or venv if present, then runs the script with any passed arguments.

cd /d "%~dp0"

if exist ".venv\Scripts\activate.bat" (
  call ".venv\Scripts\activate.bat"
) else if exist "venv\Scripts\activate.bat" (
  call "venv\Scripts\activate.bat"
)

REM Run the Python script with any arguments passed to this .bat file
python azure_pr_report.py %*

echo.
echo Finished. Press any key to close...
pause >nul
