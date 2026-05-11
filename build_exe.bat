@echo off
REM Build standalone executable for tfs_report\azure_pr_report.py using PyInstaller
REM Ensure you have PyInstaller installed: pip install pyinstaller
set SCRIPT=azure_pr_report.py
set NAME=azure_pr_report_tfs

pyinstaller --onefile --noconsole --name %NAME% %SCRIPT%
if %ERRORLEVEL% neq 0 (
  echo Build failed.
  pause
  exit /b %ERRORLEVEL%
)
echo Build succeeded. Executable is in the dist\%NAME%\%NAME%.exe
pause
