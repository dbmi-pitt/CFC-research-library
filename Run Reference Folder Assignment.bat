@echo off
setlocal

cd /d "%~dp0"
if not exist "logs" mkdir "logs"
for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd-HHmmss"') do set RUN_STAMP=%%I
set LOG_FILE=logs\reference_folder_assignment_%RUN_STAMP%.log

echo Reference folder assignment started at %DATE% %TIME% > "%LOG_FILE%"
echo Working folder: %CD% >> "%LOG_FILE%"
echo. >> "%LOG_FILE%"

echo Starting reference folder assignment...
echo.

if not exist ".env" (
  echo Missing .env file. Add OPENAI_API_KEY and other keys before running.
  echo Missing .env file. >> "%LOG_FILE%"
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo Missing Python environment. Run Repair CFC Research Setup.bat first.
  echo Missing Python environment. >> "%LOG_FILE%"
  pause
  exit /b 1
)

".venv\Scripts\python.exe" reference_folder_assignment.py --input reports\CFC_All_Categories_Master_Review_Report.xlsx --output reports\CFC_Reference_Folder_Assignment.xlsx >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

echo.
echo Done.
echo Workbook created at:
echo reports\CFC_Reference_Folder_Assignment.xlsx
echo Full log saved to:
echo %LOG_FILE%
echo.
pause
exit /b 0

:failed
echo.
echo Reference folder assignment did not finish. Full log saved to:
echo %LOG_FILE%
echo.
pause
exit /b 1
