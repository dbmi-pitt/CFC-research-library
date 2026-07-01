@echo off
setlocal

cd /d "%~dp0"
if not exist "logs" mkdir "logs"
for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd-HHmmss"') do set RUN_STAMP=%%I
set LOG_FILE=logs\category_comparison_matrix_%RUN_STAMP%.log

echo Category comparison matrix started at %DATE% %TIME% > "%LOG_FILE%"
echo Working folder: %CD% >> "%LOG_FILE%"
echo. >> "%LOG_FILE%"

echo Starting category comparison matrix...
echo.

if not exist ".venv\Scripts\python.exe" (
  echo Missing Python environment. Run Repair CFC Research Setup.bat first.
  echo Missing Python environment. >> "%LOG_FILE%"
  pause
  exit /b 1
)

if not exist "reports\CFC_Reference_Folder_Assignment.xlsx" (
  echo Missing reports\CFC_Reference_Folder_Assignment.xlsx.
  echo Run Reference Folder Assignment first.
  echo Missing reference assignment workbook. >> "%LOG_FILE%"
  pause
  exit /b 1
)

".venv\Scripts\python.exe" category_comparison_matrix.py --input reports\CFC_Reference_Folder_Assignment.xlsx --output reports\CFC_Category_Comparison_Matrix.xlsx >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

echo.
echo Done.
echo Workbook created at:
echo reports\CFC_Category_Comparison_Matrix.xlsx
echo Full log saved to:
echo %LOG_FILE%
echo.
pause
exit /b 0

:failed
echo.
echo Category comparison matrix did not finish. Full log saved to:
echo %LOG_FILE%
echo.
pause
exit /b 1
