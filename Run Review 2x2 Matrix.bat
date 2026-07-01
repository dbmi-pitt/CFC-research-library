@echo off
setlocal

cd /d "%~dp0"
if not exist "logs" mkdir "logs"
for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd-HHmmss"') do set RUN_STAMP=%%I
set LOG_FILE=logs\review_2x2_matrix_%RUN_STAMP%.log

echo Review 2x2 matrix started at %DATE% %TIME% > "%LOG_FILE%"
echo Working folder: %CD% >> "%LOG_FILE%"
echo. >> "%LOG_FILE%"

echo Starting human-vs-OpenAI 2x2 matrix...
echo.

if not exist "reports\CFC_2017_2022_Review_Comparison.xlsx" (
  echo Missing reports\CFC_2017_2022_Review_Comparison.xlsx.
  echo Run the 2017-2022 review comparison first.
  echo Missing review comparison workbook. >> "%LOG_FILE%"
  pause
  exit /b 1
)

if exist ".uv-cache" set UV_CACHE_DIR=%CD%\.uv-cache
if exist ".uv-python" set UV_PYTHON_INSTALL_DIR=%CD%\.uv-python

uv run python review_2x2_matrix.py --input reports\CFC_2017_2022_Review_Comparison.xlsx --output reports\CFC_2017_2022_2x2_Matrix.xlsx >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

echo.
echo Done.
echo Workbook created at:
echo reports\CFC_2017_2022_2x2_Matrix.xlsx
echo Full log saved to:
echo %LOG_FILE%
echo.
pause
exit /b 0

:failed
echo.
echo Review 2x2 matrix did not finish. Full log saved to:
echo %LOG_FILE%
echo.
pause
exit /b 1
