@echo off
setlocal

cd /d "%~dp0"
if not exist "logs" mkdir "logs"
set LOG_FILE=logs\repair_setup.log

echo CFC Research Library repair started at %DATE% %TIME% > "%LOG_FILE%"
echo Working folder: %CD% >> "%LOG_FILE%"
echo. >> "%LOG_FILE%"

echo Repairing CFC Research Library setup...
echo.

if exist ".venv" (
  set BACKUP_NAME=.venv-replaced-%RANDOM%
  echo Moving old Python environment to %BACKUP_NAME%...
  echo Moving old Python environment to %BACKUP_NAME%... >> "%LOG_FILE%"
  ren ".venv" "%BACKUP_NAME%" >> "%LOG_FILE%" 2>&1
  if errorlevel 1 goto failed
)

set BUNDLED_PY=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe

echo Creating a fresh Python environment...
if exist "%BUNDLED_PY%" (
  "%BUNDLED_PY%" -m venv .venv >> "%LOG_FILE%" 2>&1
) else (
  where py >nul 2>nul
  if not errorlevel 1 (
    py -3 -m venv .venv >> "%LOG_FILE%" 2>&1
  ) else (
    python -m venv .venv >> "%LOG_FILE%" 2>&1
  )
)
if errorlevel 1 goto failed

echo Installing all required libraries...
".venv\Scripts\python.exe" -m pip install --upgrade pip >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
".venv\Scripts\python.exe" -m pip install -r requirements.txt >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

echo.
echo Repair complete.
echo Now double-click:
echo Run CFC Research Update.bat
echo.
echo Full repair log saved to:
echo %LOG_FILE%
echo.
pause
exit /b 0

:failed
echo.
echo Repair did not finish. Full repair log saved to:
echo %LOG_FILE%
echo.
pause
exit /b 1
