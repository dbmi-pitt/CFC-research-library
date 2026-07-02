@echo off
setlocal

cd /d "%~dp0"
if not exist "logs" mkdir "logs"
if not exist ".uv-cache" mkdir ".uv-cache"
if not exist ".uv-python" mkdir ".uv-python"
set LOG_FILE=logs\last_run.log
set UV_CACHE_DIR=%CD%\.uv-cache
set UV_PYTHON_INSTALL_DIR=%CD%\.uv-python

echo CFC Research Library update started at %DATE% %TIME% > "%LOG_FILE%"
echo Working folder: %CD% >> "%LOG_FILE%"
echo UV cache: %UV_CACHE_DIR% >> "%LOG_FILE%"
echo UV Python install dir: %UV_PYTHON_INSTALL_DIR% >> "%LOG_FILE%"
echo. >> "%LOG_FILE%"

echo Starting CFC Research Library update...
echo.

if not exist ".env" (
  echo Missing .env file.
  echo.
  echo Create a file named .env in this folder using .env.example as the template.
  echo Add ENTREZ_EMAIL, ZOTERO_GROUP_ID, ZOTERO_API_KEY, and OPENAI_API_KEY.
  echo Missing .env file. >> "%LOG_FILE%"
  echo.
  pause
  exit /b 1
)

if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" -c "print('venv ok')" >> "%LOG_FILE%" 2>&1
  if errorlevel 1 (
    echo Existing Python environment is broken. Rebuilding it...
    echo Existing Python environment is broken. Rebuilding it... >> "%LOG_FILE%"
    ren ".venv" ".venv-broken-%RANDOM%" >> "%LOG_FILE%" 2>&1
    if errorlevel 1 goto failed
  )
)

if not exist ".venv\Scripts\python.exe" (
  echo Creating local Python environment...
  set BUNDLED_PY=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe
  if exist "%BUNDLED_PY%" (
    "%BUNDLED_PY%" -m venv .venv >> "%LOG_FILE%" 2>&1
    if errorlevel 1 goto failed
  ) else (
    where py >nul 2>nul
    if not errorlevel 1 (
      py -3 -m venv .venv >> "%LOG_FILE%" 2>&1
      if errorlevel 1 goto failed
    ) else (
      where python >nul 2>nul
      if not errorlevel 1 (
        python -m venv .venv >> "%LOG_FILE%" 2>&1
        if errorlevel 1 goto failed
      ) else (
        where uv >nul 2>nul
        if not errorlevel 1 (
          uv venv >> "%LOG_FILE%" 2>&1
          if errorlevel 1 goto failed
        ) else (
          echo No Python installation was found.
          echo Install Python from https://www.python.org/downloads/ and run this again.
          echo No Python installation was found. >> "%LOG_FILE%"
          pause
          exit /b 1
        )
      )
    )
  )
)

echo Installing or updating required libraries...
where uv >nul 2>nul
if not errorlevel 1 (
  uv pip install -r requirements.txt >> "%LOG_FILE%" 2>&1
  if errorlevel 1 goto failed
) else (
  ".venv\Scripts\python.exe" -m pip install --upgrade pip >> "%LOG_FILE%" 2>&1
  if errorlevel 1 goto failed
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt >> "%LOG_FILE%" 2>&1
  if errorlevel 1 goto failed
)

echo.
echo Running all-category research update...
where uv >nul 2>nul
if not errorlevel 1 (
  uv run python cfc_research_library.py --all-categories --since-year 2025 --output reports\CFC_All_Categories_Master_Review_Report.xlsx >> "%LOG_FILE%" 2>&1
  if errorlevel 1 goto failed
) else (
  ".venv\Scripts\python.exe" cfc_research_library.py --all-categories --since-year 2025 --output reports\CFC_All_Categories_Master_Review_Report.xlsx >> "%LOG_FILE%" 2>&1
  if errorlevel 1 goto failed
)

echo.
echo Done.
echo Workbook created at:
echo reports\CFC_All_Categories_Master_Review_Report.xlsx
echo Full log saved to:
echo %LOG_FILE%
echo.
pause
exit /b 0

:failed
echo.
echo The update did not finish. Read the message above for what needs attention.
echo Full log saved to:
echo %LOG_FILE%
echo.
pause
exit /b 1
