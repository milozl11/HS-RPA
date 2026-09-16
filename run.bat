@echo off
REM =====================================================================
REM  FT6AUTO - Portable launcher (no Python installation required)
REM  Uses the bundled python-embed\ interpreter.
REM =====================================================================
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

REM ---- Locate the embedded Python ----
set "PYTHON_EXE=%~dp0python-embed\python.exe"
set "PLAYWRIGHT_BROWSERS_PATH=%~dp0playwright-browsers"
set "CHROMIUM_EXE=%PLAYWRIGHT_BROWSERS_PATH%\chromium-1134\chrome-win\chrome.exe"

REM ---- Auto-setup: download Python + deps + browser if missing ----
set "NEED_SETUP=0"
if not exist "%PYTHON_EXE%" set "NEED_SETUP=1"
if not exist "%CHROMIUM_EXE%" set "NEED_SETUP=1"
if "%NEED_SETUP%"=="0" (
  "%PYTHON_EXE%" -c "import flask, openpyxl, waitress, playwright" >nul 2>nul
  if errorlevel 1 set "NEED_SETUP=1"
)

if "%NEED_SETUP%"=="1" (
  echo.
  echo  First run detected - running automatic setup...
  echo  This requires internet access and will take a few minutes.
  echo.
  if not exist "%~dp0setup.bat" (
    echo ERROR: setup.bat not found next to this script.
    pause
    exit /b 1
  )
  call "%~dp0setup.bat" --no-pause
  if errorlevel 1 (
    echo.
    echo ERROR: Setup failed. Check the output above.
    pause
    exit /b 1
  )
  REM Re-check after setup
  if not exist "%PYTHON_EXE%" (
    echo ERROR: Setup completed but Python was not installed correctly.
    pause
    exit /b 1
  )
)

echo.
echo ===================================================
echo   FT6AUTO - SAP GTS Customs Description Uploader
echo   Starting server on http://localhost:5000
echo ===================================================
echo.

REM ---- Start the server before opening the UI ----
start "FT6AUTO Server" /min "%PYTHON_EXE%" server\app.py
set "SERVER_READY=0"
for /l %%N in (1,1,30) do (
  if "!SERVER_READY!"=="0" (
    "%PYTHON_EXE%" -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:5000/', timeout=1)" >nul 2>nul
    if not errorlevel 1 set "SERVER_READY=1"
    if "!SERVER_READY!"=="0" ping 127.0.0.1 -n 2 >nul
  )
)
if "!SERVER_READY!"=="0" (
  echo.
  echo ERROR: Serverul nu a devenit disponibil in 30 de secunde.
  echo Verifica output-ul procesului Python si portul 5000.
  exit /b 1
)

start "" http://localhost:5000
echo Serverul este disponibil. Browserul a fost deschis.
exit /b 0
