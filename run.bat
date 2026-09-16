@echo off
REM =====================================================================
REM  FT6AUTO - Portable launcher (no Python installation required)
REM  Uses the bundled python-embed\ interpreter.
REM =====================================================================
setlocal EnableExtensions EnableDelayedExpansion
title FT6AUTO Launcher
set "APP_ROOT=%~dp0"
echo Application folder: %APP_ROOT%

REM ---- Locate the embedded Python ----
set "PYTHON_EXE=%APP_ROOT%python-embed\python.exe"
set "PLAYWRIGHT_BROWSERS_PATH=%APP_ROOT%playwright-browsers"
set "CHROMIUM_EXE=%PLAYWRIGHT_BROWSERS_PATH%\chromium-1134\chrome-win\chrome.exe"

REM ---- Validate the complete offline bundle ----
if not exist "%PYTHON_EXE%" (
  echo.
  echo ERROR: Offline bundle is incomplete. Missing:
  echo   %PYTHON_EXE%
  echo Extract the complete ZIP before running. No download is required.
  pause
  exit /b 1
)
if not exist "%CHROMIUM_EXE%" (
  echo.
  echo ERROR: Offline bundle is incomplete. Missing:
  echo   %CHROMIUM_EXE%
  echo Check whether antivirus quarantined chrome.exe, then extract again.
  pause
  exit /b 1
)
if not exist "%APP_ROOT%server\app.py" (
  echo.
  echo ERROR: Offline bundle is incomplete. Missing server\app.py.
  pause
  exit /b 1
)
"%PYTHON_EXE%" -c "import flask, openpyxl, waitress, playwright" >nul 2>nul
if errorlevel 1 (
  echo.
  echo ERROR: Embedded Python packages are incomplete or blocked.
  echo No installation is required; extract a fresh copy of the full ZIP.
  pause
  exit /b 1
)

echo.
echo ===================================================
echo   FT6AUTO - SAP GTS Customs Description Uploader
echo   Starting server on http://localhost:5000
echo ===================================================
echo.

REM ---- Start the server before opening the UI ----
start "FT6AUTO Server" /min /d "%APP_ROOT%" "%PYTHON_EXE%" "%APP_ROOT%server\app.py"
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
