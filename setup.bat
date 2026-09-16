@echo off
REM =====================================================================
REM  FT6AUTO - One-time setup for the bundled Python environment
REM
REM  Run this ONCE on a machine that has internet access.
REM  No Python installation is needed - it downloads everything.
REM  It creates and populates the python-embed\ folder
REM  with all dependencies. After this, run.bat works on any PC
REM  fully offline.
REM =====================================================================
setlocal
cd /d "%~dp0"

set "PY_EMBED=%~dp0python-embed"
set "PY_EXE=%PY_EMBED%\python.exe"
set "PLAYWRIGHT_BROWSERS_PATH=%~dp0playwright-browsers"

echo.
echo ===================================================
echo   FT6AUTO Setup - Preparing portable environment
echo ===================================================
echo.

REM ---- Step 1: Check if python-embed exists ----
if not exist "%PY_EXE%" (
  echo [1/4] Downloading Python Embedded ...
  REM Try curl first (available on Windows 10+)
  set "PY_ZIP=%~dp0_python-embed.zip"
  curl.exe -L -o "%PY_ZIP%" "https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip" --progress-bar 2>nul
  if not exist "%PY_ZIP%" (
    echo curl failed, trying PowerShell ...
    powershell -NoProfile -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip' -OutFile '%PY_ZIP%' -UseBasicParsing"
  )
  if not exist "%PY_ZIP%" (
    echo.
    echo ERROR: Could not download Python Embedded.
    echo Download manually from:
    echo   https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip
    echo Extract into python-embed\ next to this script.
    pause
    exit /b 1
  )
  echo Extracting...
  powershell -NoProfile -Command "Expand-Archive -Path '%PY_ZIP%' -DestinationPath '%PY_EMBED%' -Force"
  del "%PY_ZIP%" 2>nul
  if not exist "%PY_EXE%" (
    echo ERROR: Extraction failed.
    pause
    exit /b 1
  )
  REM Enable site-packages
  (
    echo python311.zip
    echo .
    echo Lib\site-packages
    echo import site
  ) > "%PY_EMBED%\python311._pth"
  echo [1/4] Python Embedded ready.
) else (
  echo [1/4] Python Embedded already present.
)

REM ---- Step 2: Bootstrap pip ----
"%PY_EXE%" -m pip --version >nul 2>nul
if errorlevel 1 (
  echo [2/4] Installing pip ...
  REM Use get-pip.py
  set "GET_PIP=%~dp0_get-pip.py"
  curl.exe -L -o "%GET_PIP%" "https://bootstrap.pypa.io/get-pip.py" --silent 2>nul
  if not exist "%GET_PIP%" (
    powershell -NoProfile -Command "Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile '%GET_PIP%' -UseBasicParsing"
  )
  if exist "%GET_PIP%" (
    "%PY_EXE%" "%GET_PIP%" --no-warn-script-location
    del "%GET_PIP%" 2>nul
  ) else (
    REM Fallback: copy pip from local Python venv if available
    echo get-pip.py download failed; attempting local copy...
    if exist ".venv\Scripts\python.exe" (
      ".venv\Scripts\python.exe" -m pip install pip setuptools --target "%PY_EMBED%\Lib\site-packages" --no-warn-script-location --quiet
    ) else (
      where python >nul 2>nul
      if not errorlevel 1 (
        python -m pip install pip setuptools --target "%PY_EMBED%\Lib\site-packages" --no-warn-script-location --quiet
      ) else (
        echo ERROR: Cannot install pip. No internet and no local Python found.
        pause
        exit /b 1
      )
    )
  )
  echo [2/4] pip installed.
) else (
  echo [2/4] pip already available.
)

REM ---- Step 3: Install application dependencies ----
echo [3/4] Installing application dependencies ...
"%PY_EXE%" -m pip install -r requirements.txt --no-warn-script-location --quiet
if errorlevel 1 (
  echo ERROR: Failed to install dependencies.
  pause
  exit /b 1
)
echo [3/4] Dependencies installed.

REM ---- Step 4: Install Playwright Chromium ----
echo [4/4] Installing Playwright Chromium browser ...
"%PY_EXE%" -m playwright install chromium
if errorlevel 1 (
  echo WARNING: Playwright Chromium installation had issues.
  echo The browser may need to be installed manually.
)
echo [4/4] Playwright browser ready.

echo.
echo ===================================================
echo   Setup complete!
echo   You can now run the application with: run.bat
echo ===================================================
echo.
REM Only pause if run directly (not called from run.bat)
if "%~1"=="--no-pause" goto :eof
pause
