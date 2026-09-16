@echo off
REM =====================================================================
REM  FT6AUTO - Bundle creator
REM
REM  Run this on a machine where setup.bat has already been executed.
REM  It packages EVERYTHING (Python + packages + Chromium + app code)
REM  into a single FT6AUTO.zip that works on any Windows 10+ PC
REM  WITHOUT internet, WITHOUT Python installed, WITHOUT pip.
REM
REM  Distribute the ZIP via network share, Teams, USB, etc.
REM  Recipient: extract -> double-click run.bat -> done.
REM =====================================================================
setlocal
cd /d "%~dp0"

set "BUNDLE_NAME=FT6AUTO_v1.2"
set "OUT_ZIP=%~dp0%BUNDLE_NAME%.zip"

echo.
echo ===================================================
echo   FT6AUTO - Creating portable bundle
echo ===================================================
echo.

REM ---- Validate everything is present ----
set "OK=1"

if not exist "python-embed\python.exe" (
  echo [ERROR] python-embed\python.exe not found.
  echo         Run setup.bat first.
  set "OK=0"
)

if not exist "playwright-browsers\chromium-1134\chrome-win\chrome.exe" (
  echo [ERROR] Chromium browser not found.
  echo         Run setup.bat first.
  set "OK=0"
)

if "%OK%"=="0" (
  echo.
  echo Cannot create bundle. Fix errors above and retry.
  pause
  exit /b 1
)

REM ---- Verify imports work ----
echo [1/3] Verifying environment...
"python-embed\python.exe" -c "import flask, openpyxl, waitress, playwright; print('OK')" 2>nul
if errorlevel 1 (
  echo [ERROR] Python packages are incomplete.
  echo         Run setup.bat first.
  pause
  exit /b 1
)
echo       Environment OK.

REM ---- Clean temporary files ----
echo [2/3] Cleaning temporary files...
if exist "uploads" rd /s /q "uploads" 2>nul
if exist "user-data" rd /s /q "user-data" 2>nul
for /d /r "python-embed" %%d in (__pycache__) do rd /s /q "%%d" 2>nul
for /d /r "server" %%d in (__pycache__) do rd /s /q "%%d" 2>nul
del /q "*.pyc" 2>nul
echo       Done.

REM ---- Create ZIP ----
echo [3/3] Creating %BUNDLE_NAME%.zip ...
echo       This may take a few minutes (bundling ~500 MB)...

if exist "%OUT_ZIP%" del "%OUT_ZIP%"

powershell -NoProfile -Command ^
  "$files = @('config.json','requirements.txt','run.bat','setup.bat','bundle.bat','README.md','server','tests','Referinta','python-embed','playwright-browsers'); " ^
  "$files = $files | Where-Object { Test-Path $_ }; " ^
  "Compress-Archive -Path $files -DestinationPath '%OUT_ZIP%' -CompressionLevel Optimal -Force"

if not exist "%OUT_ZIP%" (
  echo.
  echo [ERROR] Failed to create ZIP file.
  pause
  exit /b 1
)

REM ---- Report ----
for %%A in ("%OUT_ZIP%") do set "SIZE=%%~zA"
set /a "SIZE_MB=%SIZE% / 1048576"

echo.
echo ===================================================
echo   Bundle created successfully!
echo.
echo   File: %OUT_ZIP%
echo   Size: ~%SIZE_MB% MB
echo.
echo   Distribution:
echo     1. Copy the ZIP to the target PC
echo     2. Extract to any folder
echo     3. Double-click run.bat
echo.
echo   No Python, no internet, no pip needed.
echo ===================================================
echo.
pause
