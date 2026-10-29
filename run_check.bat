@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==================================================
echo   Shop image checker
echo ==================================================
rem --- Find a real Python (skip the Microsoft Store "python" shortcut) ---
set "PY="
py -3 -c "import sys" >nul 2>nul
if not errorlevel 1 set "PY=py -3"
if not defined PY (
  python -c "import sys" >nul 2>nul
  if not errorlevel 1 set "PY=python"
)
if not defined PY (
  echo.
  echo [!] Python is not installed on this PC.
  echo     The "python" command here is only a Microsoft Store shortcut.
  echo.
  echo     1. Download Python from https://www.python.org/downloads/
  echo     2. On the first install screen, CHECK "Add python.exe to PATH"
  echo     3. Close this window and run this file again
  echo.
  start "" "https://www.python.org/downloads/"
  pause
  exit /b 1
)
echo Using Python:
%PY% --version
echo.
echo [1/2] Installing required packages...
%PY% -m pip install --disable-pip-version-check --quiet --upgrade requests beautifulsoup4 pillow openpyxl
if errorlevel 1 (
  echo [!] Package install failed. Check your internet connection.
  pause
  exit /b 1
)
echo.
set /p SHOP=Shop URL (press Enter for https://vivacious.kr/): 
if "%SHOP%"=="" set SHOP=https://vivacious.kr/
echo.
echo [2/2] Checking %SHOP% ...
%PY% shop_image_check.py %SHOP%
echo.
echo Done. The Excel report is in this folder.
start "" "%~dp0"
pause
