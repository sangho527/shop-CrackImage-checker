@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==================================================
echo   Build: Shop Image Checker (.exe)
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
echo [1/3] Installing build tools...
%PY% -m pip install --disable-pip-version-check --quiet --upgrade requests beautifulsoup4 pillow openpyxl pyinstaller
if errorlevel 1 (
  echo.
  echo [!] Package install failed. See the messages above.
  echo     - Check your internet connection.
  echo     - On a company network, a firewall may block pypi.org.
  pause
  exit /b 1
)
echo [2/3] Building the program (1-3 minutes)...
%PY% -m PyInstaller --noconfirm --clean --onefile --windowed --name ShopImageChecker image_checker_app.py
if not exist "dist\ShopImageChecker.exe" (
  echo [!] Build failed. Please send the messages above to the person who gave you this file.
  pause
  exit /b 1
)
echo [3/3] Copying the program...
copy /y "dist\ShopImageChecker.exe" "ShopImageChecker.exe" >nul
rmdir /s /q build >nul 2>nul
rmdir /s /q dist >nul 2>nul
del /q ShopImageChecker.spec >nul 2>nul
echo.
echo Done! "ShopImageChecker.exe" is ready in this folder.
echo Share that single file (with the usage guide in the docs folder) with your team.
start "" "%~dp0"
pause
