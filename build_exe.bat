@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==================================================
echo   Build: Shop Image Checker (.exe)
echo ==================================================
where python >nul 2>nul
if errorlevel 1 (
  echo [!] Python is not installed. Install it from https://www.python.org/downloads/
  echo     and check "Add python.exe to PATH" during install.
  pause
  exit /b 1
)
echo [1/3] Installing build tools...
python -m pip install --quiet --upgrade pip
python -m pip install --quiet --upgrade requests beautifulsoup4 pillow openpyxl pyinstaller
if errorlevel 1 (
  echo [!] Package install failed. Check your internet connection.
  pause
  exit /b 1
)
echo [2/3] Building the program (1-3 minutes)...
python -m PyInstaller --noconfirm --clean --onefile --windowed --name ShopImageChecker image_checker_app.py
if not exist "dist\ShopImageChecker.exe" (
  echo [!] Build failed. Please send the messages above to the person who gave you this file.
  pause
  exit /b 1
)
echo [3/3] Copying the program...
copy /y "dist\ShopImageChecker.exe" "ShopImageChecker.exe" >nul
rmdir /s /q build >nul 2>nul
del /q ShopImageChecker.spec >nul 2>nul
echo.
echo Done! "ShopImageChecker.exe" is ready in this folder.
echo Share that single file (with the usage guide) with your team.
start "" "%~dp0"
pause
