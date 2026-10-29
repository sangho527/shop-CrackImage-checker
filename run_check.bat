@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==================================================
echo   Shop image checker
echo ==================================================
where python >nul 2>nul
if errorlevel 1 (
  echo [!] Python is not installed. Install it from https://www.python.org/downloads/
  echo     and check "Add python.exe to PATH" during install.
  pause
  exit /b 1
)
echo [1/2] Installing required packages...
python -m pip install --quiet --upgrade requests beautifulsoup4 pillow openpyxl
echo.
set /p SHOP=Shop URL (press Enter for https://vivacious.kr/): 
if "%SHOP%"=="" set SHOP=https://vivacious.kr/
echo.
echo [2/2] Checking %SHOP% ...
python shop_image_check.py %SHOP%
echo.
echo Done. The Excel report is in this folder.
start "" "%~dp0"
pause
