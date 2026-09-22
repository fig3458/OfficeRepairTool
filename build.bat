@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo [1/2] Installing dependencies...
python -m pip install -r requirements.txt -q
if errorlevel 1 exit /b 1

echo [2/2] Building lightweight tkinter EXE (UAC admin)...
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name "OfficeRepairTool" ^
  --uac-admin ^
  --exclude-module PIL --exclude-module Pillow --exclude-module numpy --exclude-module PySide6 ^
  --add-data "assets\word_safe_mode_error.png;assets" ^
  main.py

if errorlevel 1 exit /b 1
echo.
for %%A in (dist\OfficeRepairTool.exe) do echo Done: %%~fA  (%%~zA bytes)
explorer dist
pause
