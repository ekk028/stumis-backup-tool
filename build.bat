@echo off
echo ========================================
echo  STU-MIS Backup Tool — PyInstaller Build
echo ========================================

cd /d "%~dp0"

echo [1/3] Installing dependencies...
pip install -r requirements.txt

echo [2/3] Building executable...
pyinstaller ^
  --onefile ^
  --windowed ^
  --name "STU-MIS Backup Tool" ^
  --add-data "src;src" ^
  --paths "src" ^
  --hidden-import customtkinter ^
  --hidden-import lxml ^
  --hidden-import lxml.etree ^
  --hidden-import xlsxwriter ^
  --collect-all customtkinter ^
  src\main.py

echo [3/3] Done!
echo Executable is at: dist\STU-MIS Backup Tool.exe
pause
