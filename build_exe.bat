@echo off
REM Build a single-file executable using PyInstaller
REM Usage: run this script inside the extracted folder where main.py is located.
REM Make sure you have activated your venv and installed requirements.

pip install pyinstaller
pyinstaller --onefile --add-data "dzair.db;." main.py
echo Build finished. Check the dist folder for main.exe (rename as needed).
pause
