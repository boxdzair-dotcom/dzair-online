
# DZAIR Sales & Profit Manager - Improved Package

This improved package includes:
- main.py : Improved Tkinter app with charts, filters, validation
- dzair.db : SQLite database (created automatically)
- build_exe.bat : Windows batch script to build a single-file exe using PyInstaller
- requirements.txt : Python dependencies
- README.md : Instructions (this file)

## How to run locally
1. Ensure Python 3.9+ is installed.
2. Create and activate a virtual environment:
   python -m venv venv
   venv\Scripts\activate
3. Install dependencies:
   pip install -r requirements.txt
4. Run the app:
   python main.py

## How to build a Windows executable (.exe)
1. Install PyInstaller (inside your venv):
   pip install pyinstaller
2. Run the batch script included (build_exe.bat) from the folder:
   build_exe.bat
3. After building, the exe will be in the `dist` folder. Copy `dzair.db` next to the exe if needed.

Notes:
- If matplotlib is not installed, the dashboard will show a message; install matplotlib to get charts.
- reportlab is required to export PDF invoices.
