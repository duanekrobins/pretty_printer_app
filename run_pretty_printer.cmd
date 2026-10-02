@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Virtual environment not found. Run setup_windows.cmd first.
    exit /b 1
)

call ".venv\Scripts\activate.bat"
python pretty_printer.py --config pretty_print_config.json
exit /b %errorlevel%
