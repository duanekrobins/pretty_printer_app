@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Creating Python virtual environment...
    py -3 -m venv .venv 2>nul
    if errorlevel 1 python -m venv .venv
)

call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo.
where java >nul 2>&1
if errorlevel 1 (
    echo WARNING: Java was not found on PATH.
    echo TestSavvy execution-result PDF processing can still run.
    echo General/Vantage PDF-to-Markdown conversion requires Java 11 or newer.
) else (
    echo Java runtime found:
    java -version
)

if errorlevel 1 (
    echo.
    echo Setup failed.
    exit /b 1
)

echo.
echo Pretty Printer environment is ready.
echo Run: run_pretty_printer.cmd
exit /b 0
