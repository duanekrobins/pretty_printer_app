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
if errorlevel 1 goto :setup_failed

python -m pip install -r requirements.txt
if errorlevel 1 goto :setup_failed

echo.
echo Pretty Printer base environment is ready.
echo Run TestSavvy processing: run_pretty_printer.cmd
echo.
echo For general Vantage/Advantage PDF-to-Markdown support, also run:
echo setup_vantage_pdf_to_markdown.cmd
exit /b 0

:setup_failed
echo.
echo Setup failed.
exit /b 1
