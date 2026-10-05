@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Pretty Printer environment not found.
    echo Run setup_windows.cmd first.
    exit /b 1
)

call ".venv\Scripts\activate.bat"
python vantage_pdf_to_markdown.py %*
exit /b %ERRORLEVEL%
