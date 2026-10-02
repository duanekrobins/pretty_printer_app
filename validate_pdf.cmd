@echo off
setlocal
cd /d "%~dp0"
if "%~1"=="" (
    echo Usage: validate_pdf.cmd "C:\path\to\execution_results.pdf"
    exit /b 1
)
if not exist ".venv\Scripts\python.exe" (
    echo Virtual environment not found. Run setup_windows.cmd first.
    exit /b 1
)
call ".venv\Scripts\activate.bat"
python tools\validate_pdf_corpus.py "%~1" --json output\pdf_validation.json
exit /b %errorlevel%
