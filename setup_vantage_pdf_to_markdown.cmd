@echo off
setlocal
cd /d "%~dp0"

where java >nul 2>&1
if errorlevel 1 (
    echo.
    echo ERROR: Java was not found on PATH.
    echo OpenDataLoader general PDF conversion requires Java 11 or newer.
    echo Install Java and verify with: java -version
    exit /b 1
)

echo Java runtime found:
java -version
if errorlevel 1 exit /b 1

if not exist ".venv\Scripts\python.exe" (
    echo.
    echo Base Pretty Printer environment not found.
    echo Running setup_windows.cmd first...
    call setup_windows.cmd
    if errorlevel 1 exit /b 1
)

call ".venv\Scripts\activate.bat"

echo.
echo Installing general PDF Markdown dependencies...
python -m pip install -r requirements_general_pdf.txt
if errorlevel 1 (
    echo.
    echo General PDF setup failed.
    exit /b 1
)

echo.
python -c "import opendataloader_pdf; print('OpenDataLoader PDF Python package is ready.')"
if errorlevel 1 (
    echo.
    echo OpenDataLoader import verification failed.
    exit /b 1
)

if not exist "inputs" mkdir "inputs"
if not exist "outputs" mkdir "outputs"

echo.
echo Vantage/Advantage PDF-to-Markdown environment is ready.
echo Put PDF files in: %CD%\inputs
echo Run: run_vantage_pdf_to_markdown.cmd
exit /b 0
