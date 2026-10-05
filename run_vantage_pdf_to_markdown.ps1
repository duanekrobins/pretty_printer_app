$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$python = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"

if (-not (Test-Path $python)) {
    Write-Error "Pretty Printer environment not found. Run setup_windows.cmd first."
    exit 1
}

& $python "$PSScriptRoot\vantage_pdf_to_markdown.py" @args
exit $LASTEXITCODE
