# One-time setup for GSE MPR report project (Windows PowerShell)
# Run from your MPR folder:
#   cd "C:\Users\533406\Desktop\SRUSHTI LINGE\MPR"
#   powershell -ExecutionPolicy Bypass -File setup.ps1

$ErrorActionPreference = "Stop"
$Root = $PSScriptRoot

Write-Host "=== GSE MPR Project Setup ===" -ForegroundColor Cyan
Write-Host "Project root: $Root"

@("data", "templates", "output", "scripts") | ForEach-Object {
    $dir = Join-Path $Root $_
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir | Out-Null
        Write-Host "Created folder: $_"
    }
}

$Venv = Join-Path $Root ".venv"
if (-not (Test-Path $Venv)) {
    Write-Host "Creating virtual environment..." -ForegroundColor Cyan
    if (Get-Command py -ErrorAction SilentlyContinue) {
        py -m venv $Venv
    } elseif (Get-Command python -ErrorAction SilentlyContinue) {
        python -m venv $Venv
    } else {
        Write-Host "ERROR: Python not found. Install Python from https://python.org" -ForegroundColor Red
        exit 1
    }
}

$Pip = Join-Path $Venv "Scripts\pip.exe"
$Python = Join-Path $Venv "Scripts\python.exe"

Write-Host "Installing Python packages..." -ForegroundColor Cyan
& $Pip install -r (Join-Path $Root "requirements.txt")

$Excel = Join-Path $Root "data\MPR Actuals and Goals_v2.xlsx"
$Template = Join-Path $Root "templates\GSE MPR - Template.pptx"

Write-Host ""
Write-Host "=== Setup complete ===" -ForegroundColor Green

if (-not (Test-Path $Excel)) {
    Write-Host "MISSING: Download Excel to data\MPR Actuals and Goals_v2.xlsx" -ForegroundColor Yellow
} else {
    Write-Host "OK: Excel file found"
    & $Python (Join-Path $Root "scripts\list_sheets.py")
}

if (-not (Test-Path $Template)) {
    Write-Host "MISSING: Add templates\GSE MPR - Template.pptx" -ForegroundColor Yellow
} else {
    Write-Host "OK: PowerPoint template found"
}

Write-Host ""
& $Python (Join-Path $Root "scripts\verify_setup.py")

Write-Host ""
Write-Host "To generate a report:" -ForegroundColor Cyan
Write-Host '  .venv\Scripts\activate'
Write-Host '  python main.py'