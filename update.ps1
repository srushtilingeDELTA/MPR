# Sync latest project files from GitHub (Windows PowerShell)
# Run: powershell -ExecutionPolicy Bypass -File update.ps1

$ErrorActionPreference = "Stop"
$Root = $PSScriptRoot
$RepoBase = "https://raw.githubusercontent.com/srushtilingeDELTA/MPR/main"

$Files = @(
    "main.py",
    "config.yaml",
    "template_map.yaml",
    "requirements.txt",
    "setup.ps1",
    "update.ps1",
    "README.md",
    "mpr_data.py",
    "ppt_builder.py",
    "ppt_format.py",
    "report_utils.py",
    "workbook_store.py",
    "scorecard_screenshots.py",
    "sharepoint_live.py",
    "sharepoint_selenium.py",
    "scripts/verify_setup.py",
    "scripts/inspect_kpis.py",
    "scripts/inspect_scorecard_system.py",
    "scripts/slide_review.py",
    "scripts/dump_template_inventory.py",
    "scripts/list_workbook_sheets.py",
    "scripts/list_scorecard_sheets.py",
    "scripts/download_from_sharepoint.py",
    "scripts/sync_sharepoint_files.py",
    "scripts/list_sharepoint_folder.py",
    "scripts/upload_report_to_sharepoint.py"
)

Write-Host "=== Updating GSE MPR project files ===" -ForegroundColor Cyan
Write-Host "Source: $RepoBase"
Write-Host "Target: $Root"

@("scripts", "data", "templates", "output") | ForEach-Object {
    $dir = Join-Path $Root $_
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
}

foreach ($relPath in $Files) {
    $dest = Join-Path $Root $relPath
    $url = "$RepoBase/$($relPath -replace '\\', '/')"
    Write-Host "Downloading $relPath ..."
    try {
        Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing
    } catch {
        Write-Host "  WARNING: Could not download $relPath (may not be on GitHub yet)" -ForegroundColor Yellow
    }
}

$MainPath = Join-Path $Root "main.py"
$MainText = Get-Content $MainPath -Raw
if ($MainText -notmatch "SCRIPT_VERSION" -or $MainText -notmatch "Report settings") {
    Write-Host "ERROR: main.py did not update correctly." -ForegroundColor Red
    exit 1
}
Write-Host "OK: main.py updated" -ForegroundColor Green

Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host '  .venv\Scripts\activate'
Write-Host '  pip install -r requirements.txt'
Write-Host '  python scripts\verify_setup.py'
Write-Host '  python main.py'
