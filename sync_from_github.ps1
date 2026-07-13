# Sync MPR project files from GitHub branch cursor/system-scorecard-slide-e0c3
# Run from your MPR folder:
#   cd "C:\Users\533406\Desktop\SRUSHTI LINGE\MPR"
#   powershell -ExecutionPolicy Bypass -File sync_from_github.ps1
#
# Or one-liner (downloads script then runs it):
#   cd "C:\Users\533406\Desktop\SRUSHTI LINGE\MPR"
#   Invoke-WebRequest -Uri "https://raw.githubusercontent.com/srushtisl20/DELTA/cursor/system-scorecard-slide-e0c3/sync_from_github.ps1" -OutFile sync_from_github.ps1
#   powershell -ExecutionPolicy Bypass -File sync_from_github.ps1

$ErrorActionPreference = "Stop"
$Root = $PSScriptRoot
$Branch = "cursor/system-scorecard-slide-e0c3"
$RepoBase = "https://raw.githubusercontent.com/srushtisl20/DELTA/$Branch"

Write-Host "=== Syncing GSE MPR from GitHub ===" -ForegroundColor Cyan
Write-Host "Branch: $Branch"
Write-Host "Target: $Root"
Write-Host ""

@("scripts", "data", "templates", "output", "tests") | ForEach-Object {
    $dir = Join-Path $Root $_
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir | Out-Null
        Write-Host "Created folder: $_"
    }
}

$Files = @(
    "main.py",
    "config.yaml",
    "template_map.yaml",
    "requirements.txt",
    "mpr_data.py",
    "data_lookup.py",
    "workbook_store.py",
    "ppt_builder.py",
    "ppt_format.py",
    "ppt_missing.py",
    "report_utils.py",
    "gir_slide.py",
    "people_slide.py",
    "safety_compliance.py",
    "scorecard_data.py",
    "scorecard_style.py",
    "scorecard_layout.py",
    "narrative_boxes.py",
    "picture_replace.py",
    "sharepoint_live.py",
    "sharepoint_excel.py",
    "sharepoint_selenium.py",
    "DESKTOP_COPY_CHECKLIST.md",
    "SLIDE_CHECKLIST.md",
    "scripts/slide_review.py",
    "scripts/inspect_kpis.py",
    "scripts/inspect_scorecards.py",
    "scripts/dump_template_inventory.py",
    "scripts/list_workbook_sheets.py",
    "scripts/verify_setup.py",
    "scripts/list_sheets.py",
    "scripts/inspect_excel.py",
    "scripts/inspect_template.py",
    "scripts/download_from_sharepoint.py",
    "scripts/sync_sharepoint_files.py",
    "templates/GSE MPR - Template.pptx"
)

$ok = 0
$fail = 0
foreach ($rel in $Files) {
    $dest = Join-Path $Root ($rel -replace "/", [IO.Path]::DirectorySeparatorChar)
    $parent = Split-Path $dest -Parent
    if ($parent -and -not (Test-Path $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    $url = "$RepoBase/$($rel -replace '\\','/')"
    Write-Host "  $rel"
    try {
        Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing
        $ok++
    } catch {
        Write-Host "    FAILED: $_" -ForegroundColor Red
        $fail++
    }
}

Write-Host ""
if ($fail -eq 0) {
    Write-Host "Synced $ok files successfully." -ForegroundColor Green
} else {
    Write-Host "Synced $ok files; $fail failed." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host '  .venv\Scripts\activate'
Write-Host '  pip install -r requirements.txt'
Write-Host '  python main.py'

$Required = @(
    "scorecard_layout.py",
    "scorecard_style.py",
    "picture_replace.py",
    "ppt_builder.py",
    "narrative_boxes.py"
)
$Missing = @()
foreach ($rel in $Required) {
    $path = Join-Path $Root $rel
    if (-not (Test-Path $path)) {
        $Missing += $rel
    }
}
if ($Missing.Count -gt 0) {
    Write-Host ""
    Write-Host "Missing required files after sync:" -ForegroundColor Red
    $Missing | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    Write-Host "Re-run this script or download the missing files from GitHub." -ForegroundColor Red
    exit 1
}
