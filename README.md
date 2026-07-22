# GSE MPR Report Generator

Automates monthly performance reports: Excel (SharePoint) → PowerPoint.

## Quick start (Windows)

```powershell
cd "C:\Users\533406\Desktop\SRUSHTI LINGE\MPR"
.venv\Scripts\activate
python scripts\inspect_kpis.py      # see KPI names in your Excel
python main.py                      # generate report
```

Output: `output/GSE_MPR_Report_2026_05.pptx`

## SharePoint connection (download Excel automatically)

Install packages:

```powershell
pip install -r requirements.txt
```

| Method | Command | When to use |
|--------|---------|-------------|
| **Live read (Edge)** | `python main.py` with `excel.source: sharepoint` | One Edge login; Excel + template in memory |
| **Sync all files** | `python scripts\sync_sharepoint_files.py` | Test multi-file fetch |
| **List folder** | `python scripts\list_sharepoint_folder.py` | List `6 - TESTING` after login |
| **Live download** | `python scripts\download_from_sharepoint.py --method live` | Save workbooks to `data/` |
| **Selenium download** | `python scripts\download_from_sharepoint.py --method selenium` | Browser UI download fallback |

**Do not put your Delta password in code.** Use browser login.

SharePoint (`config.yaml`):
- Site: `https://deltaairlines.sharepoint.com/sites/DL002488`
- Library: `GSE MPR Documents` → folder `6 - TESTING`
- Files: under `sharepoint.files` (Excel, template, optional scorecards)

```yaml
excel:
  source: "sharepoint"

sharepoint:
  live_read: true
  cache_to_disk: false
  files:
    - name: "MPR Actuals and Goals_v2.xlsx"
      dest: "data/MPR Actuals and Goals_v2.xlsx"
    - name: "GSE MPR - Template.pptx"
      dest: "templates/GSE MPR - Template.pptx"
    - name: "2026 - GSE Scorecards.xlsx"
      optional: true
```

```powershell
python main.py
```

## What the generator fills today

| Slide | Content updated |
|-------|-----------------|
| 0 | Report month text |
| 1 | Planned Discussion agenda topics + times |
| 2 | System scorecard screenshot: Safety & Security + Customer Experience + Operations |
| 3 | System scorecard screenshot: People + Finance + Overall and Opportunities |
| 4 | GIR native tables + chart filled from workings workbook (`GIR` tab); Leading Issues / Action Plan cleared |
| 5 | EA / ASAP table screenshot from workings workbook; Leading Issues / Action Plan cleared |
| 6 | People: first PEOPLE table screenshot + 3 Excel graph screenshots (Leadership / Psych Safety / Accountability) from Workings `PEOPLE` tab |
| 7 | Finance: Regions/BUDGET/OVERTIME/TOTAL HOURS table screenshot from Workings `FINANCE` tab |
| 8–9 | Finance comments: Leading Issues / Action Plan text boxes emptied (boxes kept for manual entry) |
| 10 | PMI: Regions MOTORIZED/STATIONARY + NON-MOTORIZED table + Mot/Stat graphs from Workings `PMI`; narrative emptied |
| 11 | ISR: Regions RELIABILITY + SEVERITY table + Rel/Sev graphs from Workings `ISR`; narrative emptied |
| 13 | ISR comments: Leading Issues / Action Plan text boxes emptied (boxes kept) |
| 14 | North Scorecard Summary from `GSE MPR Visualizations.xlsx` → Scorecard Summaries: GSE MPR category table + KPI metrics table + 2 legends (score bands / KPI status; legends captured at 200% Excel zoom + upscaled for readability) |
| 15 | North Scorecard Comparison from `GSE MPR Visualizations.xlsx` (or `… - TESTING`) → Scorecard Comparison: Entity/Period KPI table (May'26 / vs LM / vs LYSM) |
| 16–18 | South / Stationary scorecard sheet screenshots |
| 19 | Jam Rate / Clear Times charts |
| 21 | VOS chart |
| 23–24 | Galley scorecard sheet screenshots |

Narrative / comment slides stay as template text for manual entry.

On Windows with Excel installed, screenshots use Excel **CopyPicture**. Elsewhere Pillow renders cells as a fallback.

```powershell
python scripts\list_scorecard_sheets.py      # see scorecard/workings sheet names
python scripts\inspect_scorecard_system.py   # System section detection
python main.py
```

## Project files

```text
MPR/
├── main.py
├── mpr_data.py
├── ppt_builder.py
├── ppt_format.py
├── scorecard_screenshots.py
├── north_summary.py
├── scorecard_comparison.py
├── sharepoint_live.py
├── sharepoint_selenium.py
├── workbook_store.py
├── report_utils.py
├── config.yaml
├── template_map.yaml
├── data/
├── templates/
├── output/
└── scripts/
    ├── inspect_kpis.py
    ├── inspect_scorecard_system.py
    ├── slide_review.py
    ├── dump_template_inventory.py
    ├── list_workbook_sheets.py
    ├── list_scorecard_sheets.py
    ├── verify_setup.py
    ├── sync_sharepoint_files.py
    ├── list_sharepoint_folder.py
    ├── download_from_sharepoint.py
    └── upload_report_to_sharepoint.py
```

## GitHub

Repo: https://github.com/srushtilingeDELTA/MPR

```powershell
git pull
python main.py
```

## May 2026 config

```yaml
excel:
  sheet_name: "May Actuals"
report:
  use_previous_month: false
  year: 2026
  month: 5
```
