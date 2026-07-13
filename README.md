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
| 0 | Report month text (e.g. May 2026) |
| 1 | Planned Discussion agenda topics + times |
| 2 | System Scorecard: Safety & Security + Customer Experience + Operations |
| 3 | System Scorecard: People + Finance + Overall and Opportunities |
| 4 | GIR tables + monthly trend chart |
| 6 | People scorecard table (where KPIs match) |
| 10 | PMI Motorized / Stationary charts |
| 11 | ISR Reliability / Severity charts |
| 19 | Jam Rate / Clear Times charts |
| 21 | VOS chart |

Other scorecard image slides stay as template images until entity screenshots are added.

```powershell
python scripts\inspect_scorecard_system.py
```

## Project files

```text
MPR/
├── main.py
├── mpr_data.py
├── ppt_builder.py
├── ppt_format.py
├── scorecard_screenshots.py
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
