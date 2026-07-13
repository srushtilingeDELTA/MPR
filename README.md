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

Three methods — try in this order:

| Method | Command | When to use |
|--------|---------|-------------|
| **Live read (Edge)** | `python main.py` with `excel.source: sharepoint` | One Edge login; Excel + template loaded in memory (no disk copy) |
| **Sync all files** | `python scripts\sync_sharepoint_files.py` | Test multi-file fetch before running the report |
| **List folder** | `python scripts\list_sharepoint_folder.py` | See every file in `6 - TESTING` after login |
| **SDK** | `python scripts\download_from_sharepoint.py --method sdk` | Opens Microsoft login in browser; handles MFA |
| **Local sync** | `python scripts\download_from_sharepoint.py --method local` | Library synced via OneDrive — no login |
| **Selenium download** | `python scripts\download_from_sharepoint.py --method selenium` | Saves files to `data/` and `templates/` |

**Do not put your Delta password in code.** Use browser login (Edge live read, SDK, or Selenium).

SharePoint location (in `config.yaml`):
- Site: `https://deltaairlines.sharepoint.com/sites/DL002488`
- Library: `GSE MPR Documents` → folder `6 - TESTING`
- Files: configured under `sharepoint.files` (Excel, template, optional scorecards)

Multi-file live read (`config.yaml`):

```yaml
excel:
  source: "sharepoint"

sharepoint:
  live_read: true
  cache_to_disk: false   # true = also save copies under data/ and templates/
  files:
    - name: "MPR Actuals and Goals_v2.xlsx"
      dest: "data/MPR Actuals and Goals_v2.xlsx"
    - name: "GSE MPR - Template.pptx"
      dest: "templates/GSE MPR - Template.pptx"
    - name: "2026 - GSE Scorecards.xlsx"
      optional: true
```

Full pipeline (live read — recommended on Delta network):

```powershell
python main.py
```

Or download to disk first:

```powershell
python scripts\download_from_sharepoint.py --method selenium --browser edge
python main.py
```

## What the generator fills today

| Slide | Content updated from Excel |
|-------|----------------------------|
| 0 | Report month text (e.g. May 2026) |
| 2 | System Scorecard (PPT slide 3): Safety & Security + Customer Experience + Operations |
| 3 | System Scorecard (PPT slide 4): People + Finance + Overall and Opportunities |
| 4 | GIR tables + monthly trend chart |
| 6 | People scorecard table (where KPIs match) |
| 10 | PMI Motorized / Stationary charts |
| 11 | ISR Reliability / Severity charts |
| 19 | Jam Rate / Clear Times charts |
| 21 | VOS chart |

**Other scorecard image slides** (13–18, 23–24) are still images from the template. Updating those requires the same screenshot approach per entity tab (next phase).

On Windows with Excel installed, System screenshots use Excel **CopyPicture** (true screenshots). Elsewhere, tables are rendered with Pillow as a fallback.

```powershell
python scripts\inspect_scorecard_system.py   # see detected System sections
```

## Project files

```text
MPR/
├── main.py
├── mpr_data.py          # reads May Actuals + other month sheets
├── ppt_builder.py       # fills template tables/charts/text
├── scorecard_screenshots.py  # System tab screenshots → slides 3/4
├── report_utils.py
├── config.yaml
├── data/                # Excel (local only)
├── templates/           # GSE MPR - Template.pptx
├── output/              # generated reports
└── scripts/
    ├── inspect_kpis.py      # list KPI names in Excel
    ├── inspect_scorecard_system.py  # list System tab sections
    ├── inspect_template.py  # list slides/tables/charts in template
    ├── verify_setup.py
    ├── sync_sharepoint_files.py   # one Edge login, all sharepoint.files
    └── list_sharepoint_folder.py  # list folder contents via REST
```

## GitHub

Your repo: https://github.com/srushtilingeDELTA/MPR

After updating files locally:
```powershell
git add main.py mpr_data.py ppt_builder.py report_utils.py config.yaml scripts/
git commit -m "Add template data filling from Excel"
git push origin main
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

## SharePoint

Site: `https://deltaairlines.sharepoint.com/sites/DL002488`