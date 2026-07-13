# Desktop copy checklist — DO NOT copy until generation is verified

Use this list **once** after you confirm `python main.py` produces an accurate deck locally.  
Copy the entire project folder from the GitHub branch, or copy only the files listed below into:

`C:\Users\533406\Desktop\SRUSHTI LINGE\MPR`

---

## One-time setup (if not already done)

| Item | Action |
|------|--------|
| Python 3.11+ | Installed |
| Virtual env | `python -m venv .venv` then `.venv\Scripts\activate` |
| Dependencies | `pip install -r requirements.txt` |
| SharePoint files | Run `python scripts\sync_sharepoint_files.py` OR place files in `data\` manually |
| Template | `templates\GSE MPR - Template.pptx` (from repo — **do not use old local copy**) |

---

## Excel files required (data folder)

| File | Used on slides |
|------|----------------|
| `data\MPR Actuals and Goals_v2.xlsx` | GIR, EA/ASAP, People, PMI, ISR, Operations, VOS |
| `data\2026 - GSE Scorecards.xlsx` | System scorecards (PPT 3–4), entity scorecards (PPT 14–19, 24–25), GIR scores |
| `data\New GSE MPR Workings.xlsx` | Finance overview (PPT 8) — **confirm sheet name with team** |

---

## Python files to copy (generation engine)

Copy all of these together — partial copy causes import errors.

```
main.py
config.yaml
template_map.yaml
mpr_data.py
data_lookup.py
workbook_store.py
ppt_builder.py
ppt_format.py
ppt_missing.py
report_utils.py
gir_slide.py
people_slide.py
safety_compliance.py
scorecard_data.py
scorecard_style.py          ← NEW: color-coded table styling (template colors)
scorecard_layout.py         ← NEW: fit scorecard tables to slide placeholder size
narrative_boxes.py          ← NEW: blank Leading Issues / Action Plan boxes
picture_replace.py          ← removes screenshots, inserts styled native tables
sharepoint_live.py          (if using SharePoint live read)
sharepoint_excel.py
sharepoint_selenium.py      (if using Selenium upload)
templates\GSE MPR - Template.pptx
```

Optional but recommended:

```
scripts\slide_review.py
scripts\inspect_kpis.py
scripts\inspect_scorecards.py
scripts\dump_template_inventory.py
SLIDE_CHECKLIST.md
DESKTOP_COPY_CHECKLIST.md     ← this file
```

---

## What the generator does now (screenshot → data)

| PPT | Slide | Old template | New behavior |
|----:|-------|--------------|--------------|
| 3 | System Scorecard 1 | Large **picture** | Removed → native **table** from scorecards `summary_1` block |
| 4 | System Scorecard 2 | Large **picture** | Removed → native **table** from scorecards `summary_2` block |
| 5 | GIR | Native tables + chart | Filled from actuals (DNF if missing) |
| 6 | EA/ASAP | **Embedded OLE** (Excel screenshot) | Removed → native **EA/ASAP table** from actuals |
| 7 | People | Table + charts + small logo pic | Filled from actuals + scorecard scores |
| 8 | Finance | Large **picture** | Removed → native **table** from workings workbook |
| 11 | PMI | Charts only | Charts from actuals |
| 12 | ISR | **Picture** + charts | Picture removed; Reliability/Severity charts from actuals |
| 14–19 | Entity scorecards | **Pictures** | Removed → native **table** pasted from scorecard sheet |
| 20 | Operations | **Picture** + charts | Picture removed; Jam/Clear Times charts from actuals |
| 22 | VOS | Chart | Chart from actuals |
| 24–25 | Galley scorecards | **Pictures** | Removed → native **table** from scorecard sheet |

Slides with **manual narrative only** (Leading Issues / Action Plan / Finance comments): body text cleared, headers kept.

---

## Config to verify before copy (`config.yaml`)

```yaml
report:
  year: 2026
  month: 5          # change each month

excel:
  sheet_name: "May Actuals"   # must match report month

kpi_mappings:
  gir: "GIR"
  injury_count: "Injury Count"
  eac: "EAC"
  asap: "ASAP"
  # … etc
```

---

## Verify before you copy (run on desktop)

```powershell
cd "C:\Users\533406\Desktop\SRUSHTI LINGE\MPR"
.venv\Scripts\activate
python -m pytest tests\ -q --ignore=tests\test_sharepoint_live.py
python main.py
```

Open `output\GSE MPR - May 2026.pptx` and check slides 3–7, 11–12, 20, 22 against the reference deck.

---

## Still needs your input (cannot auto-fill without mapping)

| Slide | What we need |
|-------|----------------|
| PPT 8 Finance | Exact **sheet name** in Workings file for Budget/OT/Hours table |
| PPT 14–19 Entity scorecards | Confirm **sheet index → entity** mapping in `template_map.yaml` |
| PPT 23 Station Winners | **Data source** (no Excel mapping yet) |
| PPT 24–25 Galley | Confirm scorecard **sheet names** |
| Any DNF cells | Run `python scripts\inspect_kpis.py` — KPI name may differ in Excel |

---

## After verified — single sync command

**Recommended** — downloads every file from the latest GitHub branch:

```powershell
cd "C:\Users\533406\Desktop\SRUSHTI LINGE\MPR"
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/srushtisl20/DELTA/cursor/system-scorecard-slide-e0c3/sync_from_github.ps1" -OutFile sync_from_github.ps1
powershell -ExecutionPolicy Bypass -File sync_from_github.ps1
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

If you already have `sync_from_github.ps1` in the folder, just run:

```powershell
cd "C:\Users\533406\Desktop\SRUSHTI LINGE\MPR"
powershell -ExecutionPolicy Bypass -File sync_from_github.ps1
```

**Alternative** — if the MPR folder is a git clone:

```powershell
cd "C:\Users\533406\Desktop\SRUSHTI LINGE\MPR"
git fetch origin cursor/system-scorecard-slide-e0c3
git checkout cursor/system-scorecard-slide-e0c3
git pull origin cursor/system-scorecard-slide-e0c3
```

---

## Branch / PR

Latest work: branch `cursor/system-scorecard-slide-e0c3` on https://github.com/srushtisl20/DELTA

Do not copy until you confirm the generated PPT matches the reference May deck.
