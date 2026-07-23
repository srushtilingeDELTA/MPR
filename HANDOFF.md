# GSE MPR Project — Full Handoff Summary for Kiro IDE

## 1. What this project is

**Repo:** https://github.com/srushtilingeDELTA/MPR  
**Active branch:** `cursor/scorecard-system-screenshots-67e6`  
**Open PR:** https://github.com/srushtilingeDELTA/MPR/pull/1  
**Base branch:** `main`  
**Current script version:** `2026.07.22-north-summary-layout` (printed when you run `main.py`)

This is **Project 1: Monthly GSE MPR PowerPoint automation**.

**Objective:** Every month, automatically build the Delta GSE Monthly Performance Report deck by:

1. Logging into SharePoint (Edge browser, no passwords in code)
2. Reading live Excel workbooks from `6 - TESTING`
3. Opening the PowerPoint template `GSE MPR - Template.pptx`
4. Filling / replacing mapped slide content with **current month data**
5. Saving the report under `output/`
6. Optionally uploading the finished deck back to SharePoint

The guiding design rule is: **preserve the template look**. Do not redesign slides. Prefer Excel **screenshots** (or native PPT table/chart fills where already built) so formatting, colors, and sparklines match Excel.

---

## 2. Overall goals of the workstream

### Primary goal

Replace manual copy/paste from Excel into PowerPoint with a repeatable Windows pipeline:

`SharePoint Excel → detect ranges → screenshot / fill → template PPT → SharePoint upload`

### Quality bar (what “done” means for a slide)

- Content comes from the **correct workbook + sheet**
- Capture is the **right table/graph**, not neighboring junk
- Placement matches the **template composition** (size, position, no overlap)
- Narrative boxes (“Leading Issues” / “Action Plan”) are **cleared** where automation should leave room for manual entry
- On Windows + Excel: use **COM `CopyPicture`** (true Excel screenshot)
- Fallback: Pillow cell render (usable off Windows, but lower fidelity)

### Secondary goals already pursued

- Robust sheet name matching (fuzzy / `sheet_match`)
- Optional workbooks so a missing Visualizations file doesn’t kill the whole run
- Debug PNG previews in `output/` for visual QA
- Verify hooks in `ppt_builder.py` that print picture counts after build

---

## 3. Architecture (how the code is organized)

### Entry point

- `main.py` — loads config, syncs SharePoint, builds PPT, uploads

### Config

- `config.yaml` — workbooks, SharePoint site/folder/files, report month, KPI name mappings
- `template_map.yaml` — **slide index → element type → workbook/sheet/options**  
  This is the control plane for what each slide does.

### Data / IO

- `workbook_store.py` — loads workbook bytes (disk or SharePoint cache)
- `mpr_data.py` — report period + workbook accessors
- `sharepoint_live.py` — Edge login + REST download/upload; supports `match_contains` for fuzzy filenames
- `sharepoint_selenium.py` — UI download fallback

### PowerPoint assembly

- `ppt_builder.py` — applies `template_map.yaml` elements; runs `_verify_output`
- `ppt_format.py` — formatting helpers
- `report_utils.py` — config/period helpers

### Screenshot / panel modules (slide-specific)

| Module | Purpose |
|--------|---------|
| `scorecard_screenshots.py` | Core Excel COM CopyPicture + Pillow fallback + System scorecard section capture + placement helpers (`capture_range_png`, `place_picture_on_slide`, zoom, upscale) |
| `gir_workings.py` / `gir_panels.py` | GIR native tables/charts + narrative clear |
| `ea_asap_panels.py` | EA/ASAP combined table screenshot |
| `people_workings.py` | PEOPLE table + 3 Excel graph screenshots |
| `finance_workings.py` | FINANCE Regions/Budget/OT/Hours table |
| `pmi_workings.py` | PMI Mot/Stat + Non-Mot table + graphs |
| `isr_workings.py` | ISR Rel/Sev table + graphs |
| `north_summary.py` | **Slide 14** North Scorecard Summary (4 panels) |
| `scorecard_comparison.py` | **Slide 15** North Scorecard Comparison table |

### Important shared capture APIs (`scorecard_screenshots.py`)

- `capture_range_png(..., zoom=, render_scale=, min_width=)`
  - `zoom` bumps Excel ActiveWindow zoom before CopyPicture (critical for small legends)
  - `min_width` LANCZOS-upscales tiny PNGs so they don’t look pixelated on the slide
- `place_picture_on_slide(..., fit="fill"|"contain", align=...)`
- `resolve_sheet_name(...)` — exact sheet, index, or substring match

---

## 4. Source Excel workbooks

Configured under `config.yaml → workbooks` and `sharepoint.files`:

| Key | Typical SharePoint file | Used for |
|-----|-------------------------|----------|
| `actuals` | `MPR Actuals and Goals_v2.xlsx` | KPI charts (Jam/Clear/VOS), some mappings |
| `workings` | `New GSE MPR Workings.xlsx` | GIR, EA/ASAP, People, Finance, PMI, ISR |
| `scorecards` | `2026 - GSE Scorecards.xlsx` | System scorecards (slides 3–4), South/Stationary/Galley placeholders |
| `visualizations` | `GSE MPR Visualizations - TESTING.xlsx` (matched via `match_contains: "Visualizations"`) | **North slides 14–15** |

SharePoint folder:  
`https://deltaairlines.sharepoint.com/sites/DL002488` → library `GSE MPR Documents` → `6 - TESTING`

Report period currently pinned in config:

```yaml
report:
  use_previous_month: false
  year: 2026
  month: 5
```

---

## 5. Slide-by-slide status (what’s built)

Indexing note: **PPT slide N = `template_map` index N−1**.

| PPT # | Name | Status | Source |
|------:|------|--------|--------|
| 1 | Cover | Done | month tokens |
| 2 | Agenda | Done | planned discussion fill |
| 3 | System scorecard (Safety/CX/Ops) | Done | Scorecards `System` sections screenshot |
| 4 | System scorecard (People/Finance/Overall) | Done | Scorecards `System` sections screenshot |
| 5 | GIR | Done | Workings `GIR` **native** tables/chart (not screenshot) |
| 6 | EA / ASAP | Done | Workings EA/ASAP table screenshot |
| 7 | People | Done | Workings PEOPLE table + 3 graphs |
| 8 | Finance overview | Done | Workings FINANCE table |
| 9–10 | Finance comments | Done (clear only) | empty Leading Issues / Action Plan |
| 11 | PMI | Done | Workings PMI Mot/Stat+Non-Mot + graphs |
| 12 | ISR | Done | Workings ISR Rel/Sev + graphs |
| 13 | ISR comments | Done (clear only) | empty narrative boxes |
| **14** | **North Scorecard Summary** | **Active / iteratively refined** | Visualizations `Scorecard Summaries` |
| **15** | **North Scorecard Comparison** | **Newly wired** | Visualizations `Scorecard Comparison` |
| 16–18 | South / Stationary | Placeholder screenshots from scorecards sheets (likely wrong long-term; still mapped to ACTUALS/Weighting/Stationary/GOALS) |
| 19–20 | Ops Jam/Clear + comments | Charts from actuals; comments manual |
| 21–22 | VOS + Station of Quarter | Chart / manual |
| 23–24 | Galley | Placeholder scorecard sheet screenshots |

---

## 6. Deep dive: recent work (North slides) — what we struggled with and why

This is the focus of the latest agent work and what Kiro should treat carefully.

### 6.1 Slide 14 — North Scorecard Summary

**Objective:** Reproduce the live Monthly Performance North slide that shows:

1. **GSE MPR category summary table** (SAFETY / CUSTOMER EXPERIENCE / OPERATIONS / FINANCE / PEOPLE / TOTAL × stations like NORTH, BOS, DTW, JFK/LGA or NY, MSP, SLC)
2. **KPI metrics table** underneath (Global Injury Rate, EA Compliance, ASAP, ISR%, SEV, PMI, PMI NME, QC Compliance, Budget $000s, Overtime, Total Hours, Lead Input)
3. **Two legends**
   - Score bands: Above 4 / Between 3–4 / Below 3
   - KPI status: Better than Goal / Worse / Not Applicable / Goal Pending

**Source (important):**  
Not the Scorecards workbook.  
**`GSE MPR Visualizations[- TESTING].xlsx` → sheet matching `"scorecard summar"`** (Scorecard Summaries).

**Module:** `north_summary.py`  
**Element type:** `north_summary_panels` in `template_map.yaml` index `13`

#### Evolution of the implementation (lessons learned)

1. **Wrong source first** — early attempts used Motorized/System-style scorecards. Corrected to Visualizations Scorecard Summaries.

2. **Detection must find four ranges**
   - Summary block by category labels
   - Metrics block by KPI labels
   - Score legend by exact body labels (`Score Above 4`, etc.)
   - KPI legend by exact body labels (`KPI Better than Goal`, etc.)
   - Had to ignore legend-adjacent junk (`OT`, `Hours`, `HC`) that short aliases falsely matched as KPIs (`_cluster_hits_by_column`, `_is_legend_junk_neighbor`)

3. **Legend readability**
   - Tiny Excel ranges produced ~400–600px bitmaps → looked pixelated on PPT
   - Fix: capture legends at **Excel zoom 200%**, LANCZOS upscale to **min width ~1800**, higher Pillow render scale for fallback

4. **Layout / overlap (latest user feedback)**
   - Problem: legends overlapped tables; composition sat too far left with empty right space
   - Root cause contributors:
     - Left-column legend boxes colliding with left-biased table placement
     - `_resolve_placement_boxes` previously reused **template picture slots**, which pulled tables left into the legend area
   - **Current intended layout** (as of latest commit):
     - Wide content band shifted right (`_BAND_LEFT ≈ 1.65M EMU`, width ≈ 9.7M)
     - Summary table on top
     - **Large score legend under summary**, flush left with table
     - Metrics table below
     - **Large KPI legend under metrics**, flush left with table
     - Template picture slots **ignored by default** (`use_template_slots` only if explicitly set)

**User visual target:** tables further right / more centered; legends large and readable; no overlap.  
**Verify by eye after `python main.py` on Windows** — layout constants may still need tuning against the real desired photo.

### 6.2 Slide 15 — North Scorecard Comparison

**Objective:** Screenshot the Entity/Period comparison grid that looks like:

- Columns: Entity, Period, GIR, EAC, ASAP, ISR, PMI, PMI(N), QC, Budget, OT, Total Hours, LIC
- Rows grouped by entity (NORTH, BOS, DTW, MSP, NY, SLC), each with May'26 / vs LM / vs LYSM

**Source:** Visualizations → sheet matching `"scorecard comparison"`  
**Module:** `scorecard_comparison.py`  
**Element type:** `scorecard_comparison`  
**Placement:** template slide 15 often has **no picture slot**, so code places into a default wide content box under the title.

Detection finds the header row containing Entity + Period (+ KPI headers), then expands down through period rows.

---

## 7. How a normal run works (Windows)

```powershell
cd <repo>
.venv\Scripts\activate
git checkout cursor/scorecard-system-screenshots-67e6
git pull
python main.py
```

Expected flow:

1. Edge opens for SharePoint login (if live_read)
2. Files synced/read (Actuals, Template, Scorecards, Workings, Visualizations)
3. Template loaded; each mapped slide applied
4. Debug PNGs written under `output/` (e.g. `_debug_north_*.png`)
5. VERIFY lines printed (picture counts for key slides)
6. Report saved; optional upload to SharePoint

**Excel must be installed** for high-quality screenshots. Without COM, Pillow fallback runs and quality drops (especially legends).

---

## 8. What is still incomplete / likely next work for Kiro

### High priority / recently unstable

1. **Slide 14 layout QA** — confirm the under-table large-legend layout matches the user’s “second image.” If the true live template still wants **left-side legends beside tables**, constants/placement in `north_summary.py` need another pass. Do not reintroduce template-slot reuse without checking overlap.
2. **Slide 15 visual QA** — confirm Scorecard Comparison capture crops only the Entity/Period table and fills the slide cleanly.

### Medium priority (placeholders / wrong sources)

3. **South / Stationary / Galley comparison slides (16–18, 23–24)** still point at generic Scorecards sheets (`ACTUALS`, `Weighting Lookups`, `GOALS`, `Galley`). They likely need the same Visualizations-style treatment as North (or correct regional sheets), with proper detection—not whole-sheet dumps.
4. **Narrative slides** intentionally left manual (Finance comments, ISR comments, Ops comments, Station of the Quarter). Keep clearing boxes only unless product owners want auto text.

### Platform / ops

5. Excel COM reliability on Windows (pywin32 cache corruption was fixed earlier—clear gen_py and retry if COM fails).
6. Month rollover: update `config.yaml` `report.year/month` and Actuals sheet name when moving past May 2026.
7. Filename drift: Visualizations file may be named `… - TESTING.xlsx`; `match_contains: "Visualizations"` handles this—keep that pattern for other optional files.

---

## 9. Key files to open first in Kiro

If continuing layout/screenshot work:

1. `template_map.yaml` — what each slide is supposed to do
2. `north_summary.py` — Slide 14 detection + placement constants at top of file
3. `scorecard_comparison.py` — Slide 15
4. `scorecard_screenshots.py` — capture engine (zoom/upscale/placement)
5. `ppt_builder.py` — element dispatch + VERIFY checks
6. `config.yaml` — workbook/SharePoint wiring
7. `README.md` — operator quick start
8. `HANDOFF.md` — this document

Useful scripts:

- `scripts/list_scorecard_sheets.py` — sheet inventory + suggested mappings
- `scripts/inspect_scorecard_system.py` — System section detection
- `scripts/slide_review.py` — review helpers
- `scripts/dump_template_inventory.py` — template shape inventory

---

## 10. Design conventions the team has been following

- **Prefer screenshots of Excel** over rebuilding tables in PPT (except GIR native fill, which was an explicit choice).
- **Detect by labels**, not hard-coded A1 ranges (Excel layouts shift).
- **Tight crops** — exclude neighboring OT/Hours/HC/legend junk.
- **Optional workbooks** for Visualizations so missing files warn instead of hard-failing the whole deck.
- **Don’t clear template content until a replacement screenshot succeeds.**
- **Preserve brand/template chrome** (title, speaker box, Delta logo, red underline).
- Placement uses **EMU** coordinates (widescreen 13.333"×7.5" = 12,192,000 × 6,858,000 EMU).

---

## 11. Practical Kiro startup checklist

1. Clone/pull `cursor/scorecard-system-screenshots-67e6` (or merge PR #1 into `main` first if that’s the new base).
2. Create/activate `.venv`, `pip install -r requirements.txt` (needs `pywin32`, `openpyxl`, `python-pptx`, `Pillow`, Office365/SharePoint deps).
3. Confirm Excel + Edge on the machine.
4. Ensure SharePoint `6 - TESTING` has:
   - Actuals, Template, Scorecards, Workings, Visualizations (-TESTING OK)
5. Run `python main.py` for May 2026.
6. Open output PPT and **visually inspect slides 14 and 15 first** (most recently changed).
7. Use `output/_debug_*.png` to see raw captures before placement.

---

## 12. One-sentence north star

**Build a reliable monthly pipeline that turns live SharePoint Excel (Scorecards + Workings + Visualizations) into a polished GSE MPR PowerPoint that looks like the hand-built deck—especially the North Scorecard Summary/Comparison slides—without manual screenshot pasting.**
