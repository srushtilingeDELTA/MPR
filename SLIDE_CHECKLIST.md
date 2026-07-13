# GSE MPR — Slide-by-slide data audit

Template source: `GSE MPR - Template.pptx` on [GitHub main](https://github.com/srushtisl20/DELTA/blob/main/GSE%20MPR%20-%20Template.pptx)  
Local copy: `templates/GSE MPR - Template.pptx` (25 slides)

**PowerPoint slide # = template index + 1**

## Workbooks

| Workbook | File | Used for |
|----------|------|----------|
| actuals | `MPR Actuals and Goals_v2.xlsx` | KPI rows: KPI, Yr_Nb, Mo_Nb, Entity, Num, Den, Actual, Goal |
| scorecards | `2026 - GSE Scorecards.xlsx` | System + entity scorecard grids |
| workings | `New GSE MPR Workings.xlsx` | Finance tables (future) |

---

## Slide status

| PPT | Idx | Title | Data source | Auto-fill | Blocker |
|----:|----:|-------|-------------|-----------|---------|
| 1 | 0 | Cover | config report month | ✅ month tokens | — |
| 2 | 1 | Agenda | static section list | ✅ clears detail bullets | — |
| 3 | 2 | System Scorecard (1) | scorecards → summary_1 | ✅ | Screenshot removed → native table from Excel |
| 4 | 3 | System Scorecard (2) | scorecards → summary_2 | ✅ | Same |
| 5 | 4 | **GIR** | actuals + scorecards (scores) | ✅ | DNF if missing; injury KPI names must match Excel |
| 6 | 5 | **EA/ASAP** | actuals (EAC, ASAP by region) | ✅ | OLE/screenshot removed → native table + fill |
| 7 | 6 | **People** | actuals + scorecards | ✅ | Confirm survey KPI names |
| 8 | 7 | Finance overview | workings workbook | ✅ | Screenshot removed → workings sheet pasted |
| 9 | 8 | Finance comments | manual | ✅ clears narrative | — |
| 10 | 9 | Finance comments | manual | ✅ clears narrative | — |
| 11 | 10 | PMI | actuals PM (M), PM (S) | ✅ 2 charts | — |
| 12 | 11 | ISR | actuals REL, SEV | ✅ 2 charts | Remove template pictures on slide |
| 13 | 12 | ISR comments | manual | ✅ clears narrative | — |
| 14 | 13 | North scorecard summary | scorecards sheet | ⚠️ | Picture only — need sheet name → index mapping |
| 15 | 14 | North comparison | scorecards sheet | ⚠️ | Same |
| 16 | 15 | South scorecard summary | scorecards sheet | ⚠️ | Same |
| 17 | 16 | South comparison | scorecards sheet | ⚠️ | Same |
| 18 | 17 | Stationary summary | scorecards sheet | ⚠️ | Same |
| 19 | 18 | Stationary comparison | scorecards sheet | ⚠️ | Same |
| 20 | 19 | Operations | actuals Jams, Clear Times | ✅ 2 charts | Tighten KPI pattern for Clear Times |
| 21 | 20 | Ops comments | manual | ✅ clears narrative | — |
| 22 | 21 | VOS | actuals VOS (S) | ✅ chart | — |
| 23 | 22 | Station Winners | manual / ? | ❌ | **No Excel source identified** — point us to data |
| 24 | 23 | Galley scorecard | scorecards sheet | ⚠️ | Picture only |
| 25 | 24 | Galley comparison | scorecards sheet | ⚠️ | Picture only |

---

## GIR slide (PPT 5) — expected values from template

| Element | Excel source |
|---------|----------------|
| MTD Actual / Goal | actuals, KPI=GIR, Entity=System, report month |
| YTD Actual | weighted sum(Num)/sum(Den) Jan–report month |
| Yo1Y / Yo2Y | GIR actual, same month, prior years |
| Monthly chart | System GIR Actual per month + P1Y series |
| Injury Count row | sum Injury Count by month |
| Injury Breakdown | Recordable / NonRec / DART by entity 587, 613 |
| Score MTD/YTD | scorecards summary_1 → Global Injury Rate → Score |
| Leading Issues / Action Plan | **manual** (cleared to blank) |

Missing values show **DNF**.

---

## Commands

```powershell
python scripts\slide_review.py 4 --report    # GIR
python scripts\inspect_kpis.py                 # list KPI names in actuals
python scripts\inspect_scorecards.py         # scorecard sheet names
python scripts\dump_template_inventory.py      # refresh shape inventory
```

## What we need from you

1. **EA/ASAP slide (PPT 6)** — template version with native **tables** (not screenshot), or confirm Excel sheet/tab for EAC/ASAP layout  
2. **System scorecard slides (PPT 3–4)** — add formatted **tables** to template (code cannot edit pictures)  
3. **Finance slide (PPT 8)** — which Excel tab has Budget / OT / Hours round-robin data?  
4. **Entity scorecards (PPT 14–19)** — map each slide to exact sheet name in `2026 - GSE Scorecards.xlsx`  
5. **Station Winners (PPT 23)** — data source?  
6. **Galley (PPT 24–25)** — sheet names in scorecards workbook  
7. Run `python scripts\inspect_kpis.py` and share output if any slide shows DNF
