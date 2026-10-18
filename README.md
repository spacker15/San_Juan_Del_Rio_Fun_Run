# Saints Run Wild: Fun Run Fundraising Workbook

This is the fundraising tracker for the **San Juan del Rio Catholic School Fun Run**. It's a single Excel file: no database, no website, no macros. The goal is **$65,000**, the stretch goal is **$60,000 (slime)**, and the campaign runs **Oct 30 – Nov 13, 2026**.

| File | Use |
|---|---|
| `output/SaintsRunWild_2026.xlsx` | **Blank template.** Load the real roster and use this one. |
| `output/SaintsRunWild_2026_DEMO.xlsx` | The same workbook filled with fictional data (424 students, about $60k), for training and demos. |
| `docs/Volunteer_Guide.md` | One-page daily runbook and campaign-close checklist |
| `docs/OneCause_Import.md` | Export mapping, name matching, and reconciliation rules |
| `templates/roster_import.csv` | Roster column order. Paste it into Roster **B2** (Student IDs in column A are pre-assigned). |

## What's inside
| Tab group | Tabs |
|---|---|
| **Dashboards** (red) | TV Dashboard (thermometer, KPIs, grade chart, top 10s, daily trend vs last year) · Leaderboards (grade filter) · Challenge Winners (auto daily winners + detail + Wild Card drawing) · Prize Report · Data Check |
| **Inputs** (blue) | Settings · Classrooms · Roster · Prizes · Challenges · OneCause_Paste · Cash_Check · History |
| **Reports** (green) | Rpt_Summary · Rpt_Grades · Rpt_Classes · Rpt_Students · Rpt_Families · Rpt_Donors · Rpt_Daily (print-ready, PDF export) |
| **Engine** (hidden, locked) | OC_Norm · Ledger · Calc_Students · Calc_Classes · Calc_Grades · Calc_Families · Calc_Donors · Calc_Daily · Calc_Challenges · Calc_KPI |

Data flows in one direction: inputs → Ledger → Calc_* → dashboards and reports. Every output reads from the same calculation sheets, so every number agrees.

The **daily incentives** from the flyer are pre-loaded on the Challenges tab and scored automatically:
- top class per grade over a weekend
- class participation, including $20 Tuesday's $20 minimum
- the Wild Card drawing
- top student per class
- grade vs. grade
- the 50% participation target
- the $60k slime goal

Capacity: 800 students, 40 classrooms, 5,000 OneCause rows, 2,000 cash/check rows, 15 challenges, and 12 prize tiers. Formulas use only Excel 2010+ functions, so the file works in Excel desktop, Excel for the web, Mac, and LibreOffice.

## Rebuild and test (maintainers)
```bash
pip install openpyxl pytest            # LibreOffice (soffice) is needed for tests
python build/build_workbook.py         # writes both files to output/
python -m pytest tests -q              # 40 tests, about 3 min
```
- **Change defaults** (goal, dates, prizes, challenges, OneCause headers, capacity) in `build/config.py`.
- **Test method:** each scenario workbook is recalculated by LibreOffice. Its values are compared with an independent Python model in `tests/reference.py`.
- **Test coverage:** totals, participation, ranks, prize thresholds, every challenge type (including ties and date-window boundaries), families, donors, daily series, OneCause status rules, the drawing, the thermometer at 0/50/92/100/123%, and a scan for formula errors.
