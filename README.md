# Saints Run Wild: Fun Run Tracker

This is a single Excel file (no macros) for the San Juan del Rio Catholic School Fun Run. The goal is $65,000, the stretch goal is $60,000 (slime), and the campaign runs Oct 30 – Nov 13, 2026.

| File | Use |
|---|---|
| `output/SaintsRunWild_2026.xlsx` | Blank. Load the real roster into this one. |
| `output/SaintsRunWild_2026_DEMO.xlsx` | The same file filled with fictional data (424 students, about $61k) |
| `docs/Volunteer_Guide.md` | Daily steps |
| `templates/roster_import.csv` | Roster column order. Paste it at Roster **B2**. |

## Tabs
| Tab | Shows |
|---|---|
| **Dashboard** | TV screen: thermometer, raised / % of goal, students donated, raised today, today's challenge, top 10 classes, top 5 fundraisers |
| **Daily Challenges** | Each flyer day with its result, plus winner details (top class per grade, class participation, Wild Card eligible list, top student per class, grade ranking) |
| **By Class** | Each teacher/class: students, donated, participation, total, $ per day. Grade totals are at the bottom. |
| **By Day** | OneCause vs cash/check per day, gifts, running total |
| **Students** | Each student: total, OneCause, cash/check, gifts, $ per day |
| **All Donations** | Every gift (who gave, for whom, when, how much), with a status |
| **Settings, Roster, OneCause_Paste, Cash_Check** | The only tabs you type or paste into |

## Rebuild and test (maintainers)
```bash
pip install openpyxl pytest           # LibreOffice (soffice) is needed for the tests
python build/build_workbook.py        # writes both files to output/
python -m pytest tests -q             # 25 tests, about 75 seconds
```
- **Defaults** (goal, dates, challenge list, OneCause headers) live in `build/config.py`.
- **How the tests work:** each test workbook is recalculated by LibreOffice. Its values are compared with an independent Python model in `tests/reference.py`.
- **What's covered:** totals, class/grade/day breakdowns, every challenge result (including ties and date-window edges), donation status rules, and the thermometer fill.
