# Saints Run Wild: Fun Run Spreadsheet

This is a simple Excel file for the San Juan del Rio Catholic School Fun Run. It has two lists you fill in and three dashboards that update on their own.

| Tab | What it is |
|---|---|
| **By Teacher** | Dashboard: raised, goal, % of goal, students donated. Every teacher's students, donated, participation, and $, plus grade totals and a chart. |
| **By Day** | Dashboard: $ per day, gifts, running total, with charts |
| **Daily Challenges** | Dashboard: each flyer day and its winner. Class and grade scores are below. |
| **Donations** | You fill this in: Date, Student, Donor, Amount, Method. Teacher and grade fill in automatically. |
| **Students** | You fill this in: Student, Teacher, Grade. Totals fill in automatically. |

- **Files:** `output/SaintsRunWild_2026.xlsx` (blank) · `output/SaintsRunWild_2026_DEMO.xlsx` (fictional data)
- **How to use it:** `docs/Volunteer_Guide.md`

## Maintainers
```bash
python build/build_workbook.py     # regenerate both files (defaults in build/config.py)
python -m pytest tests -q          # LibreOffice recalculation vs a Python model, about 10 s
```
