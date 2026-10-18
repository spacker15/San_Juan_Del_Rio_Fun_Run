# Saints Run Wild: Volunteer Guide (1 page)

**File:** `SaintsRunWild_2026.xlsx` · **Owner:** Fun Run admin · **Time:** about 5 minutes a day

## One-time setup (before Fri 10/30)
| # | Tab | Do this |
|---|-----|---------|
| 1 | **Settings** | Check the goal ($65,000), the stretch goal ($60,000 slime), the dates (10/30–11/13), and the privacy mode. Leave **"As of" override blank**. |
| 2 | **Classrooms** | One row per Teacher + Grade, for example `Mrs. Smith` / `3`. The Classroom name fills itself in. |
| 3 | **Roster** | Paste First, Last, Grade, Teacher, and Family at **B2** (use the column order in `templates/roster_import.csv`). **Don't edit Student IDs.** Column K must show ✔. |
| 4 | **Prizes** | Replace the sample prizes. List them smallest to largest threshold. |
| 5 | **Challenges** | Already filled in from the Daily Incentives flyer. Change dates or rewards here only. |
| 6 | **History** *(optional)* | Enter last year's totals (row 4) and the daily cumulative amounts (row 16 = year). |
| 7 | **Settings → OneCause mapping** | Paste one real OneCause export, then confirm that every required row says **✔ found**. |

## Every day
1. **OneCause_Paste.** Export the donation report from OneCause. Click cell A1 on this tab, press **Ctrl+A**, then **Delete**. Paste the full export at **A1**, with headers in row 1. *Always paste the full export, not just the new rows. Duplicate transaction IDs are ignored automatically.*
2. **Cash_Check.** Add one row per cash or check gift. Pick the student from the dropdown, or type their ID (for example `S0042`). Enter the amount and method.
3. **Data Check.** The banner must say **ALL CLEAR ✔**. Common fixes:
   | Flag | Fix |
   |---|---|
   | Unmatched student name | Roster column H (**OneCause name**): type the name exactly as OneCause shows it |
   | Negative / refund | Confirm the refund with OneCause. The refund is already excluded. |
   | Cash_Check NO STUDENT | Re-pick the student from the dropdown |
   | Roster "not on Classrooms" | Add that Teacher + Grade on Classrooms, or fix the typo |
4. **Challenge Winners.** Read today's result. Use **Challenge Detail** (the dropdown) for per-grade or per-class winners.
   - **Wild Card drawing:** pick the Wild Card challenge. Copy the winners in **C28:D37**, then use **Paste Values** into **I28** to lock them. The live list reshuffles on every recalculation.
5. **TV Dashboard.** Save the file. On the TV, open the TV Dashboard tab, hide the ribbon (**Ctrl+F1**), and zoom to fit.

## Reports (File → Export → PDF)
`Rpt_Summary` · `Rpt_Grades` · `Rpt_Classes` · `Rpt_Students`\* · `Rpt_Families`\* · `Rpt_Donors`\* · `Rpt_Daily` · `Prize Report` (pickup checklist)
\*Confidential: these show full names and amounts. Never put them on a public screen.

## Privacy rules
- TV Dashboard and Leaderboards show names using the Settings privacy mode. The default is **First name + last initial**.
- Donor names and per-family amounts appear only on the confidential reports.

## Close the campaign (after 11/13)
1. Print Rpt_Summary, Rpt_Grades, Rpt_Daily, and Prize Report to PDF and archive them.
2. In **History**, add a row: year, goal, raised, participation, students, and the 11 grade totals (from Rpt_Grades). Paste the Rpt_Daily **Cumulative** column under a new year column (row 16).
3. **Save As** `SaintsRunWild_2027.xlsx`. Clear OneCause_Paste, Cash_Check, and the Roster names (keep the IDs). Update the Settings dates, goal, and compare-year, then update Challenges.
