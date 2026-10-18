# Saints Run Wild: Volunteer Guide

**File:** `SaintsRunWild_2026.xlsx` · **Time:** about 5 minutes a day

## Setup (once, before Fri 10/30)
1. **Settings.** Check the goal ($65,000), the stretch goal ($60,000 slime), the dates, and the daily challenge dates and rewards.
2. **Roster.** Paste First, Last, Grade, and Teacher starting at **B2** (column order in `templates/roster_import.csv`). Don't touch the IDs in column A.
3. **OneCause headers.** Paste one OneCause export. On Settings, each header row should say **✔ found**. If one doesn't, type the exact column name OneCause uses.

## Every day
| Step | Tab | Do this |
|---|---|---|
| 1 | **OneCause_Paste** | Click A1, press Ctrl+A, then Delete. Paste the **full** OneCause export at A1. Repeats are ignored automatically. |
| 2 | **Cash_Check** | Add one row per cash or check gift: date, student (pick from the list or type an ID like `S0042`), amount. |
| 3 | **All Donations** | Look for rows in red. **NAME NOT ON ROSTER** means you should type the name exactly as OneCause spells it into **Roster column F**. |
| 4 | **Daily Challenges** | Today's winner is already calculated. Details are below the table. |
| 5 | **Dashboard** | Save the file. Put this tab on the TV (Ctrl+F1 hides the ribbon). |

## Where to find things
| I want… | Tab |
|---|---|
| Who donated (every gift: date, donor, student, class, amount) | **All Donations** (use the filter arrows) |
| Totals by teacher/class, participation, $ per day | **By Class** (grade totals are at the bottom) |
| $ per day and the running total | **By Day** |
| Every student's total and $ per day | **Students** |
| Daily incentive winners | **Daily Challenges** |

## Rules
- **Donated** means a student has at least one counted gift.
- **Gifts that don't count:** refunds, $0, unreadable dates or amounts, duplicate transaction IDs, and names not on the roster. These show in red on All Donations until fixed (duplicates are just ignored).
- **Wild Card Wednesday:** Daily Challenges lists every eligible student. Draw 10 names from that list.
