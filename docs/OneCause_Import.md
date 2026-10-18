# OneCause → Workbook Import

## How it works
```
OneCause export (CSV/XLSX) ──paste──► OneCause_Paste (raw, unchanged)
                                         │  header names looked up via Settings rows 19–26
                                         ▼
                                  OC_Norm (hidden): date, amount, name → Student ID, status
                                         ▼
                                  Ledger (hidden) ◄── Cash_Check
                                         ▼
                              every total, leaderboard, challenge, and report
```
Columns are matched by **header text**, not by position, so the order of OneCause columns doesn't matter.

## Header mapping (Settings tab)
| Field | Default header | Required | If OneCause uses a different name |
|---|---|---|---|
| Date | `Transaction Date` | ✔ | Type the exact header in Settings B19 |
| Amount | `Amount` | ✔ | B20 |
| Student full name | `Participant Name` | ✔ (or first + last) | B21 |
| Student first / last (if split) | *(blank)* | — | B22 / B23, and clear B21 |
| Donor name | `Donor Name` | — | B24 |
| Transaction ID | `Transaction ID` | recommended | B25 (enables duplicate protection) |
| Payment type | `Payment Method` | — | B26 |

Column D shows **✔ found** or **✖ NOT FOUND** as soon as an export is pasted.

## Name matching rules (in order)
1. Roster column H, **OneCause name**: an exact override, e.g. a nickname ("Bobby Tables").
2. `First Last` from the Roster.
3. `Last, First` from the Roster.

Matching ignores case and extra spaces. Inactive students (Roster Active = N) don't match. Their gifts show as UNMATCHED so the money is never silently dropped.

## Row status rules (first match wins)
| Status | Counted? | Meaning |
|---|---|---|
| BAD AMOUNT | No | Amount is blank or not a number |
| NEGATIVE / REFUND | No | Amount < 0 |
| ZERO AMOUNT | No | Amount = 0 |
| BAD DATE | No | Date can't be read |
| UNMATCHED | No | No roster student matches the name |
| DUPLICATE TXN | No | The same transaction ID appeared earlier in the paste |
| OK | **Yes** | Credited to the student on that date |

Amounts like `$1,250.00` and dates like `11/03/2026 10:15 AM` are converted automatically.

## Reconciliation (Data Check)
OneCause $ in the file = Counted + Unmatched + Duplicates + Refunds + Bad dates. The difference must be **$0.00**.
The totals tie-out (sum of students = counted gifts) must also be **$0.00**.
