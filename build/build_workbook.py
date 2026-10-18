#!/usr/bin/env python3
"""Generate the Saints Run Wild fundraising workbook (.xlsx, no macros).

    python build/build_workbook.py          # blank template + demo workbook

Tabs
    Dashboard          TV screen: thermometer, totals, today's challenge, top classes
    Daily Challenges   the flyer's daily incentives, scored automatically
    By Class           every teacher/class: students, donated, %, total, $ per day (+ grade totals)
    By Day             $ per day, running total
    Students           every student: total, gifts, $ per day
    All Donations      who donated: every OneCause + cash/check gift with its status
    Settings*, Roster*, OneCause_Paste*, Cash_Check*   (* = you type/paste here)
"""
import argparse
import os
import sys

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Border, Protection
from openpyxl.utils import get_column_letter as CL
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config as C  # noqa: E402
import styles as S  # noqa: E402

STU_LAST = 1 + C.N_STU
CC_LAST = 1 + C.N_CC
DON_LAST = 1 + C.N_OC + C.N_CC
CLS0, CLS_LAST = 5, 4 + C.N_CLS          # By Class rows
GR0 = CLS_LAST + 4                       # By Class grade summary first row
CH0 = 23                                 # Settings challenge rows start
CH_LAST = CH0 + len(C.CHALLENGES) - 1

DAY_COLS = [CL(9 + i) for i in range(C.N_DAY)]           # Students I..AC
CLS_DAY_COLS = [CL(10 + i) for i in range(C.N_DAY)]      # By Class J..AD
STU_CH = [CL(41 + i) for i in range(len(C.CHALLENGES))]  # Students: $ counted for each challenge
CLS_MET = [CL(33 + i) for i in range(len(C.CHALLENGES))]  # By Class: class score per challenge
CLS_KEY = [CL(46 + i) for i in range(len(C.CHALLENGES))]  # By Class: tie-broken ranking key

WINDOW_RULES = {"CLASS_TOTAL_PER_GRADE", "CLASS_PARTICIPATION_MIN", "DRAWING", "TOP_STUDENT_PER_CLASS",
                "GRADE_TOTAL"}
TODATE_RULES = {"CLASS_PARTICIPATION_CUM", "SCHOOL_PARTICIPATION_TARGET", "SCHOOL_TOTAL_TARGET"}
ZERO_BLANK = '"$"#,##0;-"$"#,##0;""'


def R(sheet, col, r1, r2):
    return f"'{sheet}'!${col}${r1}:${col}${r2}"


def stu(col):
    return R("Students", col, 2, STU_LAST)


def don(col):
    return R("All Donations", col, 2, DON_LAST)


def bc(col):
    return R("By Class", col, CLS0, CLS_LAST)


def ch_cell(col, k):
    """Settings cell for challenge k (1-based)."""
    return f"'Settings'!${col}${CH0 + k - 1}"


# ---------------------------------------------------------------- helpers
def put(ws, ref, value=None, *, font=None, fill=None, fmt=None, align=None, border=None):
    c = ws[ref]
    if value is not None:
        c.value = value
    if font:
        c.font = font
    if fill:
        c.fill = S.fill(fill)
    if fmt:
        c.number_format = fmt
    if align:
        c.alignment = align
    if border:
        c.border = border
    return c


def arr(ws, ref, formula):
    ws[ref] = ArrayFormula(ref, formula)


def header(ws, row, labels, col0=1):
    for i, lab in enumerate(labels):
        put(ws, f"{CL(col0 + i)}{row}", lab, font=S.HEADER_FONT, fill=S.NAVY, align=S.CENTER, border=S.BOX)
    ws.row_dimensions[row].height = 30


def title(ws, text, last_col, sub=None):
    ws.merge_cells(f"A1:{last_col}1")
    put(ws, "A1", text, font=S.display(16, S.WHITE), fill=S.NAVY, align=S.CENTER)
    ws.row_dimensions[1].height = 30
    if sub:
        put(ws, "A2", sub, font=S.font(10, italic=True, color="555555"))


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def hide(ws, cols):
    for col in cols:
        ws.column_dimensions[col].hidden = True


def rules_where(ws, cells, key):
    """Light row lines only on rows that have data (blank formula rows stay clean)."""
    ws.conditional_formatting.add(cells, FormulaRule(formula=[f'{key}<>""'], border=S.BOX))


def dv_list(ws, cells, formula, strict=True):
    dv = DataValidation(type="list", formula1=formula, allow_blank=True, showErrorMessage=True,
                        errorStyle="stop" if strict else "warning")
    ws.add_data_validation(dv)
    dv.add(cells)


def name(wb, nm, target):
    wb.defined_names[nm] = DefinedName(nm, attr_text=target)


def protect(ws):
    ws.protection.sheet = True
    ws.protection.autoFilter = False
    ws.protection.sort = False


def landscape(ws, rows=None):
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.oddFooter.right.text = "Page &P of &N"
    if rows:
        ws.print_title_rows = rows


# ---------------------------------------------------------------- inputs
def build_settings(wb, ws, st, challenges):
    ws.sheet_properties.tabColor = S.TAB_INPUT
    widths(ws, {"A": 26, "B": 30, "C": 12, "D": 12, "E": 40, "F": 14, "G": 52})
    title(ws, "SETTINGS", "G", "Sand cells are editable.")
    rows = [(3, "School", st["school"], None), (4, "Campaign name", st["campaign"], None),
            (5, "Year", st["year"], "0"), (6, "Goal ($)", st["goal"], S.FMT_USD),
            (7, "Stretch goal ($)", st["stretch"], S.FMT_USD), (8, "Stretch goal prize", st["stretch_label"], None),
            (9, "First day", st["start"], S.FMT_DATE), (10, "Last day (Fun Run)", st["end"], S.FMT_DATE),
            (11, "'As of' date (leave blank)", st["asof"], S.FMT_DATE)]
    for r, lab, val, fmt in rows:
        put(ws, f"A{r}", lab, font=S.font(11, True), border=S.BOX)
        put(ws, f"B{r}", val, fill=S.INPUT, border=S.BOX, fmt=fmt, align=S.LEFT)
    for nm, r in [("SchoolName", 3), ("CampaignName", 4), ("CampaignYear", 5), ("Goal", 6), ("Stretch", 7),
                  ("StretchLabel", 8), ("StartDate", 9), ("EndDate", 10)]:
        name(wb, nm, f"'Settings'!$B${r}")

    header(ws, 13, ["OneCause field", "Header in the export", "Found in col", "Status"])
    for i, (field, hdr) in enumerate(C.ONECAUSE_HEADERS):
        r = 14 + i
        put(ws, f"A{r}", field, border=S.BOX)
        put(ws, f"B{r}", hdr, fill=S.INPUT, border=S.BOX)
        put(ws, f"C{r}", f"=IF(TRIM(B{r})=\"\",0,IFERROR(MATCH(TRIM(B{r}),'OneCause_Paste'!$1:$1,0),0))",
            fill=S.CALC, border=S.BOX, align=S.CENTER)
        put(ws, f"D{r}", f"=IF(C{r}>0,\"✔ found\",IF(TRIM(B{r})=\"\",\"(not used)\",\"✖ not found\"))",
            fill=S.CALC, border=S.BOX)
        name(wb, f"OC_{['Date', 'Amt', 'Name', 'Donor', 'Txn', 'Pay'][i]}", f"'Settings'!$C${r}")

    header(ws, CH0 - 1, ["Day", "Challenge", "Start", "End", "Reward", "Min $ / Target", "How it's scored"])
    for k, (day, nm, s, e, reward, rule, val) in enumerate(challenges, start=1):
        r = CH0 + k - 1
        for col, v in zip("ABCDEF", [day, nm, s, e, reward, val]):
            put(ws, f"{col}{r}", v, fill=S.INPUT, border=S.BOX)
        ws[f"C{r}"].number_format = ws[f"D{r}"].number_format = "ddd m/d"
        put(ws, f"G{r}", C.RULE_TEXT[rule], fill=S.CALC, border=S.BOX, font=S.font(10, italic=True))
    for i, g in enumerate(C.GRADES):
        ws[f"J{3 + i}"] = g
    for i, m in enumerate(C.METHODS):
        ws[f"K{3 + i}"] = m
    hide(ws, "JK")
    name(wb, "GradeList", f"'Settings'!$J$3:$J${2 + len(C.GRADES)}")
    name(wb, "MethodList", f"'Settings'!$K$3:$K${2 + len(C.METHODS)}")


def build_roster(ws, roster):
    ws.sheet_properties.tabColor = S.TAB_INPUT
    header(ws, 1, ["ID", "First name", "Last name", "Grade", "Teacher", "OneCause name (only if different)",
                   "Class (auto)"])
    widths(ws, {"A": 8, "B": 16, "C": 18, "D": 8, "E": 22, "F": 30, "G": 26})
    for r in range(2, STU_LAST + 1):
        put(ws, f"A{r}", f"S{r - 1:04d}", fill=S.CALC)
        for col in "BCDEF":
            put(ws, f"{col}{r}", fill=S.INPUT)
        ws[f"D{r}"].number_format = "@"
        put(ws, f"G{r}", f"=IF(TRIM(E{r})=\"\",\"\",TRIM(E{r})&\" – \"&D{r})", fill=S.CALC)
    for i, s in enumerate(roster):
        r = 2 + i
        ws[f"B{r}"], ws[f"C{r}"], ws[f"D{r}"] = s["first"], s["last"], str(s["grade"])
        ws[f"E{r}"], ws[f"F{r}"] = s["teacher"], s.get("alias")
    dv_list(ws, f"D2:D{STU_LAST}", "=GradeList")
    ws.freeze_panes = "B2"


def build_onecause(ws, headers, rows):
    ws.sheet_properties.tabColor = S.TAB_INPUT
    for i, h in enumerate(headers):
        put(ws, f"{CL(i + 1)}1", h, font=S.font(11, True))
        ws.column_dimensions[CL(i + 1)].width = 20
    for r, row in enumerate(rows, start=2):
        for i, v in enumerate(row):
            if v not in (None, ""):
                ws.cell(r, i + 1, v)
    ws.freeze_panes = "A2"


def build_cash(ws, cash):
    ws.sheet_properties.tabColor = S.TAB_INPUT
    header(ws, 1, ["Date", "Student (pick from list or type ID)", "Donor", "Amount", "Method", "Notes"])
    widths(ws, {"A": 12, "B": 46, "C": 24, "D": 11, "E": 10, "F": 30})
    for r in range(2, CC_LAST + 1):
        for col in "ABCDEF":
            put(ws, f"{col}{r}", fill=S.INPUT)
        ws[f"A{r}"].number_format = "m/d/yyyy"
        ws[f"D{r}"].number_format = S.FMT_USD2
    for i, g in enumerate(cash):
        r = 2 + i
        for col, key in zip("ABCDEF", ["date", "student", "donor", "amount", "method", "notes"]):
            ws[f"{col}{r}"] = g.get(key)
    dv_list(ws, f"B2:B{CC_LAST}", f"={stu('AM')}", strict=False)
    dv_list(ws, f"E2:E{CC_LAST}", "=MethodList")
    ws.freeze_panes = "A2"


# ---------------------------------------------------------------- all donations (the ledger)
def build_donations(ws):
    """One row per OneCause row (2..N_OC+1) then one per Cash_Check row. Hidden cols I-U do the parsing."""
    ws.sheet_properties.tabColor = S.TAB_OUTPUT
    header(ws, 1, ["Date", "Donor", "Student", "Class", "Grade", "Amount", "Source", "Status",
                   "Raw date", "Raw amount", "Raw student", "Txn ID", "Pay type", "Date #", "Amount #",
                   "Roster row", "", "Student ID", "Counted $", "Counted date", "Donor raw"])
    widths(ws, {"A": 11, "B": 24, "C": 24, "D": 24, "E": 7, "F": 11, "G": 11, "H": 22})

    def pull(col, r):
        idx = f"INDEX('OneCause_Paste'!$A${r}:$AZ${r},1,{col})"
        return f"IF({col}=0,\"\",IF({idx}=\"\",\"\",{idx}))"

    for r in range(2, DON_LAST + 1):
        f = {}
        if r <= 1 + C.N_OC:
            f["I"] = "=" + pull("OC_Date", r)
            f["J"] = "=" + pull("OC_Amt", r)
            f["K"] = "=" + pull("OC_Name", r)
            f["L"] = "=" + pull("OC_Txn", r)
            f["M"] = "=" + pull("OC_Pay", r)
            f["U"] = "=" + pull("OC_Donor", r)
            key = f"LOWER(TRIM(K{r}))"
            f["P"] = (f"=IF(TRIM(K{r})=\"\",\"\",IFERROR(MATCH({key},{stu('AH')},0),IFERROR(MATCH({key},"
                      f"{stu('AF')},0),IFERROR(MATCH({key},{stu('AG')},0),\"\"))))")
            f["G"] = f"=IF(H{r}=\"\",\"\",\"OneCause\")"
            f["H"] = (f"=IF(AND(I{r}=\"\",J{r}=\"\",K{r}=\"\"),\"\",IF(NOT(ISNUMBER(O{r})),\"Bad amount\","
                      f"IF(O{r}<0,\"Refund (not counted)\",IF(O{r}=0,\"Zero amount\",IF(NOT(ISNUMBER(N{r})),"
                      f"\"Bad date\",IF(P{r}=\"\",\"NAME NOT ON ROSTER\",IF(L{r}=\"\",\"OK\",IF(MATCH(L{r},"
                      f"$L$2:$L${1 + C.N_OC},0)<ROW()-1,\"Duplicate (ignored)\",\"OK\"))))))))")
        else:
            c = r - C.N_OC
            f["I"] = f"=IF('Cash_Check'!A{c}=\"\",\"\",'Cash_Check'!A{c})"
            f["J"] = f"=IF('Cash_Check'!D{c}=\"\",\"\",'Cash_Check'!D{c})"
            f["K"] = f"=TRIM('Cash_Check'!B{c})"
            f["M"] = f"=IF(TRIM('Cash_Check'!E{c})=\"\",\"Cash/check\",TRIM('Cash_Check'!E{c}))"
            f["U"] = f"=TRIM('Cash_Check'!C{c})"
            sid = f"IF(ISNUMBER(FIND(\"#\",K{r})),TRIM(MID(K{r},FIND(\"#\",K{r})+1,20)),UPPER(K{r}))"
            f["P"] = f"=IF(K{r}=\"\",\"\",IFERROR(MATCH({sid},{stu('AD')},0),\"\"))"
            f["G"] = f"=IF(H{r}=\"\",\"\",M{r})"
            f["H"] = (f"=IF(AND(I{r}=\"\",J{r}=\"\",K{r}=\"\"),\"\",IF(OR(NOT(ISNUMBER(O{r})),N(O{r})<=0),"
                      f"\"Bad amount\",IF(NOT(ISNUMBER(N{r})),\"Bad date\",IF(P{r}=\"\",\"NAME NOT ON ROSTER\","
                      f"\"OK\"))))")
        f["N"] = (f"=IF(I{r}=\"\",\"\",IF(ISNUMBER(I{r}),INT(I{r}),IFERROR(INT(DATEVALUE(I{r})),"
                  f"IFERROR(INT(VALUE(I{r})),\"bad\"))))")
        f["O"] = (f"=IF(J{r}=\"\",\"\",IF(ISNUMBER(J{r}),ROUND(J{r},2),IFERROR(ROUND(VALUE(SUBSTITUTE(SUBSTITUTE("
                  f"SUBSTITUTE(J{r},\"$\",\"\"),\",\",\"\"),\" \",\"\")),2),\"bad\")))")
        f["R"] = f"=IF(P{r}=\"\",\"\",INDEX({stu('AD')},P{r}))"
        f["S"] = f"=IF(H{r}=\"OK\",O{r},0)"
        f["T"] = f"=IF(H{r}=\"OK\",N{r},\"\")"
        f["A"] = f"=IF(H{r}=\"\",\"\",IF(ISNUMBER(N{r}),N{r},I{r}))"
        f["B"] = f"=IF(H{r}=\"\",\"\",U{r})"
        f["C"] = f"=IF(H{r}=\"\",\"\",IF(P{r}=\"\",K{r},INDEX({stu('A')},P{r})))"
        f["D"] = f"=IF(OR(H{r}=\"\",P{r}=\"\"),\"\",INDEX({stu('B')},P{r}))"
        f["E"] = f"=IF(OR(H{r}=\"\",P{r}=\"\"),\"\",INDEX({stu('C')},P{r}))"
        f["F"] = f"=IF(H{r}=\"\",\"\",IF(ISNUMBER(O{r}),O{r},J{r}))"
        for col, v in f.items():
            ws[f"{col}{r}"] = v
        ws[f"A{r}"].number_format = "m/d/yyyy"
        ws[f"F{r}"].number_format = S.FMT_USD2
    hide(ws, [CL(i) for i in range(9, 22)])
    ws.conditional_formatting.add(f"A2:H{DON_LAST}", FormulaRule(
        formula=['AND($H2<>"",$H2<>"OK",LEFT($H2,9)<>"Duplicate")'], fill=S.fill("F8D7DA"),
        font=S.font(11, True, S.RED)))
    rules_where(ws, f"A2:H{DON_LAST}", "$H2")
    ws.auto_filter.ref = f"A1:H{DON_LAST}"
    ws.freeze_panes = "A2"
    landscape(ws, "1:1")
    protect(ws)


# ---------------------------------------------------------------- students
def build_students(ws, challenges, days):
    ws.sheet_properties.tabColor = S.TAB_OUTPUT
    header(ws, 1, ["Student", "Class", "Grade", "Total", "OneCause", "Cash/check", "Gifts", "Donated?"])
    for i, col in enumerate(DAY_COLS):
        put(ws, f"{col}1", f"=IF(StartDate+{i}<=EndDate,StartDate+{i},\"\")", font=S.HEADER_FONT, fill=S.NAVY,
            align=S.CENTER, fmt="ddd m/d", border=S.BOX)
        ws.column_dimensions[col].width = 10
    hide(ws, DAY_COLS[days:])  # spare day columns (unhide if the campaign gets longer)
    widths(ws, {"A": 22, "B": 24, "C": 7, "D": 11, "E": 11, "F": 11, "G": 7, "H": 9})
    for i, lab in enumerate(["ID", "Valid", "key first last", "key last, first", "key OneCause", "class first",
                             "class #", "top key", "display", "pick label", "teacher"]):
        ws.cell(1, 30 + i, lab)
    for k, ch in enumerate(challenges, start=1):
        ws[f"{STU_CH[k - 1]}1"] = f"ch{k} {ch[1]}"
    for r in range(2, STU_LAST + 1):
        ro = lambda c: f"'Roster'!{c}{r}"  # noqa: E731
        f = {
            "AE": f"=IF(TRIM({ro('B')}&{ro('C')})<>\"\",1,0)",
            "AD": f"=IF(AE{r}=1,{ro('A')},\"\")",
            "A": f"=IF(AE{r}=1,TRIM(TRIM({ro('B')})&\" \"&TRIM({ro('C')})),\"\")",
            "AN": f"=IF(AE{r}=1,TRIM({ro('E')}),\"\")",
            "C": f"=IF(AE{r}=1,TRIM({ro('D')}&\"\"),\"\")",
            "B": f"=IF(AE{r}=1,IF(AN{r}=\"\",\"(no teacher)\",AN{r}&\" – \"&C{r}),\"\")",
            "AF": f"=IF(AE{r}=1,LOWER(A{r}),\"#\")",
            "AG": f"=IF(AE{r}=1,LOWER(TRIM(TRIM({ro('C')})&\", \"&TRIM({ro('B')}))),\"#\")",
            "AH": f"=IF(AND(AE{r}=1,TRIM({ro('F')})<>\"\"),LOWER(TRIM({ro('F')})),\"#\")",
            "D": f"=IF(AE{r}=1,ROUND(SUMIFS({don('S')},{don('R')},AD{r}),2),\"\")",
            "E": f"=IF(AE{r}=1,ROUND(SUMIFS({don('S')},{don('R')},AD{r},{don('G')},\"OneCause\"),2),\"\")",
            "F": f"=IF(AE{r}=1,D{r}-E{r},\"\")",
            "G": f"=IF(AE{r}=1,COUNTIFS({don('R')},AD{r},{don('H')},\"OK\"),\"\")",
            "H": f"=IF(AE{r}<>1,\"\",IF(D{r}>0,\"Yes\",\"\"))",
            "AI": f"=IF(AE{r}=1,IF(MATCH(B{r},$B$2:$B${STU_LAST},0)=ROW()-1,1,0),0)",
            "AJ": f"=AI{r}" if r == 2 else f"=AJ{r - 1}+AI{r}",
            "AK": f"=IF(AE{r}=1,D{r}+(1000-ROW())/10^7,-1)",
            "AL": f"=IF(AE{r}=1,TRIM({ro('B')})&\" \"&LEFT(TRIM({ro('C')}),1)&\".\",\"\")",
            "AM": f"=IF(AE{r}=1,TRIM({ro('C')})&\", \"&TRIM({ro('B')})&\" (\"&B{r}&\") #\"&AD{r},\"\")",
        }
        for col in DAY_COLS:
            f[col] = f"=IF(OR($AE{r}<>1,{col}$1=\"\"),\"\",SUMIFS({don('S')},{don('R')},$AD{r},{don('T')},{col}$1))"
        for k, ch in enumerate(challenges, start=1):
            rule = ch[5]
            if rule in WINDOW_RULES:
                crit = f"{don('T')},\">=\"&{ch_cell('C', k)},{don('T')},\"<=\"&{ch_cell('D', k)}"
            elif rule in TODATE_RULES:
                crit = f"{don('T')},\"<=\"&{ch_cell('D', k)}"
            else:
                continue
            f[STU_CH[k - 1]] = f"=IF(AE{r}=1,ROUND(SUMIFS({don('S')},{don('R')},AD{r},{crit}),2),0)"
        for col, v in f.items():
            ws[f"{col}{r}"] = v
        for col in "DEF":
            ws[f"{col}{r}"].number_format = S.FMT_USD2
        for col in DAY_COLS:
            ws[f"{col}{r}"].number_format = ZERO_BLANK
    hide(ws, [CL(i) for i in range(30, 41 + len(challenges))])
    rules_where(ws, f"A2:{DAY_COLS[-1]}{STU_LAST}", "$A2")
    ws.auto_filter.ref = f"A1:H{STU_LAST}"
    ws.freeze_panes = "B2"
    landscape(ws, "1:1")
    protect(ws)


# ---------------------------------------------------------------- by class (+ grade totals)
def build_by_class(ws, challenges, days):
    ws.sheet_properties.tabColor = S.TAB_OUTPUT
    title(ws, "BY TEACHER / CLASS", "I")
    put(ws, "A2", "=\"As of \"&TEXT(AsOfDate,\"ddd m/d\")&\"   •   $ per day to the right   •   grade totals below\"",
        font=S.font(10, italic=True, color="555555"))
    header(ws, 4, ["Class", "Teacher", "Grade", "Students", "Donated", "Participation", "Total $", "Avg / student",
                   "Rank"])
    for i, col in enumerate(CLS_DAY_COLS):
        put(ws, f"{col}4", f"='Students'!{DAY_COLS[i]}1", font=S.HEADER_FONT, fill=S.NAVY, align=S.CENTER,
            fmt="ddd m/d", border=S.BOX)
        ws.column_dimensions[col].width = 10
    hide(ws, CLS_DAY_COLS[days:])
    widths(ws, {"A": 24, "B": 20, "C": 7, "D": 9, "E": 9, "F": 12, "G": 12, "H": 12, "I": 7})
    for r in range(CLS0, CLS_LAST + 1):
        k = r - CLS0 + 1
        f = {
            "AE": f"=IFERROR(MATCH({k},{stu('AJ')},0),\"\")",
            "A": f"=IF(AE{r}=\"\",\"\",INDEX({stu('B')},AE{r}))",
            "B": f"=IF(AE{r}=\"\",\"\",INDEX({stu('AN')},AE{r}))",
            "C": f"=IF(AE{r}=\"\",\"\",INDEX({stu('C')},AE{r}))",
            "D": f"=IF(A{r}=\"\",\"\",COUNTIFS({stu('B')},A{r},{stu('AE')},1))",
            "E": f"=IF(A{r}=\"\",\"\",COUNTIFS({stu('B')},A{r},{stu('D')},\">0\"))",
            "F": f"=IF(A{r}=\"\",\"\",IF(D{r}>0,E{r}/D{r},0))",
            "G": f"=IF(A{r}=\"\",\"\",ROUND(SUMIFS({stu('D')},{stu('B')},A{r}),2))",
            "H": f"=IF(A{r}=\"\",\"\",IF(D{r}>0,G{r}/D{r},0))",
            "I": f"=IF(A{r}=\"\",\"\",COUNTIF($G${CLS0}:$G${CLS_LAST},\">\"&G{r})+1)",
            "AF": f"=IF(A{r}=\"\",-1,G{r}+(100-ROW())/10^7)",
        }
        for i, col in enumerate(CLS_DAY_COLS):
            f[col] = (f"=IF(OR($A{r}=\"\",{col}$4=\"\"),\"\",SUMIFS('Students'!{DAY_COLS[i]}$2:{DAY_COLS[i]}"
                      f"${STU_LAST},{stu('B')},$A{r}))")
        for kk, ch in enumerate(challenges, start=1):
            rule, h = ch[5], stu(STU_CH[kk - 1])
            met, key = CLS_MET[kk - 1], CLS_KEY[kk - 1]
            if rule == "CLASS_TOTAL_PER_GRADE":
                f[met] = f"=IF($A{r}=\"\",-1,ROUND(SUMIFS({h},{stu('B')},$A{r}),2))"
            elif rule in ("CLASS_PARTICIPATION_MIN", "CLASS_PARTICIPATION_CUM"):
                mn = f"MAX(N({ch_cell('F', kk)}),0.01)" if rule.endswith("MIN") else "0.01"
                f[met] = (f"=IF($A{r}=\"\",-1,IF(N($D{r})=0,0,ROUND(COUNTIFS({stu('B')},$A{r},{h},\">=\"&{mn})"
                          f"/$D{r},6)))")
                f[key] = f"=IF({met}{r}<0,-1,{met}{r}+N($G{r})/10^9+(100-ROW())/10^13)"
        for col, v in f.items():
            ws[f"{col}{r}"] = v
        ws[f"F{r}"].number_format = S.FMT_PCT
        ws[f"G{r}"].number_format = ws[f"H{r}"].number_format = S.FMT_USD2
        for col in CLS_DAY_COLS:
            ws[f"{col}{r}"].number_format = ZERO_BLANK
    rules_where(ws, f"A{CLS0}:{CLS_DAY_COLS[-1]}{CLS_LAST}", f"$A{CLS0}")

    put(ws, f"A{GR0 - 2}", "BY GRADE", font=S.display(13, S.RED))
    header(ws, GR0 - 1, ["Grade", "", "Code", "Students", "Donated", "Participation", "Total $", "Avg / student",
                         "Rank"])
    for col in CLS_DAY_COLS:
        put(ws, f"{col}{GR0 - 1}", f"={col}$4", font=S.HEADER_FONT, fill=S.NAVY, align=S.CENTER, fmt="ddd m/d")
    last_g = GR0 + len(C.GRADES) - 1
    for i, g in enumerate(C.GRADES):
        r = GR0 + i
        put(ws, f"A{r}", f"Grade {g}", font=S.font(11, True), border=S.BOX)
        put(ws, f"C{r}", g, fmt="@", border=S.BOX)
        vals = {
            "D": (f"=COUNTIFS({stu('C')},C{r},{stu('AE')},1)", "0"),
            "E": (f"=COUNTIFS({stu('C')},C{r},{stu('D')},\">0\")", "0"),
            "F": (f"=IF(D{r}>0,E{r}/D{r},0)", S.FMT_PCT),
            "G": (f"=ROUND(SUMIFS({stu('D')},{stu('C')},C{r}),2)", S.FMT_USD2),
            "H": (f"=IF(D{r}>0,G{r}/D{r},0)", S.FMT_USD2),
            "I": (f"=COUNTIF($G${GR0}:$G${last_g},\">\"&G{r})+1", "0"),
        }
        for col, (v, fmt) in vals.items():
            put(ws, f"{col}{r}", v, border=S.BOX, fmt=fmt)
        for j, col in enumerate(CLS_DAY_COLS):
            put(ws, f"{col}{r}", f"=IF({col}$4=\"\",\"\",SUMIFS('Students'!{DAY_COLS[j]}$2:{DAY_COLS[j]}"
                f"${STU_LAST},{stu('C')},$C{r}))", fmt=ZERO_BLANK, border=S.BOX)
    t = last_g + 1
    put(ws, f"A{t}", "SCHOOL TOTAL", font=S.font(11, True), border=S.BOX)
    for col, v, fmt in [("D", f"=SUM(D{GR0}:D{last_g})", "0"), ("E", f"=SUM(E{GR0}:E{last_g})", "0"),
                        ("F", f"=IF(D{t}>0,E{t}/D{t},0)", S.FMT_PCT), ("G", f"=SUM(G{GR0}:G{last_g})", S.FMT_USD2),
                        ("H", f"=IF(D{t}>0,G{t}/D{t},0)", S.FMT_USD2)]:
        put(ws, f"{col}{t}", v, fmt=fmt, font=S.font(11, True), border=S.BOX)
    for col in CLS_DAY_COLS:
        put(ws, f"{col}{t}", f"=IF({col}$4=\"\",\"\",SUM({col}{GR0}:{col}{last_g}))", fmt=ZERO_BLANK,
            font=S.font(11, True), border=S.BOX)
    hide(ws, [CL(i) for i in range(31, 46 + len(challenges))])
    ws.freeze_panes = "B5"
    landscape(ws, "4:4")
    protect(ws)


# ---------------------------------------------------------------- by day
def build_by_day(ws):
    ws.sheet_properties.tabColor = S.TAB_OUTPUT
    title(ws, "BY DAY", "G")
    header(ws, 4, ["Date", "OneCause $", "Cash/check $", "Total $", "Gifts", "Running total", "% of goal"])
    widths(ws, {"A": 14, "B": 13, "C": 13, "D": 13, "E": 8, "F": 15, "G": 11})
    before = 6 + C.N_DAY
    for i in range(C.N_DAY):
        r = 5 + i
        put(ws, f"A{r}", f"='Students'!{DAY_COLS[i]}$1", fmt="ddd m/d")
        f = {
            "B": f"=IF(A{r}=\"\",\"\",ROUND(SUMIFS({don('S')},{don('T')},A{r},{don('G')},\"OneCause\"),2))",
            "C": f"=IF(A{r}=\"\",\"\",ROUND(SUMIFS({don('S')},{don('T')},A{r}),2)-B{r})",
            "D": f"=IF(A{r}=\"\",\"\",B{r}+C{r})",
            "E": f"=IF(A{r}=\"\",\"\",COUNTIFS({don('T')},A{r}))",
            "F": f"=IF(A{r}=\"\",\"\",ROUND($D${before}+SUM($D$5:D{r}),2))",
            "G": f"=IF(A{r}=\"\",\"\",F{r}/Goal)",
        }
        for col, v in f.items():
            put(ws, f"{col}{r}", v)
        for col in "BCDF":
            ws[f"{col}{r}"].number_format = S.FMT_USD2
        ws[f"G{r}"].number_format = S.FMT_PCT
    rules_where(ws, f"A5:G{4 + C.N_DAY}", "$A5")
    put(ws, f"A{before}", "Before first day", font=S.font(10, italic=True))
    put(ws, f"D{before}", f"=ROUND(SUMIFS({don('S')},{don('T')},\"<\"&StartDate),2)", fmt=S.FMT_USD2)
    put(ws, f"A{before + 1}", "After last day", font=S.font(10, italic=True))
    put(ws, f"D{before + 1}", f"=ROUND(SUMIFS({don('S')},{don('T')},\">\"&EndDate),2)", fmt=S.FMT_USD2)
    put(ws, f"A{before + 2}", "TOTAL", font=S.font(11, True))
    put(ws, f"D{before + 2}", "=Raised", fmt=S.FMT_USD2, font=S.font(11, True))
    landscape(ws)
    protect(ws)


# ---------------------------------------------------------------- daily challenges
def build_challenges(ws, challenges):
    """Summary table (one row per flyer day) + a detail block for each challenge that has winners."""
    ws.sheet_properties.tabColor = S.TAB_OUTPUT
    ws.sheet_view.showGridLines = False
    title(ws, "DAILY CHALLENGES", "F")
    put(ws, "A2", "=\"Updates automatically   •   as of \"&TEXT(AsOfDate,\"ddd m/d\")",
        font=S.font(10, italic=True, color="555555"))
    widths(ws, {"A": 28, "B": 36, "C": 18, "D": 44, "E": 10, "F": 52})
    header(ws, 4, ["Challenge", "Reward", "Dates", "Result", "Status", "How it's scored"])
    n = len(challenges)
    hid = 20  # hidden helper columns T..W
    row = 6 + n
    for k, ch in enumerate(challenges, start=1):
        r = 4 + k
        rule = ch[5]
        s, e, mn = ch_cell("C", k), ch_cell("D", k), ch_cell("F", k)
        h = stu(STU_CH[k - 1])
        put(ws, f"A{r}", f"={ch_cell('B', k)}", font=S.font(11, True), border=S.BOX)
        put(ws, f"B{r}", f"={ch_cell('E', k)}", border=S.BOX)
        put(ws, f"C{r}", f"=TEXT({s},\"ddd m/d\")&IF({e}>{s},\" – \"&TEXT({e},\"ddd m/d\"),\"\")", border=S.BOX)
        put(ws, f"E{r}", f"=IF(AsOfDate<{s},\"Upcoming\",IF(AsOfDate<={e},\"LIVE\",\"Final\"))", border=S.BOX,
            align=S.CENTER)
        put(ws, f"F{r}", f"={ch_cell('G', k)}", border=S.BOX, font=S.font(10, italic=True))
        res = "\"—\""
        if rule in ("CLASS_PARTICIPATION_MIN", "CLASS_PARTICIPATION_CUM"):
            met = R("By Class", CLS_MET[k - 1], CLS0, CLS_LAST)
            T, U, V, W = (CL(hid + i) for i in range(4))
            ws[f"{T}{r}"] = f"=MAX({met})"
            ws[f"{U}{r}"] = f"=COUNTIF({met},{T}{r})"
            ws[f"{V}{r}"] = f"=IFERROR(INDEX({bc('A')},MATCH({T}{r},{met},0)),\"\")"
            arr(ws, f"{W}{r}", (f"=IF({U}{r}<2,\"\",INDEX({bc('A')},SMALL(IF({met}={T}{r},ROW({met})-{CLS0 - 1},"
                                f"999),2)))"))
            res = (f"IF({T}{r}<=0,\"No gifts yet\",IF({U}{r}>1,\"TIE: \"&{V}{r}&\" / \"&{W}{r}&IF({U}{r}>2,\" +\"&"
                   f"({U}{r}-2)&\" more\",\"\"),{V}{r})&\" — \"&TEXT({T}{r},\"0%\"))")
        elif rule == "SCHOOL_PARTICIPATION_TARGET":
            v = f"IF(StudentsN=0,0,COUNTIF({h},\">0\")/StudentsN)"
            res = (f"IF({v}>=N({mn}),\"HIT ✔ \",IF(AsOfDate>{e},\"Missed: \",\"Not yet: \"))&TEXT({v},\"0%\")&"
                   f"\" of \"&TEXT(N({mn}),\"0%\")")
        elif rule == "SCHOOL_TOTAL_TARGET":
            v = f"SUM({h})"
            res = (f"IF({v}>=N({mn}),\"HIT ✔ \",IF(AsOfDate>{e},\"Missed: \",\"Not yet: \"))&TEXT({v},\"$#,##0\")&"
                   f"\" of \"&TEXT(N({mn}),\"$#,##0\")")
        elif rule == "DRAWING":
            res = (f"IF(COUNTIF({h},\">0\")=0,\"No gifts yet\",COUNTIF({h},\">0\")&\" students eligible — draw \"&"
                   f"N({mn})&\" names (list below)\")")

        if rule == "CLASS_TOTAL_PER_GRADE":
            met = R("By Class", CLS_MET[k - 1], CLS0, CLS_LAST)
            put(ws, f"A{row}", f"={ch_cell('B', k)}&\": top class in each grade\"", font=S.display(12, S.RED))
            header(ws, row + 1, ["Grade", "Winning class", "Raised"])
            for gi, g in enumerate(C.GRADES):
                rr = row + 2 + gi
                put(ws, f"A{rr}", f"Grade {g}", border=S.BOX)
                put(ws, f"{CL(hid)}{rr}", g)
                gc = f"${CL(hid)}{rr}"
                arr(ws, f"C{rr}", f"=MAX(IF({bc('C')}={gc},{met},-1))")
                cnt = f"COUNTIFS({bc('C')},{gc},{met},C{rr})"
                pick = f"INDEX({bc('A')},SMALL(IF(({bc('C')}={gc})*({met}=C{rr}),ROW({met})-{CLS0 - 1},999),{{n}}))"
                arr(ws, f"B{rr}", (f"=IF(C{rr}<=0,\"—\",IF({cnt}>1,\"TIE: \",\"\")&{pick.format(n=1)}&IF({cnt}>1,"
                                   f"\" / \"&{pick.format(n=2)},\"\"))"))
                ws[f"B{rr}"].border = ws[f"C{rr}"].border = S.BOX
                ws[f"C{rr}"].number_format = '"$"#,##0.00;;"—"'
            amt = f"C{row + 2}:C{row + 1 + len(C.GRADES)}"
            res = f"IF(COUNTIF({amt},\">0\")=0,\"No gifts yet\",COUNTIF({amt},\">0\")&\" grade winners (see below)\")"
            row += len(C.GRADES) + 3
        elif rule == "TOP_STUDENT_PER_CLASS":
            put(ws, f"A{row}", f"={ch_cell('B', k)}&\": top student in each class\"", font=S.display(12, S.RED))
            header(ws, row + 1, ["Class", "Top student", "Raised"])
            for ci in range(C.N_CLS):
                rr = row + 2 + ci
                put(ws, f"A{rr}", f"='By Class'!A{CLS0 + ci}")
                arr(ws, f"C{rr}", f"=IF(A{rr}=\"\",\"\",MAX(IF({stu('B')}=A{rr},{h},-1)))")
                cnt = f"COUNTIFS({stu('B')},A{rr},{h},C{rr})"
                pick = f"INDEX({stu('A')},SMALL(IF(({stu('B')}=A{rr})*({h}=C{rr}),ROW({h})-1,99999),{{n}}))"
                arr(ws, f"B{rr}", (f"=IF(A{rr}=\"\",\"\",IF(C{rr}<=0,\"—\",IF({cnt}>1,\"TIE: \",\"\")&"
                                   f"{pick.format(n=1)}&IF({cnt}>1,\" / \"&{pick.format(n=2)},\"\")))"))
                ws[f"C{rr}"].number_format = '"$"#,##0.00;;"—"'
            rules_where(ws, f"A{row + 2}:C{row + 1 + C.N_CLS}", f"$A{row + 2}")
            amt = f"C{row + 2}:C{row + 1 + C.N_CLS}"
            res = f"IF(COUNTIF({amt},\">0\")=0,\"No gifts yet\",COUNTIF({amt},\">0\")&\" class winners (see below)\")"
            row += C.N_CLS + 3
        elif rule in ("CLASS_PARTICIPATION_MIN", "CLASS_PARTICIPATION_CUM"):
            key = R("By Class", CLS_KEY[k - 1], CLS0, CLS_LAST)
            met = R("By Class", CLS_MET[k - 1], CLS0, CLS_LAST)
            put(ws, f"A{row}", f"={ch_cell('B', k)}&\": class participation (top 10)\"", font=S.display(12, S.RED))
            header(ws, row + 1, ["Class", "Participation", "Rank"])
            for j in range(1, 11):
                rr = row + 1 + j
                v = f"LARGE({key},{j})"
                put(ws, f"A{rr}", f"=IFERROR(IF({v}<0,\"\",INDEX({bc('A')},MATCH({v},{key},0))),\"\")", border=S.BOX)
                put(ws, f"B{rr}", f"=IFERROR(IF({v}<0,\"\",INDEX({met},MATCH({v},{key},0))),\"\")", fmt=S.FMT_PCT,
                    border=S.BOX)
                put(ws, f"C{rr}", f"=IF(A{rr}=\"\",\"\",{j})", border=S.BOX, align=S.CENTER)
            row += 13
        elif rule == "GRADE_TOTAL":
            put(ws, f"A{row}", f"={ch_cell('B', k)}&\": grade ranking\"", font=S.display(12, S.RED))
            header(ws, row + 1, ["Grade", "Raised", "Rank"])
            g0 = row + 2
            gl = g0 + len(C.GRADES) - 1
            amt_r, key_r, lab_r = (f"${CL(hid + i)}${g0}:${CL(hid + i)}${gl}" for i in range(3))
            for gi, g in enumerate(C.GRADES):
                rr = g0 + gi
                put(ws, f"{CL(hid)}{rr}", f"=ROUND(SUMIFS({h},{stu('C')},\"{g}\"),2)")
                put(ws, f"{CL(hid + 1)}{rr}", f"={CL(hid)}{rr}+(100-ROW())/10^7")
                put(ws, f"{CL(hid + 2)}{rr}", f"Grade {g}")
                hit = f"MATCH(LARGE({key_r},{gi + 1}),{key_r},0)"
                put(ws, f"A{rr}", f"=INDEX({lab_r},{hit})", border=S.BOX)
                put(ws, f"B{rr}", f"=INDEX({amt_r},{hit})", fmt=S.FMT_USD2, border=S.BOX)
                put(ws, f"C{rr}", gi + 1, border=S.BOX, align=S.CENTER)
            top, second = f"B{g0}", f"B{g0 + 1}"
            res = (f"IF({top}<=0,\"No gifts yet\",IF({top}={second},\"TIE: \"&A{g0}&\" / \"&A{g0 + 1},A{g0})&\" — \"&"
                   f"TEXT({top},\"$#,##0\"))")
            row += len(C.GRADES) + 3
        elif rule == "DRAWING":
            put(ws, f"A{row}", f"={ch_cell('B', k)}&\": eligible students (gave that day)\"",
                font=S.display(12, S.RED))
            header(ws, row + 1, ["Student", "Class", "#"])
            for j in range(1, C.N_ELIG + 1):
                rr = row + 1 + j
                arr(ws, f"{CL(hid)}{rr}", f"=IFERROR(SMALL(IF({h}>0,ROW({h})-1,99999),{j}),99999)")
                put(ws, f"A{rr}", f"=IF({CL(hid)}{rr}=99999,\"\",INDEX({stu('A')},{CL(hid)}{rr}))")
                put(ws, f"B{rr}", f"=IF({CL(hid)}{rr}=99999,\"\",INDEX({stu('B')},{CL(hid)}{rr}))")
                put(ws, f"C{rr}", f"=IF(A{rr}=\"\",\"\",{j})", align=S.CENTER)
            rules_where(ws, f"A{row + 2}:C{row + 1 + C.N_ELIG}", f"$A{row + 2}")
            put(ws, f"D{row + 2}", "Draw names from this list (pick random numbers from the # column).",
                font=S.font(10, italic=True))
            row += C.N_ELIG + 3
        put(ws, f"D{r}", "=" + res, border=S.BOX, font=S.font(11, True, S.NAVY))
    ws.conditional_formatting.add(f"E5:E{4 + n}", FormulaRule(formula=['E5="LIVE"'], fill=S.fill(S.GOLD)))
    ws.conditional_formatting.add(f"D5:D{4 + n}", FormulaRule(formula=['LEFT(D5,3)="HIT"'], fill=S.fill("D4EDDA")))
    hide(ws, [CL(hid + i) for i in range(4)])
    landscape(ws)
    protect(ws)


# ---------------------------------------------------------------- calc + dashboard
def build_calc(wb, ws, challenges):
    n = len(challenges)
    starts, ends = f"'Settings'!$C${CH0}:$C${CH_LAST}", f"'Settings'!$D${CH0}:$D${CH_LAST}"
    items = [
        ("AsOfDate", "=IF(ISNUMBER('Settings'!B11),INT('Settings'!B11),TODAY())"),
        ("Raised", f"=ROUND(SUM({don('S')}),2)"),
        ("Remaining", "=MAX(0,Goal-Raised)"),
        ("PctGoal", "=IF(Goal>0,Raised/Goal,0)"),
        ("StudentsN", f"=COUNTIF({stu('AE')},1)"),
        ("DonatedN", f"=COUNTIF({stu('D')},\">0\")"),
        ("PartPct", "=IF(StudentsN>0,DonatedN/StudentsN,0)"),
        ("Gifts", f"=COUNTIF({don('H')},\"OK\")"),
        ("DaysLeft", "=MAX(0,EndDate-AsOfDate)"),
        ("RaisedToday", f"=ROUND(SUMIFS({don('S')},{don('T')},AsOfDate),2)"),
        ("ThermoCells", "=MIN(20,MAX(0,ROUND(PctGoal*20,0)))"),
        ("StretchCell", "=MIN(20,MAX(1,ROUND(Stretch/Goal*20,0)))"),
        ("TodayIdx", None),
        ("NextIdx", None),
        ("ShowIdx", "=IF(TodayIdx<99,TodayIdx,IF(NextIdx<99,NextIdx,0))"),
        ("ShowLabel", f"=IF(TodayIdx<99,\"TODAY'S CHALLENGE\",IF(NextIdx<99,\"NEXT: \"&TEXT(INDEX({starts},NextIdx),"
                      "\"dddd m/d\"),\"THANK YOU, SAINTS!\"))"),
        ("ShowName", f"=IF(ShowIdx=0,\"\",INDEX('Settings'!$B${CH0}:$B${CH_LAST},ShowIdx))"),
        ("ShowReward", f"=IF(ShowIdx=0,\"\",INDEX('Settings'!$E${CH0}:$E${CH_LAST},ShowIdx))"),
        ("ShowResult", f"=IF(ShowIdx=0,\"\",INDEX('Daily Challenges'!$D$5:$D${4 + n},ShowIdx))"),
        ("Unmatched", f"=COUNTIF({don('H')},\"NAME NOT ON ROSTER\")"),
        ("Problems", f"=COUNTIF({don('H')},\"?*\")-COUNTIF({don('H')},\"OK\")-COUNTIF({don('H')},\"Duplicate*\")"),
        ("Banner", "=IF(Raised>=Goal,\"GOAL REACHED! THANK YOU, SAINTS!\",IF(Raised>=Stretch,\"WE HIT \"&"
                   "TEXT(Stretch,\"$#,##0\")&\": \"&StretchLabel,TEXT(Stretch-Raised,\"$#,##0\")&\" to go until \"&"
                   "StretchLabel))"),
    ]
    for i, (nm, formula) in enumerate(items):
        r = 2 + i
        ws[f"A{r}"] = nm
        if formula:
            ws[f"B{r}"] = formula
        name(wb, nm, f"'Calc'!$B${r}")
    tr = 2 + [x[0] for x in items].index("TodayIdx")
    arr(ws, f"B{tr}", f"=MIN(IF(({starts}<=AsOfDate)*({ends}>=AsOfDate),ROW({starts})-{CH0 - 1},99))")
    arr(ws, f"B{tr + 1}", f"=MIN(IF({starts}>AsOfDate,ROW({starts})-{CH0 - 1},99))")
    ws.sheet_state = "hidden"
    protect(ws)


def build_dashboard(ws):
    ws.sheet_properties.tabColor = S.GOLD
    ws.sheet_view.showGridLines = False
    ws.sheet_view.showRowColHeaders = False
    widths(ws, {"A": 2, "B": 9, "C": 9, "D": 22, "E": 2, "F": 13, "G": 13, "H": 13, "I": 13, "J": 13, "K": 13,
                "L": 2, "M": 5, "N": 30, "O": 13, "P": 9, "Q": 2})
    for r in range(1, 32):
        for c in range(1, 18):
            ws.cell(r, c).fill = S.fill(S.CREAM)
    ws.merge_cells("B2:P2")
    put(ws, "B2", "=UPPER(CampaignName)&\" \"&CampaignYear", font=S.display(36, S.RED), align=S.CENTER)
    ws.row_dimensions[2].height = 50
    ws.merge_cells("B3:P3")
    put(ws, "B3", "=SchoolName&\"  •  \"&TEXT(StartDate,\"mmm d\")&\" – \"&TEXT(EndDate,\"mmm d\")",
        font=S.display(14, S.NAVY), align=S.CENTER)
    ws.merge_cells("B4:P4")
    put(ws, "B4", "=Banner", font=S.display(13, S.WHITE), fill=S.GREEN, align=S.CENTER)
    ws.row_dimensions[4].height = 24

    for i in range(20):  # thermometer: 20 cells of 5%, rows 6-25
        r = 6 + i
        idx = 20 - i
        ws.row_dimensions[r].height = 20
        put(ws, f"A{r}", idx, font=S.font(6, color=S.CREAM))
        put(ws, f"B{r}", f"=IF(MOD(A{r},4)=0,TEXT(Goal*A{r}/20,\"$#,##0\"),\"\")", font=S.font(10, True, S.NAVY),
            align=S.RIGHT)
        c = put(ws, f"C{r}", fill="E8E8E8")
        c.border = Border(left=S.thick_navy, right=S.thick_navy, top=S.thick_navy if idx == 20 else None)
        put(ws, f"D{r}", f"=IF(A{r}=20,\"◄ GOAL\",IF(A{r}=StretchCell,\"◄ \"&TEXT(Stretch,\"$#,##0\")&\" SLIME!\",\"\"))",
            font=S.font(11, True, S.RED))
    ws.conditional_formatting.add("C6:C25", FormulaRule(formula=["$A6<=ThermoCells"], fill=S.fill(S.RED)))
    ws.merge_cells("B26:D28")
    put(ws, "B26", "=TEXT(PctGoal,\"0%\")&\" of goal\"", font=S.display(20, S.WHITE), fill=S.RED, align=S.CENTER)

    def tile(lab_rng, val_rng, label, formula, fmt, size, color):
        ws.merge_cells(lab_rng)
        put(ws, lab_rng.split(":")[0], label, font=S.font(12, True, S.WHITE), fill=color, align=S.CENTER)
        ws.merge_cells(val_rng)
        put(ws, val_rng.split(":")[0], formula, font=S.display(size, color), fill=S.WHITE, fmt=fmt, align=S.CENTER)

    tile("F6:H6", "F7:H10", "RAISED SO FAR", "=Raised", S.FMT_USD, 36, S.RED)
    tile("I6:K6", "I7:K10", "OF OUR GOAL", "=PctGoal", S.FMT_PCT, 36, S.NAVY)
    ws.merge_cells("F11:K11")
    put(ws, "F11", "=\"Goal \"&TEXT(Goal,\"$#,##0\")&\"   •   \"&TEXT(Remaining,\"$#,##0\")&\" to go\"",
        font=S.font(12, True, S.NAVY), align=S.CENTER)
    tile("F13:G13", "F14:G16", "STUDENTS DONATED", "=PartPct", S.FMT_PCT, 22, S.GREEN)
    tile("H13:I13", "H14:I16", "RAISED TODAY", "=RaisedToday", S.FMT_USD, 22, S.RED)
    tile("J13:K13", "J14:K16", "DAYS LEFT", "=DaysLeft", "0", 22, S.NAVY)
    ws.merge_cells("F18:K18")
    put(ws, "F18", "=ShowLabel", font=S.display(13, S.WHITE), fill=S.NAVY, align=S.CENTER)
    for rng_, f, fnt in [("F19:K21", "=ShowName", S.display(18, S.RED)),
                         ("F22:K23", "=ShowReward", S.font(13, True, S.NAVY)),
                         ("F24:K25", "=ShowResult", S.font(13, True, S.GREEN))]:
        ws.merge_cells(rng_)
        put(ws, rng_.split(":")[0], f, font=fnt, fill=S.SAND, align=S.CENTER)

    ws.merge_cells("M6:P6")
    put(ws, "M6", "TOP 10 CLASSES", font=S.display(13, S.WHITE), fill=S.RED, align=S.CENTER)
    for k in range(1, 11):
        r = 6 + k
        v = f"LARGE({bc('AF')},{k})"
        hit = f"MATCH({v},{bc('AF')},0)"
        put(ws, f"M{r}", k, font=S.display(11, S.NAVY), align=S.CENTER)
        put(ws, f"N{r}", f"=IF({v}<0.005,\"\",INDEX({bc('A')},{hit}))", font=S.font(12, True))
        put(ws, f"O{r}", f"=IF({v}<0.005,\"\",INDEX({bc('G')},{hit}))", font=S.font(12, True), fmt=S.FMT_USD)
        put(ws, f"P{r}", f"=IF({v}<0.005,\"\",INDEX({bc('F')},{hit}))", font=S.font(11, color=S.GREEN), fmt=S.FMT_PCT)
    ws.merge_cells("M18:P18")
    put(ws, "M18", "TOP 5 FUNDRAISERS", font=S.display(13, S.WHITE), fill=S.NAVY, align=S.CENTER)
    for k in range(1, 6):
        r = 18 + k
        v = f"LARGE({stu('AK')},{k})"
        hit = f"MATCH({v},{stu('AK')},0)"
        put(ws, f"M{r}", k, font=S.display(11, S.NAVY), align=S.CENTER)
        put(ws, f"N{r}", f"=IF({v}<0.005,\"\",INDEX({stu('AL')},{hit})&\"  (\"&INDEX({stu('B')},{hit})&\")\")",
            font=S.font(11, True))
        ws.merge_cells(f"O{r}:P{r}")
        put(ws, f"O{r}", f"=IF({v}<0.005,\"\",INDEX({stu('D')},{hit}))", font=S.font(12, True), fmt=S.FMT_USD)
    ws.merge_cells("B30:P30")
    put(ws, "B30", "=\"Updated \"&TEXT(AsOfDate,\"dddd, mmmm d\")&IF(Problems>0,\"   •   \"&Problems&\" donation \""
        "&\"row(s) need attention on All Donations\",\"\")", font=S.font(10, italic=True, color="777777"),
        align=S.CENTER)
    ws.print_area = "A1:Q30"
    landscape(ws)
    ws.page_setup.fitToHeight = 1
    protect(ws)


def build_howto(ws):
    ws.sheet_properties.tabColor = S.GOLD
    ws.sheet_view.showGridLines = False
    widths(ws, {"A": 3, "B": 10, "C": 100})
    ws.merge_cells("B1:C1")
    put(ws, "B1", "HOW TO USE (about 5 minutes a day)", font=S.display(16, S.WHITE), fill=S.RED, align=S.CENTER)
    ws.row_dimensions[1].height = 30
    steps = [
        ("Setup", "Settings: check goal and dates. Roster: paste First, Last, Grade, Teacher starting at B2 "
                  "(IDs in column A are fixed)."),
        ("1", "OneCause_Paste: click A1, Ctrl+A, Delete, then paste the full OneCause export at A1."),
        ("2", "Cash_Check: one row per cash/check gift (date, pick the student, amount)."),
        ("3", "All Donations: rows in red need a fix. NAME NOT ON ROSTER → type the name exactly as OneCause "
              "spells it in Roster column F."),
        ("4", "Daily Challenges: each day's winner is calculated for you (details below the table)."),
        ("5", "Save. Put the Dashboard tab on the TV (Ctrl+F1 hides the ribbon)."),
        ("Look up", "Who donated → All Donations   •   teacher/class & grade → By Class   •   by day → By Day   •   "
                    "each student → Students"),
    ]
    for i, (k, t) in enumerate(steps):
        r = 3 + i
        put(ws, f"B{r}", k, font=S.font(11, True, S.RED), align=S.CENTER)
        put(ws, f"C{r}", t, align=S.LEFT_WRAP)
        ws.row_dimensions[r].height = 32
    protect(ws)


# ---------------------------------------------------------------- assembly
ORDER = ["Dashboard", "Daily Challenges", "By Class", "By Day", "Students", "All Donations", "How To",
         "Settings", "Roster", "OneCause_Paste", "Cash_Check", "Calc"]


def build(path, data=None):
    data = data or {}
    st = dict(C.SETTINGS, **data.get("settings", {}))
    challenges = data.get("challenges", C.CHALLENGES)
    headers = data.get("oc_headers", [h for _, h in C.ONECAUSE_HEADERS])

    wb = Workbook()
    wb.remove(wb.active)
    sh = {n: wb.create_sheet(n) for n in ORDER}
    build_settings(wb, sh["Settings"], st, challenges)
    build_roster(sh["Roster"], data.get("roster", []))
    build_onecause(sh["OneCause_Paste"], headers, data.get("oc_rows", []))
    build_cash(sh["Cash_Check"], data.get("cash", []))
    build_donations(sh["All Donations"])
    days = min(C.N_DAY, (st["end"] - st["start"]).days + 1)
    build_students(sh["Students"], challenges, days)
    build_by_class(sh["By Class"], challenges, days)
    build_calc(wb, sh["Calc"], challenges)
    build_by_day(sh["By Day"])
    build_challenges(sh["Daily Challenges"], challenges)
    build_dashboard(sh["Dashboard"])
    build_howto(sh["How To"])
    if data.get("preview_only"):  # render check: show just these tabs
        for n in ORDER:
            if n not in data["preview_only"]:
                sh[n].sheet_state = "hidden"
        wb.active = ORDER.index(data["preview_only"][0])
    else:
        wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    wb.save(path)
    return path


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "output"))
    args = ap.parse_args()
    year = C.SETTINGS["year"]
    print("built", build(os.path.join(args.out, f"SaintsRunWild_{year}.xlsx")))
    import demo_data
    print("built", build(os.path.join(args.out, f"SaintsRunWild_{year}_DEMO.xlsx"), demo_data.generate()))


if __name__ == "__main__":
    main()
