#!/usr/bin/env python3
"""Generate the Saints Run Wild workbook: 3 dashboards + 2 input lists, plain Excel formulas.

    python build/build_workbook.py      # writes output/SaintsRunWild_2026.xlsx and _DEMO.xlsx

Tabs: By Teacher, By Day, Daily Challenges (dashboards) · Donations, Students (you type/paste here)
"""
import os
import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as CL
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.formula import ArrayFormula

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config as C  # noqa: E402
import styles as S  # noqa: E402

# Whole-column references keep the formulas short and readable.
DON = {k: f"Donations!${c}:${c}" for k, c in
       {"date": "A", "student": "B", "amount": "D", "day": "F", "match": "G", "teacher": "H", "grade": "I"}.items()}
STU_LAST = 1 + C.N_STU
STUD = {k: f"Students!${c}$2:${c}${STU_LAST}" for k, c in {"name": "A", "teacher": "B", "total": "D"}.items()}
T0, T_LAST = 8, 7 + C.N_TCH            # By Teacher teacher rows
CH0 = 5                                 # Daily Challenges summary rows start
STUDENT_RULES = {"CLASS_PCT_MIN", "DRAWING", "CLASS_PCT_TO_DATE", "TOP_STUDENT_PER_CLASS", "SCHOOL_PCT_TARGET"}
TO_DATE_RULES = {"CLASS_PCT_TO_DATE", "SCHOOL_PCT_TARGET"}


def put(ws, ref, value=None, *, font=None, fill=None, fmt=None, align=None, border=None):
    c = ws[ref]
    if value is not None:
        c.value = value
    for attr, v in (("font", font), ("number_format", fmt), ("alignment", align), ("border", border)):
        if v is not None:
            setattr(c, attr, v)
    if fill:
        c.fill = S.fill(fill)
    return c


def header(ws, row, labels, col0=1):
    for i, lab in enumerate(labels):
        put(ws, f"{CL(col0 + i)}{row}", lab, font=S.HEADER_FONT, fill=S.NAVY, align=S.CENTER, border=S.BOX)
    ws.row_dimensions[row].height = 32


def title(ws, text, last_col):
    ws.merge_cells(f"A1:{last_col}1")
    put(ws, "A1", text, font=S.display(18, S.WHITE), fill=S.RED, align=S.CENTER)
    ws.row_dimensions[1].height = 34
    ws.sheet_view.showGridLines = False


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def page(ws, landscape=True):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0


def ch_cols(challenges):
    """Students-sheet helper column letter for each challenge that needs per-student $."""
    cols, n = {}, 0
    for k, ch in enumerate(challenges):
        if ch[5] in STUDENT_RULES:
            cols[k] = CL(6 + n)
            n += 1
    return cols


def ch_ref(k, col):
    """Daily Challenges cell for challenge k (0-based): C=start, D=end, G=min/target."""
    return f"'Daily Challenges'!${col}${CH0 + k}"


# ---------------------------------------------------------------- inputs
def build_donations(ws, rows):
    header(ws, 1, ["Date", "Student (as on OneCause)", "Donor", "Amount", "Method",
                   "Day (auto)", "Student on list (auto)", "Teacher (auto)", "Grade (auto)"])
    widths(ws, {"A": 18, "B": 26, "C": 24, "D": 11, "E": 11, "F": 11, "G": 26, "H": 20, "I": 9})
    for col in "FGHI":
        ws[f"{col}1"].fill = S.fill("5B6B9A")
    for r in range(2, 2 + C.N_DON):
        for col in "ABCDE":
            ws[f"{col}{r}"].fill = S.fill(S.INPUT)
        ws[f"A{r}"].number_format = "m/d/yyyy h:mm AM/PM"
        ws[f"D{r}"].number_format = S.FMT_USD2
        ws[f"F{r}"] = f'=IF(A{r}="","",IFERROR(INT(A{r}),"check date"))'
        ws[f"F{r}"].number_format = "ddd m/d"
        ws[f"G{r}"] = f'=IF(B{r}="","",IFERROR(INDEX(Students!$A:$A,MATCH(TRIM(B{r}),Students!$A:$A,0)),"NOT ON LIST"))'
        ws[f"H{r}"] = f'=IF(OR(G{r}="",G{r}="NOT ON LIST"),"",VLOOKUP(G{r},Students!$A:$C,2,FALSE))'
        ws[f"I{r}"] = f'=IF(OR(G{r}="",G{r}="NOT ON LIST"),"",VLOOKUP(G{r},Students!$A:$C,3,FALSE)&"")'
        for col in "FGHI":
            ws[f"{col}{r}"].fill = S.fill(S.CALC)
    for i, d in enumerate(rows):
        for col, key in zip("ABCDE", ["date", "student", "donor", "amount", "method"]):
            ws[f"{col}{2 + i}"] = d.get(key)
    ws.conditional_formatting.add(f"A2:I{1 + C.N_DON}", FormulaRule(
        formula=['$G2="NOT ON LIST"'], fill=S.fill("F8D7DA"), font=S.font(11, True, S.RED)))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:I{1 + C.N_DON}"
    ws.sheet_properties.tabColor = S.TAB_INPUT


def build_students(ws, roster, challenges):
    cols = ch_cols(challenges)
    header(ws, 1, ["Student", "Teacher", "Grade", "Total $ (auto)", "Donated? (auto)"])
    widths(ws, {"A": 24, "B": 20, "C": 8, "D": 13, "E": 11})
    for k, col in cols.items():
        start, end = ch_ref(k, "C"), ch_ref(k, "D")
        label = (f'="$ thru "&TEXT({end},"m/d")' if challenges[k][5] in TO_DATE_RULES else
                 f'="$ "&TEXT({start},"m/d")&IF({end}>{start},"–"&TEXT({end},"m/d"),"")')
        put(ws, f"{col}1", label, font=S.HEADER_FONT, fill="5B6B9A", align=S.CENTER, border=S.BOX)
        ws.column_dimensions[col].width = 12
    for col in "DE":
        ws[f"{col}1"].fill = S.fill("5B6B9A")
    for r in range(2, STU_LAST + 1):
        for col in "ABC":
            ws[f"{col}{r}"].fill = S.fill(S.INPUT)
        ws[f"D{r}"] = f'=IF(A{r}="","",SUMIFS({DON["amount"]},{DON["match"]},A{r}))'
        ws[f"E{r}"] = f'=IF(A{r}="","",IF(D{r}>0,"Yes",""))'
        ws[f"D{r}"].number_format = S.FMT_USD2
        for k, col in cols.items():
            start, end = ch_ref(k, "C"), ch_ref(k, "D")
            when = (f'{DON["day"]},"<="&{end}' if challenges[k][5] in TO_DATE_RULES else
                    f'{DON["day"]},">="&{start},{DON["day"]},"<="&{end}')
            ws[f"{col}{r}"] = f'=IF($A{r}="","",SUMIFS({DON["amount"]},{DON["match"]},$A{r},{when}))'
            ws[f"{col}{r}"].number_format = '"$"#,##0;-"$"#,##0;""'
        for col in ["D", "E"] + list(cols.values()):
            ws[f"{col}{r}"].fill = S.fill(S.CALC)
    for i, s in enumerate(roster):
        ws[f"A{2 + i}"], ws[f"B{2 + i}"], ws[f"C{2 + i}"] = s["name"], s["teacher"], s["grade"]
        ws[f"C{2 + i}"].number_format = "@"
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{CL(5 + len(cols))}{STU_LAST}"
    ws.sheet_properties.tabColor = S.TAB_INPUT


# ---------------------------------------------------------------- dashboards
def build_by_teacher(wb, ws, teachers):
    title(ws, "SAINTS RUN WILD  •  BY TEACHER", "H")
    widths(ws, {"A": 22, "B": 14, "C": 12, "D": 12, "E": 14, "F": 14, "G": 9, "H": 3})
    for i, (lab, f, fmt) in enumerate([
            ("Raised", "=SUM(Donations!$D:$D)", S.FMT_USD), ("Goal", C.GOAL, S.FMT_USD),
            ("% of goal", "=IF(B4>0,A4/B4,0)", S.FMT_PCT), ("To go", "=MAX(0,B4-A4)", S.FMT_USD),
            ("Students donated", '=IFERROR(COUNTIF(Students!$E:$E,"Yes")/(COUNTA(Students!$A:$A)-1),0)',
             S.FMT_PCT)]):
        col = "ABCDE"[i]
        put(ws, f"{col}3", lab, font=S.font(11, True, S.WHITE), fill=S.NAVY, align=S.CENTER)
        put(ws, f"{col}4", f, font=S.display(18, S.RED if i != 1 else S.NAVY), fmt=fmt, align=S.CENTER,
            fill=S.INPUT if i == 1 else S.WHITE, border=S.BOX)
    ws.row_dimensions[4].height = 34
    wb.defined_names["Goal"] = DefinedName("Goal", attr_text="'By Teacher'!$B$4")
    nf = f'COUNTIF({DON["match"]},"NOT ON LIST")'
    put(ws, "A5", f'=IF({nf}=0,"✔ Every donation matches a student","⚠ "&{nf}&" donation(s) ("&TEXT(SUMIFS('
        f'{DON["amount"]},{DON["match"]},"NOT ON LIST"),"$#,##0")&") have a student not on the Students list '
        f'(red rows on Donations)")', font=S.font(11, True, S.GREEN))
    ws.conditional_formatting.add("A5", FormulaRule(formula=['LEFT($A$5,1)="⚠"'], font=S.font(11, True, S.RED)))

    header(ws, 7, ["Teacher (type once)", "Grade", "Students", "Donated", "Participation", "Total $", "Rank"])
    for r in range(T0, T_LAST + 1):
        put(ws, f"A{r}", fill=S.INPUT, border=S.BOX)
        f = {"B": f'=IF(A{r}="","",IFERROR(INDEX(Students!$C:$C,MATCH(A{r},Students!$B:$B,0))&"",""))',
             "C": f'=IF(A{r}="","",COUNTIF(Students!$B:$B,A{r}))',
             "D": f'=IF(A{r}="","",COUNTIFS(Students!$B:$B,A{r},Students!$D:$D,">0"))',
             "E": f'=IF(OR(A{r}="",N(C{r})=0),"",D{r}/C{r})',
             "F": f'=IF(A{r}="","",SUMIFS({DON["amount"]},{DON["teacher"]},A{r}))',
             "G": f'=IF(A{r}="","",COUNTIF($F${T0}:$F${T_LAST},">"&F{r})+1)'}
        for col, v in f.items():
            put(ws, f"{col}{r}", v, border=S.BOX, align=S.CENTER if col in "BCDG" else None)
        ws[f"E{r}"].number_format = S.FMT_PCT
        ws[f"F{r}"].number_format = S.FMT_USD
    for i, t in enumerate(teachers):
        ws[f"A{T0 + i}"] = t
    ws.conditional_formatting.add(f"A{T0}:G{T_LAST}", FormulaRule(formula=[f"$G{T0}=1"], fill=S.fill("FFF2CC")))

    g0 = T_LAST + 3
    header(ws, g0 - 1, ["Grade", "", "Students", "Donated", "Participation", "Total $", "Rank"])
    for i, g in enumerate(C.GRADES):
        r = g0 + i
        put(ws, f"A{r}", f"Grade {g}", font=S.font(11, True), border=S.BOX)
        put(ws, f"B{r}", g, fmt="@", font=S.font(9, color="999999"), border=S.BOX)
        f = {"C": f'=COUNTIF(Students!$C:$C,B{r})', "D": f'=COUNTIFS(Students!$C:$C,B{r},Students!$D:$D,">0")',
             "E": f'=IF(C{r}=0,"",D{r}/C{r})', "F": f'=SUMIFS({DON["amount"]},{DON["grade"]},B{r})',
             "G": f'=COUNTIF($F${g0}:$F${g0 + len(C.GRADES) - 1},">"&F{r})+1'}
        for col, v in f.items():
            put(ws, f"{col}{r}", v, border=S.BOX, align=S.CENTER if col in "CDG" else None)
        ws[f"E{r}"].number_format = S.FMT_PCT
        ws[f"F{r}"].number_format = S.FMT_USD

    n = max(len(teachers), 1)
    bar = BarChart()
    bar.type, bar.title, bar.legend = "bar", "Raised by teacher", None
    bar.add_data(Reference(ws, min_col=6, min_row=T0, max_row=T0 + n - 1))
    bar.set_categories(Reference(ws, min_col=1, min_row=T0, max_row=T0 + n - 1))
    bar.series[0].graphicalProperties.solidFill = S.RED
    bar.x_axis.scaling.orientation = "maxMin"
    bar.y_axis.numFmt = '"$"#,##0'
    bar.y_axis.majorGridlines = None
    bar.width, bar.height = 16, max(8, 0.55 * n)
    ws.add_chart(bar, "I3")
    ws.freeze_panes = "A8"
    page(ws)
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.tabColor = S.RED


def build_by_day(ws):
    title(ws, "SAINTS RUN WILD  •  BY DAY", "F")
    widths(ws, {"A": 14, "B": 14, "C": 9, "D": 16, "E": 11, "F": 3})
    put(ws, "A3", "First day:", font=S.font(11, True))
    put(ws, "B3", C.FIRST_DAY, fill=S.INPUT, fmt="ddd m/d/yyyy", border=S.BOX)
    header(ws, 5, ["Date", "$ raised", "Gifts", "Running total", "% of goal"])
    for i in range(C.N_DAY):
        r = 6 + i
        put(ws, f"A{r}", f"=$B$3+{i}", fmt="ddd m/d", border=S.BOX)
        put(ws, f"B{r}", f'=SUMIFS({DON["amount"]},{DON["day"]},A{r})', fmt=S.FMT_USD, border=S.BOX)
        put(ws, f"C{r}", f'=COUNTIFS({DON["day"]},A{r})', border=S.BOX, align=S.CENTER)
        put(ws, f"D{r}", f'=SUMIFS({DON["amount"]},{DON["day"]},"<="&A{r})', fmt=S.FMT_USD, border=S.BOX)
        put(ws, f"E{r}", f"=IF(Goal>0,D{r}/Goal,0)", fmt=S.FMT_PCT, border=S.BOX)
    last = 5 + C.N_DAY
    put(ws, f"A{last + 1}", "TOTAL", font=S.font(11, True))
    put(ws, f"B{last + 1}", f"=SUM(B6:B{last})", fmt=S.FMT_USD, font=S.font(11, True))
    put(ws, f"C{last + 1}", f"=SUM(C6:C{last})", font=S.font(11, True), align=S.CENTER)
    put(ws, f"A{last + 2}", "Gifts dated outside these days are in the running total only.",
        font=S.font(9, italic=True, color="777777"))

    bar = BarChart()
    bar.title, bar.legend = "$ raised per day", None
    bar.add_data(Reference(ws, min_col=2, min_row=6, max_row=last))
    bar.set_categories(Reference(ws, min_col=1, min_row=6, max_row=last))
    bar.series[0].graphicalProperties.solidFill = S.NAVY
    bar.y_axis.numFmt = '"$"#,##0'
    bar.y_axis.majorGridlines = None
    bar.width, bar.height = 18, 7.5
    ws.add_chart(bar, "G3")
    line = LineChart()
    line.title, line.legend = "Running total", None
    line.add_data(Reference(ws, min_col=4, min_row=6, max_row=last))
    line.set_categories(Reference(ws, min_col=1, min_row=6, max_row=last))
    line.series[0].graphicalProperties.line.solidFill = S.RED
    line.series[0].graphicalProperties.line.width = 32000
    line.y_axis.numFmt = '"$"#,##0'
    line.y_axis.majorGridlines = None
    line.width, line.height = 18, 7.5
    ws.add_chart(line, "G19")
    page(ws)
    ws.sheet_properties.tabColor = S.RED


def build_challenges(ws, challenges):
    title(ws, "SAINTS RUN WILD  •  DAILY CHALLENGES", "G")
    put(ws, "A2", "Dates and Min $ / Target are editable. Winners update automatically.",
        font=S.font(10, italic=True, color="555555"))
    widths(ws, {"A": 28, "B": 8, "C": 11, "D": 11, "E": 38, "F": 50, "G": 12, "H": 14})
    header(ws, 4, ["Challenge", "Day", "Start", "End", "Reward", "Winner / result", "Min $ / Target"])
    scols = ch_cols(challenges)
    n_students = "(COUNTA(Students!$A:$A)-1)"

    # class-scores grid (one row per teacher) and grade grid, built first so results can point at them
    c0 = CH0 + len(challenges) + 4
    # fixed columns so names land in the wide columns: D $, E/F top student, G.. percentages
    grid_cols, pct = {}, iter("GHIJ")
    for k, ch in enumerate(challenges):
        if ch[5] == "CLASS_TOTAL_PER_GRADE":
            grid_cols[k] = ["D"]
        elif ch[5] == "TOP_STUDENT_PER_CLASS":
            grid_cols[k] = ["E", "F"]
        elif ch[5] in ("CLASS_PCT_MIN", "CLASS_PCT_TO_DATE"):
            grid_cols[k] = [next(pct)]
    put(ws, f"A{c0 - 2}", "CLASS SCORES", font=S.display(13, S.RED))
    header(ws, c0 - 1, ["Teacher", "Grade", "Students", "", "", "", ""])
    for k, cols in grid_cols.items():
        nm, rule = challenges[k][1], challenges[k][5]
        labs = ([f"{nm}: top student", f"{nm} $"] if rule == "TOP_STUDENT_PER_CLASS" else
                [f"{nm} {'$' if rule == 'CLASS_TOTAL_PER_GRADE' else '%'}"])
        for col, lab in zip(cols, labs):
            put(ws, f"{col}{c0 - 1}", lab, font=S.HEADER_FONT, fill=S.NAVY, align=S.CENTER, border=S.BOX)
    ws.row_dimensions[c0 - 1].height = 45
    c_last = c0 + C.N_TCH - 1
    for i in range(C.N_TCH):
        r, tr = c0 + i, T0 + i
        put(ws, f"A{r}", f"=IF('By Teacher'!A{tr}=\"\",\"\",'By Teacher'!A{tr})", border=S.BOX)
        put(ws, f"B{r}", f"=IF(A{r}=\"\",\"\",'By Teacher'!B{tr})", border=S.BOX, align=S.CENTER)
        put(ws, f"C{r}", f"=IF(A{r}=\"\",\"\",'By Teacher'!C{tr})", border=S.BOX, align=S.CENTER)
        for k, cols in grid_cols.items():
            rule, s, e = challenges[k][5], ch_ref(k, "C"), ch_ref(k, "D")
            if rule == "CLASS_TOTAL_PER_GRADE":
                f = (f'=IF(A{r}="","",SUMIFS({DON["amount"]},{DON["teacher"]},A{r},{DON["day"]},">="&{s},'
                     f'{DON["day"]},"<="&{e}))')
                put(ws, f"{cols[0]}{r}", f, fmt=S.FMT_USD, border=S.BOX)
            elif rule in ("CLASS_PCT_MIN", "CLASS_PCT_TO_DATE"):
                sc = f"Students!${scols[k]}:${scols[k]}"
                floor = f'">="&MAX(0.01,N({ch_ref(k, "G")}))' if rule == "CLASS_PCT_MIN" else '">0"'
                f = f'=IF(OR(A{r}="",N(C{r})=0),"",COUNTIFS(Students!$B:$B,A{r},{sc},{floor})/C{r})'
                put(ws, f"{cols[0]}{r}", f, fmt=S.FMT_PCT, border=S.BOX, align=S.CENTER)
            elif rule == "TOP_STUDENT_PER_CLASS":
                sc = f"Students!${scols[k]}$2:${scols[k]}${STU_LAST}"
                top = f"{cols[1]}{r}"
                put(ws, top, f'=IF(A{r}="","",_xlfn.MAXIFS({sc},{STUD["teacher"]},A{r}))', fmt=S.FMT_USD,
                    border=S.BOX)
                ref = f"{cols[0]}{r}"
                ws[ref] = ArrayFormula(ref, (
                    f'=IF(A{r}="","",IF(N({top})<=0,"—",INDEX({STUD["name"]},MATCH(1,({STUD["teacher"]}=A{r})*'
                    f'({sc}={top}),0))&IF(COUNTIFS({STUD["teacher"]},A{r},{sc},{top})>1," (tie)","")))'))
                ws[ref].border = S.BOX

    g0 = c_last + 4
    gcols = {}  # C grade $, E winning class (wide), D its $
    for k, ch in enumerate(challenges):
        if ch[5] == "GRADE_TOTAL":
            gcols[k] = ["C"]
        elif ch[5] == "CLASS_TOTAL_PER_GRADE":
            gcols[k] = ["E", "D"]
    put(ws, f"A{g0 - 2}", "GRADE SCORES", font=S.display(13, S.RED))
    header(ws, g0 - 1, ["Grade", "", "", "", ""])
    for k, cols in gcols.items():
        nm = challenges[k][1]
        labs = [f"{nm} $"] if len(cols) == 1 else [f"{nm}: winning class", f"{nm} $"]
        for col, lab in zip(cols, labs):
            put(ws, f"{col}{g0 - 1}", lab, font=S.HEADER_FONT, fill=S.NAVY, align=S.CENTER, border=S.BOX)
    ws.row_dimensions[g0 - 1].height = 45
    g_last = g0 + len(C.GRADES) - 1
    for i, g in enumerate(C.GRADES):
        r = g0 + i
        put(ws, f"A{r}", g, fmt="@", font=S.font(11, True), border=S.BOX, align=S.CENTER)
        for k, cols in gcols.items():
            s, e = ch_ref(k, "C"), ch_ref(k, "D")
            if len(cols) == 1:
                f = f'=SUMIFS({DON["amount"]},{DON["grade"]},A{r},{DON["day"]},">="&{s},{DON["day"]},"<="&{e})'
                put(ws, f"{cols[0]}{r}", f, fmt=S.FMT_USD, border=S.BOX)
            else:
                amt = grid_cols[k][0]
                arng, grng, trng = (f"${c}${c0}:${c}${c_last}" for c in (amt, "B", "A"))
                top = f"{cols[1]}{r}"
                put(ws, top, f"=_xlfn.MAXIFS({arng},{grng},A{r})", fmt=S.FMT_USD, border=S.BOX)
                ref = f"{cols[0]}{r}"
                ws[ref] = ArrayFormula(ref, (
                    f'=IF(N({top})<=0,"—",INDEX({trng},MATCH(1,({grng}&""=A{r})*({arng}={top}),0))&'
                    f'IF(COUNTIFS({grng},A{r},{arng},{top})>1," (tie)",""))'))
                ws[ref].border = S.BOX

    # summary rows
    for k, (day, nm, start, end, reward, rule, val) in enumerate(challenges):
        r = CH0 + k
        for col, v in zip("ABCDEG", [nm, day, start, end, reward, val]):
            put(ws, f"{col}{r}", v, border=S.BOX)
        for col in "CDG":
            ws[f"{col}{r}"].fill = S.fill(S.INPUT)
        ws[f"C{r}"].number_format = ws[f"D{r}"].number_format = "m/d"
        ws[f"A{r}"].font = S.font(11, True)
        target, s, e = f"N(G{r})", f"C{r}", f"D{r}"
        if rule == "CLASS_TOTAL_PER_GRADE":
            amt = f"${gcols[k][1]}${g0}:${gcols[k][1]}${g_last}"
            res = f'=IF(COUNTIF({amt},">0")=0,"No gifts yet","See GRADE SCORES below")'
        elif rule in ("CLASS_PCT_MIN", "CLASS_PCT_TO_DATE"):
            col = f"${grid_cols[k][0]}${c0}:${grid_cols[k][0]}${c_last}"
            names = f"$A${c0}:$A${c_last}"
            res = (f'=IF(N(MAX({col}))<=0,"No gifts yet",INDEX({names},MATCH(MAX({col}),{col},0))&'
                   f'IF(COUNTIF({col},MAX({col}))>1," (tie)","")&" — "&TEXT(MAX({col}),"0%"))')
        elif rule == "TOP_STUDENT_PER_CLASS":
            amt = f"${grid_cols[k][1]}${c0}:${grid_cols[k][1]}${c_last}"
            res = f'=IF(COUNTIF({amt},">0")=0,"No gifts yet","See CLASS SCORES below")'
        elif rule == "GRADE_TOTAL":
            col = f"${gcols[k][0]}${g0}:${gcols[k][0]}${g_last}"
            names = f"$A${g0}:$A${g_last}"
            res = (f'=IF(MAX({col})<=0,"No gifts yet","Grade "&INDEX({names},MATCH(MAX({col}),{col},0))&'
                   f'IF(COUNTIF({col},MAX({col}))>1," (tie)","")&" — "&TEXT(MAX({col}),"$#,##0"))')
        elif rule == "DRAWING":
            sc = f"Students!${scols[k]}:${scols[k]}"
            res = f'=COUNTIF({sc},">0")&" eligible: draw "&{target}&" (Students col {scols[k]} > $0)"'
        elif rule == "SCHOOL_PCT_TARGET":
            sc = f"Students!${scols[k]}:${scols[k]}"
            v = f'IFERROR(COUNTIF({sc},">0")/{n_students},0)'
            res = f'=IF({v}>={target},"HIT! ","Not yet: ")&TEXT({v},"0%")&" of "&TEXT({target},"0%")'
        elif rule == "SCHOOL_TOTAL_TARGET":
            v = f'SUMIFS({DON["amount"]},{DON["day"]},"<="&{e})'
            res = f'=IF({v}>={target},"HIT! ","Not yet: ")&TEXT({v},"$#,##0")&" of "&TEXT({target},"$#,##0")'
        else:
            res = "—"
        put(ws, f"F{r}", res, font=S.font(11, True, S.NAVY), border=S.BOX)
        ws[f"G{r}"].number_format = "0%" if rule == "SCHOOL_PCT_TARGET" else S.FMT_USD
        if val is None:
            ws[f"G{r}"].fill = S.fill(S.WHITE)
    ws.conditional_formatting.add(f"F{CH0}:F{CH0 + len(challenges) - 1}",
                                  FormulaRule(formula=[f'LEFT(F{CH0},3)="HIT"'], fill=S.fill("D4EDDA")))
    page(ws)
    ws.sheet_properties.tabColor = S.RED


def build(path, data=None):
    data = data or {}
    challenges = data.get("challenges", C.CHALLENGES)
    wb = Workbook()
    wb.remove(wb.active)
    names = ["By Teacher", "By Day", "Daily Challenges", "Donations", "Students"]
    sh = {n: wb.create_sheet(n) for n in names}
    build_by_teacher(wb, sh["By Teacher"], data.get("teachers", []))
    build_by_day(sh["By Day"])
    build_challenges(sh["Daily Challenges"], challenges)
    build_donations(sh["Donations"], data.get("donations", []))
    build_students(sh["Students"], data.get("roster", []), challenges)
    if data.get("preview_only"):
        for n in names:
            if n not in data["preview_only"]:
                sh[n].sheet_state = "hidden"
        wb.active = names.index(data["preview_only"][0])
    wb.calculation.fullCalcOnLoad = True
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    wb.save(path)
    return path


if __name__ == "__main__":
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "output")
    print("built", build(os.path.join(out, "SaintsRunWild_2026.xlsx")))
    import demo_data
    print("built", build(os.path.join(out, "SaintsRunWild_2026_DEMO.xlsx"), demo_data.generate()))
