"""Workbook formulas (recalculated by LibreOffice) vs the independent Python reference model."""
import pytest

from conftest import num
from reference import compute, serial
from scenarios import THERMO_AMOUNTS

import config as C

MONEY = dict(abs=0.005)


@pytest.fixture(scope="module")
def demo(books):
    return books["demo"]


@pytest.fixture(scope="module")
def ref(demo):
    return compute(demo.data)


def col_values(ws, col, r1, r2):
    return [ws[f"{col}{r}"].value for r in range(r1, r2 + 1)]


# ---------------------------------------------------------------- integrity
@pytest.mark.parametrize("key", ["blank", "demo", "draw", "ties"])
def test_no_formula_errors(books, key):
    allowed_na = {("Calc_Daily", c) for c in "JKL"}  # #N/A on purpose: chart gaps
    bad = []
    for ws in books[key].wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v[:1] == "#" and v.rstrip("!?0/").upper() in (
                        "#REF", "#VALUE", "#N/A", "#NAME", "#DIV", "#NUM", "#NULL", "ERR:502", "#DIV/"):
                    if (ws.title, c.column_letter) not in allowed_na:
                        bad.append(f"{ws.title}!{c.coordinate}={v}")
    assert not bad, bad[:20]


def test_blank_template_is_clean(books):
    b = books["blank"]
    assert b.kpi("Raised") == 0
    assert b.kpi("Students") == 0
    assert b.kpi("IssueCount") == 0
    assert b["Data Check"]["B3"].value == "ALL CLEAR ✔"
    assert b["Challenge Winners"]["G5"].value == "No gifts yet"


# ---------------------------------------------------------------- demo: totals
@pytest.mark.parametrize("key", ["Raised", "RaisedOnline", "RaisedOffline", "Students", "Participants", "Gifts",
                                 "Families", "Donors", "ThermoCells", "RaisedToday", "Remaining"])
def test_kpis(demo, ref, key):
    assert demo.kpi(key) == pytest.approx(ref["kpi"][key], **MONEY)


def test_demo_is_realistic(ref):
    assert 60000 <= ref["kpi"]["Raised"] <= 65000, "demo should sit between the slime line and the goal"


def test_onecause_statuses(demo, ref):
    got = col_values(demo["OC_Norm"], "L", 2, 1 + len(ref["oc_status"]))
    assert got == [s for s, _ in ref["oc_status"]]
    statuses = set(got)
    for expected in ["OK", "UNMATCHED", "DUPLICATE TXN", "NEGATIVE / REFUND", "ZERO AMOUNT", "BAD DATE",
                     "BAD AMOUNT"]:
        assert expected in statuses


def test_cash_statuses(demo, ref):
    got = col_values(demo["Cash_Check"], "J", 2, 1 + len(ref["cash_status"]))
    assert got == ref["cash_status"]
    assert {"OK", "BAD AMOUNT", "NO STUDENT", "BAD DATE"} <= set(got)


def test_students(demo, ref):
    ws = demo["Calc_Students"]
    for s in ref["students"]:
        r = s["row"] + 2
        assert ws[f"B{r}"].value == s["id"]
        assert ws[f"L{r}"].value == pytest.approx(s["online"], **MONEY), s["id"]
        assert ws[f"M{r}"].value == pytest.approx(s["offline"], **MONEY), s["id"]
        assert ws[f"N{r}"].value == pytest.approx(s["total"], **MONEY), s["id"]
        assert ws[f"O{r}"].value == s["gifts"], s["id"]
        assert ws[f"P{r}"].value == s["part"], s["id"]
        assert ws[f"R{r}"].value == s["rank"], s["id"]
        assert ws[f"T{r}"].value == s["prize"], s["id"]
        fg = ws[f"AA{r}"].value
        assert (num(fg) if fg not in ("", None) else None) == s["first_gift"], s["id"]
    # inactive student and blank rows are excluded
    assert ws["A12"].value == 0


def test_alias_and_name_variants_match(demo, ref):
    s3 = ref["students"][3]
    assert s3["online"] > 0, "OneCause alias 'Bobby Tables' should credit roster row 4"


def test_exact_threshold_earns_prize(ref):
    exact = [s for s in ref["students"] if s["total"] == 50]
    assert exact and all(s["prize"] >= 2 for s in exact)


def test_classes(demo, ref):
    ws = demo["Calc_Classes"]
    for i, c in enumerate(ref["classes"]):
        r = i + 2
        assert ws[f"B{r}"].value == c["name"]
        assert ws[f"F{r}"].value == c["students"]
        assert ws[f"G{r}"].value == c["participants"]
        assert ws[f"H{r}"].value == pytest.approx(c["pct"], abs=1e-6)
        assert ws[f"I{r}"].value == pytest.approx(c["total"], **MONEY)


def test_grades(demo, ref):
    ws = demo["Calc_Grades"]
    for i, g in enumerate(C.GRADES):
        r = i + 2
        exp = ref["grades"][g]
        assert str(ws[f"A{r}"].value) == g
        assert ws[f"C{r}"].value == exp["classes"]
        assert ws[f"D{r}"].value == exp["students"]
        assert ws[f"E{r}"].value == exp["participants"]
        assert ws[f"G{r}"].value == pytest.approx(exp["total"], **MONEY)
    assert sum(ref["grades"][g]["total"] for g in C.GRADES) == pytest.approx(ref["kpi"]["Raised"], **MONEY)


def test_families(demo, ref):
    ws = demo["Calc_Families"]
    got = {}
    for r in range(2, 2 + C.N_FAM):
        nm = ws[f"B{r}"].value
        if nm:
            got[nm.lower()] = ws[f"E{r}"].value
    assert set(got) == set(ref["families"])
    for k, v in ref["families"].items():
        assert got[k] == pytest.approx(v, **MONEY), k


def test_donors(demo, ref):
    ws = demo["Calc_Donors"]
    got = {}
    for r in range(2, 2 + C.N_DON):
        key = ws[f"C{r}"].value
        if key:
            got[key] = (ws[f"E{r}"].value, ws[f"D{r}"].value)
    assert set(got) == set(ref["donors"])
    for k, (amt, n) in ref["donors"].items():
        assert got[k][0] == pytest.approx(amt, **MONEY) and got[k][1] == n, k


def test_daily(demo, ref):
    ws = demo["Calc_Daily"]
    days = sorted(ref["daily"])
    for i, d in enumerate(days):
        r = i + 2
        assert num(ws[f"B{r}"].value) == d
        assert ws[f"F{r}"].value == pytest.approx(ref["daily"][d], **MONEY)
        assert ws[f"H{r}"].value == pytest.approx(ref["cum"][d], **MONEY)
    assert ws[f"B{len(days) + 2}"].value in ("", None)
    assert ws["P2"].value == pytest.approx(ref["pre"], **MONEY) and ref["pre"] > 0
    assert ws["P3"].value == pytest.approx(ref["post"], **MONEY) and ref["post"] > 0
    assert ref["cum"][days[-1]] + ref["post"] == pytest.approx(ref["kpi"]["Raised"], **MONEY)


def test_leaderboard_top10(demo, ref):
    ws = demo["TV Dashboard"]
    order = sorted(ref["students"], key=lambda s: (-s["total"], s["row"]))[:10]
    for k, s in enumerate(order, start=1):
        assert ws[f"R{18 + k}"].value == pytest.approx(s["total"], **MONEY)
        assert ws[f"P{18 + k}"].value == f"{s['first']} {s['last'][0]}.", "privacy: first name + last initial"
    classes = sorted(ref["classes"], key=lambda c: -c["total"])[:10]
    for k, c in enumerate(classes, start=1):
        assert ws[f"P{6 + k}"].value == c["name"]


def test_prize_report(demo, ref):
    ws = demo["Prize Report"]
    hr = 7 + C.N_PR + 1
    names = [ws[f"C{r}"].value for r in range(hr + 1, hr + 1 + C.N_STU) if ws[f"C{r}"].value]
    expected = [s for s in ref["students"] if s["prize"] >= 1]
    assert len(names) == len(expected)
    assert set(names) == {s["full"] for s in expected}


def test_data_check(demo, ref):
    ws = demo["Data Check"]
    labels = {ws[f"B{r}"].value.strip(): r for r in range(7, 24) if ws[f"B{r}"].value}
    assert ws[f"D{labels['OneCause reconciliation difference']}"].value == pytest.approx(0, abs=0.005)
    assert ws[f"D{labels['Totals tie-out (students = counted gifts)']}"].value == pytest.approx(0, abs=0.005)
    unmatched = sum(1 for s, _ in ref["oc_status"] if s == "UNMATCHED")
    assert ws[f"C{labels['Unmatched student names']}"].value == unmatched
    assert demo.kpi("IssueCount") == 6  # unmatched, refunds, bad date, bad amount, cash problems, roster row
    listed = [ws[f"C{r}"].value for r in range(29, 69) if ws[f"C{r}"].value]
    assert len(listed) == min(40, sum(1 for s, _ in ref["oc_status"] if s != "OK"))


def test_demo_challenges(demo, ref):
    ws = demo["Calc_Challenges"]
    for i, res in enumerate(ref["challenges"]):
        r = i + 2
        assert ws[f"D{r}"].value == res["name"]
        assert ws[f"K{r}"].value == res["status"]
        m = res["metric"]
        if m in ("CLASS_PARTICIPATION_MIN", "CLASS_PARTICIPATION_CUM"):
            assert ws[f"L{r}"].value == pytest.approx(res["class_max"], abs=1e-6)
            assert ws[f"N{r}"].value == res["class_winners"][0]
            assert ws[f"M{r}"].value == len(res["class_winners"])
        if m == "GRADE_TOTAL":
            assert ws[f"P{r}"].value == pytest.approx(res["grade_max"], **MONEY)
            assert str(ws[f"R{r}"].value) == res["grade_winners"][0]
        if m in ("SCHOOL_TOTAL_TARGET", "SCHOOL_PARTICIPATION_TARGET"):
            assert ws[f"T{r}"].value == pytest.approx(res["value"], abs=0.005)
            assert (ws[f"U{r}"].value == "HIT") == res["hit"]
        if m == "DRAWING":
            assert ws[f"V{r}"].value == len(res["eligible"])
        if m == "CLASS_TOTAL_PER_GRADE":
            for gi, g in enumerate(C.GRADES):
                winners, amt = res["per_grade"][g]
                cell = ws.cell(20 + gi, 2 + 2 * i).value
                assert ws.cell(20 + gi, 3 + 2 * i).value == pytest.approx(amt, **MONEY)
                assert winners[0] in cell
        if m == "TOP_STUDENT_PER_CLASS":
            for ci, c in enumerate(ref["classes"]):
                winners, amt = res["per_class"][c["name"]]
                assert ws.cell(40 + ci, 3 + 2 * i).value == pytest.approx(amt, **MONEY)
                assert winners[0] in ws.cell(40 + ci, 2 + 2 * i).value


def test_slime_milestone_hit(demo, ref):
    ultimate = [r for r in ref["challenges"] if r["metric"] == "SCHOOL_TOTAL_TARGET"][0]
    assert ultimate["hit"]
    assert demo.kpi("Banner").startswith("WE HIT $60,000")


# ---------------------------------------------------------------- drawing
def test_wild_card_drawing(books):
    b = books["draw"]
    ref = compute(b.data)
    res = [r for r in ref["challenges"] if r["metric"] == "DRAWING"][0]
    ws = b["Challenge Winners"]
    winners = [ws[f"C{r}"].value for r in range(28, 68) if ws[f"C{r}"].value]
    assert len(winners) == min(res["target"], len(res["eligible"])) == 10
    assert len(set(winners)) == len(winners), "no duplicate winners"
    eligible_names = {ref["by_id"][sid]["full"] for sid in res["eligible"]}
    assert set(winners) <= eligible_names


# ---------------------------------------------------------------- ties & boundaries
def test_ties_and_window_boundaries(books):
    b = books["ties"]
    ref = compute(b.data)
    ws = b["Calc_Challenges"]
    # 1. class total per grade: grade 3 tie (window end date included), grade 4 has $0 (day after window)
    g3, g4 = 20 + C.GRADES.index("3"), 20 + C.GRADES.index("4")
    assert ws[f"B{g3}"].value == "TIE: Ms. A – 3 / Ms. B – 3" and ws[f"C{g3}"].value == 50
    assert ws[f"B{g4}"].value == "—" and ws[f"C{g4}"].value == 0
    assert ws["B20"].value == "—" and ws["C20"].value == -1, "grade with no classes"
    # 2. top student per class: tie inside Ms. A
    assert ws["D40"].value == "TIE: Ann One / Ben Two" and ws["E40"].value == 30
    assert ws["D41"].value == "Dee Four" and ws["D42"].value == "Eve Five"
    # 3. grade total tie
    assert ws["Y4"].value == "TIE: Grades 3 / 4: $40"
    # 4. participation with $20 minimum
    assert ws["Y5"].value == "Ms. C – 4: 100%"
    # 5/6. HIT when exactly on target
    assert ws["U6"].value == "HIT" and ws["T6"].value == 255
    assert ws["U7"].value == "HIT" and ws["T7"].value == pytest.approx(0.4)
    # 7. three-way tie
    assert ws["Y8"].value == "TIE: Ms. A – 3 / Ms. B – 3 +1 more: 100%"
    for i, res in enumerate(ref["challenges"]):
        assert ws[f"K{i + 2}"].value == res["status"]
    assert b.kpi("Raised") == 255


@pytest.mark.parametrize("amount", THERMO_AMOUNTS)
def test_thermometer(books, amount):
    b = books[f"thermo_{amount}"]
    expected_cells = {0: 0, 32500: 20, 60000: 37, 65000: 40, 80000: 40}[amount]
    assert b.kpi("ThermoCells") == expected_cells
    assert b.kpi("StretchCell") == 37
    banner = b.kpi("Banner")
    if amount >= 65000:
        assert banner == "GOAL REACHED! THANK YOU, SAINTS!"
    elif amount >= 60000:
        assert banner.startswith("WE HIT $60,000")
    else:
        assert banner.startswith(f"${60000 - amount:,} to go")
    assert b.kpi("PctGoal") == pytest.approx(amount / 65000)


def test_reference_serial_helper():
    from datetime import date
    assert serial(date(2026, 10, 30)) == 46325
