"""Workbook values (recalculated by LibreOffice) vs the independent Python model in reference.py."""
import pytest

import config as C
from conftest import num
from reference import compute
from scenarios import THERMO_AMOUNTS

MONEY = dict(abs=0.005)
ERRORS = ("#REF!", "#VALUE!", "#N/A", "#NAME?", "#DIV/0!", "#NUM!", "#NULL!", "Err:")


@pytest.fixture(scope="module")
def demo(books):
    return books["demo"]


@pytest.fixture(scope="module")
def ref(demo):
    return compute(demo.data)


@pytest.mark.parametrize("key", ["blank", "demo", "ties"])
def test_no_formula_errors(books, key):
    bad = [f"{ws.title}!{c.coordinate}={c.value}" for ws in books[key].wb.worksheets for row in ws.iter_rows()
           for c in row if isinstance(c.value, str) and c.value.startswith(ERRORS)]
    assert not bad, bad[:20]


def test_blank_template(books):
    b = books["blank"]
    assert b.kpi("Raised") == 0 and b.kpi("StudentsN") == 0 and b.kpi("Problems") == 0
    assert b["Daily Challenges"]["D5"].value == "No gifts yet"


@pytest.mark.parametrize("key", ["Raised", "Remaining", "StudentsN", "DonatedN", "Gifts", "RaisedToday",
                                 "ThermoCells"])
def test_totals(demo, ref, key):
    assert demo.kpi(key) == pytest.approx(ref["kpi"][key], **MONEY)


def test_demo_lands_between_slime_and_goal(ref):
    assert 60000 <= ref["kpi"]["Raised"] < 65000


def test_who_donated_statuses(demo, ref):
    ws = demo["All Donations"]
    oc = [ws[f"H{r}"].value for r in range(2, 2 + len(ref["oc_status"]))]
    assert oc == ref["oc_status"]
    cash0 = 2 + C.N_OC
    cash = [ws[f"H{r}"].value for r in range(cash0, cash0 + len(ref["cash_status"]))]
    assert cash == ref["cash_status"]
    for s in ["OK", "NAME NOT ON ROSTER", "Duplicate (ignored)", "Refund (not counted)", "Zero amount",
              "Bad date", "Bad amount"]:
        assert s in oc + cash


def test_students(demo, ref):
    ws = demo["Students"]
    days = sorted(ref["daily"])
    for s in ref["students"]:
        r = s["row"] + 2
        assert ws[f"A{r}"].value == s["name"]
        assert ws[f"B{r}"].value == s["cls"]
        assert ws[f"D{r}"].value == pytest.approx(s["total"], **MONEY), s["name"]
        assert ws[f"E{r}"].value == pytest.approx(s["online"], **MONEY), s["name"]
        assert ws[f"G{r}"].value == s["gifts"], s["name"]
        for i, d in enumerate(days):
            assert (ws.cell(r, 9 + i).value or 0) == pytest.approx(s["by_day"][d], **MONEY), (s["name"], d)
    assert ref["students"][3]["online"] > 0, "OneCause nickname credited via Roster column F"


def test_by_class(demo, ref):
    ws = demo["By Class"]
    for i, c in enumerate(ref["classes"]):
        r = 5 + i
        assert ws[f"A{r}"].value == c["name"]
        assert (ws[f"D{r}"].value, ws[f"E{r}"].value) == (c["students"], c["donated"])
        assert ws[f"F{r}"].value == pytest.approx(c["pct"])
        assert ws[f"G{r}"].value == pytest.approx(c["total"], **MONEY)
    assert ws[f"A{5 + len(ref['classes'])}"].value in ("", None)


def test_by_grade(demo, ref):
    ws = demo["By Class"]
    r0 = demo.find("By Class", "A", "Grade P3")
    for i, g in enumerate(C.GRADES):
        exp = ref["grades"][g]
        assert ws[f"D{r0 + i}"].value == exp["students"]
        assert ws[f"E{r0 + i}"].value == exp["donated"]
        assert ws[f"G{r0 + i}"].value == pytest.approx(exp["total"], **MONEY)
    assert ws[f"G{r0 + len(C.GRADES)}"].value == pytest.approx(ref["kpi"]["Raised"], **MONEY)


def test_by_day(demo, ref):
    ws = demo["By Day"]
    run = ref["pre"]
    for i, d in enumerate(sorted(ref["daily"])):
        r = 5 + i
        run += ref["daily"][d]
        assert num(ws[f"A{r}"].value) == d
        assert ws[f"D{r}"].value == pytest.approx(ref["daily"][d], **MONEY)
        assert ws[f"F{r}"].value == pytest.approx(run, **MONEY)
    assert ref["pre"] > 0 and ref["post"] > 0
    assert run + ref["post"] == pytest.approx(ref["kpi"]["Raised"], **MONEY)


def test_dashboard_top_classes(demo, ref):
    ws = demo["Dashboard"]
    top = sorted(ref["classes"], key=lambda c: -c["total"])[:10]
    assert [ws[f"N{7 + i}"].value for i in range(10)] == [c["name"] for c in top]
    best = max(ref["students"], key=lambda s: s["total"])
    assert ws["N19"].value.startswith(f"{best['first']} {best['last'][0]}.")


def test_demo_challenges(demo, ref):
    ws = demo["Daily Challenges"]
    for k, res in enumerate(ref["challenges"]):
        r = 5 + k
        assert ws[f"A{r}"].value == res["name"]
        assert ws[f"E{r}"].value == res["status"]
        text = ws[f"D{r}"].value
        if res["rule"] in ("CLASS_PARTICIPATION_MIN", "CLASS_PARTICIPATION_CUM"):
            assert res["winners"][0] in text and f"{round(res['max'] * 100)}%" in text
        elif res["rule"] == "GRADE_TOTAL":
            assert f"Grade {res['winners'][0]}" in text
        elif res["rule"] in ("SCHOOL_PARTICIPATION_TARGET", "SCHOOL_TOTAL_TARGET"):
            assert text.startswith("HIT") == res["hit"]
        elif res["rule"] == "DRAWING":
            assert text.startswith(f"{len(res['eligible'])} students eligible")
            r0 = demo.find("Daily Challenges", "A", f"{res['name']}: eligible") + 2
            listed = [ws[f"A{i}"].value for i in range(r0, r0 + C.N_ELIG) if ws[f"A{i}"].value]
            assert listed == res["eligible"][:C.N_ELIG]
        elif res["rule"] == "CLASS_TOTAL_PER_GRADE":
            r0 = demo.find("Daily Challenges", "A", f"{res['name']}: top class") + 2
            for i, g in enumerate(C.GRADES):
                winners, amt = res["per_grade"][g]
                assert ws[f"C{r0 + i}"].value == pytest.approx(amt, **MONEY)
                assert winners[0] in ws[f"B{r0 + i}"].value
        elif res["rule"] == "TOP_STUDENT_PER_CLASS":
            r0 = demo.find("Daily Challenges", "A", f"{res['name']}: top student") + 2
            for i, c in enumerate(ref["classes"]):
                winners, amt = res["per_class"][c["name"]]
                assert ws[f"A{r0 + i}"].value == c["name"]
                assert ws[f"C{r0 + i}"].value == pytest.approx(amt, **MONEY)
                assert winners[0] in ws[f"B{r0 + i}"].value


def test_ties_and_date_edges(books):
    b = books["ties"]
    ws = b["Daily Challenges"]
    g = b.find("Daily Challenges", "A", "Weekend: top class") + 2
    g3, g4 = g + C.GRADES.index("3"), g + C.GRADES.index("4")
    assert ws[f"B{g3}"].value == "TIE: Ms. A – 3 / Ms. B – 3" and ws[f"C{g3}"].value == 50
    assert ws[f"B{g4}"].value == "—"                         # Eve's gift was the day after the window
    t = b.find("Daily Challenges", "A", "Top student: top student") + 2
    assert ws[f"B{t}"].value == "TIE: Ann One / Ben Two" and ws[f"C{t}"].value == 30
    assert ws["D7"].value == "TIE: Grade 3 / Grade 4 — $40"
    assert ws["D8"].value == "Ms. C – 4 — 100%"
    assert ws["D9"].value == "HIT ✔ $255 of $255"             # exactly on target counts
    assert ws["D10"].value == "HIT ✔ 40% of 40%"
    assert ws["D11"].value == "TIE: Ms. A – 3 / Ms. B – 3 +1 more — 100%"
    assert b.kpi("Raised") == 255


@pytest.mark.parametrize("amount", THERMO_AMOUNTS)
def test_thermometer(books, amount):
    b = books[f"thermo_{amount}"]
    assert b.kpi("ThermoCells") == {0: 0, 32500: 10, 60000: 18, 65000: 20, 80000: 20}[amount]
    assert b.kpi("StretchCell") == 18
    banner = b.kpi("Banner")
    if amount >= 65000:
        assert banner == "GOAL REACHED! THANK YOU, SAINTS!"
    elif amount >= 60000:
        assert banner.startswith("WE HIT $60,000")
    else:
        assert banner.startswith(f"${60000 - amount:,} to go")
