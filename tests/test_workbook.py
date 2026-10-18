"""Recalculated workbook values vs the Python model in reference.py."""
from datetime import datetime

import pytest

import config as C
from reference import compute

ERRORS = ("#REF!", "#VALUE!", "#N/A", "#NAME?", "#DIV/0!", "#NUM!", "#NULL!", "Err:")
T0 = 8


@pytest.fixture(scope="module")
def demo(books):
    return books["demo"]


@pytest.fixture(scope="module")
def ref(demo):
    return compute(demo.data)


@pytest.mark.parametrize("key", ["blank", "demo", "ties"])
def test_no_formula_errors(books, key):
    bad = [f"{ws.title}!{c.coordinate}={c.value}" for ws in books[key].wb for row in ws.iter_rows() for c in row
           if isinstance(c.value, str) and c.value.startswith(ERRORS)]
    assert not bad, bad[:20]


def test_blank(books):
    ws = books["blank"]["By Teacher"]
    assert ws["A4"].value == 0 and ws["A5"].value.startswith("✔")
    assert books["blank"]["Daily Challenges"]["F5"].value == "No gifts yet"


def test_summary(demo, ref):
    ws = demo["By Teacher"]
    assert ws["A4"].value == pytest.approx(ref["raised"])
    assert 60000 <= ref["raised"] < 65000
    assert ws["E4"].value == pytest.approx(ref["donated_pct"])
    assert ws["A5"].value.startswith(f"⚠ {ref['not_on_list']} donation")


def test_students(demo, ref):
    ws = demo["Students"]
    for i, s in enumerate(demo.data["roster"]):
        assert ws[f"D{2 + i}"].value == pytest.approx(ref["total"][s["name"]]), s["name"]


def test_by_teacher(demo, ref):
    ws = demo["By Teacher"]
    for i, (t, exp) in enumerate(ref["teachers"].items()):
        r = T0 + i
        assert ws[f"A{r}"].value == t
        assert str(ws[f"B{r}"].value) == exp["grade"]
        assert (ws[f"C{r}"].value, ws[f"D{r}"].value) == (exp["students"], exp["donated"])
        assert ws[f"F{r}"].value == pytest.approx(exp["total"])
    g0 = demo.find("By Teacher", "A", "Grade P3")
    for i, g in enumerate(C.GRADES):
        assert ws[f"C{g0 + i}"].value == ref["grades"][g]["students"]
        assert ws[f"F{g0 + i}"].value == pytest.approx(ref["grades"][g]["total"])


def test_by_day(demo, ref):
    ws = demo["By Day"]
    for i, d in enumerate(sorted(ref["daily"])):
        v = ws[f"A{6 + i}"].value
        assert (v.date() if isinstance(v, datetime) else v) == d
        assert ws[f"B{6 + i}"].value == pytest.approx(ref["daily"][d])
        assert ws[f"D{6 + i}"].value == pytest.approx(ref["running"][d])


def test_challenges(demo, ref):
    ws = demo["Daily Challenges"]
    for k, res in enumerate(ref["challenges"]):
        text = ws[f"F{5 + k}"].value
        rule = res["rule"]
        if rule in ("CLASS_PCT_MIN", "CLASS_PCT_TO_DATE"):
            assert text.startswith(res["winners"][0]) and f"{round(res['max'] * 100)}%" in text
            assert ("(tie)" in text) == (len(res["winners"]) > 1)
        elif rule == "GRADE_TOTAL":
            assert text.startswith(f"Grade {res['winners'][0]}")
        elif rule in ("SCHOOL_PCT_TARGET", "SCHOOL_TOTAL_TARGET"):
            assert text.startswith("HIT") == res["hit"]
        elif rule == "DRAWING":
            assert text.startswith(f"{res['eligible']} eligible")
        elif rule == "CLASS_TOTAL_PER_GRADE":
            g0 = demo.find("Daily Challenges", "A", "P3")
            for i, g in enumerate(C.GRADES):
                winners, amt = res["per_grade"][g]
                cell = ws[f"E{g0 + i}"].value
                assert ws[f"D{g0 + i}"].value == pytest.approx(amt)
                assert (winners[0] in cell) if winners else cell == "—"
        elif rule == "TOP_STUDENT_PER_CLASS":
            c0 = demo.find("Daily Challenges", "A", "Teacher") + 1
            col = next(c for c in range(4, 12) if "top student" in str(ws.cell(c0 - 1, c).value))
            for i, (t, (winners, amt)) in enumerate(res["per_class"].items()):
                assert ws.cell(c0 + i, 1).value == t
                assert ws.cell(c0 + i, col + 1).value == pytest.approx(amt)
                assert winners[0] in ws.cell(c0 + i, col).value


def test_ties_and_date_edges(books):
    b = books["ties"]
    ws = b["Daily Challenges"]
    assert [ws[f"F{r}"].value for r in range(5, 12)] == [
        "See GRADE SCORES below",
        "See CLASS SCORES below",
        "Grade 3 (tie) — $40",
        "Ms. C — 100%",
        "HIT! $255 of $255",
        "HIT! 40% of 40%",
        "Ms. A (tie) — 100%",
    ]
    g0 = b.find("Daily Challenges", "A", "P3")
    g3, g4 = g0 + C.GRADES.index("3"), g0 + C.GRADES.index("4")
    assert ws[f"E{g3}"].value == "Ms. A (tie)"     # Cal's gift on the window's last day counts
    assert ws[f"E{g4}"].value == "—"               # Eve's gift the day after does not
    c0 = b.find("Daily Challenges", "A", "Teacher") + 1
    assert ws.cell(c0, 5).value == "Ann One (tie)" and ws.cell(c0, 6).value == 30
