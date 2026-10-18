"""Independent Python version of the workbook's rules; tests compare it with the recalculated workbook."""
from datetime import date, datetime

import config as C


def day(v):
    return v.date() if isinstance(v, datetime) else v


def norm(s):
    return " ".join(str(s).split()).lower()


def compute(data):
    challenges = data.get("challenges", C.CHALLENGES)
    roster = data.get("roster", [])
    by_name = {norm(s["name"]): s for s in roster}
    gifts = []
    for d in data.get("donations", []):
        s = by_name.get(norm(d["student"]))
        gifts.append({"day": day(d["date"]), "amt": d["amount"], "student": s["name"] if s else None,
                      "teacher": s["teacher"] if s else None, "grade": s["grade"] if s else None})
    total = {s["name"]: sum(g["amt"] for g in gifts if g["student"] == s["name"]) for s in roster}

    teachers = {}
    for t in data.get("teachers", []):
        mem = [s for s in roster if s["teacher"] == t]
        teachers[t] = {"students": len(mem), "donated": sum(1 for s in mem if total[s["name"]] > 0),
                       "total": sum(g["amt"] for g in gifts if g["teacher"] == t),
                       "grade": mem[0]["grade"] if mem else ""}
    grades = {g: {"students": sum(1 for s in roster if s["grade"] == g),
                  "total": sum(x["amt"] for x in gifts if x["grade"] == g)} for g in C.GRADES}
    raised = sum(g["amt"] for g in gifts)

    def student_amt(s, start, end, to_date):
        return sum(g["amt"] for g in gifts if g["student"] == s["name"] and
                   (g["day"] <= end if to_date else start <= g["day"] <= end))

    results = []
    for (_d, nm, start, end, _r, rule, val) in challenges:
        res = {"name": nm, "rule": rule}
        if rule == "CLASS_TOTAL_PER_GRADE":
            per = {}
            for g in C.GRADES:
                vals = {t: sum(x["amt"] for x in gifts if x["teacher"] == t and start <= x["day"] <= end)
                        for t, info in teachers.items() if info["grade"] == g}
                mx = max(vals.values(), default=0)
                per[g] = ([t for t, v in vals.items() if v == mx] if mx > 0 else [], mx)
            res["per_grade"] = per
        elif rule in ("CLASS_PCT_MIN", "CLASS_PCT_TO_DATE"):
            floor = max(0.01, val or 0) if rule == "CLASS_PCT_MIN" else 0.01
            vals = {t: sum(1 for s in roster if s["teacher"] == t and
                           student_amt(s, start, end, rule == "CLASS_PCT_TO_DATE") >= floor) / info["students"]
                    for t, info in teachers.items() if info["students"]}
            mx = max(vals.values())
            res["max"], res["winners"] = mx, [t for t, v in vals.items() if v == mx]
        elif rule == "TOP_STUDENT_PER_CLASS":
            per = {}
            for t in teachers:
                amts = {s["name"]: student_amt(s, start, end, False) for s in roster if s["teacher"] == t}
                mx = max(amts.values(), default=0)
                per[t] = ([n for n, v in amts.items() if v == mx] if mx > 0 else [], mx)
            res["per_class"] = per
        elif rule == "GRADE_TOTAL":
            vals = {g: sum(x["amt"] for x in gifts if x["grade"] == g and start <= x["day"] <= end) for g in C.GRADES}
            mx = max(vals.values())
            res["max"], res["winners"] = mx, [g for g, v in vals.items() if v == mx]
        elif rule == "DRAWING":
            res["eligible"] = sum(1 for s in roster if student_amt(s, start, end, False) > 0)
        elif rule == "SCHOOL_PCT_TARGET":
            res["value"] = sum(1 for s in roster if student_amt(s, start, end, True) > 0) / len(roster)
            res["hit"] = res["value"] >= val
        elif rule == "SCHOOL_TOTAL_TARGET":
            res["value"] = sum(g["amt"] for g in gifts if g["day"] <= end)
            res["hit"] = res["value"] >= val
        results.append(res)

    first = C.FIRST_DAY
    days = [date.fromordinal(first.toordinal() + i) for i in range(C.N_DAY)]
    return {
        "raised": raised, "total": total, "teachers": teachers, "grades": grades, "challenges": results,
        "not_on_list": sum(1 for g in gifts if g["student"] is None),
        "donated_pct": sum(1 for v in total.values() if v > 0) / len(roster) if roster else 0,
        "daily": {d: sum(g["amt"] for g in gifts if g["day"] == d) for d in days},
        "running": {d: sum(g["amt"] for g in gifts if g["day"] <= d) for d in days},
    }
