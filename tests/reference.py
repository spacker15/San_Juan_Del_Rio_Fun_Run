"""Independent Python re-implementation of the workbook's rules.

Tests feed the same inputs to this model and to the workbook (recalculated by LibreOffice)
and assert the numbers agree. Nothing here reads workbook formulas.
"""
from datetime import date, datetime

import config as C

EPOCH = date(1899, 12, 30)


def serial(d):
    if isinstance(d, datetime):
        d = d.date()
    return (d - EPOCH).days


def norm(s):
    return " ".join(str(s).split()).lower()


def parse_date(v):
    if isinstance(v, (datetime, date)):
        return serial(v)
    for fmt in ("%m/%d/%Y %I:%M %p", "%m/%d/%Y"):
        try:
            return serial(datetime.strptime(str(v).strip(), fmt))
        except ValueError:
            pass
    return None


def parse_amount(v):
    if isinstance(v, (int, float)):
        return round(float(v), 2)
    try:
        return round(float(str(v).replace("$", "").replace(",", "").replace(" ", "")), 2)
    except ValueError:
        return None


def compute(data):
    st = dict(C.SETTINGS, **data.get("settings", {}))
    challenges = data.get("challenges", C.CHALLENGES)
    start, end = serial(st["start"]), serial(st["end"])
    asof = serial(st["asof"]) if st.get("asof") else serial(date.today())
    goal = st["goal"]

    students = []
    for i, s in enumerate(data.get("roster", [])):
        first, last = str(s.get("first") or "").strip(), str(s.get("last") or "").strip()
        if not first + last:
            continue
        teacher, grade = str(s.get("teacher") or "").strip(), str(s["grade"]).strip()
        students.append({"row": i, "id": f"S{i + 1:04d}", "name": " ".join(f"{first} {last}".split()),
                         "first": first, "last": last, "grade": grade,
                         "cls": f"{teacher} – {grade}" if teacher else "(no teacher)",
                         "alias": norm(s.get("alias") or ""), "fl": norm(f"{first} {last}"),
                         "lf": norm(f"{last}, {first}")})
    by_id = {s["id"]: s for s in students}

    def match(key):
        for k in ("alias", "fl", "lf"):
            for s in students:
                if key and s[k] == key:
                    return s["id"]
        return None

    ledger, oc_status, cash_status = [], [], []
    h = {x: data["oc_headers"].index(x) for x in data["oc_headers"]}
    rows = data.get("oc_rows", [])
    first_txn = {}
    for idx, row in enumerate(rows):
        t = str(row[h["Transaction ID"]] or "").lower()
        if t and t not in first_txn:
            first_txn[t] = idx
    for idx, row in enumerate(rows):
        d, a = parse_date(row[h["Transaction Date"]]), parse_amount(row[h["Amount"]])
        sid = match(norm(row[h["Participant Name"]]))
        t = str(row[h["Transaction ID"]] or "").lower()
        st_ = ("Bad amount" if a is None else "Refund (not counted)" if a < 0 else "Zero amount" if a == 0
               else "Bad date" if d is None else "NAME NOT ON ROSTER" if sid is None
               else "Duplicate (ignored)" if t and first_txn[t] < idx else "OK")
        oc_status.append(st_)
        if st_ == "OK":
            ledger.append({"src": "OneCause", "date": d, "sid": sid, "amt": a})
    for g in data.get("cash", []):
        stud = str(g.get("student") or "").strip()
        sid = stud.split("#", 1)[1].strip() if "#" in stud else stud.upper()
        a = parse_amount(g.get("amount")) if g.get("amount") is not None else None
        d = parse_date(g.get("date"))
        st_ = ("Bad amount" if a is None or a <= 0 else "Bad date" if d is None
               else "NAME NOT ON ROSTER" if sid not in by_id else "OK")
        cash_status.append(st_)
        if st_ == "OK":
            ledger.append({"src": g.get("method") or "Cash/check", "date": d, "sid": sid, "amt": a})

    for s in students:
        mine = [x for x in ledger if x["sid"] == s["id"]]
        s["online"] = round(sum(x["amt"] for x in mine if x["src"] == "OneCause"), 2)
        s["total"] = round(sum(x["amt"] for x in mine), 2)
        s["gifts"] = len(mine)
        s["by_day"] = {d: round(sum(x["amt"] for x in mine if x["date"] == d), 2) for d in range(start, end + 1)}

    classes = []
    for s in students:
        if s["cls"] not in [c["name"] for c in classes]:
            classes.append({"name": s["cls"], "grade": s["grade"]})
    for c in classes:
        mem = [s for s in students if s["cls"] == c["name"]]
        c["students"], c["donated"] = len(mem), sum(1 for s in mem if s["total"] > 0)
        c["pct"] = c["donated"] / c["students"]
        c["total"] = round(sum(s["total"] for s in mem), 2)
    grades = {g: {"students": sum(1 for s in students if s["grade"] == g),
                  "donated": sum(1 for s in students if s["grade"] == g and s["total"] > 0),
                  "total": round(sum(s["total"] for s in students if s["grade"] == g), 2)} for g in C.GRADES}
    days = range(start, end + 1)
    daily = {d: round(sum(x["amt"] for x in ledger if x["date"] == d), 2) for d in days}
    pre = round(sum(x["amt"] for x in ledger if x["date"] < start), 2)
    post = round(sum(x["amt"] for x in ledger if x["date"] > end), 2)

    results = []
    for (_day, nm, cs_, ce, _reward, rule, val) in challenges:
        s0, e0 = serial(cs_), serial(ce)

        def amt(s, todate):
            return round(sum(x["amt"] for x in ledger if x["sid"] == s["id"] and
                             (x["date"] <= e0 if todate else s0 <= x["date"] <= e0)), 2)

        res = {"name": nm, "rule": rule, "status": "Upcoming" if asof < s0 else "LIVE" if asof <= e0 else "Final"}
        if rule == "CLASS_TOTAL_PER_GRADE":
            res["per_grade"] = {}
            for g in C.GRADES:
                vals = {c["name"]: round(sum(amt(s, False) for s in students if s["cls"] == c["name"]), 2)
                        for c in classes if c["grade"] == g}
                mx = max(vals.values()) if vals else -1
                res["per_grade"][g] = ([k for k, v in vals.items() if v == mx], mx)
        elif rule in ("CLASS_PARTICIPATION_MIN", "CLASS_PARTICIPATION_CUM"):
            thr = max(val or 0, 0.01) if rule.endswith("MIN") else 0.01
            vals = {c["name"]: round(sum(1 for s in students if s["cls"] == c["name"] and
                                         amt(s, rule.endswith("CUM")) >= thr) / c["students"], 6) for c in classes}
            mx = max(vals.values())
            res["max"], res["winners"] = mx, [k for k, v in vals.items() if v == mx]
        elif rule == "TOP_STUDENT_PER_CLASS":
            res["per_class"] = {}
            for c in classes:
                mem = [(s["name"], amt(s, False)) for s in students if s["cls"] == c["name"]]
                mx = max(v for _, v in mem)
                res["per_class"][c["name"]] = ([n for n, v in mem if v == mx], mx)
        elif rule == "GRADE_TOTAL":
            vals = {g: round(sum(amt(s, False) for s in students if s["grade"] == g), 2) for g in C.GRADES}
            mx = max(vals.values())
            res["max"], res["winners"] = mx, [g for g, v in vals.items() if v == mx]
        elif rule == "SCHOOL_PARTICIPATION_TARGET":
            res["value"] = sum(1 for s in students if amt(s, True) > 0) / len(students)
            res["hit"] = res["value"] >= val
        elif rule == "SCHOOL_TOTAL_TARGET":
            res["value"] = round(sum(amt(s, True) for s in students), 2)
            res["hit"] = res["value"] >= val
        elif rule == "DRAWING":
            res["eligible"] = [s["name"] for s in students if amt(s, False) > 0]
        results.append(res)

    raised = round(sum(s["total"] for s in students), 2)
    return {"students": students, "by_id": by_id, "classes": classes, "grades": grades, "daily": daily,
            "pre": pre, "post": post, "oc_status": oc_status, "cash_status": cash_status, "challenges": results,
            "kpi": {"Raised": raised, "Remaining": max(0, round(goal - raised, 2)),
                    "StudentsN": len(students), "DonatedN": sum(1 for s in students if s["total"] > 0),
                    "Gifts": len(ledger), "RaisedToday": daily.get(asof, 0),
                    "ThermoCells": min(20, max(0, int(raised / goal * 20 + 0.5)))}}
