"""Independent Python re-implementation of the workbook's business rules.

Tests feed the same input data to this model and to the workbook (recalculated by
LibreOffice) and assert that every number agrees. Nothing here imports workbook formulas.
"""
from collections import defaultdict
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
    if isinstance(v, (int, float)):
        return int(v)
    for fmt in ("%m/%d/%Y %I:%M %p", "%m/%d/%Y"):
        try:
            return serial(datetime.strptime(v.strip(), fmt))
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
    settings = dict(C.SETTINGS, **data.get("settings", {}))
    prizes = data.get("prizes", C.PRIZES)
    challenges = data.get("challenges", C.CHALLENGES)
    part_min = settings["part_min"]
    start, end = serial(settings["start"]), serial(settings["end"])
    asof = serial(settings["asof"]) if settings.get("asof") else serial(date.today())

    # ---- roster ----
    students = []
    for i, s in enumerate(data.get("roster", [])):
        first, last = str(s.get("first") or "").strip(), str(s.get("last") or "").strip()
        active = str(s.get("active") or "Y").strip().upper()[:1] != "N"
        if not (first + last) or not active:
            continue
        teacher, grade = str(s["teacher"]).strip(), str(s["grade"]).strip()
        students.append({
            "row": i, "id": f"S{i + 1:04d}", "first": first, "last": last, "grade": grade,
            "teacher": teacher, "classroom": f"{teacher} – {grade}" if teacher else "(No teacher)",
            "family": str(s.get("family") or "").strip(), "alias": norm(s.get("alias") or ""),
            "fl": norm(f"{first} {last}"), "lf": norm(f"{last}, {first}"),
        })
    by_id = {s["id"]: s for s in students}

    def match(name_key):
        for key in ("alias", "fl", "lf"):
            for s in students:
                if name_key and s[key] == name_key:
                    return s["id"]
        return None

    # ---- OneCause ----
    headers = data["oc_headers"]
    hi = {h: headers.index(h) for h in headers}
    oc_status, ledger = [], []
    seen_txn = {}
    rows = data.get("oc_rows", [])
    for idx, row in enumerate(rows):
        txn = row[hi["Transaction ID"]] or ""
        if txn and str(txn).lower() not in seen_txn:
            seen_txn[str(txn).lower()] = idx
    for idx, row in enumerate(rows):
        rd, ra, rn = row[hi["Transaction Date"]], row[hi["Amount"]], row[hi["Participant Name"]]
        txn = row[hi["Transaction ID"]] or ""
        d, a = parse_date(rd), parse_amount(ra)
        sid = match(norm(rn)) if norm(rn) else None
        if a is None:
            st = "BAD AMOUNT"
        elif a < 0:
            st = "NEGATIVE / REFUND"
        elif a == 0:
            st = "ZERO AMOUNT"
        elif d is None:
            st = "BAD DATE"
        elif sid is None:
            st = "UNMATCHED"
        elif txn and seen_txn[str(txn).lower()] < idx:
            st = "DUPLICATE TXN"
        else:
            st = "OK"
        oc_status.append((st, a))
        if st == "OK":
            ledger.append({"src": "Online", "date": d, "sid": sid, "amt": a,
                           "donor": " ".join(str(row[hi["Donor Name"]] or "").split())})

    # ---- Cash / check ----
    cash_status = []
    for g in data.get("cash", []):
        stud = str(g.get("student") or "").strip()
        sid = stud.split("#", 1)[1].strip()[:20] if "#" in stud else stud.upper()
        amt, dt = g.get("amount"), g.get("date")
        if not isinstance(amt, (int, float)) or amt <= 0:
            st = "BAD AMOUNT"
        elif not isinstance(dt, (date, datetime)):
            st = "BAD DATE"
        elif sid not in by_id:
            st = "NO STUDENT"
        else:
            st = "OK"
        cash_status.append(st)
        if st == "OK":
            ledger.append({"src": g.get("method") or "Offline", "date": serial(dt), "sid": sid,
                           "amt": round(amt, 2), "donor": " ".join(str(g.get("donor") or "").split())})

    # ---- students ----
    thresholds = [p[1] for p in prizes if p[1] and p[1] > 0]
    for s in students:
        mine = [x for x in ledger if x["sid"] == s["id"]]
        s["online"] = round(sum(x["amt"] for x in mine if x["src"] == "Online"), 2)
        s["offline"] = round(sum(x["amt"] for x in mine if x["src"] != "Online"), 2)
        s["total"] = round(s["online"] + s["offline"], 2)
        s["gifts"] = len(mine)
        s["part"] = 1 if s["total"] > 0 and s["total"] >= part_min else 0
        s["prize"] = sum(1 for t in thresholds if t <= s["total"] + 1e-9)
        s["first_gift"] = min((x["date"] for x in mine), default=None)
        s["full"] = f"{s['first']} {s['last']}"
    for s in students:
        s["rank"] = 1 + sum(1 for o in students if o["total"] > s["total"])

    # ---- classes / grades / families / donors ----
    classes = []
    seen = set()
    for teacher, grade, goal in data.get("classrooms", []):
        name = f"{teacher.strip()} – {str(grade).strip()}"
        members = [s for s in students if s["classroom"] == name]
        n = len(members)
        p = sum(s["part"] for s in members)
        classes.append({"name": name, "grade": str(grade), "students": n, "participants": p,
                        "pct": round(p / n, 6) if n else 0, "total": round(sum(s["total"] for s in members), 2)})
        seen.add(name)
    grades = {}
    for g in C.GRADES:
        mem = [s for s in students if s["grade"] == g]
        grades[g] = {"students": len(mem), "participants": sum(s["part"] for s in mem),
                     "total": round(sum(s["total"] for s in mem), 2),
                     "classes": sum(1 for c in classes if c["grade"] == g)}
    families = defaultdict(float)
    for s in students:
        if s["family"]:
            families[s["family"].lower()] += s["total"]
    donors = defaultdict(lambda: [0.0, 0])
    for x in ledger:
        if x["donor"]:
            donors[x["donor"].lower()][0] += x["amt"]
            donors[x["donor"].lower()][1] += 1

    # ---- daily ----
    days = list(range(start, end + 1))
    daily = {d: round(sum(x["amt"] for x in ledger if x["date"] == d), 2) for d in days}
    pre = round(sum(x["amt"] for x in ledger if x["date"] < start), 2)
    post = round(sum(x["amt"] for x in ledger if x["date"] > end), 2)
    cum, run = {}, pre
    for d in days:
        run += daily[d]
        cum[d] = round(run, 2)

    # ---- challenges ----
    ch_results = []
    for (cs_, ce, _day, name, _desc, _reward, metric, mn, target) in challenges:
        s0, e0 = serial(cs_), serial(ce)
        mn = mn or 0
        win = {s["id"]: round(sum(x["amt"] for x in ledger if x["sid"] == s["id"] and s0 <= x["date"] <= e0), 2)
               for s in students}
        asof_amt = {s["id"]: round(sum(x["amt"] for x in ledger if x["sid"] == s["id"] and x["date"] <= e0), 2)
                    for s in students}
        res = {"name": name, "metric": metric, "status": "Upcoming" if asof < s0 else ("LIVE" if asof <= e0 else
                                                                                         "Final")}
        if metric == "CLASS_TOTAL_PER_GRADE":
            per_grade = {}
            for g in C.GRADES:
                cl = [c for c in classes if c["grade"] == g]
                vals = {c["name"]: round(sum(win[s["id"]] for s in students if s["classroom"] == c["name"]), 2)
                        for c in cl}
                if not vals:
                    per_grade[g] = (None, -1)
                    continue
                mx = max(vals.values())
                per_grade[g] = ([k for k, v in vals.items() if v == mx], mx)
            res["per_grade"] = per_grade
        elif metric in ("CLASS_PARTICIPATION_MIN", "CLASS_PARTICIPATION_CUM"):
            thr = max(mn, 0.01) if metric.endswith("MIN") else max(mn, part_min, 0.01)
            src = win if metric.endswith("MIN") else asof_amt
            vals = {}
            for c in classes:
                mem = [s for s in students if s["classroom"] == c["name"]]
                vals[c["name"]] = round(sum(1 for s in mem if src[s["id"]] >= thr) / len(mem), 6) if mem else 0
            mx = max(vals.values())
            res["class_max"], res["class_winners"] = mx, [k for k, v in vals.items() if v == mx]
        elif metric == "TOP_STUDENT_PER_CLASS":
            per_class = {}
            for c in classes:
                mem = [s for s in students if s["classroom"] == c["name"]]
                if not mem:
                    per_class[c["name"]] = (None, -1)
                    continue
                mx = max(win[s["id"]] for s in mem)
                per_class[c["name"]] = ([s["full"] for s in mem if win[s["id"]] == mx], mx)
            res["per_class"] = per_class
        elif metric == "GRADE_TOTAL":
            vals = {g: round(sum(win[s["id"]] for s in students if s["grade"] == g), 2) for g in C.GRADES}
            mx = max(vals.values())
            res["grade_max"], res["grade_winners"] = mx, [g for g, v in vals.items() if v == mx]
        elif metric == "SCHOOL_PARTICIPATION_TARGET":
            thr = max(mn, part_min, 0.01)
            res["value"] = round(sum(1 for s in students if asof_amt[s["id"]] >= thr) / len(students), 6)
            res["hit"] = res["value"] >= target
        elif metric == "SCHOOL_TOTAL_TARGET":
            res["value"] = round(sum(asof_amt.values()), 2)
            res["hit"] = res["value"] >= target
        elif metric == "DRAWING":
            res["eligible"] = sorted(s["id"] for s in students if win[s["id"]] > 0)
            res["target"] = target
        ch_results.append(res)

    raised = round(sum(s["total"] for s in students), 2)
    goal = settings["goal"]
    return {
        "students": students, "by_id": by_id, "classes": classes, "grades": grades,
        "families": dict(families), "donors": dict(donors), "daily": daily, "cum": cum, "pre": pre, "post": post,
        "oc_status": oc_status, "cash_status": cash_status, "challenges": ch_results, "ledger": ledger,
        "kpi": {
            "Raised": raised,
            "RaisedOnline": round(sum(s["online"] for s in students), 2),
            "RaisedOffline": round(sum(s["offline"] for s in students), 2),
            "Students": len(students),
            "Participants": sum(s["part"] for s in students),
            "Gifts": sum(s["gifts"] for s in students),
            "Families": len(families),
            "Donors": len(donors),
            "ThermoCells": min(40, max(0, int(raised / goal * 40 + 0.5))),
            "RaisedToday": round(sum(x["amt"] for x in ledger if x["date"] == asof), 2),
            "Remaining": max(0, round(goal - raised, 2)),
        },
    }
