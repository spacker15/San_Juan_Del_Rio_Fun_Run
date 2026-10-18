"""Deterministic synthetic data for the DEMO workbook and the test suite.

Every name is fictional. Deliberate edge cases (all asserted by tests/):
  * OneCause name variants: "Last, First", odd case/spacing, OneCause alias on the roster
  * unmatched names, duplicate transaction IDs, refunds (negative), $0, unparseable date/amount
  * text amounts ("$1,250.00") and text dates ("11/03/2026 10:15 AM")
  * gifts before the campaign start and after the end date
  * a student whose total lands exactly on a prize threshold
  * an inactive student (Active = N) and a roster row with a teacher not on Classrooms
  * cash/check rows: picked label, typed ID, bad amount, unknown student
"""
import random
from datetime import date, datetime, timedelta

import config as C

FIRST = ["Sofia", "Mateo", "Isabella", "Santiago", "Camila", "Lucas", "Valentina", "Diego", "Emma", "Liam",
         "Olivia", "Noah", "Ava", "Elijah", "Mia", "James", "Amelia", "Benjamin", "Harper", "Henry", "Evelyn",
         "Leo", "Luna", "Gabriel", "Chloe", "Daniel", "Grace", "Samuel", "Lily", "Joseph", "Zoe", "David",
         "Nora", "Isaac", "Aria", "Adrian", "Ella", "Julian", "Maya", "Nicolas", "Elena", "Thomas", "Clara",
         "Anthony", "Lucia", "Andres", "Natalia", "Xavier", "Paige", "Owen"]
LAST = ["Garcia", "Martinez", "Lopez", "Hernandez", "Smith", "Johnson", "Williams", "Brown", "Rodriguez",
        "Gonzalez", "Perez", "Sanchez", "Ramirez", "Torres", "Flores", "Rivera", "Gomez", "Diaz", "Reyes",
        "Cruz", "Morales", "Ortiz", "Gutierrez", "Chavez", "Ramos", "Ruiz", "Alvarez", "Mendoza", "Castillo",
        "Vasquez", "Kim", "Nguyen", "Patel", "O'Brien", "Murphy", "Sullivan", "Kelly", "Walsh", "Byrne", "Ryan",
        "Russo", "Romano", "Bianchi", "Costa", "Silva", "Santos", "Moreno", "Jimenez", "Navarro", "Delgado"]
TEACHERS = ["Ms. Alvarez", "Mrs. Bennett", "Mr. Carter", "Mrs. Dominguez", "Ms. Ellison", "Mrs. Fischer",
            "Mr. Galloway", "Mrs. Hayes", "Ms. Ingram", "Mrs. Jensen", "Mr. Kowalski", "Mrs. Lindqvist",
            "Ms. Moreau", "Mrs. Novak", "Mr. Okafor", "Mrs. Pruitt", "Ms. Quinlan", "Mrs. Ravel", "Mr. Stanton",
            "Mrs. Thorne", "Ms. Underwood", "Mrs. Valdez"]

START, END = C.SETTINGS["start"], C.SETTINGS["end"]


def _ts(day, rng):
    return datetime(day.year, day.month, day.day, rng.randint(7, 21), rng.choice([0, 5, 15, 30, 45]))


def generate(seed=2026):
    rng = random.Random(seed)
    classrooms = []
    for gi, g in enumerate(C.GRADES):
        for j in range(2):
            classrooms.append((TEACHERS[gi * 2 + j], g, 2500 if g in ("7", "8") else None))

    roster, families = [], []
    used = set()
    for teacher, grade, _ in classrooms:
        n = 16 if grade in ("P3", "P4") else 20
        for _ in range(n):
            if families and rng.random() < 0.3:          # sibling of an existing family
                fam_last, fam = rng.choice(families)
                last = fam_last
            else:
                last = rng.choice(LAST)
                fam = f"{last} ({rng.choice(['A', 'B', 'C', 'D', 'E', 'F'])})"
                families.append((last, fam))
            first = rng.choice(FIRST)
            while (first, last) in used:
                first = rng.choice(FIRST)
            used.add((first, last))
            roster.append({"first": first, "last": last, "grade": grade, "teacher": teacher, "family": fam,
                           "email": f"{first.lower()}.{last.lower().replace(chr(39), '')}@example.org"})
    # Edge rows
    roster[3]["alias"] = "Bobby Tables"                    # OneCause uses a nickname
    roster[10]["active"] = "N"                             # inactive student: gifts become UNMATCHED
    roster.append({"first": "Pat", "last": "Newkid", "grade": "3", "teacher": "Mr. Nobody", "family": "Newkid (Z)"})

    def sid(i):
        return f"S{i + 1:04d}"

    active = [i for i, s in enumerate(roster) if s.get("active") != "N" and s["teacher"] != "Mr. Nobody"]
    days = [START + timedelta(d) for d in range((END - START).days + 1)]
    weights = [6, 3, 3, 5, 4, 5, 6, 5, 3, 3, 6, 6, 2, 8, 9][:len(days)]

    oc_rows = []
    txn = 100000
    for _ in range(1040):
        i = rng.choice(active)
        s = roster[i]
        day = rng.choices(days, weights)[0]
        amt = rng.choice([10, 20, 20, 25, 25, 25, 25, 40, 50, 50, 50, 75, 100, 100, 150])
        name_fmt = rng.random()
        if i == 3:
            name = "Bobby Tables"
        elif name_fmt < 0.1:
            name = f"{s['last']}, {s['first']}"
        elif name_fmt < 0.15:
            name = f"  {s['first'].upper()}   {s['last'].lower()} "
        else:
            name = f"{s['first']} {s['last']}"
        txn += rng.randint(1, 9)
        amount = f"${amt:,.2f}" if rng.random() < 0.05 else amt
        when = _ts(day, rng)
        if rng.random() < 0.04:
            when = when.strftime("%m/%d/%Y %I:%M %p")
        oc_rows.append([when, f"T{txn}", name, f"{rng.choice(FIRST)} {s['last']}", amount,
                        rng.choice(["Credit Card", "Credit Card", "Apple Pay", "PayPal"])])
    # Big text amount with thousands separator
    oc_rows.append([_ts(days[5], rng), "T900001", f"{roster[20]['first']} {roster[20]['last']}", "Grandma Rivera",
                    "$1,250.00", "Credit Card"])
    # Edge cases
    s5 = roster[5]
    oc_rows += [
        [_ts(days[2], rng), "T900002", "Unknown Kiddo", "Mystery Donor", 40, "Credit Card"],      # unmatched
        [_ts(days[3], rng), "T900003", "Someone Else", "A Donor", 60, "Credit Card"],              # unmatched
        [_ts(days[4], rng), "T900004", f"{roster[10]['first']} {roster[10]['last']}", "Inactive Fan", 30,
         "PayPal"],                                                                                # inactive
        oc_rows[0][:],                                                                             # duplicate
        [_ts(days[6], rng), "T900005", f"{s5['first']} {s5['last']}", "Refund Donor", -25, "Credit Card"],
        [_ts(days[6], rng), "T900006", f"{s5['first']} {s5['last']}", "Zero Donor", 0, "Credit Card"],
        ["N/A", "T900007", f"{s5['first']} {s5['last']}", "Bad Date Donor", 35, "Credit Card"],
        [_ts(days[7], rng), "T900008", f"{s5['first']} {s5['last']}", "Bad Amount Donor", "abc", "Credit Card"],
        [datetime(2026, 10, 28, 9, 0), "T900009", f"{s5['first']} {s5['last']}", "Early Bird", 45, "Credit Card"],
        [datetime(2026, 11, 14, 9, 0), "T900010", f"{s5['first']} {s5['last']}", "Late Owl", 55, "Credit Card"],
        [_ts(days[8], rng), "", f"{s5['first']} {s5['last']}", "No Txn Id", 15, "Credit Card"],
    ]
    rng.shuffle(oc_rows)

    cash = []
    for _ in range(140):
        i = rng.choice(active)
        s = roster[i]
        day = rng.choice(days[:-1])
        label = f"{s['last']}, {s['first']} — {s['teacher']} – {s['grade']} #{sid(i)}"
        student = label if rng.random() < 0.85 else sid(i).lower()
        method = rng.choice(["Cash", "Check", "Check", "Other"])
        cash.append({"date": day, "student": student, "donor": f"{s['last']} family",
                     "amount": rng.choice([5, 10, 20, 20, 25, 40, 50, 100]), "method": method,
                     "check": rng.randint(1000, 9999) if method == "Check" else None, "by": "Volunteer"})
    cash += [
        {"date": days[3], "student": f"#{sid(7)}", "donor": "Bad Amount", "amount": "twenty", "method": "Cash"},
        {"date": days[3], "student": "S0999", "donor": "Ghost", "amount": 20, "method": "Cash"},
        {"date": "next tues", "student": sid(7), "donor": "No Date", "amount": 20, "method": "Cash"},
    ]
    # Student with no other gifts lands exactly on the $50 prize threshold
    exact = max(active)
    cash.append({"date": days[1], "student": sid(exact), "donor": "Exact Fifty", "amount": 50, "method": "Cash"})
    ex = roster[exact]
    variants = {f"{ex['first']} {ex['last']}".lower(), f"{ex['last']}, {ex['first']}".lower()}
    for row in oc_rows:  # move any random online gifts off that student so the total is exactly $50
        if isinstance(row[2], str) and " ".join(row[2].split()).lower() in variants:
            row[2] = f"{roster[0]['first']} {roster[0]['last']}"
    for g in cash[:-1]:
        if sid(exact) in str(g["student"]).upper():
            g["student"] = sid(0)

    history = {
        "summary": [[2025, 55000, 51240.5, 0.71, 412] + [2100, 2600, 3900, 4400, 4850, 5100, 5300, 5600, 5900,
                                                          5800, 5690.5]],
        "daily": {2025: [round(51240.5 * min(1, (d / 15) ** 0.85), 2) if d <= 15 else None for d in range(1, 32)]},
    }
    return {
        "settings": {"asof": date(2026, 11, 13)},
        "classrooms": classrooms,
        "roster": roster,
        "oc_headers": ["Transaction Date", "Transaction ID", "Participant Name", "Donor Name", "Amount",
                       "Payment Method"],
        "oc_rows": oc_rows,
        "cash": cash,
        "history": history,
    }
