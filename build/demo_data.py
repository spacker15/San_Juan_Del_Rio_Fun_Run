"""Deterministic fictional data for the DEMO workbook and the tests.

Deliberate edge cases: OneCause name variants ("Last, First", odd spacing, a OneCause-only
nickname), a name not on the roster, a duplicate transaction, a refund, $0, an unreadable date and
amount, "$1,250.00" text amounts, text dates, gifts before/after the campaign, and cash rows with a
typed ID, a bad amount, an unknown student and a bad date.
"""
import random
from datetime import datetime, timedelta

import config as C

FIRST = ["Sofia", "Mateo", "Isabella", "Santiago", "Camila", "Lucas", "Valentina", "Diego", "Emma", "Liam",
         "Olivia", "Noah", "Ava", "Elijah", "Mia", "James", "Amelia", "Benjamin", "Harper", "Henry", "Evelyn",
         "Leo", "Luna", "Gabriel", "Chloe", "Daniel", "Grace", "Samuel", "Lily", "Joseph", "Zoe", "David",
         "Nora", "Isaac", "Aria", "Adrian", "Ella", "Julian", "Maya", "Nicolas", "Elena", "Thomas", "Clara"]
LAST = ["Garcia", "Martinez", "Lopez", "Hernandez", "Smith", "Johnson", "Williams", "Brown", "Rodriguez",
        "Gonzalez", "Perez", "Sanchez", "Ramirez", "Torres", "Flores", "Rivera", "Gomez", "Diaz", "Reyes",
        "Cruz", "Morales", "Ortiz", "Gutierrez", "Chavez", "Ramos", "Ruiz", "Alvarez", "Mendoza", "Castillo",
        "Vasquez", "Kim", "Nguyen", "Patel", "O'Brien", "Murphy", "Sullivan", "Kelly", "Walsh", "Byrne", "Ryan"]
TEACHERS = ["Ms. Alvarez", "Mrs. Bennett", "Mr. Carter", "Mrs. Dominguez", "Ms. Ellison", "Mrs. Fischer",
            "Mr. Galloway", "Mrs. Hayes", "Ms. Ingram", "Mrs. Jensen", "Mr. Kowalski", "Mrs. Lindqvist",
            "Ms. Moreau", "Mrs. Novak", "Mr. Okafor", "Mrs. Pruitt", "Ms. Quinlan", "Mrs. Ravel", "Mr. Stanton",
            "Mrs. Thorne", "Ms. Underwood", "Mrs. Valdez"]
HEADERS = ["Transaction Date", "Transaction ID", "Participant Name", "Donor Name", "Amount", "Payment Method"]


def generate(seed=2026):
    rng = random.Random(seed)
    start, end = C.SETTINGS["start"], C.SETTINGS["end"]
    classrooms = [(TEACHERS[gi * 2 + j], g) for gi, g in enumerate(C.GRADES) for j in range(2)]

    roster, used = [], set()
    for teacher, grade in classrooms:
        for _ in range(16 if grade in ("P3", "P4") else 20):
            first, last = rng.choice(FIRST), rng.choice(LAST)
            while (first, last) in used:
                first, last = rng.choice(FIRST), rng.choice(LAST)
            used.add((first, last))
            roster.append({"first": first, "last": last, "grade": grade, "teacher": teacher})
    roster[3]["alias"] = "Bobby Tables"  # OneCause uses a nickname for this student

    def sid(i):
        return f"S{i + 1:04d}"

    def ts(day):
        return datetime(day.year, day.month, day.day, rng.randint(7, 21), rng.choice([0, 15, 30, 45]))

    days = [start + timedelta(d) for d in range((end - start).days + 1)]
    weights = [6, 3, 3, 5, 4, 5, 6, 5, 3, 3, 6, 6, 2, 8, 9][:len(days)]
    oc, txn = [], 100000
    for _ in range(1080):
        i = rng.randrange(len(roster))
        s = roster[i]
        amt = rng.choice([10, 20, 20, 25, 25, 25, 25, 40, 50, 50, 50, 75, 100, 100, 150])
        p = rng.random()
        name = ("Bobby Tables" if i == 3 else f"{s['last']}, {s['first']}" if p < 0.1
                else f"  {s['first'].upper()}   {s['last'].lower()} " if p < 0.15 else f"{s['first']} {s['last']}")
        txn += rng.randint(1, 9)
        when = ts(rng.choices(days, weights)[0])
        if rng.random() < 0.04:
            when = when.strftime("%m/%d/%Y %I:%M %p")
        oc.append([when, f"T{txn}", name, f"{rng.choice(FIRST)} {s['last']}",
                   f"${amt:,.2f}" if rng.random() < 0.05 else amt, rng.choice(["Credit Card", "Apple Pay"])])
    s5, s20 = roster[5], roster[20]
    n5 = f"{s5['first']} {s5['last']}"
    oc += [
        [ts(days[5]), "T900001", f"{s20['first']} {s20['last']}", "Grandma Rivera", "$1,250.00", "Credit Card"],
        [ts(days[2]), "T900002", "Unknown Kiddo", "Mystery Donor", 40, "Credit Card"],
        oc[0][:],                                                          # duplicate transaction
        [ts(days[6]), "T900005", n5, "Refund Donor", -25, "Credit Card"],
        [ts(days[6]), "T900006", n5, "Zero Donor", 0, "Credit Card"],
        ["N/A", "T900007", n5, "Bad Date Donor", 35, "Credit Card"],
        [ts(days[7]), "T900008", n5, "Bad Amount Donor", "abc", "Credit Card"],
        [datetime(2026, 10, 28, 9, 0), "T900009", n5, "Early Bird", 45, "Credit Card"],
        [datetime(2026, 11, 14, 9, 0), "T900010", n5, "Late Owl", 55, "Credit Card"],
        [ts(days[8]), "", n5, "No Txn Id", 15, "Credit Card"],
    ]
    rng.shuffle(oc)

    cash = []
    for _ in range(140):
        i = rng.randrange(len(roster))
        s = roster[i]
        label = f"{s['last']}, {s['first']} ({s['teacher']} – {s['grade']}) #{sid(i)}"
        cash.append({"date": rng.choice(days[:-1]), "student": label if rng.random() < 0.85 else sid(i).lower(),
                     "donor": f"{s['last']} family", "amount": rng.choice([5, 10, 20, 20, 25, 40, 50, 100]),
                     "method": rng.choice(["Cash", "Check", "Check"])})
    cash += [
        {"date": days[3], "student": f"#{sid(7)}", "donor": "Bad Amount", "amount": "twenty", "method": "Cash"},
        {"date": days[3], "student": "S0999", "donor": "Ghost", "amount": 20, "method": "Cash"},
        {"date": "next tues", "student": sid(7), "donor": "No Date", "amount": 20, "method": "Cash"},
    ]
    return {"settings": {"asof": end}, "classrooms": classrooms, "roster": roster, "oc_headers": HEADERS,
            "oc_rows": oc, "cash": cash}
