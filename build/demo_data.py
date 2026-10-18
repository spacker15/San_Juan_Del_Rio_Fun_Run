"""Fictional demo data: 22 teachers, ~420 students, one donations list (~$61k).

Edge cases on purpose: a name not on the Students list, a refund (negative amount), odd
spacing/case in a pasted name, and timestamps on OneCause dates.
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


def generate(seed=2026, gifts=1250):
    rng = random.Random(seed)
    teachers = [(TEACHERS[gi * 2 + j], g) for gi, g in enumerate(C.GRADES) for j in range(2)]
    roster, used = [], set()
    for teacher, grade in teachers:
        for _ in range(16 if grade in ("P3", "P4") else 20):
            name = f"{rng.choice(FIRST)} {rng.choice(LAST)}"
            while name in used:
                name = f"{rng.choice(FIRST)} {rng.choice(LAST)}"
            used.add(name)
            roster.append({"name": name, "teacher": teacher, "grade": grade})

    days = [C.FIRST_DAY + timedelta(d) for d in range(15)]
    weights = [6, 3, 3, 5, 4, 5, 6, 5, 3, 3, 6, 6, 2, 8, 9]
    donations = []
    for _ in range(gifts):
        s = rng.choice(roster)
        day = rng.choices(days, weights)[0]
        online = rng.random() < 0.88
        when = datetime(day.year, day.month, day.day, rng.randint(7, 21), rng.choice([0, 15, 30, 45]))
        donations.append({"date": when if online else day, "student": s["name"],
                          "donor": f"{rng.choice(FIRST)} {s['name'].split()[-1]}",
                          "amount": rng.choice([10, 20, 20, 25, 25, 25, 25, 40, 50, 50, 50, 75, 100, 100, 150]),
                          "method": "OneCause" if online else rng.choice(["Cash", "Check"])})
    first = roster[5]["name"]
    donations += [
        {"date": datetime(2026, 11, 2, 10), "student": "Unknown Kiddo", "donor": "Mystery Donor", "amount": 40,
         "method": "OneCause"},
        {"date": datetime(2026, 11, 6, 9), "student": first, "donor": "Refund", "amount": -25, "method": "OneCause"},
        {"date": datetime(2026, 11, 4, 12), "student": "  " + first.upper() + " ", "donor": "Caps Donor",
         "amount": 30, "method": "OneCause"},
    ]
    rng.shuffle(donations)
    return {"teachers": [t for t, _ in teachers], "roster": roster, "donations": donations}
