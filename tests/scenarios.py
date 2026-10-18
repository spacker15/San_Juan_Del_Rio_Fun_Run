"""Small hand-built dataset with known answers: ties and date-window edges."""
from datetime import date, datetime

D = date


def ties():
    roster = [
        {"name": "Ann One", "teacher": "Ms. A", "grade": "3"},
        {"name": "Ben Two", "teacher": "Ms. A", "grade": "3"},
        {"name": "Cal Three", "teacher": "Ms. B", "grade": "3"},
        {"name": "Dee Four", "teacher": "Ms. B", "grade": "3"},
        {"name": "Eve Five", "teacher": "Ms. C", "grade": "4"},
    ]

    def g(d, name, amt):
        return {"date": datetime(d.year, d.month, d.day, 12), "student": name, "amount": amt, "method": "OneCause"}

    donations = [
        g(D(2026, 10, 31), "Ann One", 50),
        g(D(2026, 11, 2), "Ann One", 30),
        g(D(2026, 11, 1), "Cal Three", 50),   # last day of the weekend window: counts
        g(D(2026, 11, 3), "Cal Three", 40),
        g(D(2026, 11, 2), "Eve Five", 5),     # day after the weekend window: doesn't count
        g(D(2026, 11, 2), "Ben Two", 30),
        g(D(2026, 11, 2), "Dee Four", 10),
        g(D(2026, 11, 3), "Eve Five", 40),
    ]
    challenges = [
        ("Fri", "Weekend", D(2026, 10, 30), D(2026, 11, 1), "", "CLASS_TOTAL_PER_GRADE", None),
        ("Mon", "Top student", D(2026, 11, 2), D(2026, 11, 2), "", "TOP_STUDENT_PER_CLASS", None),
        ("Tue", "Grade day", D(2026, 11, 3), D(2026, 11, 3), "", "GRADE_TOTAL", None),
        ("Tue", "$20 day", D(2026, 11, 3), D(2026, 11, 3), "", "CLASS_PCT_MIN", 20),
        ("All", "Total goal", D(2026, 10, 30), D(2026, 11, 13), "", "SCHOOL_TOTAL_TARGET", 255),
        ("Sun", "40% by Sunday", D(2026, 11, 1), D(2026, 11, 1), "", "SCHOOL_PCT_TARGET", 0.4),
        ("Mon", "Showdown", D(2026, 11, 2), D(2026, 11, 2), "", "CLASS_PCT_TO_DATE", None),
    ]
    return {"teachers": ["Ms. A", "Ms. B", "Ms. C"], "roster": roster, "donations": donations,
            "challenges": challenges}
