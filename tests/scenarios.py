"""Small hand-built datasets with known answers (ties, date-window edges, thermometer fill)."""
from datetime import date, datetime

HEAD = ["Transaction Date", "Transaction ID", "Participant Name", "Donor Name", "Amount", "Payment Method"]
D = date


def _oc(day, txn, name, amt):
    return [datetime(day.year, day.month, day.day, 12, 0), txn, name, "Donor " + txn, amt, "Credit Card"]


def ties():
    roster = [
        {"first": "Ann", "last": "One", "grade": "3", "teacher": "Ms. A"},
        {"first": "Ben", "last": "Two", "grade": "3", "teacher": "Ms. A"},
        {"first": "Cal", "last": "Three", "grade": "3", "teacher": "Ms. B"},
        {"first": "Dee", "last": "Four", "grade": "3", "teacher": "Ms. B"},
        {"first": "Eve", "last": "Five", "grade": "4", "teacher": "Ms. C"},
    ]
    oc = [
        _oc(D(2026, 10, 31), "T1", "Ann One", 50),
        _oc(D(2026, 11, 2), "T2", "Ann One", 30),
        _oc(D(2026, 11, 1), "T3", "Cal Three", 50),   # last day of the weekend window: counts
        _oc(D(2026, 11, 3), "T4", "Cal Three", 40),
        _oc(D(2026, 11, 2), "T5", "Eve Five", 5),     # day after the weekend window: doesn't count
    ]
    cash = [
        {"date": D(2026, 11, 2), "student": "S0002", "amount": 30, "method": "Cash"},
        {"date": D(2026, 11, 2), "student": "S0004", "amount": 10, "method": "Check"},
        {"date": D(2026, 11, 3), "student": "S0005", "amount": 40, "method": "Cash"},
    ]
    challenges = [
        ("Fri", "Weekend", D(2026, 10, 30), D(2026, 11, 1), "", "CLASS_TOTAL_PER_GRADE", None),
        ("Mon", "Top student", D(2026, 11, 2), D(2026, 11, 2), "", "TOP_STUDENT_PER_CLASS", None),
        ("Tue", "Grade vs grade", D(2026, 11, 3), D(2026, 11, 3), "", "GRADE_TOTAL", None),
        ("Tue", "$20 day", D(2026, 11, 3), D(2026, 11, 3), "", "CLASS_PARTICIPATION_MIN", 20),
        ("All", "Total goal", D(2026, 10, 30), D(2026, 11, 13), "", "SCHOOL_TOTAL_TARGET", 255),
        ("Sun", "40% by Sunday", D(2026, 11, 1), D(2026, 11, 1), "", "SCHOOL_PARTICIPATION_TARGET", 0.4),
        ("Mon", "Showdown", D(2026, 11, 2), D(2026, 11, 2), "", "CLASS_PARTICIPATION_CUM", None),
    ]
    return {"settings": {"asof": D(2026, 11, 5)}, "roster": roster, "oc_headers": HEAD, "oc_rows": oc,
            "cash": cash, "challenges": challenges}


THERMO_AMOUNTS = [0, 32500, 60000, 65000, 80000]


def thermo(amount):
    data = {"settings": {"asof": D(2026, 11, 2)}, "oc_headers": HEAD, "oc_rows": [], "cash": [],
            "roster": [{"first": "Tia", "last": "Test", "grade": "K", "teacher": "Ms. T"}]}
    if amount:
        data["cash"] = [{"date": D(2026, 11, 2), "student": "S0001", "amount": amount, "method": "Check"}]
    return data
