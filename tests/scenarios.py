"""Small hand-built datasets with known answers (ties, window boundaries, thermometer fill)."""
from datetime import date, datetime

HEAD = ["Transaction Date", "Transaction ID", "Participant Name", "Donor Name", "Amount", "Payment Method"]
D = date


def _oc(day, txn, name, amt):
    return [datetime(day.year, day.month, day.day, 12, 0), txn, name, "Donor " + txn, amt, "Credit Card"]


def ties():
    roster = [
        {"first": "Ann", "last": "One", "grade": "3", "teacher": "Ms. A", "family": "One"},
        {"first": "Ben", "last": "Two", "grade": "3", "teacher": "Ms. A", "family": "Two"},
        {"first": "Cal", "last": "Three", "grade": "3", "teacher": "Ms. B", "family": "Three"},
        {"first": "Dee", "last": "Four", "grade": "3", "teacher": "Ms. B", "family": "Four"},
        {"first": "Eve", "last": "Five", "grade": "4", "teacher": "Ms. C", "family": "Five"},
    ]
    oc = [
        _oc(D(2026, 10, 31), "T1", "Ann One", 50),
        _oc(D(2026, 11, 2), "T2", "Ann One", 30),
        _oc(D(2026, 11, 1), "T3", "Cal Three", 50),   # last day of the weekend window: included
        _oc(D(2026, 11, 3), "T4", "Cal Three", 40),
        _oc(D(2026, 11, 2), "T5", "Eve Five", 5),     # one day after the weekend window: excluded
    ]
    cash = [
        {"date": D(2026, 11, 2), "student": "S0002", "donor": "Two family", "amount": 30, "method": "Cash"},
        {"date": D(2026, 11, 2), "student": "S0004", "donor": "Four family", "amount": 10, "method": "Check"},
        {"date": D(2026, 11, 3), "student": "S0005", "donor": "Five family", "amount": 40, "method": "Cash"},
    ]
    challenges = [
        (D(2026, 10, 30), D(2026, 11, 1), "Wk1", "Weekend", "", "", "CLASS_TOTAL_PER_GRADE", None, None),
        (D(2026, 11, 2), D(2026, 11, 2), "Mon", "Top student", "", "", "TOP_STUDENT_PER_CLASS", None, None),
        (D(2026, 11, 3), D(2026, 11, 3), "Tue", "Grade vs grade", "", "", "GRADE_TOTAL", None, None),
        (D(2026, 11, 3), D(2026, 11, 3), "Tue", "$20 day", "", "", "CLASS_PARTICIPATION_MIN", 20, None),
        (D(2026, 10, 30), D(2026, 11, 13), "All", "Total goal", "", "", "SCHOOL_TOTAL_TARGET", None, 255),
        (D(2026, 11, 1), D(2026, 11, 1), "Sun", "40% by Sunday", "", "", "SCHOOL_PARTICIPATION_TARGET", None, 0.4),
        (D(2026, 11, 2), D(2026, 11, 2), "Mon", "Showdown", "", "", "CLASS_PARTICIPATION_CUM", None, None),
    ]
    return {"settings": {"asof": D(2026, 11, 5)}, "classrooms": [("Ms. A", "3", None), ("Ms. B", "3", None),
                                                                ("Ms. C", "4", None)],
            "roster": roster, "oc_headers": HEAD, "oc_rows": oc, "cash": cash, "challenges": challenges}


THERMO_AMOUNTS = [0, 32500, 60000, 65000, 80000]


def thermo(amount):
    data = {"settings": {"asof": D(2026, 11, 2)}, "classrooms": [("Ms. T", "K", None)],
            "roster": [{"first": "Tia", "last": "Test", "grade": "K", "teacher": "Ms. T"}],
            "oc_headers": HEAD, "oc_rows": [], "cash": []}
    if amount:
        data["cash"] = [{"date": D(2026, 11, 2), "student": "S0001", "amount": amount, "method": "Check"}]
    return data
