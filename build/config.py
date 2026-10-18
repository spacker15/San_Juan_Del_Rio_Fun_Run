"""Campaign defaults. Edit here, then re-run build_workbook.py."""
from datetime import date

# Capacity (pre-sized ranges keep the file fast)
N_STU = 700      # roster rows
N_CLS = 40       # classrooms (teacher + grade)
N_OC = 3000      # OneCause rows per paste
N_CC = 1000      # cash / check rows
N_DAY = 21       # day columns (campaign is 15 days)
N_ELIG = 200     # rows in the Wild Card eligible list

GRADES = ["P3", "P4", "K", "1", "2", "3", "4", "5", "6", "7", "8"]
METHODS = ["Cash", "Check", "Other"]

SETTINGS = {
    "school": "San Juan del Rio Catholic School",
    "campaign": "Saints Run Wild",
    "year": 2026,
    "goal": 65000,
    "stretch": 60000,
    "stretch_label": "Father Peter & Mrs. Rhodes get SLIMED!",
    "start": date(2026, 10, 30),
    "end": date(2026, 11, 13),
    "asof": None,  # blank = today
}

# OneCause export headers (matched by name, so column order doesn't matter)
ONECAUSE_HEADERS = [
    ("Date", "Transaction Date"),
    ("Amount", "Amount"),
    ("Student name", "Participant Name"),
    ("Donor name", "Donor Name"),
    ("Transaction ID", "Transaction ID"),
    ("Payment type", "Payment Method"),
]

# Daily incentives flyer. Dates/names/rewards/min/target are editable on Settings;
# the scoring rule for each row is fixed here.
# (day, name, start, end, reward, rule, min-or-target)
D = date
CHALLENGES = [
    ("Wk1 Fri", "Kickoff Weekend Challenge", D(2026, 10, 30), D(2026, 11, 1),
     "Top class in every grade wins", "CLASS_TOTAL_PER_GRADE", None),
    ("Wk1 Mon", "Monday Blues", D(2026, 11, 2), D(2026, 11, 2),
     "Turn the blues into green", "NONE", None),
    ("Wk1 Tue", "$20 Tuesday", D(2026, 11, 3), D(2026, 11, 3),
     "Silly string the teacher", "CLASS_PARTICIPATION_MIN", 20),
    ("Wk1 Wed", "Wild Card Wednesday", D(2026, 11, 4), D(2026, 11, 4),
     "10 Wild Card winners", "DRAWING", 10),
    ("Wk1 Thu", "Classroom Showdown", D(2026, 11, 5), D(2026, 11, 5),
     "Doughnuts for the class, coffee for the teacher", "CLASS_PARTICIPATION_CUM", None),
    ("Wk2 Fri", "Weekend Challenge", D(2026, 11, 6), D(2026, 11, 8),
     "Top fundraiser in every class pies a teacher", "TOP_STUDENT_PER_CLASS", None),
    ("Wk2 Mon", "Grade vs. Grade", D(2026, 11, 9), D(2026, 11, 9),
     "Pizza lunch for the winning grade", "GRADE_TOTAL", None),
    ("Wk2 Tue", "50% Participation", D(2026, 11, 10), D(2026, 11, 10),
     "Mrs. Rhodes & Mrs. Sites dress up", "SCHOOL_PARTICIPATION_TARGET", 0.5),
    ("Wk2 Wed", "Veterans Day (no school)", D(2026, 11, 11), D(2026, 11, 11),
     "Thank you, Veterans!", "NONE", None),
    ("Wk2 Thu", "Mystery Day", D(2026, 11, 12), D(2026, 11, 12),
     "The final WILD push", "NONE", None),
    ("Wk2 Fri", "Run Wild Fun Run!", D(2026, 11, 13), D(2026, 11, 13),
     "Celebrate!", "NONE", None),
    ("All", "Ultimate Wild Goal", D(2026, 10, 30), D(2026, 11, 13),
     "Father Peter & Mrs. Rhodes get SLIMED", "SCHOOL_TOTAL_TARGET", 60000),
]

RULE_TEXT = {
    "CLASS_TOTAL_PER_GRADE": "Most $ per class, winner in each grade",
    "CLASS_PARTICIPATION_MIN": "Highest % of class with a gift of at least Min $ during the dates",
    "CLASS_PARTICIPATION_CUM": "Highest % of class that has donated (campaign to date)",
    "DRAWING": "Students with a gift that day are eligible; draw Target names",
    "TOP_STUDENT_PER_CLASS": "Top student $ in each class during the dates",
    "GRADE_TOTAL": "Grade with the most $ during the dates",
    "SCHOOL_PARTICIPATION_TARGET": "School participation reaches Target (0.5 = 50%)",
    "SCHOOL_TOTAL_TARGET": "School total reaches Target $",
    "NONE": "No winner to calculate",
}
