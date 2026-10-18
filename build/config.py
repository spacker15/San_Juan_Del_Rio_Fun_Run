"""Campaign defaults. Edit here, then re-run build_workbook.py."""
from datetime import date

N_DON = 3000    # donation rows
N_STU = 600     # student rows
N_TCH = 30      # teacher rows
N_DAY = 15      # rows on By Day

GOAL = 65000
FIRST_DAY = date(2026, 10, 30)
GRADES = ["P3", "P4", "K", "1", "2", "3", "4", "5", "6", "7", "8"]

# Daily incentives flyer: (day, challenge, start, end, reward, rule, min $ or target)
D = date
CHALLENGES = [
    ("Fri", "Kickoff Weekend Challenge", D(2026, 10, 30), D(2026, 11, 1), "Top class in every grade wins",
     "CLASS_TOTAL_PER_GRADE", None),
    ("Mon", "Monday Blues", D(2026, 11, 2), D(2026, 11, 2), "Turn the blues into green", "NONE", None),
    ("Tue", "$20 Tuesday", D(2026, 11, 3), D(2026, 11, 3), "Silly string the teacher", "CLASS_PCT_MIN", 20),
    ("Wed", "Wild Card Wednesday", D(2026, 11, 4), D(2026, 11, 4), "10 Wild Card winners", "DRAWING", 10),
    ("Thu", "Classroom Showdown", D(2026, 11, 5), D(2026, 11, 5), "Doughnuts + coffee for the teacher",
     "CLASS_PCT_TO_DATE", None),
    ("Fri", "Weekend Challenge", D(2026, 11, 6), D(2026, 11, 8), "Top fundraiser in every class pies a teacher",
     "TOP_STUDENT_PER_CLASS", None),
    ("Mon", "Grade vs. Grade", D(2026, 11, 9), D(2026, 11, 9), "Pizza lunch for the winning grade",
     "GRADE_TOTAL", None),
    ("Tue", "50% Participation", D(2026, 11, 10), D(2026, 11, 10), "Mrs. Rhodes & Mrs. Sites dress up",
     "SCHOOL_PCT_TARGET", 0.5),
    ("Wed", "Veterans Day (no school)", D(2026, 11, 11), D(2026, 11, 11), "Thank you, Veterans!", "NONE", None),
    ("Thu", "Mystery Day", D(2026, 11, 12), D(2026, 11, 12), "The final WILD push", "NONE", None),
    ("Fri", "Run Wild Fun Run!", D(2026, 11, 13), D(2026, 11, 13), "Celebrate!", "NONE", None),
    ("All", "Ultimate Wild Goal", D(2026, 10, 30), D(2026, 11, 13), "Father Peter & Mrs. Rhodes get SLIMED",
     "SCHOOL_TOTAL_TARGET", 60000),
]
