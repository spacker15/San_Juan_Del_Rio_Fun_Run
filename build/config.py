"""Campaign defaults and workbook capacity limits.

Edit this file (not the workbook generator) to change next year's defaults.
"""
from datetime import date

# ---- Capacity (pre-sized ranges keep formulas fast and predictable) ----
N_STU = 800      # roster rows
N_CLS = 40       # classrooms
N_OC = 5000      # OneCause rows per paste
N_CC = 2000      # cash / check rows
N_CH = 15        # challenge slots
N_PR = 12        # prize tiers
N_DAY = 31       # campaign days
N_FAM = 600      # unique families
N_DON = 1500     # unique donors
N_LIST = 40      # rows in Data Check problem lists

GRADES = ["P3", "P4", "K", "1", "2", "3", "4", "5", "6", "7", "8"]

PRIVACY_OPTIONS = ["First name + last initial", "Full name", "Initials only"]
METHODS = ["Cash", "Check", "Other"]

METRICS = [
    "CLASS_TOTAL_PER_GRADE",
    "CLASS_PARTICIPATION_MIN",
    "CLASS_PARTICIPATION_CUM",
    "DRAWING",
    "TOP_STUDENT_PER_CLASS",
    "GRADE_TOTAL",
    "SCHOOL_PARTICIPATION_TARGET",
    "SCHOOL_TOTAL_TARGET",
    "NONE",
]

SETTINGS = {
    "campaign": "Saints Run Wild",
    "school": "San Juan del Rio Catholic School",
    "year": 2026,
    "goal": 65000,
    "stretch": 60000,
    "stretch_label": "SLIME! Father Peter & Mrs. Rhodes",
    "start": date(2026, 10, 30),
    "end": date(2026, 11, 13),
    "part_min": 0.01,
    "privacy": PRIVACY_OPTIONS[0],
    "tagline": "GO WILD • RAISE BIG • WIN FUN!",
    "comp_year": 2025,
    "asof": None,  # blank = TODAY()
}

# OneCause export header names (matched by text, so column order doesn't matter).
ONECAUSE_MAP = [
    ("Date", "Transaction Date", True),
    ("Amount", "Amount", True),
    ("Student full name", "Participant Name", False),
    ("Student first name (if split)", "", False),
    ("Student last name (if split)", "", False),
    ("Donor name", "Donor Name", False),
    ("Transaction ID", "Transaction ID", False),
    ("Payment type", "Payment Method", False),
]

# Individual milestone prizes (SAMPLE: edit on the Prizes tab). Ascending thresholds.
PRIZES = [
    ("Wild Sticker Pack", 25, "Raise $25+"),
    ("Lion Sunglasses", 50, "Raise $50+"),
    ("Jungle Water Bottle", 100, "Raise $100+"),
    ("Run Wild Hat", 150, "Raise $150+"),
    ("Saints Hoodie", 250, "Raise $250+"),
    ("Principal for an Hour", 500, "Raise $500+"),
]

# Daily incentive calendar from the "Saints Run Wild Daily Incentives" flyer.
# (start, end, day label, name, description, reward, metric, min $, target)
D = date
CHALLENGES = [
    (D(2026, 10, 30), D(2026, 11, 1), "Wk1 Fri-Sun", "Kickoff Day: Weekend Challenge",
     "Which class in each grade raises the most over the weekend?", "Top class in every grade wins",
     "CLASS_TOTAL_PER_GRADE", None, None),
    (D(2026, 11, 2), D(2026, 11, 2), "Wk1 Mon", "Monday Blues",
     "Turn the blues into green: get a donation today", "Bragging rights", "NONE", None, None),
    (D(2026, 11, 3), D(2026, 11, 3), "Wk1 Tue", "$20 Tuesday",
     "Find ONE person to donate $20", "Silly string the teacher (highest class participation)",
     "CLASS_PARTICIPATION_MIN", 20, None),
    (D(2026, 11, 4), D(2026, 11, 4), "Wk1 Wed", "Wild Card Wednesday",
     "Any donation today = entry into the Wild Card drawing", "10 Wild Card winners",
     "DRAWING", None, 10),
    (D(2026, 11, 5), D(2026, 11, 5), "Wk1 Thu", "Classroom Showdown",
     "Help your class reach the highest participation", "Doughnuts for the class + coffee for the teacher",
     "CLASS_PARTICIPATION_CUM", None, None),
    (D(2026, 11, 6), D(2026, 11, 8), "Wk2 Fri-Sun", "Weekend Challenge",
     "Which student in each class raises the most over the weekend?", "Top fundraiser in every class pies a teacher",
     "TOP_STUDENT_PER_CLASS", None, None),
    (D(2026, 11, 9), D(2026, 11, 9), "Wk2 Mon", "Grade vs. Grade",
     "Which grade can bring in the most today?", "Pizza lunch for the winning grade",
     "GRADE_TOTAL", None, None),
    (D(2026, 11, 10), D(2026, 11, 10), "Wk2 Tue", "50% Participation",
     "If we reach 50% participation today...", "Mrs. Rhodes & Mrs. Sites dress up",
     "SCHOOL_PARTICIPATION_TARGET", None, 0.5),
    (D(2026, 11, 11), D(2026, 11, 11), "Wk2 Wed", "Veterans Day: No School",
     "Today we pause to honor all who have served", "Thank you, Veterans!", "NONE", None, None),
    (D(2026, 11, 12), D(2026, 11, 12), "Wk2 Thu", "Mystery Day!",
     "Wait and find out: the final WILD push", "Mystery prize", "NONE", None, None),
    (D(2026, 11, 13), D(2026, 11, 13), "Wk2 Fri", "Run Wild Fun Run!",
     "Celebrate, it's time to RUN WILD!", "Fun Run", "NONE", None, None),
    (D(2026, 10, 30), D(2026, 11, 13), "Campaign", "Ultimate Wild Goal",
     "If we hit $60,000...", "Father Peter & Mrs. Rhodes get SLIMED", "SCHOOL_TOTAL_TARGET", None, 60000),
]
