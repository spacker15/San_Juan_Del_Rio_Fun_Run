"""Saints Run Wild theme: colors, fonts, fills, number formats."""
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

RED = "D7262B"
NAVY = "1F2F6B"
GREEN = "2E7D32"
LIME = "8BC34A"
SAND = "F6E7C1"
CREAM = "FFFBF0"
GOLD = "F2B705"
GRAY = "D9D9D9"
DARK = "222222"
WHITE = "FFFFFF"
INPUT = "FFF4D6"     # pale sand: cells volunteers type into
CALC = "EEF1F6"      # pale gray-blue: auto-calculated, don't type

TAB_INPUT = "2F6FD0"
TAB_OUTPUT = RED
TAB_REPORT = GREEN
TAB_ENGINE = "808080"

DISPLAY_FONT = "Arial Black"
BODY_FONT = "Calibri"

FMT_USD = '"$"#,##0'
FMT_USD2 = '"$"#,##0.00'
FMT_PCT = "0%"
FMT_PCT1 = "0.0%"
FMT_DATE = "ddd m/d/yyyy"
FMT_DATE_S = "m/d"


def fill(color):
    return PatternFill("solid", start_color=color, end_color=color)


def font(size=11, bold=False, color=DARK, name=BODY_FONT, italic=False):
    return Font(name=name, size=size, bold=bold, color=color, italic=italic)


def display(size=14, color=NAVY):
    return Font(name=DISPLAY_FONT, size=size, color=color)


thin = Side(style="thin", color="BFBFBF")
thick_navy = Side(style="thick", color=NAVY)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=False)
LEFT_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")

HEADER_FILL = fill(NAVY)
HEADER_FONT = font(11, True, WHITE)
