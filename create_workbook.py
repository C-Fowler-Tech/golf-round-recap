"""
create_workbook.py
Generates (or regenerates) the Golf Round Recap Excel workbook.
Run this once to create the file, or again to reset the structure.
WARNING: running again will overwrite the file and lose any data.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import date
import os
import shutil
import pathlib

# Template saved to the repo for source control
REPO_DIR      = pathlib.Path(__file__).parent
TEMPLATE_PATH = REPO_DIR / "Golf Round Recap.xlsx"

# Live working file on Google Drive
DRIVE_PATH = pathlib.Path(r"G:\My Drive\Project_Outputs\Golf Round Recap\Golf Round Recap.xlsx")

# ── Styles ────────────────────────────────────────────────────────────────────
HDR_FILL  = PatternFill("solid", fgColor="1F4E79")
HDR_FONT  = Font(color="FFFFFF", bold=True, name="Calibri")
BODY_FONT = Font(name="Calibri", size=11)
ALT_FILL  = PatternFill("solid", fgColor="D9E8F5")

def style_header(cell):
    cell.font = HDR_FONT
    cell.fill = HDR_FILL
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def write_row(ws, row_num, values, fill=None):
    for col, val in enumerate(values, 1):
        cell = ws.cell(row=row_num, column=col, value=val)
        cell.font = BODY_FONT
        if fill:
            cell.fill = fill


# ============================================================
# TAB 1 -- ROUNDS  (24 columns)
# ============================================================
#  1  Date
#  2  Course
#  3  Note Type
#  4  Hole
#  5  Par
#  6  Distance (m)
#  7  Stroke Index        <- populated from Courses tab for the hole
#  8  Score
#  9  Strokes
# 10  Putts
# 11  Penalties
# 12  Tee Club
# 13  Pick Up
# 14  Sentiment
# 15  Driver              }
# 16  Woods               }
# 17  Hybrids             } Overall rows only -- ball striking ratings
# 18  Long Irons (5-7)    }
# 19  Short Irons (8-P)   }
# 20  Wedges (GW/SW/LW)   }
# 21  Putter              }
# 22  Playing Handicap
# 23  Tee Colour
# 24  Notes
# ============================================================
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Rounds"

ROUND_HEADERS = [
    "Date", "Course", "Note Type", "Hole", "Par", "Distance (m)", "Stroke Index",
    "Score", "Strokes", "Putts", "Penalties",
    "Tee Club", "Pick Up", "Sentiment",
    "Driver", "Woods", "Hybrids", "Long Irons\n(5-7)", "Short Irons\n(8-P)", "Wedges\n(GW/SW/LW)", "Putter",
    "Playing\nHandicap", "Tee Colour",
    "Notes",
]
ROUND_COL_WIDTHS = [
    12, 20, 12, 7, 6, 14, 13,
    16, 10, 8, 11,
    14, 10, 12,
    11, 10, 11, 13, 13, 13, 10,
    11, 12,
    60,
]

ws.row_dimensions[1].height = 36
for col, h in enumerate(ROUND_HEADERS, 1):
    style_header(ws.cell(row=1, column=col, value=h))
    ws.column_dimensions[get_column_letter(col)].width = ROUND_COL_WIDTHS[col - 1]

ws.freeze_panes = "A2"

# Data validations
def dv(formula, sqref):
    v = DataValidation(type="list", formula1=formula, allow_blank=True, showErrorMessage=False)
    v.sqref = sqref
    return v

STRIKE_RATING = '"Great,Good,Average,Poor"'
ws.add_data_validation(dv('"Overall,Hole"',                                           "C2:C9999"))
ws.add_data_validation(dv('"Eagle,Birdie,Par,Bogey,Double Bogey,Triple Bogey,Other"', "H2:H9999"))
ws.add_data_validation(dv('"Y,N"',                                                    "M2:M9999"))
ws.add_data_validation(dv('"Positive,Neutral,Negative"',                              "N2:N9999"))
ws.add_data_validation(dv(STRIKE_RATING,                                              "O2:O9999"))
ws.add_data_validation(dv(STRIKE_RATING,                                              "P2:P9999"))
ws.add_data_validation(dv(STRIKE_RATING,                                              "Q2:Q9999"))
ws.add_data_validation(dv(STRIKE_RATING,                                              "R2:R9999"))
ws.add_data_validation(dv(STRIKE_RATING,                                              "S2:S9999"))
ws.add_data_validation(dv(STRIKE_RATING,                                              "T2:T9999"))
ws.add_data_validation(dv(STRIKE_RATING,                                              "U2:U9999"))
ws.add_data_validation(dv('"White,Yellow,Red,Blue,Black"',                            "W2:W9999"))

# ── Sample round (Pupuke, 22-Feb-2026, White tees) ───────────────────────────
write_row(ws, 2, [
    date(2026, 2, 22), "Pupuke", "Overall", 0, 70, 5188, None,
    85, 85, 32, 2, "", "", "Positive",
    "Good", "Average", "", "Good", "Average", "Good", "Good",
    18, "White",
    "Tee time 8:30am. Fine autumn morning, light wind. Course in great condition. "
    "Happy with the round - hit it well off the tee, short game let me down on a few holes.",
], fill=ALT_FILL)

sample_holes = [
    [date(2026, 2, 22), "Pupuke", "Hole", 1, 4, 300,  7, "Bogey",  5, 2, 0, "Driver", "N", "Neutral",  "", "", "", "", "", "", "", "", "", "Good drive, approach came up short. Chip to 4m, two-putted."],
    [date(2026, 2, 22), "Pupuke", "Hole", 2, 3, 139, 15, "Par",    3, 1, 0, "7 Iron", "N", "Positive", "", "", "", "", "", "", "", "", "", "Solid tee shot to 3m, holed the putt. Exactly the plan."],
    [date(2026, 2, 22), "Pupuke", "Hole", 3, 4, 335,  3, "Birdie", 3, 1, 0, "Driver", "N", "Positive", "", "", "", "", "", "", "", "", "", "Big drive, wedge to 1.5m and holed it. Best hole of the day."],
]
for i, row in enumerate(sample_holes):
    write_row(ws, 3 + i, row)


# ============================================================
# TAB 2 -- COURSES  (7 columns)
# ============================================================
#  1  Course
#  2  Hole
#  3  Tee Colour    <- par and distance vary by tee
#  4  Par
#  5  Distance (m)
#  6  Stroke Index
#  7  Notes
# ============================================================
ws_c = wb.create_sheet("Courses")

COURSE_HEADERS = ["Course", "Hole", "Tee Colour", "Par", "Distance (m)", "Stroke Index", "Notes"]
COURSE_COL_WIDTHS = [22, 7, 12, 6, 14, 14, 40]

ws_c.row_dimensions[1].height = 30
for col, h in enumerate(COURSE_HEADERS, 1):
    style_header(ws_c.cell(row=1, column=col, value=h))
    ws_c.column_dimensions[get_column_letter(col)].width = COURSE_COL_WIDTHS[col - 1]

ws_c.freeze_panes = "A2"

# Pupuke Golf Club -- White tees (verified)
# Par 70, 5188m total
PUPUKE = [
    # hole, tee,     par, dist_m, stroke_index
    ( 1, "White",  4,  300,  7),
    ( 2, "White",  3,  139, 15),
    ( 3, "White",  4,  335,  3),
    ( 4, "White",  4,  304, 13),
    ( 5, "White",  5,  431,  9),
    ( 6, "White",  3,  165, 11),
    ( 7, "White",  4,  363,  1),
    ( 8, "White",  4,  333,  5),
    ( 9, "White",  3,  147, 17),
    (10, "White",  5,  422, 14),
    (11, "White",  5,  398, 16),
    (12, "White",  4,  362,  2),
    (13, "White",  3,  167,  8),
    (14, "White",  4,  285, 10),
    (15, "White",  4,  299,  6),
    (16, "White",  4,  238, 18),
    (17, "White",  3,  143, 12),
    (18, "White",  4,  357,  4),
]

for row_i, (hole, tee, par, dist, si) in enumerate(PUPUKE, 2):
    fill = ALT_FILL if row_i % 2 == 1 else None
    write_row(ws_c, row_i, ["Pupuke", hole, tee, par, dist, si, ""], fill=fill)

total_par  = sum(h[2] for h in PUPUKE)
total_dist = sum(h[3] for h in PUPUKE)
bold = Font(bold=True, name="Calibri")
ws_c.cell(row=20, column=1, value="Pupuke").font = bold
ws_c.cell(row=20, column=2, value="TOTAL").font  = bold
ws_c.cell(row=20, column=3, value="White").font  = bold
ws_c.cell(row=20, column=4, value=total_par).font  = bold
ws_c.cell(row=20, column=5, value=total_dist).font = bold


# ── Save ─────────────────────────────────────────────────────────────────────
# 1. Save template to repo (source control)
wb.save(TEMPLATE_PATH)
print(f"Template saved : {TEMPLATE_PATH}")

# 2. Copy to Google Drive (live working file)
DRIVE_PATH.parent.mkdir(parents=True, exist_ok=True)
shutil.copy2(TEMPLATE_PATH, DRIVE_PATH)
print(f"Copied to Drive: {DRIVE_PATH}")

print(f"  Pupuke (White): Par {total_par}, {total_dist}m")
print(f"  Rounds tab: 24 columns | Courses tab: 7 columns")
