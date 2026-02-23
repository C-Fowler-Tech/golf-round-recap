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


# ============================================================
# TAB 3 -- GUIDE
# ============================================================
ws_g = wb.create_sheet("Guide")
ws_g.sheet_view.showGridLines = False

GUIDE_HDR = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
GUIDE_SEC = Font(bold=True, name="Calibri", size=12)
GUIDE_BODY = Font(name="Calibri", size=11)
SEC_FILL = PatternFill("solid", fgColor="1F4E79")
SUB_FILL = PatternFill("solid", fgColor="2E75B6")
SUB_FONT = Font(color="FFFFFF", bold=True, name="Calibri", size=10)

ws_g.column_dimensions["A"].width = 30
ws_g.column_dimensions["B"].width = 18
ws_g.column_dimensions["C"].width = 50

def guide_section(row, title):
    cell = ws_g.cell(row=row, column=1, value=title)
    cell.font = GUIDE_SEC
    cell.fill = SEC_FILL
    cell.font = Font(color="FFFFFF", bold=True, name="Calibri", size=12)
    ws_g.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    ws_g.row_dimensions[row].height = 20

def guide_subhdr(row, cols):
    for col, val in enumerate(cols, 1):
        cell = ws_g.cell(row=row, column=col, value=val)
        cell.font = SUB_FONT
        cell.fill = SUB_FILL

def guide_row(row, vals):
    for col, val in enumerate(vals, 1):
        cell = ws_g.cell(row=row, column=col, value=val)
        cell.font = GUIDE_BODY
        if row % 2 == 0:
            cell.fill = ALT_FILL

r = 1
guide_section(r, "SCORE LABELS -- use these in the Score column for Hole rows"); r += 1
guide_subhdr(r, ["Result", "Strokes vs Par", "Example (par 4)"]); r += 1
for vals in [
    ("Eagle",        "-2 or better", "2 strokes on a par 4"),
    ("Birdie",       "-1",           "3 strokes on a par 4"),
    ("Par",          "0",            "4 strokes on a par 4"),
    ("Bogey",        "+1",           "5 strokes on a par 4"),
    ("Double Bogey", "+2",           "6 strokes on a par 4"),
    ("Triple Bogey", "+3",           "7 strokes on a par 4"),
    ("Other",        "+4 or worse, or pick up", "8+ strokes, or picked up"),
]:
    guide_row(r, vals); r += 1

r += 1
guide_section(r, "SENTIMENT -- how you felt the hole or round went"); r += 1
guide_subhdr(r, ["Sentiment", "Your words might include...", ""]); r += 1
for vals in [
    ("Positive",  "great, best hole, happy, love it, exactly the plan, nice, solid, holed it", ""),
    ("Neutral",   "ok, sensible, got away with it, fine, average, recovered, not bad", ""),
    ("Negative",  "disappointed, disaster, terrible, nightmare, duffed, dire, struggled, awful, hack, lucky to escape, poor", ""),
]:
    guide_row(r, vals); r += 1

r += 1
guide_section(r, "BALL STRIKING RATINGS -- for Overall rows, one rating per club category"); r += 1
guide_subhdr(r, ["Rating", "Your words might include...", ""]); r += 1
for vals in [
    ("Great",   "monster, perfect, very good, flushed it, exactly where I wanted", ""),
    ("Good",    "good, solid, nice, decent, hit it well",                          ""),
    ("Average", "ok, bit fadey, slight fade/slice, average, could be better",      ""),
    ("Poor",    "duffed, hacked, bladed, hooked, sliced, below average, dire, terrible, sprayed", ""),
]:
    guide_row(r, vals); r += 1

r += 1
guide_section(r, "PICK UP -- did you finish the hole?"); r += 1
guide_subhdr(r, ["Value", "Meaning", ""]); r += 1
guide_row(r, ["N", "Holed out -- counted every stroke", ""]); r += 1
guide_row(r, ["Y", "Picked up / did not finish -- score is an estimate", ""]); r += 1

r += 1
guide_section(r, "OVERALL ROW TIPS -- one per round, Hole = 0"); r += 1
ws_g.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
tips = [
    "Notes should cover: tee time, weather, wind, course conditions, overall impressions.",
    "Score / Strokes = gross total for the round.",
    "Ball striking columns = how each club category felt across the whole round.",
    "Putts = total putts for the round (count all holes).",
    "Penalties = total penalty strokes for the round.",
    "Tee Colour = which tees played (affects par, distance, stroke index).",
]
for tip in tips:
    cell = ws_g.cell(row=r, column=1, value=tip)
    cell.font = GUIDE_BODY
    ws_g.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    r += 1


# ── Save ─────────────────────────────────────────────────────────────────────
# Saves template to repo only -- does NOT touch the live Drive file.
# To initialise a fresh Drive file (e.g. new season), manually copy:
#   copy "Golf Round Recap.xlsx" "G:\My Drive\Project_Outputs\Golf Round Recap\Golf Round Recap.xlsx"
wb.save(TEMPLATE_PATH)
print(f"Template saved : {TEMPLATE_PATH}")
print(f"  Pupuke (White): Par {total_par}, {total_dist}m")
print(f"  Tabs: Rounds (24 cols) | Courses (7 cols) | Guide")
print(f"  Drive file NOT touched -- copy manually only when a full reset is needed.")
