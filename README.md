# Golf Round Recap

A personal golf database tracking rounds hole-by-hole. Each round is stored as a set of rows in Excel — one **Overall** row capturing round-level context, and one **Hole** row per hole played. Over time this builds a dataset for analysis: what clubs work, where strokes are lost, which courses suit your game.

---

## File

**`Golf Round Recap.xlsx`** — two tabs:

| Tab | Purpose |
|---|---|
| `Rounds` | One row per hole + one Overall row per round |
| `Courses` | Course reference data (par, distance, stroke index per tee colour) |
| `Guide` | Scoring labels, sentiment guide, ball striking ratings — quick reference for data entry |

**Locations:**
- `C:\Users\craig.f\Home_Projects\Golf Round Recap\Golf Round Recap.xlsx` — base template in source control
- `G:\My Drive\Project_Outputs\Golf Round Recap\Golf Round Recap.xlsx` — live working file (enter data here)

---

## Rounds Tab — Column Guide (29 columns)

| # | Column | Type | Notes |
|---|---|---|---|
| 1 | **Date** | Date | Date of the round |
| 2 | **Course** | Text | Match the name used in the Courses tab |
| 3 | **Note Type** | Dropdown | `Overall` or `Hole` |
| 4 | **Hole** | Number | Hole number 1–18. Use **0** for the Overall row |
| 5 | **Par** | Number | Par for this hole. For Overall: total course par for the tee played |
| 6 | **Distance (m)** | Number | Hole distance in metres. For Overall: total course distance |
| 7 | **Stroke Index** | Number | Stroke index for this hole (from Courses tab). Leave blank for Overall |
| 8 | **Score** | Dropdown | Hole result: `Hole in One` / `Eagle` / `Birdie` / `Par` / `Bogey` / `Double Bogey` / `Triple Bogey` / `Pick Up` / `Other`. For Overall: gross total score (e.g. `85`) |
| 9 | **Strokes** | Number | Actual strokes taken. For Overall: total round strokes |
| 10 | **Putts** | Number | Number of putts. For Overall: total round putts |
| 11 | **Penalties** | Number | Penalty strokes. For Overall: total round penalties |
| 12 | **FIR** | Dropdown | **Hole rows, par 4/5 only.** `Y` = tee shot on fairway. `N` = missed. Leave blank on par 3s and Overall |
| 13 | **GIR** | Dropdown | **Hole rows only.** `Y` = ball on green within regulation shots (par 3: tee shot / par 4: within 2 / par 5: within 3). Leave blank on pick-ups and Overall |
| 14 | **Tee Club** | Text | Club used off the tee (e.g. `Driver`, `3 Hybrid`, `5 Wood`) |
| 15 | **Pick Up** | Dropdown | `Y` if you picked up / didn't finish the hole. `N` otherwise |
| 16 | **Sentiment** | Dropdown | `Positive` / `Neutral` / `Negative` — how you felt the hole went |
| 17 | **Driver** | Dropdown | **Overall rows only.** Ball striking: `Great` / `Good` / `Average` / `Poor` |
| 18 | **Woods** | Dropdown | Overall rows only. Fairway woods striking rating |
| 19 | **Hybrids** | Dropdown | Overall rows only. Hybrid striking rating |
| 20 | **Long Irons (5-7)** | Dropdown | Overall rows only. Long iron striking rating |
| 21 | **Short Irons (8-P)** | Dropdown | Overall rows only. Short iron striking rating |
| 22 | **Wedges (GW/SW/LW)** | Dropdown | Overall rows only. Wedge striking rating |
| 23 | **Putter** | Dropdown | Overall rows only. Putting feel rating |
| 24 | **Playing Handicap** | Number | Your playing handicap for the round |
| 25 | **Tee Colour** | Dropdown | `White` / `Yellow` / `Red` / `Blue` / `Black` |
| 26 | **Course Rating** | Number | **Overall rows only.** Course rating at time of play (e.g. `68.2`) — persisted so changes over time are captured |
| 27 | **Slope** | Number | Overall rows only. Slope rating at time of play (e.g. `119`) |
| 28 | **WHS Index** | Number | Overall rows only. Your WHS handicap index at time of play (e.g. `24.4`) |
| 29 | **Notes** | Text | Free text detail (see prompts below) |

---

## Data Entry — Q&A Prompts

### Starting a new round

Add the **Overall row first** (Note Type = `Overall`, Hole = `0`):

> **Date?** Date played.
> **Course?** e.g. `Pupuke`
> **Tee Colour?** e.g. `White` — determines which par/distance/stroke index to use from Courses tab. **Default: White if not specified.**
> **Par / Distance?** Copy totals from the Courses tab for the tee colour played.
> **Score / Strokes?** Your gross total for the round.
> **Putts / Penalties?** Round totals.
> **Playing Handicap?** Your handicap for this round.
> **Course Rating / Slope / WHS Index?** Fill from the Courses tab and your current WHS index — persisted here so it's captured if they change over time.
> **Sentiment?** How did the round feel overall?
> **Ball striking (Driver / Woods / Hybrids / Long Irons / Short Irons / Wedges / Putter)?**
> Rate each category: `Great` / `Good` / `Average` / `Poor`. Leave blank if you didn't use that club type.
> **Notes?** Cover:
> - Tee time
> - Weather (sun, wind, rain, temperature)
> - Course condition (fairways, greens, rough)
> - Overall impressions — what worked, what didn't

### For each hole (Note Type = `Hole`)

> **Hole / Par / Distance / Stroke Index?** Look up from the Courses tab for the tee colour played.
> **FIR?** Par 4/5 only — did the tee shot find the fairway? `Y` / `N`. Leave blank on par 3s.
> **GIR?** Did you hit the green in regulation? `Y` / `N`. Leave blank on pick-ups.
> **Score?** Select result from dropdown (Birdie, Bogey, etc.)
> **Strokes?** Actual count including penalties.
> **Putts?** How many putts did you take?
> **Penalties?** Any penalty strokes (OB, water, lost ball)?
> **Tee Club?** What did you hit off the tee?
> **Pick Up?** Did you finish the hole? `Y` = picked up, `N` = holed out.
> **Sentiment?** How did the hole feel — did it go to plan?
> **Notes?** Be specific — this is where the value is:
> - What happened off the tee?
> - What club/shot for approach?
> - Any misses — why and where?
> - Putting — distances, number of putts, any lips or misreads?
> - What would you do differently?

---

## Guide Tab — Sentiment & Rating Reference

### Score labels

| Result | Strokes vs Par | Example (par 4) |
|---|---|---|
| Hole in One | Par 3, 1 stroke | The dream |
| Albatross | -3 | 2 strokes on a par 5 |
| Eagle | -2 | 2 strokes on a par 4 |
| Birdie | -1 | 3 strokes |
| Par | 0 | 4 strokes |
| Bogey | +1 | 5 strokes |
| Double Bogey | +2 | 6 strokes |
| Triple Bogey | +3 | 7 strokes |
| Pick Up | Did not finish hole | Use with Pick Up column = Y |
| Other | +4 or worse | 8+ strokes, holed out |

### Sentiment

| Value | Your words might include... |
|---|---|
| **Positive** | great, best hole, happy, exactly the plan, nice, solid, holed it, love it |
| **Neutral** | ok, sensible, got away with it, fine, recovered, not bad, average |
| **Negative** | disappointed, disaster, terrible, nightmare, duffed, dire, struggled, hack, awful, poor, lucky to escape |

### Ball striking ratings (Overall rows)

| Rating | Your words might include... |
|---|---|
| **Great** | monster, perfect, very good, flushed it, exactly where I wanted |
| **Good** | good, solid, nice, decent, hit it well |
| **Average** | ok, bit fadey, slight fade/slice, could be better |
| **Poor** | duffed, hacked, bladed, hooked, sliced, below average, dire, terrible, sprayed |

---

## Courses Tab — Column Guide (7 columns)

| # | Column | Notes |
|---|---|---|
| 1 | **Course** | Course name — must match exactly what you use in the Rounds tab |
| 2 | **Hole** | Hole number 1–18, then TOTAL row |
| 3 | **Tee Colour** | Tee the data applies to — add separate rows per tee if needed |
| 4 | **Par** | Par for the hole from this tee |
| 5 | **Distance (m)** | Distance in metres from this tee |
| 6 | **Stroke Index** | Handicap stroke index (1 = hardest, 18 = easiest) |
| 7 | **Course Rating** | On the TOTAL row — course rating for this tee (e.g. `68.2`) |
| 8 | **Slope** | On the TOTAL row — slope rating for this tee (e.g. `119`) |
| 9 | **Notes** | Any notes about the hole (shape, hazards, typical miss) |

### Pupuke Golf Club (White tees) — verified data

| Hole | Par | Distance (m) | Stroke Index |
|------|-----|-------------|--------------|
| 1 | 4 | 300 | 7 |
| 2 | 3 | 139 | 15 |
| 3 | 4 | 335 | 3 |
| 4 | 4 | 304 | 13 |
| 5 | 5 | 431 | 9 |
| 6 | 3 | 165 | 11 |
| 7 | 4 | 363 | 1 |
| 8 | 4 | 333 | 5 |
| 9 | 3 | 147 | 17 |
| 10 | 5 | 422 | 14 |
| 11 | 5 | 398 | 16 |
| 12 | 4 | 362 | 2 |
| 13 | 3 | 167 | 8 |
| 14 | 4 | 285 | 10 |
| 15 | 4 | 299 | 6 |
| 16 | 4 | 238 | 18 |
| 17 | 3 | 143 | 12 |
| 18 | 4 | 357 | 4 |
| **Total** | **70** | **5188** | |

### Adding a new course

Add 18 rows per tee colour (e.g. 18 rows for White, 18 for Yellow) with the course name matching exactly what you'll type in the Rounds tab.

---

## Analysis Ideas (future)

Once you have a few rounds in, useful things to look at:

- **Scoring average by hole** — which holes consistently cost you strokes?
- **Scoring by stroke index** — do you score better on easier holes as expected?
- **Scoring by club off the tee** — driver vs. fairway wood vs. iron
- **Putts per round trend** — is the short game improving?
- **Sentiment vs. score correlation** — do positive holes actually score better?
- **Pick-up rate** — how often are you not finishing holes and on which?
- **Ball striking trends** — which club categories are improving over time?

---

## File locations

| File | Path |
|---|---|
| Live data file (edit here) | `G:\My Drive\Project_Outputs\Golf Round Recap\Golf Round Recap.xlsx` |
| Base template (source control) | `C:\Users\craig.f\Home_Projects\Golf Round Recap\Golf Round Recap.xlsx` |
| GitHub repo | https://github.com/C-Fowler-Tech/golf-round-recap |

> The **How To** tab in the workbook mirrors this section — both are kept in sync via source control.

---

## Workflow — entering a new round

1. Open the live file on Google Drive.
2. Add the **Overall row** (Note Type = `Overall`, Hole = `0`) — date, course, tee colour, gross score, playing handicap, Course Rating, Slope, WHS Index, ball striking ratings, notes (tee time, weather, conditions, impressions).
3. Add a **Hole row** per hole (Note Type = `Hole`) — par/distance/stroke index from Courses tab, FIR, GIR, score, strokes, putts, penalties, tee club, pick up, sentiment, notes.
4. OneDrive AutoSave handles sync — no manual save needed.

---

## Workflow — changing the schema (adding/removing columns)

> **Important:** The Drive file must have the new template copied to it *before* reloading data. If data is loaded into the old schema, values land in the wrong columns.

1. Update `create_workbook.py` — add column to `ROUND_HEADERS`, `ROUND_COL_WIDTHS`, sample rows, data validations, and Guide/How To tab content.
2. Run `create_workbook.py` — saves new template to repo only, does **not** touch the Drive file.
3. **Copy template to Drive** before reloading any data:
```bash
python -c "import shutil; shutil.copy2('Golf Round Recap.xlsx', r'G:\My Drive\Project_Outputs\Golf Round Recap\Golf Round Recap.xlsx')"
```
4. Run any data loading scripts — data will now align to the new schema.
5. Update this README to keep the column guide in sync.
6. Commit everything together in one commit.

---

## Setup — regenerating the template from scratch

Saves to repo only. Does **not** overwrite the live Drive file.

```bash
cd "C:\Users\craig.f\Home_Projects\Golf Round Recap"
python create_workbook.py
```

To also reset the Drive file (e.g. new season — **will wipe data**):
```bash
python -c "import shutil; shutil.copy2('Golf Round Recap.xlsx', r'G:\My Drive\Project_Outputs\Golf Round Recap\Golf Round Recap.xlsx')"
```

---

## Status & Next Steps

### What's been built
- `Golf Round Recap.xlsx` — Rounds (29 cols), Courses (9 cols), Guide tab
- Rounds tab has dropdowns for Note Type, Score, Pick Up, Sentiment, Tee Colour, and all 7 ball striking ratings
- Courses tab includes Tee Colour column — par and distance vary per tee
- Pupuke Golf Club pre-loaded with verified White tee data (par 70, 5188m, correct stroke indexes)
- Sample round in the Rounds tab showing correct structure for Overall + Hole rows
- `create_workbook.py` — regenerates structure, saves to repo and copies to Google Drive

### Ready to use
- [ ] **Delete the sample round** — rows 2–5 in the Rounds tab are example data, clear them before entering real rounds
- [ ] **Add other tee colours for Pupuke** — add Yellow/Red rows if you play those tees
- [ ] **Add other courses** — 18 rows per tee colour per course

### Ideas for later
- **Analysis script** — Python or Excel pivot: stroke average by hole, scoring by club, putts trend, sentiment vs. score, ball striking trends over time
- **Handicap tracking** — net score column or separate tab for handicap differential history
- **Stroke play / Stableford scoring** — add a Stableford points column for alternative scoring view
- **Course notes tab** — general notes per course (layout tips, local rules, favourite holes)
