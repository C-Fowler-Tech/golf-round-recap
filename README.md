# Golf Round Recap

A personal golf database tracking rounds hole-by-hole. Each round is stored as a set of rows in Excel — one **Overall** row capturing round-level context, and one **Hole** row per hole played. Over time this builds a dataset for analysis: what clubs work, where strokes are lost, which courses suit your game.

---

## File

**`Golf Round Recap.xlsx`** — two tabs:

| Tab | Purpose |
|---|---|
| `Rounds` | One row per hole + one Overall row per round |
| `Courses` | Course reference data (par, distance per hole) |

---

## Rounds Tab — Column Guide

| Column | Type | Notes |
|---|---|---|
| **Date** | Date | Date of the round |
| **Course** | Text | Match the name used in the Courses tab |
| **Note Type** | Dropdown | `Overall` or `Hole` |
| **Hole** | Number | Hole number 1–18. Use **0** for the Overall row |
| **Par** | Number | Par for this hole (from Courses tab). For Overall: total course par |
| **Distance (m)** | Number | Hole distance in metres. For Overall: total course distance |
| **Score** | Dropdown | Hole result: `Eagle` / `Birdie` / `Par` / `Bogey` / `Double Bogey` / `Triple Bogey` / `Other`. For Overall: enter gross total score (e.g. `85`) |
| **Strokes** | Number | Actual strokes taken. For Overall: total round strokes |
| **Putts** | Number | Number of putts. For Overall: total round putts |
| **Penalties** | Number | Penalty strokes. For Overall: total round penalties |
| **Tee Club** | Text | Club used off the tee (e.g. `Driver`, `3 Wood`, `4 Iron`) |
| **Pick Up** | Dropdown | `Y` if you picked up / didn't finish the hole. `N` otherwise |
| **Sentiment** | Dropdown | `Positive` / `Neutral` / `Negative` — how you felt the hole went |
| **Driver** | Dropdown | **Overall rows only.** Ball striking rating: `Great` / `Good` / `Average` / `Poor` |
| **Woods** | Dropdown | Overall rows only. Fairway woods striking rating |
| **Hybrids** | Dropdown | Overall rows only. Hybrid striking rating |
| **Long Irons (5-7)** | Dropdown | Overall rows only. Long iron striking rating |
| **Short Irons (8-P)** | Dropdown | Overall rows only. Short iron striking rating |
| **Wedges (GW/SW/LW)** | Dropdown | Overall rows only. Wedge striking rating |
| **Putter** | Dropdown | Overall rows only. Putting feel rating |
| **Notes** | Text | Free text detail (see prompts below) |

---

## Data Entry — Q&A Prompts

### Starting a new round

Before entering hole rows, add the **Overall row first** (Note Type = `Overall`, Hole = `0`):

> **Date?** Date played.
> **Course?** e.g. `Pupuke`
> **Par / Distance?** Copy totals from the Courses tab.
> **Score / Strokes?** Your gross total for the round.
> **Putts / Penalties?** Round totals.
> **Sentiment?** How did the round feel overall?
> **Ball striking (Driver / Woods / Hybrids / Long Irons / Short Irons / Wedges / Putter)?**
> Rate each category: `Great` / `Good` / `Average` / `Poor`. Leave blank if you didn't use that club type.
> **Notes?** Cover:
> - Tee time
> - Weather (sun, wind, rain, temperature)
> - Course condition (fairways, greens, rough)
> - Overall impressions — what worked, what didn't

### For each hole (Note Type = `Hole`)

> **Hole / Par / Distance?** Look up from the Courses tab.
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

## Courses Tab — Column Guide

| Column | Notes |
|---|---|
| **Course** | Course name — must match exactly what you use in the Rounds tab |
| **Hole** | Hole number 1–18 |
| **Par** | Par for the hole |
| **Distance (m)** | Distance in metres |
| **Stroke Index** | Handicap stroke index (1 = hardest, 18 = easiest) |
| **Notes** | Any notes about the hole (shape, hazards, typical miss) |

> **Pupuke distances and stroke index are approximate.** Update from the official club scorecard.

### Adding a new course

Add 18 rows (one per hole) with the course name matching exactly what you'll type in the Rounds tab.

---

## Analysis Ideas (future)

Once you have a few rounds in, useful things to look at:

- **Scoring average by hole** — which holes consistently cost you strokes?
- **Scoring by club off the tee** — driver vs. fairway wood vs. iron
- **Putts per round trend** — is the short game improving?
- **Sentiment vs. score correlation** — do positive holes actually score better?
- **Pick-up rate** — how often are you not finishing holes and on which?
- **Score vs. par by hole type** — par 3s, 4s, 5s

---

## Workflow

1. Play a round.
2. Open `Golf Round Recap.xlsx`.
3. Add the **Overall row** for the round.
4. Add a **Hole row** for each hole played (even picked-up holes — mark `Pick Up = Y`).
5. Save the file (OneDrive AutoSave handles sync).
6. Commit to git after a data entry session to keep a backup history.

```bash
cd "C:\Users\craig.f\Home_Projects\Golf Round Recap"
git add "Golf Round Recap.xlsx"
git commit -m "Add round: Pupuke 22-Feb-2026"
git push
```

---

## Setup

To regenerate the workbook structure from scratch (WARNING: overwrites all data):

```bash
python create_workbook.py
```

---

## Status & Next Steps

### What's been built
- `Golf Round Recap.xlsx` — Rounds tab (21 columns) and Courses tab
- Rounds tab has dropdowns for Note Type, Score, Pick Up, Sentiment, and all 7 ball striking columns
- Courses tab pre-loaded with Pupuke Golf Club (18 holes, approximate data)
- Sample round in the Rounds tab showing correct structure for Overall + Hole rows
- `create_workbook.py` — regenerates the workbook structure from scratch if needed

### Needs doing before first real use
- [ ] **Update Pupuke course data** — distances and stroke index are approximate. Open the Courses tab and update from the official Pupuke scorecard or the club website
- [ ] **Delete the sample round** — rows 2-5 in the Rounds tab are example data, clear them before entering real rounds
- [ ] **Add other courses** — add 18 rows per course to the Courses tab as you play new venues

### Ideas for later
- **Analysis tab or separate script** — once a few rounds are in, a Python script (or Excel pivot) could surface: stroke average by hole, scoring by club off tee, putts per round trend, sentiment vs. score correlation
- **Handicap tracking** — add a net score column or a separate tab for handicap differential history
- **Shot tracking** — could add a fairways hit / greens in regulation column if you want more granular stats
- **Course notes tab** — a tab for general notes per course (layout tips, local rules, favourite holes)
