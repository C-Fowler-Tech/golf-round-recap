# Golf Round Recap — Power BI Setup Guide

Connect to the **source file** (not the Analysis file) so Power BI handles the
aggregation and you keep full flexibility as more data comes in.

---

## 1. Connect to the data

1. Open **Power BI Desktop** → **Get Data** → **Excel Workbook**
2. Browse to:
   ```
   G:\My Drive\Project_Outputs\Golf Round Recap\Golf Round Recap.xlsx
   ```
3. Select the **Rounds** table → **Transform Data** (opens Power Query)

---

## 2. Power Query — split into two clean tables

The Rounds tab mixes two grains (Overall rows and Hole rows) in one sheet.
Split them in Power Query so each table has a clean, consistent shape.

**Step 1 — rename the base query**
Right-click the query → Rename → `Rounds_Raw`

**Step 2 — create the Round Summary table**
- Right-click `Rounds_Raw` → **Duplicate**
- Rename duplicate to `Round Summary`
- Add a filter step: `Note Type` = `Overall`
- Remove columns that are blank on Overall rows:
  `Stroke Index`, `FIR`, `GIR`, `Tee Club`, `Pick Up`

**Step 3 — create the Hole Detail table**
- Right-click `Rounds_Raw` → **Duplicate**
- Rename to `Hole Detail`
- Add a filter step: `Note Type` = `Hole`
- Remove columns that are blank on Hole rows:
  `Driver`, `Woods`, `Hybrids`, `Long Irons (5-7)`, `Short Irons (8-P)`,
  `Wedges (GW/SW/LW)`, `Putter`, `Course Rating`, `Slope`, `WHS Index`
- Add a step to join Tee Colour from Round Summary (merge on Date + Course)
  so Hole Detail rows carry the tee colour for slicing

**Step 4 — disable load on Rounds_Raw**
Right-click `Rounds_Raw` → uncheck **Enable Load** (keeps it as a staging query only)

Click **Close & Apply**.

---

## 3. Calculated columns (add in Data view)

### Round Summary

```dax
Sentiment Num =
SWITCH([Sentiment], "Positive", 5, "Neutral", 3, "Negative", 1)
```

```dax
Driver Num =
SWITCH([Driver], "Great", 4, "Good", 3, "Average", 2, "Poor", 1)
```
Repeat for `Woods Num`, `Hybrids Num`, `Long Irons Num`,
`Short Irons Num`, `Wedges Num`, `Putter Num`.

```dax
Front Nine =
CALCULATE(
    SUM('Hole Detail'[Strokes]),
    'Hole Detail'[Date] = [Date],
    'Hole Detail'[Hole] <= 9
)
```

```dax
Back Nine =
CALCULATE(
    SUM('Hole Detail'[Strokes]),
    'Hole Detail'[Date] = [Date],
    'Hole Detail'[Hole] >= 10
)
```

### Hole Detail

```dax
Vs Par = [Strokes] - [Par]
```

```dax
Sentiment Num =
SWITCH([Sentiment], "Positive", 5, "Neutral", 3, "Negative", 1)
```

---

## 4. Key measures (add in Data view)

```dax
Rounds Played = DISTINCTCOUNT('Round Summary'[Date])
```

```dax
Avg Score = AVERAGE('Round Summary'[Strokes])
```

```dax
Avg Handicap Diff = AVERAGE('Round Summary'[Handicap Diff])
```

```dax
Avg Strokes Per Hole = AVERAGE('Hole Detail'[Strokes])
```

```dax
FIR % =
DIVIDE(
    CALCULATE(COUNTROWS('Hole Detail'), 'Hole Detail'[FIR] = "Y"),
    CALCULATE(COUNTROWS('Hole Detail'), 'Hole Detail'[FIR] IN {"Y","N"})
)
```

```dax
GIR % =
DIVIDE(
    CALCULATE(COUNTROWS('Hole Detail'), 'Hole Detail'[GIR] = "Y"),
    CALCULATE(COUNTROWS('Hole Detail'), 'Hole Detail'[GIR] IN {"Y","N"})
)
```

```dax
Avg Putts Per Round =
AVERAGEX(
    VALUES('Round Summary'[Date]),
    CALCULATE(SUM('Hole Detail'[Putts]))
)
```

---

## 5. Suggested report pages

### Page 1 — Round Overview

| Visual | Type | X / Axis | Y / Value | Notes |
|---|---|---|---|---|
| Score per round | Line | Date | Score | Add par as constant line (70) |
| +/- Par per round | Bar (clustered) | Date | +/- Par | Conditional format: red = positive |
| Handicap diff trend | Line | Date | Handicap Diff | Useful once 5+ rounds in |
| Front vs Back nine | Clustered bar | Date | Front Nine, Back Nine | Stacked to total score |
| KPI cards | Card | — | Rounds Played, Avg Score, Avg +/- Par | Top of page |

### Page 2 — Hole Analysis

| Visual | Type | X / Axis | Y / Value | Notes |
|---|---|---|---|---|
| Avg vs Par by hole | Bar | Hole | Avg of Vs Par | Conditional format: green ≤ 0, red > 1 |
| FIR % by hole | Bar | Hole | FIR % | Par 4/5 only (FIR blank on par 3s) |
| GIR % by hole | Bar | Hole | GIR % | |
| Strokes by hole (scatter) | Scatter | Hole | Strokes | Each dot = one round, shows variance |
| Hardest holes table | Table | Hole, Par, Avg Strokes, Avg vs Par, FIR %, GIR % | Sorted by Avg vs Par desc |

### Page 3 — Scoring

| Visual | Type | Values | Notes |
|---|---|---|---|
| Score distribution | Donut | Score label counts | Filter out Pick Up if preferred |
| Score by par type | Clustered bar | Par 3 / Par 4 / Par 5 vs score label | |
| Sentiment distribution | Bar | Positive / Neutral / Negative count | Hole level |
| Penalty holes | Card | Count of holes with penalties | |

### Page 4 — Ball Striking

| Visual | Type | Notes |
|---|---|---|
| Radar / spider chart | Radar | Driver Num, Hybrids Num, Long Irons Num, Short Irons Num, Wedges Num, Putter Num — one shape per round |
| Ball striking trend | Line (multi-series) | Date vs each club category Num — shows improvement over time |
| Avg striking by category | Bar | Average of each Num column across all rounds |

> **Note:** Spider charts require a custom visual from AppSource
> (e.g. **Radar Chart** by Microsoft or **Spider Chart** by akvelon).

---

## 6. Slicers (add to all pages)

| Slicer | Field |
|---|---|
| Course | `Round Summary`[Course] |
| Tee Colour | `Round Summary`[Tee Colour] |
| Date range | `Round Summary`[Date] |

Adding Course and Tee now costs nothing and pays off immediately when
you play other courses.

---

## 7. Refreshing data

After entering a new round into `Golf Round Recap.xlsx`:

1. In Power BI Desktop: **Home → Refresh**
2. All visuals update automatically from the source file on Google Drive

> If Google Drive is mounted as `G:\`, the file path works directly.
> If not (e.g. on a different machine), update the data source path in
> **Transform Data → Data Source Settings**.

---

## 8. Saving the file

Save the `.pbix` to the same output folder as the data:
```
G:\My Drive\Project_Outputs\Golf Round Recap\Golf Round Recap.pbix
```

---

## What gets better with more rounds

| Metric | Useful from |
|---|---|
| FIR % / GIR % per hole | ~5 rounds |
| Handicap diff trend | ~5 rounds |
| Ball striking trend | ~5 rounds |
| Scoring variance by hole | ~10 rounds |
| Course vs course comparison | 2+ courses |
