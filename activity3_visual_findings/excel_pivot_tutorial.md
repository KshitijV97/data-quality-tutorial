# Excel tutorial: build both findings with a PivotTable and PivotChart

This tutorial rebuilds Finding 1 and Finding 2 by hand in Excel, so you can
do the same with any dataset. Steps are given for **Excel for Windows** and
**Excel for Mac** (Microsoft 365). Where the two differ, both are shown.

**File:** `activity3_visual_findings/visual_findings.xlsx`

| Sheet | Use |
|---|---|
| **Clean Data** | Your source: 60 cleaned respondents, stored as an Excel Table named `CleanData`. Blank cell = missing |
| **Finding 1**, **Finding 2** | The answers: pre-built summary tables and charts to check your work against |

> Python can't create PivotTables inside a file, so the script gives you clean
> data and finished answer sheets. Building the PivotTables yourself is the
> exercise.

---

## Part A – PivotTable basics (5 minutes)

A **PivotTable** groups rows and summarises them (count, sum, average) without
formulas. A **PivotChart** is a chart tied to a PivotTable: change the pivot
and the chart updates.

The **PivotTable Fields** pane has four areas:

| Area | Think of it as | Example |
|---|---|---|
| **Rows** | one row per category | `ai_tool` |
| **Columns** | split each row across categories | `satisfaction` |
| **Values** | the number in each cell (Count / Sum / Average) | Count of `respondent_id` |
| **Filters** | keep only some records | `uses_ai` = Yes |

**Golden rules**

1. Start from **one row per record** with a single header row and no merged cells.
   *Clean Data* is already set up like this.
2. To **count people**, put an ID column in Values and set it to **Count**. Excel
   defaults numeric fields to *Sum*, and a sum of IDs is meaningless.
3. Blank cells appear as **(blank)**. Decide on purpose whether to show them or
   filter them out.
4. When the data change, **Refresh** (right-click the pivot → *Refresh*).

---

## Part B – Finding 2: main AI tool share (the simpler one first)

**Target:** *ChatGPT is the main AI tool for 42% of employees; 20% use none.*

### B1. Create the PivotTable

1. Open **Clean Data** and click any cell in the table.
2. **Insert → PivotTable**.
   - Windows: *From Table/Range*. Check the table is `CleanData` and pick **New Worksheet**.
   - Mac: *PivotTable*. Check `CleanData` and pick **New sheet**.
3. Rename the new sheet `My Finding 2` (double-click the tab).

### B2. Lay out the fields

1. Drag **`ai_tool`** to **Rows**.
2. Drag **`respondent_id`** to **Values**. It will say *Sum of respondent_id*.
3. Change it to a count:
   - Windows: click *Sum of respondent_id* in Values → **Value Field Settings** → **Count** → OK.
   - Mac: click the **ⓘ** next to the field → *Summarize by* **Count** → OK.
4. You should see ChatGPT 25, Claude 8, Copilot 11, Gemini 4, None 12, Grand Total 60.

### B3. Add the percentage

1. Drag **`respondent_id`** into **Values** a second time and set it to **Count**.
2. Show it as a share:
   - Windows: right-click a number in that column → **Show Values As → % of Column Total**.
   - Mac: **ⓘ** → *Show data as* **% of total** (or *% of column*) → OK.
3. Rename the headers (click the header cell and type): `Employees`, `% of employees`.

### B4. Sort largest first

Right-click any count under *Employees* → **Sort → Sort Largest to Smallest**.

### B5. Create the PivotChart

1. Click inside the PivotTable.
2. Windows: **PivotTable Analyze → PivotChart**. Mac: **PivotTable Analyze → PivotChart**,
   or **Insert → Chart** with the pivot selected.
3. Choose **Bar → Clustered Bar** (horizontal bars suit category names).
4. The chart plots both value fields. Right-click the *% of employees* bars →
   **Delete** (the % will go in the labels instead), or remove that field from
   Values after the chart is made.

### B6. Make it a finding, not a picture

| Do this | How |
|---|---|
| **Title = the finding** | Click the title and type: *ChatGPT is the main AI tool for 42% of employees; 20% use none* |
| Largest bar on top | Excel draws bar charts bottom-up. Right-click the vertical axis → **Format Axis** → tick **Categories in reverse order** |
| Axis starts at 0 | Format the horizontal axis → *Bounds → Minimum* = **0** |
| Direct labels | **Chart Elements (+) → Data Labels** (Mac: *Chart Design → Add Chart Element → Data Labels*) |
| Remove noise | Delete the legend (one series doesn't need one) and the gridlines if you have labels. Hide field buttons: right-click a grey button on the chart → **Hide All Field Buttons on Chart** |
| No 3D, no gradients | Keep *Clustered Bar*. Format bars with a single **solid fill** |
| Make "None" different | Click the *None* bar once (selects all), click again (selects one) → solid **grey** fill |
| Rename "None" | Type `No AI tool` over the *None* label in the PivotTable. The chart follows |
| Add the context | **Insert → Text Box** under the chart: *Source: AI Workplace Survey (invented), 60 respondents, Oct 2026. One tool per person.* |

Compare with the **Finding 2** sheet.

---

## Part C – Finding 1: satisfaction by training (AI users)

**Target:** *AI users who completed training were more often satisfied with AI
tools (61% vs 37%).*

### C1. Create the PivotTable

Same as B1, in a new sheet `My Finding 1`.

### C2. Lay out the fields

| Field | Area |
|---|---|
| `uses_ai` | **Filters** → open the filter cell above the pivot, choose **Yes** |
| `training_completed` | **Rows** |
| `satisfaction` | **Columns** |
| `respondent_id` | **Values** → set to **Count** |

### C3. Remove the people we decided to exclude

- **Rows filter** (arrow on *Row Labels*): untick **Unknown**. We can't place
  them in either group.
- **Columns filter** (arrow on *Column Labels*): untick **(blank)**. These are the
  missing satisfaction scores from Activity 2. We don't impute them.

Check: Yes row = 2, 1, 4, 5, 6 (total **18**); No row = 5, 3, 4, 5, 2 (total **19**).

### C4. Turn counts into row percentages

- Windows: right-click a count → **Show Values As → % of Row Total**.
- Mac: **ⓘ** on the Values field → *Show data as* **% of row**.

The Yes row now reads 11%, 6%, 22%, 28%, 33%. **Satisfied (4–5) = 28% + 33% =
61%.** The No row gives 11% + 26% = 37%.

> **Why row %?** The groups have different sizes (18 vs 19). Comparing raw
> counts would mix "how many people" with "how satisfied". Row % puts each group
> on its own 100% scale.

### C5. Quick version: 100% stacked bar PivotChart

1. Click the pivot → **PivotChart** → **Bar → 100% Stacked Bar**.
2. Colour the segments as a **diverging scale** (click a segment → *Format Data
   Series → Fill → Solid*):

   | Rating | Colour | Hex |
   |---|---|---|
   | 1 | dark red | `#C9302C` |
   | 2 | light red | `#F0A3A2` |
   | 3 | light grey (neutral) | `#D9D8D3` |
   | 4 | light blue | `#86B6EF` |
   | 5 | blue | `#2A78D6` |

3. Add **Data Labels** and set a white border on the series to separate the segments.
4. Set the title to the finding, hide field buttons, and add the source note
   (same as B6).

This is what the **Finding 1** sheet shows. It works, but the bars all start at
the left edge, so the eye has to compare where the blue *begins*.

### C6. Better version: diverging bar centred on "neutral" (a regular chart)

Excel has no built-in diverging Likert chart. The usual trick is a helper table
where the dissatisfied side is **negative**, so it extends left of zero. Use a
normal chart for this, because PivotCharts don't allow calculated helper columns.

1. Next to your pivot, build this table. Point each cell at the pivot's % cells
   (or use `GETPIVOTDATA`). Values below are for checking:

   | | 3 (left half) | 2 | 1 | 3 (right half) | 4 | 5 |
   |---|---|---|---|---|---|---|
   | Completed AI training | **−11%** | −6% | −11% | 11% | 28% | 33% |
   | No AI training | **−11%** | −16% | −26% | 11% | 26% | 11% |

   - The neutral 3 is **split in half**, half on each side, so the bar centres on it.
   - Order the negative columns **3-left, 2, 1**: Excel stacks negatives outwards
     from zero in column order, so 1 ends up on the far left.
2. Select the table → **Insert → Bar → Stacked Bar** (not 100% stacked).
3. Colour both "3" series the same grey. Delete the duplicate "3" legend entry
   (click it once, then click it again → Delete).
4. Hide the minus signs: format the horizontal axis and the data labels with the
   custom number format **`0%;0%`**.
5. Make the axis symmetric: *Bounds* minimum **−0.6**, maximum **0.8**, so the centre
   line is honest.
6. Title, subtitle (*"Satisfaction rating of AI users, by training. Bars centred on
   neutral."*), n beside each bar (text box or include it in the category label,
   e.g. `Completed AI training (n = 18)`), and the source note.

You now have the chart in `figures/finding1_training_satisfaction.png`.

---

## Part D – Check the pivot with formulas (always verify)

Pivots are easy to get wrong: the wrong filter, *Sum* instead of *Count*, or
(blank) left in. Check one number independently with `COUNTIFS` on the table:

```excel
=COUNTIFS(CleanData[uses_ai],"Yes", CleanData[training_completed],"Yes", CleanData[satisfaction],">=4")
```
→ **11** satisfied trained users.

```excel
=COUNTIFS(CleanData[uses_ai],"Yes", CleanData[training_completed],"Yes", CleanData[satisfaction],">=1")
```
→ **18** trained users with a valid rating. 11 / 18 = **61%** ✓

```excel
=COUNTIFS(CleanData[ai_tool],"ChatGPT") / ROWS(CleanData[ai_tool])
```
→ **42%** ✓

---

## Part E – Useful extras

- **Slicers** (*PivotTable Analyze → Insert Slicer*): clickable filter buttons, e.g.
  a `department` slicer to see Finding 1 for one department. Watch n shrink to
  2–6 people per group, which shows why we didn't split by department.
- **Recommended PivotTables / Charts** (*Insert* tab): a fine starting point, but
  always fix the title, sort, axis and colours afterwards.
- **Refresh after changes:** right-click → *Refresh*, or *Data → Refresh All*.
- **Don't** chart from the raw sheet. Duplicates and `chat gpt` spellings would
  come straight through.

---

## Common mistakes and fixes

| Symptom | Cause | Fix |
|---|---|---|
| Numbers in the thousands (e.g. 61 000) | *Sum* of `respondent_id` | Value Field Settings → **Count** |
| A "(blank)" row or column | Missing values | Untick (blank) in that field's filter, and say so in the note |
| Percentages don't add to 100% across | Showing % of grand total | Use **% of Row Total** |
| Chart lists categories upside-down | Bar charts plot bottom-up | Format Axis → **Categories in reverse order** |
| Grey buttons all over the chart | PivotChart field buttons | Right-click → **Hide All Field Buttons on Chart** |
| Bars look very different but differ by 2 | Axis doesn't start at 0 | Set axis minimum to **0** |
| Chart doesn't change after editing data | Pivot cache is stale | **Refresh** |
