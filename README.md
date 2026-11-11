# ELG 5301 – Data Quality Activities: AI Workplace Survey

Team activities for GNG/ELG 5301 on one shared dataset: an invented workplace
survey on employees' use of AI tools. The `Raw Data` sheet (67 rows, 13
documented columns) has data-quality problems planted on purpose.

## Layout

```
data/
  AI_Workplace_Survey (Activity Datasheet) (1).xlsx   shared dataset + Activity 1 task sheets
common/
  survey_profile.py        loads 'Raw Data' and runs the data-quality checks (used by every activity)
activity1_know_your_dataset/
  fill_workbook.py         fills Task 1-4 sheets in the workbook
  logs/fill_workbook.log   run log: profile + every cell changed (old -> new)
activity2_largest_issue/
  largest_issue.py         ranks issues, applies the cleaning decision, measures the bias
  final_note.md            the team's write-up and the two-sentence final note
  logs/largest_issue.log   run log with every number used in the note
activity3_visual_findings/
  visual_findings.py       cleans the data, draws both charts, builds the Excel workbook
  findings.md              the two visual findings, interpretation risks, boundary sentences
  thought_process_guide.md how and why we chose these findings and chart designs
  excel_pivot_tutorial.md  step-by-step: rebuild both charts with PivotTables / PivotCharts
  figures/                 finding1_training_satisfaction.png, finding2_tool_share.png
  visual_findings.xlsx     Clean Data table (pivot-ready) + both findings as Excel charts
  logs/visual_findings.log run log with every number used
```

`common/cleaning.py` applies the Activity 1–2 cleaning decisions (de-duplication,
label fixes, impossible values to blank, no imputation, identifiers dropped) and
is shared by Activities 2 and 3.

## Setup

```bash
pip install -r requirements.txt
```

Run every script from the repository root.

## Activity 1 – Know your dataset

```bash
python3 activity1_know_your_dataset/fill_workbook.py            # fill Tasks 1-4 in place
python3 activity1_know_your_dataset/fill_workbook.py --dry-run  # profile and log only
```

The script profiles `Raw Data` (duplicates, missing and disguised-missing values,
out-of-range numbers, inconsistent labels, mixed date formats, stray columns),
writes answers into the cream input cells of the four task sheets quoting the
respondent IDs and rows it found, checks every dropdown answer against the
sheet's validation list, and refuses to save if `Raw Data` would change.

The original workbook, including the team's first notes, is in the git history
(commit `ae3a171`).

## Activity 2 – The single largest data-quality issue

```bash
python3 activity2_largest_issue/largest_issue.py
```

Reads `Raw Data` only (it never writes to the workbook). After removing
duplicate respondents, it ranks every issue by how many people it makes
unusable, picks the largest (missing or out-of-scale `satisfaction`, 10 of 60
respondents), applies the cleaning decision (complete-case analysis, no
imputation), and checks who gets dropped and how far the mean could move.
It regenerates [`final_note.md`](activity2_largest_issue/final_note.md) so the
numbers in the note always match the data.

## Activity 3 – Creating meaningful visual findings

```bash
python3 activity3_visual_findings/visual_findings.py
```

Two findings from the cleaned data, each with a descriptive title, woven text,
an interpretation risk and a boundary sentence:

1. *AI users who completed training were more often satisfied with AI tools (61% vs 37%)*,
   an association, not a cause (p = 0.19, n = 18 vs 19).
2. *ChatGPT is the main AI tool for 42% of employees; 20% use none.*

Start with [`findings.md`](activity3_visual_findings/findings.md). The reasoning is in
[`thought_process_guide.md`](activity3_visual_findings/thought_process_guide.md), and the
Excel walkthrough is in [`excel_pivot_tutorial.md`](activity3_visual_findings/excel_pivot_tutorial.md).
