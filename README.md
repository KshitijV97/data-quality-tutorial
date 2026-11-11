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
```

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
