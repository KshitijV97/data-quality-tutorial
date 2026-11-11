# ELG 5301 – Data Literacy I: AI Workplace Survey

Team activity workbook: `AI_Workplace_Survey (Activity Datasheet) (1).xlsx`.
The `Raw Data` sheet (67 rows, 13 documented columns) has data-quality problems
planted on purpose. The four task sheets record our inspection, data dictionary,
issue log and privacy check.

## Filling the task sheets

```bash
pip install -r requirements.txt
python3 fill_workbook.py            # profile Raw Data and fill Tasks 1-4 in place
python3 fill_workbook.py --dry-run  # profile and log only, save nothing
```

`fill_workbook.py`:

- profiles `Raw Data` (duplicates, missing and disguised-missing values,
  out-of-range numbers, inconsistent labels, mixed date formats, stray columns)
- writes answers into the cream input cells of the four task sheets, quoting
  the respondent IDs and rows it found
- checks every dropdown answer against the sheet's own validation list
- refuses to save if `Raw Data` would change
- writes a run log to `logs/fill_workbook.log`: the profile, plus every cell
  it changed (old value -> new value)

The original workbook, including the team's first notes, is in the git history
(commit `ae3a171`).
