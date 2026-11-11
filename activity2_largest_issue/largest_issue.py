#!/usr/bin/env python3
"""Activity 2: find the single largest data-quality issue in the shared
dataset, apply a cleaning decision, and state the bias it introduces.

Steps
  1. Load 'Raw Data' and remove duplicate respondents (exact copies and the
     re-entries under a new ID), so issues are counted per real person.
  2. Count how many respondents each issue makes unusable for analysis and
     rank the issues.
  3. For the largest issue (satisfaction), apply the cleaning decision:
     complete-case analysis, no imputation.
  4. Measure who gets dropped (by department, job level, tool, training, age)
     and how far the result could move (best/worst-case bounds).
  5. Write final_note.md (the team's note, with every number taken from this
     run) and logs/largest_issue.log.

Usage (from the repository root)
  python3 activity2_largest_issue/largest_issue.py
"""

from __future__ import annotations

import statistics as st
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

import openpyxl

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from common.survey_profile import (  # noqa: E402
    AGE_RANGE, HOURS_MAX, MISSING_TOKENS, SATISFACTION_RANGE, RunLog,
    canonical_dept, canonical_tool, ids, is_blank, profile_raw,
)

HERE = Path(__file__).resolve().parent
WORKBOOK = REPO_ROOT / "data" / "AI_Workplace_Survey (Activity Datasheet) (1).xlsx"
NOTE = HERE / "final_note.md"
LOG = HERE / "logs" / "largest_issue.log"

log = RunLog()


# --------------------------------------------------------------------------
def unique_respondents(p):
    """Drop exact duplicate rows and the near-duplicate re-entries."""
    near_copies = {id(b) for _, b in p.near_dupes}
    seen, keep = set(), []
    for r in p.records:
        key = tuple(r.values.values())
        if key in seen or id(r) in near_copies:
            continue
        seen.add(key)
        keep.append(r)
    return keep


def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def sat_ok(r) -> bool:
    s = num(r.values["satisfaction"])
    return s is not None and SATISFACTION_RANGE[0] <= s <= SATISFACTION_RANGE[1]


def rank_issues(p, people):
    """Return [(issue, column, affected respondents, analysis role)]."""
    v = lambda r, c: r.values[c]  # noqa: E731
    issues = [
        ("Missing or out-of-scale satisfaction score", "satisfaction",
         [r for r in people if not sat_ok(r)], "main outcome"),
        ("Hours not usable as a number (text, 'N/A', > 168)", "hours_per_week_ai",
         [r for r in people if num(v(r, "hours_per_week_ai")) is None
          or not 0 <= v(r, "hours_per_week_ai") <= HOURS_MAX], "main outcome"),
        ("Department missing or written inconsistently", "department",
         [r for r in people if is_blank(v(r, "department"))
          or v(r, "department") != canonical_dept(v(r, "department"))], "grouping"),
        ("Age missing or impossible", "age",
         [r for r in people if num(v(r, "age")) is None
          or not AGE_RANGE[0] <= v(r, "age") <= AGE_RANGE[1]], "grouping"),
        ("Non-ISO / text date", "response_date",
         [r for r in people if not str(v(r, "response_date")).count("-") == 2], "admin"),
        ("ai_tool spelled inconsistently", "ai_tool",
         [r for r in people if v(r, "ai_tool") != canonical_tool(v(r, "ai_tool"))], "grouping"),
        ("training_completed = 'unknown'", "training_completed",
         [r for r in people if str(v(r, "training_completed")).lower() in MISSING_TOKENS], "grouping"),
        ("-999 code in years_experience", "years_experience",
         [r for r in people if (num(v(r, "years_experience")) or 0) < 0], "grouping"),
    ]
    return sorted(issues, key=lambda x: len(x[2]), reverse=True)


def group_rates(people, dropped, col, norm=lambda x: x):
    drop_ids = {id(r) for r in dropped}
    tot, miss = Counter(), Counter()
    for r in people:
        k = norm(r.values[col]) if not is_blank(r.values[col]) else "(blank)"
        tot[k] += 1
        miss[k] += id(r) in drop_ids
    return sorted(((k, miss[k], tot[k]) for k in tot), key=lambda x: -x[1] / x[2])


def pct(a, b) -> str:
    return f"{100 * a / b:.0f}%"


# --------------------------------------------------------------------------
def main() -> int:
    log("largest_issue.py run log")
    log(f"Run at   : {datetime.now():%Y-%m-%d %H:%M:%S}")
    log(f"Workbook : {WORKBOOK.relative_to(REPO_ROOT)} (read-only)")

    wb = openpyxl.load_workbook(WORKBOOK, read_only=True)
    p = profile_raw(wb["Raw Data"])
    people = unique_respondents(p)
    n_raw, n = p.n_rows, len(people)
    n_dupe_rows = n_raw - n

    log.section("STEP 0 - DE-DUPLICATE (prerequisite, not the chosen issue)")
    log(f"Rows in Raw Data: {n_raw}; duplicate rows removed: {n_dupe_rows} "
        f"(exact repeats of {sorted(p.repeated_ids)}; re-entries "
        f"{', '.join(str(b.rid) for _, b in p.near_dupes)}); unique respondents: {n}")

    log.section("STEP 1 - RANK ISSUES BY RESPONDENTS AFFECTED (after de-duplication)")
    ranked = rank_issues(p, people)
    log(f"{'#':<3}{'Issue':<52}{'Column':<20}{'Affected':>9}{'%':>6}  Role")
    log(f"{'-':<3}{'(duplicates)':<52}{'all':<20}{n_dupe_rows:>9}{pct(n_dupe_rows, n_raw):>6}"
        "  rows, not people - removed above")
    for i, (name, col, rows, role) in enumerate(ranked, 1):
        log(f"{i:<3}{name:<52}{col:<20}{len(rows):>9}{pct(len(rows), n):>6}  {role}")
        log(f"{'':<3}  ids: {ids(rows, 12)}")
    top_name, top_col, dropped, _ = ranked[0]
    if top_col != "satisfaction":
        raise SystemExit(f"Largest issue is now {top_col}; the note text assumes satisfaction. "
                         "Review the note before re-running.")

    blanks = [r for r in dropped if is_blank(r.values["satisfaction"])]
    out_of_scale = [r for r in dropped if not is_blank(r.values["satisfaction"])]
    kept = [r for r in people if sat_ok(r)]
    sats = [r.values["satisfaction"] for r in kept]
    k = len(dropped)

    log.section("STEP 2 - CLEANING DECISION: complete-case analysis for satisfaction")
    log(f"Blank satisfaction: {len(blanks)} ({ids(blanks)})")
    oos_txt = ", ".join(f"{r.rid}={r.values['satisfaction']}" for r in out_of_scale)
    log(f"Out of scale      : {len(out_of_scale)} ({oos_txt}) -> set to missing")
    log(f"Kept for satisfaction analysis: {len(kept)} of {n} ({pct(len(kept), n)})")
    mean_kept = st.mean(sats)
    median_kept = st.median(sats)
    log(f"Satisfaction (kept): mean {mean_kept:.2f}, median {median_kept}, "
        f"distribution {dict(sorted(Counter(sats).items()))}")
    raw_vals = [num(r.values["satisfaction"]) for r in people if num(r.values["satisfaction"]) is not None]
    log(f"Mean if the out-of-scale value were left in: {st.mean(raw_vals):.2f}")

    lo = st.mean(sats + [SATISFACTION_RANGE[0]] * k)
    hi = st.mean(sats + [SATISFACTION_RANGE[1]] * k)
    log(f"Bounds if all {k} dropped people had answered 1 / 5: mean {lo:.2f} to {hi:.2f}")

    log.section("STEP 3 - WHO IS DROPPED? (missing-satisfaction rate by group)")
    tables = {
        "department": group_rates(people, dropped, "department", canonical_dept),
        "job_level": group_rates(people, dropped, "job_level"),
        "ai_tool": group_rates(people, dropped, "ai_tool", canonical_tool),
        "training_completed": group_rates(people, dropped, "training_completed"),
    }
    for col, rows in tables.items():
        log(f"{col}: " + "; ".join(f"{g} {m}/{t} ({pct(m, t)})" for g, m, t in rows))
    age = lambda rs: st.mean(a for r in rs if (a := num(r.values["age"])) and a <= AGE_RANGE[1])  # noqa: E731
    hrs = lambda rs: st.mean(h for r in rs if (h := num(r.values["hours_per_week_ai"])) is not None and h <= HOURS_MAX)  # noqa: E731
    age_kept, age_drop = age(kept), age(dropped)
    hrs_kept, hrs_drop = hrs(kept), hrs(dropped)
    log(f"Mean age: kept {age_kept:.1f} vs dropped {age_drop:.1f}")
    log(f"Mean AI hours/week: kept {hrs_kept:.1f} vs dropped {hrs_drop:.1f}")

    dept = [row for row in tables["department"] if row[0] != "(blank)"]
    worst, best = dept[0], dept[-1]
    log(f"Highest drop-out department: {worst[0]} {worst[1]}/{worst[2]}; "
        f"lowest: {best[0]} {best[1]}/{best[2]}")

    # ---------------------------------------------------------------- note
    second = ranked[1]
    rank_rows = "\n".join(
        f"| {i} | {name} | `{col}` | {len(rows)} | {pct(len(rows), n)} | {role} |"
        for i, (name, col, rows, role) in enumerate(ranked, 1)
    )
    group_lines = "\n".join(
        f"| `{col}` | " + "; ".join(f"{g} {m}/{t}" for g, m, t in rows) + " |"
        for col, rows in tables.items()
    )
    sentence1 = (
        f"After removing {n_dupe_rows} duplicate rows, we analysed satisfaction only for the "
        f"respondents with a valid 1-5 score, treating the {len(blanks)} blank answers and the "
        f"{len(out_of_scale)} out-of-range score as missing instead of imputing them, because "
        f"nothing in the data tells us what those {k} people would have answered."
    )
    sentence2 = (
        f"Satisfaction results therefore rest on {len(kept)} of {n} respondents "
        f"({pct(len(kept), n)}) and assume the missing answers are random, which is doubtful "
        f"because {worst[0]} lost {pct(worst[1], worst[2])} of its respondents and {best[0]} "
        f"{'none' if best[1] == 0 else pct(best[1], best[2])}, so the true mean could be anywhere "
        f"from {lo:.1f} to {hi:.1f} rather than the reported {mean_kept:.1f}."
    )

    note = f"""# Activity 2 – The single largest data-quality issue

GNG/ELG 5301 · AI Workplace Survey (`data/`, sheet *Raw Data*)

All numbers below come from `largest_issue.py`; see `logs/largest_issue.log`.

## 1. The largest issue: unusable satisfaction scores

After removing duplicate rows ({n_raw} rows → {n} people), we counted how many
respondents each issue makes unusable for analysis:

| # | Issue | Column | Respondents | % of {n} | Role in analysis |
|---|-------|--------|------------:|---------:|------------------|
{rank_rows}

**{top_name}** is the largest issue: **{k} of {n} respondents ({pct(k, n)})**
have no valid satisfaction score — {len(blanks)} blanks and {len(out_of_scale)}
value of {out_of_scale[0].values['satisfaction']} on a 1–5 scale
(respondent {out_of_scale[0].rid}). It is also the survey's main outcome
variable, so it affects every satisfaction result we would report.

Why not the others?

- **Duplicates** ({n_dupe_rows} of {n_raw} raw rows) distort every count, but
  they are fully fixable: the copies are identical, so removing them loses
  nothing and adds no bias. We did that first.
- **{second[0]}** ({len(second[2])} people) is the runner-up. Most of it can be
  repaired ('12 hrs' → 12); only the 'N/A' and 400-hour values are truly lost.
- The label issues (department, ai_tool, dates) can be standardised without
  losing any information.

Satisfaction is the one problem that is both large and **cannot be repaired
from the data**: nothing in the file tells us what the missing people would
have answered.

## 2. Cleaning decision

1. Remove duplicate rows first ({n_dupe_rows} rows; {n} unique respondents remain).
2. Treat the out-of-scale value ({out_of_scale[0].values['satisfaction']}) as missing — it cannot be a 1–5 answer.
3. **Complete-case analysis:** compute satisfaction results only on the
   {len(kept)} respondents with a valid score. **No imputation** (no filling
   with the mean, median or a guess).
4. Keep the {k} respondents in the dataset for all *other* analyses
   (tool use, hours, training) — they are only excluded where satisfaction is used.
5. Always report the denominator: "n = {len(kept)} of {n}".

Result: mean satisfaction **{mean_kept:.2f}** (median {median_kept}), n = {len(kept)}.

We rejected mean imputation because it would put {k} invented "average" answers
into the data, shrink the spread and make the results look more certain than
they are.

## 3. Bias and limitation this introduces

Complete-case analysis is only unbiased if the missing answers are **missing
completely at random**. The data suggest they may not be:

| Group | Missing satisfaction (dropped / total) |
|-------|----------------------------------------|
{group_lines}

- **{worst[0]}** lost {pct(worst[1], worst[2])} of its respondents and
  **{best[0]}** lost {pct(best[1], best[2])}, so department-level comparisons
  of satisfaction are skewed toward the departments that answered.
- Dropped respondents are slightly older on average ({age_drop:.1f} vs
  {age_kept:.1f} years).
- If people who were unhappy with AI tools were more likely to skip the
  question (a common pattern), the reported mean of {mean_kept:.2f} is **too high**.
  The true mean could lie anywhere from **{lo:.2f}** (all {k} would have
  answered 1) to **{hi:.2f}** (all would have answered 5).
- The sample shrinks from {n} to {len(kept)}, so estimates are less precise, and
  with only {n} people in total the group differences above are too small to
  test formally. We report them as a warning sign, not as proof.

## 4. Final note (two sentences)

> **Cleaning choice:** {sentence1}
>
> **Limitation:** {sentence2}
"""
    NOTE.write_text(note, encoding="utf-8")

    log.section("FINAL NOTE SENTENCES")
    log("1. " + sentence1)
    log("2. " + sentence2)
    log("")
    log(f"Note written: {NOTE.relative_to(REPO_ROOT)}")
    log.save(LOG)
    print(f"Log written: {LOG.relative_to(REPO_ROOT)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
