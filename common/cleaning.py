"""Cleaning decisions from Activities 1-2, applied in one place.

`clean_respondents(profile)` returns one dict per unique respondent with the
fixes from the Activity 1 issue log applied:

- duplicate rows and the +500 re-entries removed (Activity 1 issues 2-3)
- age 250 -> missing; -999 years_experience -> missing (issues 4-5)
- ai_tool and department labels standardised (issues 6, 11)
- '12 hrs' -> 12; 'N/A' and 400 hours -> missing (issues 7-8)
- satisfaction outside 1-5 -> missing, blanks stay missing (issue 9, Activity 2)
- training 'unknown' kept as its own 'Unknown' category (issue 10)
- dates parsed to real dates (issue 1)
- name, email and postal code dropped (Activity 1 privacy check)

Missing values are None. Nothing is imputed.
"""

from __future__ import annotations

import re

from common.survey_profile import (
    AGE_RANGE, HOURS_MAX, SATISFACTION_RANGE, canonical_dept, canonical_tool,
    is_blank, parse_date,
)

CLEAN_COLUMNS = [
    "respondent_id", "department", "job_level", "age", "years_experience",
    "ai_tool", "uses_ai", "hours_per_week_ai", "satisfaction",
    "training_completed", "response_date",
]


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


def _num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _hours(v):
    if isinstance(v, str):
        m = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*(hrs?|hours?)?\s*", v, re.I)
        v = float(m.group(1)) if m else None
        if v is not None and v.is_integer():
            v = int(v)
    v = _num(v)
    return v if v is not None and 0 <= v <= HOURS_MAX else None


def clean_record(r) -> dict:
    v = r.values
    age = _num(v["age"])
    exp = _num(v["years_experience"])
    sat = _num(v["satisfaction"])
    tool = canonical_tool(v["ai_tool"])
    train = v["training_completed"]
    _, date = parse_date(v["response_date"])
    return {
        "respondent_id": v["respondent_id"],
        "department": None if is_blank(v["department"]) else canonical_dept(v["department"]),
        "job_level": v["job_level"],
        "age": age if age is not None and AGE_RANGE[0] <= age <= AGE_RANGE[1] else None,
        "years_experience": exp if exp is not None and exp >= 0 else None,
        "ai_tool": tool,
        "uses_ai": "No" if tool == "None" else "Yes",
        "hours_per_week_ai": _hours(v["hours_per_week_ai"]),
        "satisfaction": sat if sat is not None and SATISFACTION_RANGE[0] <= sat <= SATISFACTION_RANGE[1] else None,
        "training_completed": train if train in ("Yes", "No") else "Unknown",
        "response_date": date.date() if date else None,
    }


def clean_respondents(p) -> list[dict]:
    return [clean_record(r) for r in unique_respondents(p)]
