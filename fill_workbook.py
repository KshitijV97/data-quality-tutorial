#!/usr/bin/env python3
"""Profile the 'Raw Data' sheet and fill the four task sheets of the
AI Workplace Survey activity workbook (GNG/ELG 5301, Data Literacy I).

What it does
  1. Reads 'Raw Data' (read-only: that sheet is never modified).
  2. Runs data-quality checks and collects concrete examples
     (respondent_id + Excel row) for every planted problem it finds.
  3. Writes answers into the cream-coloured input cells of
     'Task 1 - Inspect', 'Task 2 - Dictionary', 'Task 3 - Issues' and
     'Task 4 - Privacy', using the examples it found, so the notes always
     match the data actually in the file.
  4. Checks every dropdown cell against the sheet's own list validation,
     confirms 'Raw Data' is byte-for-byte unchanged, and saves.
  5. Writes a run log (stdout + logs/fill_workbook.log) with the profiling
     results and an old -> new record of every cell it changed.

Usage
  python3 fill_workbook.py                # fill the workbook in place
  python3 fill_workbook.py --dry-run      # profile + log only, save nothing
  python3 fill_workbook.py --output out.xlsx

The original workbook is preserved in git history, so in-place edits are safe.
"""

from __future__ import annotations

import argparse
import hashlib
import math
import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
DEFAULT_WORKBOOK = HERE / "AI_Workplace_Survey (Activity Datasheet) (1).xlsx"
DEFAULT_LOG = HERE / "logs" / "fill_workbook.log"

EXPECTED_COLUMNS = [
    "respondent_id", "full_name", "email", "age", "postal_code", "department",
    "job_level", "years_experience", "ai_tool", "hours_per_week_ai",
    "satisfaction", "training_completed", "response_date",
]
HEADER_ROW = 1          # header row in 'Raw Data'
FIRST_DATA_ROW = 2      # first data row in 'Raw Data'
SENTINELS = {-999, -99, -1, 999}
MISSING_TOKENS = {"n/a", "na", "unknown", "null", "none", "-", ""}
SATISFACTION_RANGE = (1, 5)
AGE_RANGE = (16, 80)
HOURS_MAX = 168         # hours in a week
LINE_HEIGHT_PT = 12.75  # Arial 10


# --------------------------------------------------------------------------
# Logging
# --------------------------------------------------------------------------
class RunLog:
    def __init__(self) -> None:
        self.lines: list[str] = []

    def __call__(self, msg: str = "") -> None:
        print(msg)
        self.lines.append(msg)

    def section(self, title: str) -> None:
        self("")
        self("=" * 78)
        self(title)
        self("=" * 78)

    def save(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("\n".join(self.lines) + "\n", encoding="utf-8")


log = RunLog()


# --------------------------------------------------------------------------
# Profiling
# --------------------------------------------------------------------------
@dataclass
class Record:
    excel_row: int
    values: dict

    @property
    def rid(self):
        return self.values["respondent_id"]

    def ref(self) -> str:
        return f"{self.rid} (row {self.excel_row})"


@dataclass
class Profile:
    n_rows: int = 0
    headers: list = field(default_factory=list)
    extra_columns: list = field(default_factory=list)       # (header, col letter)
    extra_values: list = field(default_factory=list)        # (cell, value, rid)
    missing: dict = field(default_factory=dict)             # col -> [Record]
    exact_dupes: list = field(default_factory=list)         # [[Record, Record..]]
    repeated_ids: dict = field(default_factory=dict)        # id -> [Record]
    near_dupes: list = field(default_factory=list)          # (orig, copy)
    age_out_of_range: list = field(default_factory=list)
    exp_sentinel: list = field(default_factory=list)
    exp_vs_age: list = field(default_factory=list)
    young_managers: list = field(default_factory=list)
    tool_raw: Counter = field(default_factory=Counter)
    tool_variants: dict = field(default_factory=dict)       # canonical -> {raw}
    tool_nonstandard: list = field(default_factory=list)    # (Record, raw)
    hours_text: list = field(default_factory=list)          # (Record, raw)
    hours_na: list = field(default_factory=list)
    hours_impossible: list = field(default_factory=list)
    hours_text_typed: int = 0
    sat_out_of_range: list = field(default_factory=list)
    training_values: Counter = field(default_factory=Counter)
    training_bad: list = field(default_factory=list)
    dept_raw: Counter = field(default_factory=Counter)
    dept_variants: dict = field(default_factory=dict)       # canonical -> {raw}
    date_formats: dict = field(default_factory=dict)        # fmt -> [Record]
    date_ambiguous: list = field(default_factory=list)
    date_min: datetime | None = None
    date_max: datetime | None = None
    postal_bad: list = field(default_factory=list)
    postal_fsas: Counter = field(default_factory=Counter)
    email_bad: list = field(default_factory=list)
    email_name_mismatch: list = field(default_factory=list)
    upper_names: list = field(default_factory=list)
    postal_unique: int = 0
    near_dupe_offset_txt: str = ""
    records: list = field(default_factory=list)
    job_levels: Counter = field(default_factory=Counter)
    unique_people: int = 0


def is_blank(v) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == "")


def refs(records, limit: int = 6) -> str:
    items = [r.ref() for r in records]
    if len(items) > limit:
        return ", ".join(items[:limit]) + f", ... (+{len(items) - limit} more)"
    return ", ".join(items)


def ids(records, limit: int = 8) -> str:
    items = [str(r.rid) for r in records]
    if len(items) > limit:
        return ", ".join(items[:limit]) + f", ... (+{len(items) - limit})"
    return ", ".join(items)


def mean_of(p: "Profile", col: str, exclude=()) -> float:
    """Mean of the numeric values in `col`, optionally excluding some records."""
    skip = {id(r) for r in exclude}
    vals = [r.values[col] for r in p.records
            if id(r) not in skip and isinstance(r.values[col], (int, float))]
    return sum(vals) / len(vals)


def mean_shift(p: "Profile", col: str, bad) -> float:
    """How much the bad records move the column mean."""
    return mean_of(p, col) - mean_of(p, col, bad)


def canonical_tool(v) -> str:
    s = re.sub(r"[\s_-]+", "", str(v)).lower()
    return {"chatgpt": "ChatGPT", "copilot": "Copilot", "claude": "Claude",
            "gemini": "Gemini", "none": "None"}.get(s, str(v).strip())


DEPT_ALIASES = {"eng": "engineering", "eng.": "engineering", "cs": "computer science",
                "comp sci": "computer science", "bus": "business", "edu": "education"}


def canonical_dept(v) -> str:
    s = re.sub(r"\s+", " ", str(v)).strip().lower()
    s = DEPT_ALIASES.get(s, s)
    return s.title()


DATE_PATTERNS = [
    ("YYYY-MM-DD", r"\d{4}-\d{2}-\d{2}", "%Y-%m-%d"),
    ("MM/DD/YYYY", r"\d{1,2}/\d{1,2}/\d{4}", "%m/%d/%Y"),
    ("Mon DD, YYYY", r"[A-Za-z]{3} \d{1,2}, \d{4}", "%b %d, %Y"),
]


def parse_date(v):
    if isinstance(v, datetime):
        return "Excel date", v
    s = str(v).strip()
    for name, rx, fmt in DATE_PATTERNS:
        if re.fullmatch(rx, s):
            try:
                return name, datetime.strptime(s, fmt)
            except ValueError:
                return name + " (invalid)", None
    return "unrecognised", None


def profile_raw(ws) -> Profile:
    p = Profile()
    p.headers = [c.value for c in ws[HEADER_ROW]]
    col_index = {h: i for i, h in enumerate(p.headers)}

    missing_cols = [c for c in EXPECTED_COLUMNS if c not in col_index]
    if missing_cols:
        raise SystemExit(f"Raw Data is missing expected columns: {missing_cols}")

    for i, h in enumerate(p.headers):
        if h not in EXPECTED_COLUMNS:
            p.extra_columns.append((h, openpyxl.utils.get_column_letter(i + 1)))

    records: list[Record] = []
    for r_idx, row in enumerate(ws.iter_rows(min_row=FIRST_DATA_ROW), start=FIRST_DATA_ROW):
        vals = [c.value for c in row]
        if all(is_blank(v) for v in vals):
            continue
        rec = Record(r_idx, {c: vals[col_index[c]] for c in EXPECTED_COLUMNS})
        records.append(rec)
        for h, letter in p.extra_columns:
            v = vals[col_index[h]] if col_index[h] < len(vals) else None
            if not is_blank(v):
                p.extra_values.append((f"{letter}{r_idx}", v, rec.rid))
        # numeric values typed as text in the hours column
        hcell = row[col_index["hours_per_week_ai"]]
        if hcell.number_format == "@" or isinstance(hcell.value, str):
            p.hours_text_typed += 1
    p.n_rows = len(records)

    # ---- missing values -------------------------------------------------
    for c in EXPECTED_COLUMNS:
        p.missing[c] = [r for r in records if is_blank(r.values[c])]

    # ---- duplicates -----------------------------------------------------
    by_full = defaultdict(list)
    by_id = defaultdict(list)
    for r in records:
        by_full[tuple(r.values[c] for c in EXPECTED_COLUMNS)].append(r)
        by_id[r.rid].append(r)
    p.exact_dupes = [g for g in by_full.values() if len(g) > 1]
    p.repeated_ids = {k: g for k, g in by_id.items() if len(g) > 1}

    def norm_key(r):
        return tuple(
            (str(r.values[c]).strip().lower() if isinstance(r.values[c], str) else r.values[c])
            for c in EXPECTED_COLUMNS if c != "respondent_id"
        )

    first_by_norm = {}
    for r in records:
        k = norm_key(r)
        if k in first_by_norm and first_by_norm[k].rid != r.rid:
            p.near_dupes.append((first_by_norm[k], r))
        first_by_norm.setdefault(k, r)
    p.unique_people = len(first_by_norm)
    offsets = {b.rid - a.rid for a, b in p.near_dupes}
    p.near_dupe_offset_txt = (str(offsets.pop()) if len(offsets) == 1
                              else "a varying amount")
    p.records = records
    pc_counts = Counter(r.values["postal_code"] for r in records)
    p.postal_unique = sum(1 for r in records if pc_counts[r.values["postal_code"]] == 1)

    # ---- per-column checks ----------------------------------------------
    for r in records:
        v = r.values
        age, exp = v["age"], v["years_experience"]
        if isinstance(age, (int, float)) and not (AGE_RANGE[0] <= age <= AGE_RANGE[1]):
            p.age_out_of_range.append(r)
        if isinstance(exp, (int, float)):
            if exp in SENTINELS or exp < 0:
                p.exp_sentinel.append(r)
            elif isinstance(age, (int, float)) and AGE_RANGE[0] <= age <= AGE_RANGE[1] and age - exp < 14:
                p.exp_vs_age.append(r)
        if (v["job_level"] == "Manager" and isinstance(age, (int, float)) and age < 26
                and isinstance(exp, (int, float)) and 0 <= exp <= 1):
            p.young_managers.append(r)
        p.job_levels[v["job_level"]] += 1

        tool = v["ai_tool"]
        if not is_blank(tool):
            p.tool_raw[repr(tool)] += 1
            p.tool_variants.setdefault(canonical_tool(tool), set()).add(tool)
            if tool != canonical_tool(tool):
                p.tool_nonstandard.append((r, tool))

        h = v["hours_per_week_ai"]
        if isinstance(h, str):
            if h.strip().lower() in MISSING_TOKENS:
                p.hours_na.append(r)
            else:
                p.hours_text.append((r, h))
        elif isinstance(h, (int, float)) and (h < 0 or h > HOURS_MAX):
            p.hours_impossible.append(r)

        s = v["satisfaction"]
        if isinstance(s, (int, float)) and not (SATISFACTION_RANGE[0] <= s <= SATISFACTION_RANGE[1]):
            p.sat_out_of_range.append(r)

        t = v["training_completed"]
        p.training_values[t] += 1
        if t not in ("Yes", "No"):
            p.training_bad.append(r)

        d = v["department"]
        if not is_blank(d):
            p.dept_raw[d] += 1
            p.dept_variants.setdefault(canonical_dept(d), set()).add(d)

        fmt, parsed = parse_date(v["response_date"])
        p.date_formats.setdefault(fmt, []).append(r)
        if parsed:
            p.date_min = min(p.date_min or parsed, parsed)
            p.date_max = max(p.date_max or parsed, parsed)
            if fmt == "MM/DD/YYYY":
                a, b, _ = str(v["response_date"]).split("/")
                if int(a) <= 12 and int(b) <= 12 and a != b:
                    p.date_ambiguous.append(r)

        pc = v["postal_code"]
        if not re.fullmatch(r"[A-Z]\d[A-Z] \d[A-Z]\d", str(pc or "")):
            p.postal_bad.append(r)
        else:
            p.postal_fsas[pc[:3]] += 1

        em, nm = v["email"], v["full_name"]
        if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[A-Za-z]{2,}", str(em or "")):
            p.email_bad.append(r)
        elif nm and str(em).split("@")[0].lower() != str(nm).lower().replace(" ", "."):
            p.email_name_mismatch.append(r)
        if isinstance(nm, str) and nm.isupper():
            p.upper_names.append(r)
    return p


def log_profile(p: Profile) -> None:
    log.section("PROFILE OF 'Raw Data'")
    log(f"Data rows: {p.n_rows}  |  header columns: {len(p.headers)}  |  "
        f"distinct people after de-duplication: {p.unique_people}")
    log(f"Extra (undocumented) columns: {p.extra_columns or 'none'}")
    for cell, val, rid in p.extra_values:
        log(f"  non-empty extra cell {cell} = {val!r} (respondent {rid})")

    log("\nBlank cells per column:")
    for c in EXPECTED_COLUMNS:
        if p.missing[c]:
            log(f"  {c:<20} {len(p.missing[c]):>3}  -> {ids(p.missing[c], 12)}")
    log("  (columns not listed have no blank cells)")

    log("\nExact duplicate rows (all 13 fields identical):")
    for g in p.exact_dupes:
        log(f"  respondent {g[0].rid}: rows {', '.join(str(r.excel_row) for r in g)}")
    log(f"Repeated respondent_id values: {sorted(p.repeated_ids)}")
    log("Near-duplicates (same person, different ID, upper-cased name/email):")
    for a, b in p.near_dupes:
        log(f"  {b.ref()} duplicates {a.ref()}  [{b.values['full_name']} / {b.values['email']}]")

    log(f"\nAge outside {AGE_RANGE}: {refs(p.age_out_of_range)} "
        f"values={[r.values['age'] for r in p.age_out_of_range]}")
    log(f"years_experience sentinel/negative: {refs(p.exp_sentinel)} "
        f"values={[r.values['years_experience'] for r in p.exp_sentinel]}")
    log(f"years_experience implausible for age: {refs(p.exp_vs_age) or 'none'}")
    log(f"Plausibility flag - Manager aged <26 with 0-1 yrs experience: {refs(p.young_managers) or 'none'}")
    log(f"job_level values: {dict(p.job_levels)}")

    log(f"\nai_tool raw values: {dict(p.tool_raw)}")
    for k, v in p.tool_variants.items():
        if len(v) > 1:
            log(f"  {k!r} written {len(v)} ways: {sorted(map(repr, v))}")
    log(f"\nhours_per_week_ai text with units: {[(r.ref(), raw) for r, raw in p.hours_text]}")
    log(f"hours_per_week_ai 'N/A': {refs(p.hours_na)}")
    log(f"hours_per_week_ai > {HOURS_MAX}: {[(r.ref(), r.values['hours_per_week_ai']) for r in p.hours_impossible]}")
    log(f"hours_per_week_ai cells typed as text: {p.hours_text_typed} of {p.n_rows}")

    log(f"\nsatisfaction outside {SATISFACTION_RANGE}: "
        f"{[(r.ref(), r.values['satisfaction']) for r in p.sat_out_of_range]}")
    log(f"training_completed values: {dict(p.training_values)}")
    log(f"department raw values: {dict(p.dept_raw)}")
    for k, v in p.dept_variants.items():
        if len(v) > 1:
            log(f"  {k!r} written {len(v)} ways: {sorted(map(repr, v))}")

    log("\nresponse_date formats:")
    for fmt, recs in p.date_formats.items():
        log(f"  {fmt:<16} {len(recs):>3}  e.g. {recs[0].values['response_date']!r} "
            f"{'' if fmt == 'YYYY-MM-DD' else '-> ' + refs(recs)}")
    log(f"  date range: {p.date_min:%Y-%m-%d} to {p.date_max:%Y-%m-%d}; "
        f"ambiguous day/month: {refs(p.date_ambiguous) or 'none'}")

    log(f"\nPostal codes failing 'A1A 1A1': {refs(p.postal_bad) or 'none'}  |  "
        f"distinct FSAs: {len(p.postal_fsas)}")
    log(f"Emails failing pattern: {refs(p.email_bad) or 'none'}  |  "
        f"email local part != name: {refs(p.email_name_mismatch) or 'none'}")
    log(f"ALL-CAPS names: {refs(p.upper_names) or 'none'}")


# --------------------------------------------------------------------------
# Content for the task sheets (built from the profile)
# --------------------------------------------------------------------------
def build_task1(p: Profile) -> dict:
    n_exact_extra = sum(len(g) - 1 for g in p.exact_dupes)
    miss = {c: len(p.missing[c]) for c in EXPECTED_COLUMNS if p.missing[c]}
    miss_txt = "; ".join(f"{c} {n}" for c, n in miss.items())
    tool_var = [sorted(map(repr, v)) for k, v in p.tool_variants.items() if len(v) > 1]
    dept_var = [sorted(map(repr, v)) for k, v in p.dept_variants.items() if len(v) > 1]
    non_iso = {f: r for f, r in p.date_formats.items() if f != "YYYY-MM-DD"}
    age_bad = ", ".join(f"{r.values['age']} ({r.rid})" for r in p.age_out_of_range)
    hrs_bad = ", ".join(f"{r.values['hours_per_week_ai']} ({r.rid})" for r in p.hours_impossible)
    sat_bad = ", ".join(f"{r.values['satisfaction']} ({r.rid})" for r in p.sat_out_of_range)
    hrs_txt = ", ".join(f"{raw!r} ({r.rid})" for r, raw in p.hours_text)
    dates = "; ".join(f"{fmt} e.g. {recs[0].values['response_date']!r} ({ids(recs)})"
                      for fmt, recs in non_iso.items())
    redup = ", ".join(f"{b.rid}={a.rid}" for a, b in p.near_dupes)

    return {
        "C4": (
            f"One survey response from one employee about their use of AI tools at work "
            f"(not a student). Each row should be one unique respondent keyed by respondent_id, "
            f"but that is not true yet: of {p.n_rows} rows, {n_exact_extra} are exact repeats and "
            f"{len(p.near_dupes)} are the same person re-entered under a new ID, leaving "
            f"{p.unique_people} distinct people."
        ),
        "C5": (
            "Identifiers: respondent_id (row key), full_name, email (contact/linking only, not analysis).\n"
            "Respondent profile: age, postal_code (home area), department, job_level, years_experience - "
            "used to compare groups.\n"
            "AI use (the variables we would actually study): ai_tool (main tool, 'None' = no tool), "
            "hours_per_week_ai, satisfaction (1-5 rating), training_completed (Yes/No).\n"
            "Admin: response_date (when the survey was submitted).\n"
            f"Plus {len(p.extra_columns)} unlabelled columns "
            f"({', '.join(repr(h) for h, _ in p.extra_columns)}) with no documented purpose."
        ),
        "C6": (
            f"Blank cells: {miss_txt}.\n"
            f"Hidden missing values (look filled but are not): years_experience = -999 "
            f"({ids(p.exp_sentinel)}); hours_per_week_ai = 'N/A' ({ids(p.hours_na)}); "
            f"training_completed = 'unknown' ({ids(p.training_bad)}).\n"
            "Missing documentation: no codebook, so we do not know the satisfaction scale "
            "direction, the period 'hours per week' refers to, how 'None' differs from blank, "
            "or how respondents were sampled."
        ),
        "C7": (
            f"Duplicates: repeated IDs {sorted(p.repeated_ids)}; re-entries {redup} "
            f"(new ID = original + {p.near_dupe_offset_txt}, ALL-CAPS name/email).\n"
            f"Out of range: age {age_bad}; hours {hrs_bad} (a week has {HOURS_MAX} h); "
            f"satisfaction {sat_bad} on a 1-5 scale.\n"
            f"Wrong type: hours as text {hrs_txt}.\n"
            f"Inconsistent labels: ai_tool {'; '.join(', '.join(v) for v in tool_var)}; "
            f"department {'; '.join(', '.join(v) for v in dept_var)}.\n"
            f"Mixed date formats: {dates}; all dates stored as text.\n"
            f"Implausible: Managers aged 21-25 with 0-1 years' experience ({ids(p.young_managers)})."
        ),
        "C8": (
            "1. A codebook: definition and units of each column, satisfaction scale (is 1 or 5 best?), "
            "hours reference period, meaning of 'None', blank, 'N/A', 'unknown' and -999.\n"
            "2. Are the 15xx IDs and the repeated IDs re-submissions or export errors? Which copy is correct?\n"
            "3. Correct values for age 250, 400 hours and satisfaction 9 (typos such as 25 / 40 / ?).\n"
            "4. What the unlabelled columns 'f 35', 'f54', 'f54-2' are, and the value 60 in one of them.\n"
            "5. How and when the survey was collected (sampling, population, date window, mandatory "
            "questions) so we know whom the results represent.\n"
            "6. Consent and ethics terms: may the data be used for research, and do we need names, "
            "emails and postal codes at all?"
        ),
    }


def build_task2(p: Profile) -> dict:
    tools = sorted(k for k in p.tool_variants)
    depts = sorted(k for k in p.dept_variants)
    levels = sorted(k for k in p.job_levels if k)
    return {
        # column: (Meaning, Data type, Expected values, Potential issue)
        "respondent_id": (
            "Survey-assigned ID for each respondent; the key that should identify one row.",
            "Number",
            "Unique positive integer, one per person (observed 1001-1060). Used as a label, never in arithmetic.",
            f"Repeated IDs {sorted(p.repeated_ids)} (exact duplicate rows); out-of-sequence IDs "
            f"{', '.join(str(b.rid) for _, b in p.near_dupes)} are copies of other people "
            f"(original ID + {p.near_dupe_offset_txt}).",
        ),
        "full_name": (
            "Respondent's first and last name.",
            "Text",
            "'First Last' in normal capitalisation; one person per name.",
            f"ALL-CAPS duplicates ({ids(p.upper_names)}). Direct identifier: not needed for analysis.",
        ),
        "email": (
            "Respondent's contact email address.",
            "Text",
            "Valid address, lower case, unique per person (pattern first.last@example.com).",
            f"Upper-case copies ({ids(p.upper_names)}) defeat simple duplicate checks. Direct identifier.",
        ),
        "age": (
            "Respondent's age in completed years at the time of the survey.",
            "Number",
            f"Whole number in a working-age range (about {AGE_RANGE[0]}-{AGE_RANGE[1]}).",
            f"Blank for {ids(p.missing['age'])}; impossible value "
            f"{', '.join(str(r.values['age']) for r in p.age_out_of_range)} ({ids(p.age_out_of_range)}).",
        ),
        "postal_code": (
            "Canadian postal code of the respondent's home or work location (all Ottawa 'K' codes).",
            "Text",
            "Format 'A1A 1A1' (letter-digit-letter space digit-letter-digit).",
            f"All {p.n_rows - len(p.postal_bad)} values are well-formed, but a full postal code "
            "narrows a person to a few streets: a privacy risk rather than a quality one.",
        ),
        "department": (
            "Academic/organisational department the respondent works in.",
            "Category",
            f"One of: {', '.join(depts)}.",
            f"Same department written differently ({'; '.join(', '.join(map(repr, sorted(v))) for v in p.dept_variants.values() if len(v) > 1)}); "
            f"blank for {ids(p.missing['department'])}.",
        ),
        "job_level": (
            "Seniority of the respondent's position.",
            "Category",
            f"Ordered category: {', '.join(l for l in ['Entry', 'Mid', 'Senior', 'Manager'] if l in levels)}.",
            f"Text labels, not numbers. Order of 'Manager' vs 'Senior' is not defined; some Managers "
            f"look implausible (aged <26 with 0-1 years' experience: {ids(p.young_managers)}).",
        ),
        "years_experience": (
            "Number of years of work experience (in the field or at the organisation - not defined).",
            "Number",
            "Whole number from 0 up to about (age - 16).",
            f"-999 used as a missing-value code ({ids(p.exp_sentinel)}); would distort averages if "
            "treated as a number. Definition (career vs current job) is unclear.",
        ),
        "ai_tool": (
            "Main AI tool the respondent uses at work.",
            "Category",
            f"One of: {', '.join(tools)} ('None' = uses no AI tool).",
            f"Same tool spelled several ways ({'; '.join(', '.join(map(repr, sorted(v))) for v in p.tool_variants.values() if len(v) > 1)}). "
            "Only one tool can be recorded per person.",
        ),
        "hours_per_week_ai": (
            "Average hours per week the respondent spends using AI tools.",
            "Number",
            f"Number of hours 0-{HOURS_MAX} (realistically 0-60); 0 when ai_tool = 'None'.",
            f"Stored as text; units typed in ({', '.join(repr(raw) for _, raw in p.hours_text)}); "
            f"'N/A' for {ids(p.hours_na)}; impossible "
            f"{', '.join(str(r.values['hours_per_week_ai']) for r in p.hours_impossible)} ({ids(p.hours_impossible)}).",
        ),
        "satisfaction": (
            "Respondent's satisfaction with AI tools at work (rating scale).",
            "Number",
            f"Whole number {SATISFACTION_RANGE[0]}-{SATISFACTION_RANGE[1]} (ordinal; direction assumed 5 = most satisfied).",
            f"{len(p.missing['satisfaction'])} blanks; out-of-scale value "
            f"{', '.join(str(r.values['satisfaction']) for r in p.sat_out_of_range)} ({ids(p.sat_out_of_range)}). "
            "Scale labels are not documented.",
        ),
        "training_completed": (
            "Whether the respondent has completed AI-tool training.",
            "Category",
            "'Yes' or 'No'.",
            f"'unknown' used instead of a blank ({ids(p.training_bad)}); unclear whether it means "
            "'did not answer' or 'not sure'.",
        ),
        "response_date": (
            "Date the respondent submitted the survey.",
            "Date",
            f"Real date in the survey window (observed {p.date_min:%Y-%m-%d} to {p.date_max:%Y-%m-%d}), "
            "one format: YYYY-MM-DD.",
            "Stored as text in mixed formats ("
            + "; ".join(f"{f} e.g. {r[0].values['response_date']!r}" for f, r in p.date_formats.items()
                        if f != "YYYY-MM-DD")
            + "); MM/DD vs DD/MM can be ambiguous; text dates cannot be sorted or filtered by date.",
        ),
    }


def build_task3(p: Profile) -> list[tuple]:
    non_iso = [r for f, recs in p.date_formats.items() if f != "YYYY-MM-DD" for r in recs]
    exact_ex = "; ".join(f"{g[0].rid} (rows {' & '.join(str(r.excel_row) for r in g)})"
                         for g in p.exact_dupes)
    near_ex = "; ".join(f"{b.ref()} = {a.ref()}" for a, b in p.near_dupes)
    tool_variants = "; ".join(", ".join(map(repr, sorted(v)))
                              for v in p.tool_variants.values() if len(v) > 1)
    dept_variants = "; ".join(", ".join(map(repr, sorted(v)))
                              for v in p.dept_variants.values() if len(v) > 1)
    extra_vals = ", ".join(f"{cell} = {val!r} (respondent {rid})" for cell, val, rid in p.extra_values)
    amb = "none are ambiguous" if not p.date_ambiguous else f"ambiguous: {ids(p.date_ambiguous)}"
    tool_ex = ", ".join(f"{r.rid} {raw!r}" for r, raw in p.tool_nonstandard)

    # (Issue found, Column(s), Example, Why it matters, Proposed action, Justification)
    return [
        (
            "Malformed / mixed date formats, stored as text",
            "response_date",
            refs(non_iso) + " - e.g. " + ", ".join(repr(r.values["response_date"]) for r in non_iso),
            "Text dates in several formats cannot be sorted, filtered or compared; MM/DD vs DD/MM "
            "can be misread, putting responses on the wrong day.",
            "Fix",
            f"Convert every value to a real date in ISO YYYY-MM-DD. The intended date is recoverable: "
            f"all responses fall in {p.date_min:%B %Y} and {amb} (day > 12). Keep the original text "
            "in a separate column for traceability.",
        ),
        (
            "Exact duplicate rows (same respondent_id repeated)",
            "All columns / respondent_id",
            exact_ex,
            "Double-counts those respondents, inflating n and biasing every average and percentage; "
            "breaks respondent_id as a unique key.",
            "Remove",
            "Every field is identical, so the copies add no information. Keep the first occurrence, "
            "delete the rest, and record how many rows were removed.",
        ),
        (
            "Same person re-entered under a new ID with name/email in ALL CAPS",
            "respondent_id, full_name, email",
            near_ex,
            "Hidden duplicates survive 'Remove Duplicates' because the ID and letter case differ, "
            "so the person is counted twice.",
            "Remove",
            f"All 12 other fields match the original exactly once case is ignored, and each new ID is "
            f"the original + {p.near_dupe_offset_txt} - a pattern that suggests an export/re-import "
            f"error. Remove the copies ({ids([b for _, b in p.near_dupes])}), keep the originals, and "
            "confirm the pattern with the creator.",
        ),
        (
            "Impossible or missing age",
            "age",
            f"{refs(p.age_out_of_range)} age = "
            f"{', '.join(str(r.values['age']) for r in p.age_out_of_range)}; blank: {refs(p.missing['age'])}",
            f"An age of 250 is impossible and would pull the mean age up by "
            f"{mean_shift(p, 'age', p.age_out_of_range):.1f} years; blanks reduce the sample for any age analysis.",
            "Ask the creator",
            "We cannot tell whether 250 was meant to be 25, 50 or something else, so guessing would "
            "invent data. Until answered, set it to missing; keep the blanks as missing (do not "
            "fill with an average).",
        ),
        (
            "-999 used as a code for missing experience",
            "years_experience",
            f"{refs(p.exp_sentinel)} = "
            f"{', '.join(str(r.values['years_experience']) for r in p.exp_sentinel)}",
            f"If read as numbers, the -999 codes drag mean experience down by "
            f"{-mean_shift(p, 'years_experience', p.exp_sentinel):.0f} years and break any correlation with age or job level.",
            "Fix",
            "-999 is a common 'no answer' code and cannot be a real number of years. Recode it to "
            "a true blank/NA so software treats it as missing; document the recoding.",
        ),
        (
            "Inconsistent ai_tool labels (case, spacing, spelling)",
            "ai_tool",
            f"Non-standard: {tool_ex}",
            "Counts by tool are split across spellings (" + tool_variants + "), so tool usage "
            "shares are wrong.",
            "Fix",
            "The variants clearly name the same product; standardise to one label each (ChatGPT, "
            "Copilot, Claude, Gemini, None) after trimming spaces. Keep 'None' as a real category "
            "meaning 'uses no AI tool', not as missing.",
        ),
        (
            "Hours stored as text, with units or 'N/A'",
            "hours_per_week_ai",
            ", ".join(f"{r.ref()} {raw!r}" for r, raw in p.hours_text)
            + f"; 'N/A': {refs(p.hours_na)}",
            f"{p.hours_text_typed} of {p.n_rows} cells hold text instead of numbers, so Excel cannot "
            "reliably sum or average the column; '12 hrs' and 'N/A' are silently dropped or cause errors.",
            "Fix",
            "'12 hrs' unambiguously means 12 - strip the unit and store as a number. 'N/A' becomes "
            "a true blank (these respondents do use a tool, so 'N/A' is not 0). Convert the column to "
            "a numeric type.",
        ),
        (
            "Impossible weekly hours value",
            "hours_per_week_ai",
            ", ".join(f"{r.ref()} = {r.values['hours_per_week_ai']}" for r in p.hours_impossible),
            f"A week only has {HOURS_MAX} hours; this one value raises the mean weekly hours of the "
            f"whole sample from {mean_of(p, 'hours_per_week_ai', p.hours_impossible):.1f} to "
            f"{mean_of(p, 'hours_per_week_ai'):.1f}.",
            "Ask the creator",
            "It may be a typo for 40 or 4.00, but we cannot know which. Treat it as missing in the "
            "meantime rather than guess, and ask the creator for the original response.",
        ),
        (
            "Out-of-scale and missing satisfaction scores",
            "satisfaction",
            ", ".join(f"{r.ref()} = {r.values['satisfaction']}" for r in p.sat_out_of_range)
            + f"; {len(p.missing['satisfaction'])} blanks e.g. {refs(p.missing['satisfaction'], 4)}",
            f"9 is outside the {SATISFACTION_RANGE[0]}-{SATISFACTION_RANGE[1]} scale and inflates the "
            f"average; {len(p.missing['satisfaction'])} blanks "
            f"({len(p.missing['satisfaction']) * 100 // p.n_rows}% of rows) may bias results if "
            "unhappy users skipped the question.",
            "Ask the creator",
            "We need the scale definition (is it 1-5? which end is 'satisfied'?) and whether the item "
            "was optional. Until then set 9 to missing, keep blanks as missing and report the "
            "response rate for this question.",
        ),
        (
            "'unknown' used in a Yes/No field",
            "training_completed",
            f"{refs(p.training_bad)} = 'unknown'",
            "A third category in a binary variable distorts the % trained; it is unclear if it means "
            "'not answered' or 'not sure'.",
            "Keep",
            "We cannot convert 'unknown' to Yes or No without inventing data. Keep it as an explicit "
            "Unknown/missing category, exclude it from Yes/No percentages and report how many there are.",
        ),
        (
            "Inconsistent and missing department names",
            "department",
            dept_variants + f"; blank: {refs(p.missing['department'])}",
            "Group comparisons split one department into several (e.g. Education vs EDUCATION) and "
            "drop rows with blanks.",
            "Fix",
            "Variants differ only in case, trailing spaces or abbreviation ('Eng.' = Engineering), "
            "so standardising is safe. Leave the 2 blanks as missing - do not guess a department.",
        ),
        (
            "Unlabelled extra columns outside the documented 13",
            ", ".join(h for h, _ in p.extra_columns),
            extra_vals or "columns empty",
            "Columns with no documented meaning cannot be interpreted and may be leftovers from "
            "another export; a stray value could be mistaken for real data.",
            "Remove",
            f"The instructions describe 13 columns; these {len(p.extra_columns)} have no meaningful "
            f"header and hold only {len(p.extra_values)} value{'' if len(p.extra_values) == 1 else 's'} in total. Drop them from the "
            "analysis copy and ask the creator what they were in case they matter.",
        ),
    ]


def build_task4(p: Profile) -> dict:
    # column: (Type of information, Privacy risk, What to remove or protect)
    return {
        "respondent_id": (
            "Direct identifier",
            "Unique per person and links each row to the name and email in this file (or in the "
            "creator's records), so anyone with the key can re-identify answers.",
            "Replace with a new random study ID before sharing; keep any lookup key separately and "
            "securely (or destroy it).",
        ),
        "full_name": (
            "Direct identifier",
            "Names the person outright - the highest-risk field. (Correction: it is a direct, not a "
            "quasi-, identifier.)",
            "Remove entirely; it is not needed for any analysis.",
        ),
        "email": (
            "Direct identifier",
            "Uniquely identifies and allows contacting the person; also reveals their name.",
            "Remove. If follow-up is needed, keep emails in a separate, access-controlled contact list.",
        ),
        "age": (
            "Quasi-identifier",
            "Exact age combined with department, job level and postal code can single out one employee "
            "(e.g. the only 57-year-old Manager in Engineering).",
            "Generalise into bands (e.g. 20-29, 30-39 ...); suppress very small groups.",
        ),
        "postal_code": (
            "Quasi-identifier",
            f"A full 6-character postal code covers only a few streets; {p.postal_unique} of {p.n_rows} "
            "rows have a code no one else shares, so with age it can point to one household.",
            "Remove, or keep only the first 3 characters (FSA) if location is needed for the research.",
        ),
        "department": (
            "Quasi-identifier",
            "Small departments make people stand out, especially combined with job level and age.",
            "Keep only if needed; merge small departments or suppress groups with < 5 people in outputs.",
        ),
        "job_level": (
            "Quasi-identifier",
            "Few people hold 'Manager' in a given department, so job level + department can identify "
            "someone.",
            "Keep for analysis but only report in combination with other fields when each group has "
            ">= 5 people; consider merging levels.",
        ),
        "years_experience": (
            "Quasi-identifier",
            "Exact tenure (e.g. 35 years) is unusual and, with age, can identify a person to colleagues.",
            "Generalise into bands (0-2, 3-5, 6-10, 11-20, 20+); top-code high values.",
        ),
        "ai_tool": (
            "Low risk",
            "Many people share each tool, so it rarely identifies anyone; could be sensitive only if "
            "an employer bans a tool.",
            "Keep; check that rare tools do not create unique combinations with other fields.",
        ),
        "hours_per_week_ai": (
            "Sensitive attribute",
            "Could be used by an employer to judge productivity or policy compliance; harmful if "
            "linked back to a named employee.",
            "Keep only in de-identified data; report as aggregates or bands, never alongside "
            "identifiers.",
        ),
        "satisfaction": (
            "Sensitive attribute",
            "An opinion about workplace tools; if linked to a person it could affect their standing "
            "with managers.",
            "Keep only in de-identified data; report aggregates for groups of >= 5.",
        ),
        "training_completed": (
            "Sensitive attribute",
            "Shows whether an employee complied with (possibly mandatory) training - could lead to "
            "consequences if identifiable.",
            "Keep in de-identified data; report only group-level completion rates.",
        ),
        "response_date": (
            "Quasi-identifier",
            "Exact submission date can be matched to who was in the office or emailed that day, "
            "narrowing down respondents in small groups.",
            "Generalise to week or month, or drop if timing is not part of the research question.",
        ),
    }


TASK4_SAFEGUARDS = [
    (
        "De-identify before analysis: delete full_name and email, replace respondent_id with a new "
        "random study ID, and store any linking key separately under restricted access.",
        "No direct identifiers remain in the analysis file; the key (if kept) is encrypted and only the "
        "data steward can open it; a search of the file for '@' or names returns nothing.",
    ),
    (
        "Generalise quasi-identifiers: age and years_experience into bands, postal code to the first "
        "3 characters (or removed), response_date to month.",
        "Check k-anonymity: every combination of age band, department, job level and postal area "
        "should contain at least 5 people; merge or suppress categories that do not.",
    ),
    (
        "Share and publish aggregates only - no row-level data - and suppress any table cell based on "
        "fewer than 5 respondents.",
        "Review every table/chart before release for small cells, and make sure two tables cannot be "
        "subtracted to reveal one person.",
    ),
    (
        "Confirm consent and ethics: check the survey's consent wording and get research-ethics (REB) "
        "approval before reusing the data; restrict access to the project team on secure storage.",
        "Respondents agreed to research use; REB approval or exemption is documented; files sit on "
        "institution-approved storage with access logged; a retention/deletion date is set.",
    ),
]


# --------------------------------------------------------------------------
# Writing
# --------------------------------------------------------------------------
changes: list[tuple] = []


def put(ws, coord: str, value) -> None:
    old = ws[coord].value
    ws[coord].value = value
    status = "unchanged" if old == value else ("filled" if is_blank(old) else "replaced")
    changes.append((ws.title, coord, status, old, value))


def list_validation_values(ws, coord: str) -> list[str] | None:
    for dv in ws.data_validations.dataValidation:
        if dv.type == "list" and coord in dv.sqref:
            return [x.strip() for x in dv.formula1.strip('"').split(",")]
    return None


def check_dropdown(ws, coord: str, value: str) -> None:
    allowed = list_validation_values(ws, coord)
    if allowed is None:
        raise SystemExit(f"{ws.title}!{coord}: expected a dropdown list validation, found none")
    if value not in allowed:
        raise SystemExit(f"{ws.title}!{coord}: {value!r} is not one of the allowed values {allowed}")


def autofit_row(ws, row: int, min_height: float) -> None:
    """Grow (never shrink) a row so wrapped text is visible."""
    lines_needed = 1
    for cell in ws[row]:
        if is_blank(cell.value):
            continue
        letter = cell.column_letter
        width = ws.column_dimensions[letter].width or 8.43
        chars_per_line = max(1, int(width * 1.15))
        lines = sum(max(1, math.ceil(len(par) / chars_per_line))
                    for par in str(cell.value).split("\n"))
        lines_needed = max(lines_needed, lines)
    target = round(lines_needed * LINE_HEIGHT_PT + 6, 2)
    current = ws.row_dimensions[row].height or min_height
    if target > current:
        ws.row_dimensions[row].height = target


def find_row(ws, col_letter: str, value, start: int) -> int:
    for r in range(start, ws.max_row + 1):
        if ws[f"{col_letter}{r}"].value == value:
            return r
    raise SystemExit(f"{ws.title}: could not find {value!r} in column {col_letter}")


def fill_task1(ws, p):
    for coord, text in build_task1(p).items():
        put(ws, coord, text)
        autofit_row(ws, int(coord[1:]), 69.75)


def fill_task2(ws, p):
    for col, (meaning, dtype, expected, issue) in build_task2(p).items():
        r = find_row(ws, "A", col, 4)
        check_dropdown(ws, f"C{r}", dtype)
        put(ws, f"B{r}", meaning)
        put(ws, f"C{r}", dtype)
        put(ws, f"D{r}", expected)
        put(ws, f"E{r}", issue)
        autofit_row(ws, r, 42)


def fill_task3(ws, p):
    issues = build_task3(p)
    slots = [r for r in range(4, ws.max_row + 1) if isinstance(ws[f"A{r}"].value, int)]
    if len(issues) > len(slots):
        raise SystemExit(f"{len(issues)} issues but only {len(slots)} rows in Task 3")
    for r, (issue, cols, example, why, action, just) in zip(slots, issues):
        check_dropdown(ws, f"F{r}", action)
        for letter, val in zip("BCDEFG", (issue, cols, example, why, action, just)):
            put(ws, f"{letter}{r}", val)
        autofit_row(ws, r, 42)


def fill_task4(ws, p):
    for col, (kind, risk, protect) in build_task4(p).items():
        r = find_row(ws, "A", col, 4)
        check_dropdown(ws, f"B{r}", kind)
        put(ws, f"B{r}", kind)
        put(ws, f"C{r}", risk)
        put(ws, f"D{r}", protect)
        autofit_row(ws, r, 30)
    hdr = find_row(ws, "A", "#", 19)
    for i, (safeguard, check) in enumerate(TASK4_SAFEGUARDS, start=1):
        r = find_row(ws, "A", i, hdr + 1)
        put(ws, f"B{r}", safeguard)
        put(ws, f"C{r}", check)
        autofit_row(ws, r, 33.75)


def sheet_fingerprint(ws) -> str:
    h = hashlib.sha256()
    for row in ws.iter_rows():
        for c in row:
            h.update(f"{c.coordinate}={c.value!r}|{c.number_format};".encode())
    return h.hexdigest()


# --------------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    ap.add_argument("--output", type=Path, help="save here instead of overwriting --workbook")
    ap.add_argument("--log", type=Path, default=DEFAULT_LOG)
    ap.add_argument("--dry-run", action="store_true", help="profile and log, but save nothing")
    args = ap.parse_args()

    out = args.output or args.workbook
    log("fill_workbook.py run log")
    log(f"Run at     : {datetime.now():%Y-%m-%d %H:%M:%S}")
    log(f"Python     : {sys.version.split()[0]}  |  openpyxl {openpyxl.__version__}")
    log(f"Workbook   : {args.workbook.name}")
    log(f"Output     : {'(dry run - not saved)' if args.dry_run else out.name}")

    wb = openpyxl.load_workbook(args.workbook)
    required = ["Raw Data", "Task 1 - Inspect", "Task 2 - Dictionary", "Task 3 - Issues", "Task 4 - Privacy"]
    missing = [s for s in required if s not in wb.sheetnames]
    if missing:
        raise SystemExit(f"Workbook is missing sheets: {missing}")
    log(f"Sheets     : {wb.sheetnames}")

    raw = wb["Raw Data"]
    before = sheet_fingerprint(raw)

    p = profile_raw(raw)
    log_profile(p)

    fill_task1(wb["Task 1 - Inspect"], p)
    fill_task2(wb["Task 2 - Dictionary"], p)
    fill_task3(wb["Task 3 - Issues"], p)
    fill_task4(wb["Task 4 - Privacy"], p)

    after = sheet_fingerprint(raw)
    if before != after:
        raise SystemExit("Safety check failed: 'Raw Data' would be modified. Nothing saved.")

    log.section("CELL CHANGES (old -> new)")
    tally = Counter(c[2] for c in changes)
    for sheet, coord, status, old, new in changes:
        log(f"[{sheet}] {coord}  {status.upper()}")
        if status == "replaced":
            log(f"    old: {old!r}")
        log(f"    new: {new!r}")

    log.section("SUMMARY")
    log(f"Cells written: {len(changes)}  ({dict(tally)})")
    per_sheet = Counter(c[0] for c in changes)
    for s, n in per_sheet.items():
        log(f"  {s:<22} {n} cells")
    log("Dropdown cells validated against each sheet's list validation: OK")
    log(f"'Raw Data' unchanged (sha256 {after[:16]}...): OK")

    if args.dry_run:
        log("Dry run: workbook NOT saved.")
    else:
        wb.save(out)
        log(f"Saved: {out.name}")
    log.save(args.log)
    print(f"\nLog written to {args.log}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
