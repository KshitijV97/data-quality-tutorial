"""Shared loading and profiling of the 'Raw Data' sheet (AI Workplace Survey).

Used by every activity script so all of them see the same checks and examples.
"""

from __future__ import annotations

import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import openpyxl

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


def log_profile(p: Profile, log: "RunLog") -> None:
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
