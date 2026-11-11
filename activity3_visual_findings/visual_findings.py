#!/usr/bin/env python3
"""Activity 3: create two visual findings from the cleaned survey data.

Outputs (all in this folder)
  figures/finding1_training_satisfaction.png   diverging stacked bar (Likert)
  figures/finding2_tool_share.png              horizontal bar, sorted
  visual_findings.xlsx                         Clean Data sheet (ready for a
                                               PivotTable) + summary tables and
                                               native Excel charts for both findings
  logs/visual_findings.log                     every number used in findings.md

Usage (from the repository root)
  python3 activity3_visual_findings/visual_findings.py
"""

from __future__ import annotations

import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import openpyxl  # noqa: E402
from openpyxl.chart import BarChart, Reference  # noqa: E402
from openpyxl.chart.label import DataLabelList  # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402
from openpyxl.worksheet.table import Table, TableStyleInfo  # noqa: E402

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from common.cleaning import CLEAN_COLUMNS, clean_respondents  # noqa: E402
from common.survey_profile import RunLog, profile_raw  # noqa: E402

HERE = Path(__file__).resolve().parent
WORKBOOK = REPO_ROOT / "data" / "AI_Workplace_Survey (Activity Datasheet) (1).xlsx"
FIG_DIR = HERE / "figures"
XLSX_OUT = HERE / "visual_findings.xlsx"
LOG = HERE / "logs" / "visual_findings.log"

# ---- palette (dataviz reference palette, light mode) ----------------------
SURFACE = "#fcfcfb"
TEXT = "#0b0b0b"
TEXT_2 = "#52514e"
GRID = "#e4e3df"
BLUE = "#2a78d6"            # single-series hue (slot 1)
# Diverging Likert ramp: red arm <- neutral gray -> blue arm, equal steps per arm
LIKERT = {1: "#c9302c", 2: "#f0a3a2", 3: "#d9d8d3", 4: "#86b6ef", 5: "#2a78d6"}
LIKERT_LABEL = {1: "1 Very dissatisfied", 2: "2", 3: "3 Neutral", 4: "4", 5: "5 Very satisfied"}
LABEL_ON = {1: "white", 2: TEXT, 3: TEXT, 4: TEXT, 5: "white"}

plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 10,
    "axes.edgecolor": GRID, "axes.labelcolor": TEXT_2, "xtick.color": TEXT_2,
    "ytick.color": TEXT, "figure.facecolor": SURFACE, "axes.facecolor": SURFACE,
    "savefig.facecolor": SURFACE,
})

log = RunLog()


def fisher_two_sided(a, b, c, d) -> float:
    """Two-sided Fisher exact test for the 2x2 table [[a, b], [c, d]]."""
    from math import comb
    r1, c1, n = a + b, a + c, a + b + c + d
    prob = lambda x: comb(c1, x) * comb(n - c1, r1 - x) / comb(n, r1)  # noqa: E731
    p_obs = prob(a)
    lo, hi = max(0, r1 + c1 - n), min(r1, c1)
    return min(1.0, sum(prob(x) for x in range(lo, hi + 1) if prob(x) <= p_obs * (1 + 1e-9)))


def pearson(xy) -> float:
    xs, ys = zip(*xy)
    mx, my = sum(xs) / len(xs), sum(ys) / len(ys)
    sxy = sum((a - mx) * (b - my) for a, b in xy)
    sxx = sum((a - mx) ** 2 for a in xs)
    syy = sum((b - my) ** 2 for b in ys)
    return sxy / (sxx * syy) ** 0.5


def group_means(pool, col, val) -> str:
    groups = {}
    for r in pool:
        if r[val] is not None:
            groups.setdefault(r[col] or "(blank)", []).append(r[val])
    return "; ".join(f"{k} {sum(v) / len(v):.2f} (n={len(v)})"
                     for k, v in sorted(groups.items(), key=lambda kv: -sum(kv[1]) / len(kv[1])))


def pct(a, b):
    return 100 * a / b if b else 0.0


# --------------------------------------------------------------------------
# Finding 1: satisfaction by training (AI users only)
# --------------------------------------------------------------------------
def finding1_data(rows):
    users = [r for r in rows if r["uses_ai"] == "Yes"]
    groups = {}
    for label, key in (("Completed AI training", "Yes"), ("No AI training", "No")):
        g = [r for r in users if r["training_completed"] == key]
        scored = [r["satisfaction"] for r in g if r["satisfaction"] is not None]
        groups[label] = {
            "n_group": len(g),
            "n": len(scored),
            "counts": Counter(scored),
            "satisfied": sum(1 for s in scored if s >= 4),
            "dissatisfied": sum(1 for s in scored if s <= 2),
            "mean": sum(scored) / len(scored),
        }
    excluded = {
        "non_users": len(rows) - len(users),
        "unknown_training": sum(1 for r in users if r["training_completed"] == "Unknown"),
        "missing_sat": sum(1 for r in users if r["training_completed"] in ("Yes", "No")
                           and r["satisfaction"] is None),
    }
    return groups, excluded, len(users)


def chart_finding1(groups, path):
    labels = list(groups)[::-1]          # top row = trained
    fig, ax = plt.subplots(figsize=(9, 3.6), dpi=200)
    for y, label in enumerate(labels):
        g = groups[label]
        # centre the neutral (3) on zero: negatives = 1-2 and half of 3
        shares = {k: pct(g["counts"].get(k, 0), g["n"]) for k in LIKERT}
        left = -(shares[1] + shares[2] + shares[3] / 2)
        for k in LIKERT:
            w = shares[k]
            if w == 0:
                continue
            ax.barh(y, w, left=left, height=0.55, color=LIKERT[k],
                    edgecolor=SURFACE, linewidth=2)
            if w >= 7:
                ax.text(left + w / 2, y, f"{w:.0f}%", ha="center", va="center",
                        fontsize=9, color=LABEL_ON[k])
            left += w
        ax.text(80, y, f"{pct(g['satisfied'], g['n']):.0f}% satisfied\n(n = {g['n']})",
                ha="left", va="center", fontsize=9.5, color=TEXT, fontweight="bold")
    ax.axvline(0, color=TEXT_2, linewidth=1)
    ax.set_yticks(range(len(labels)), labels)
    ax.set_xlim(-60, 100)
    ax.set_xticks([-60, -40, -20, 0, 20, 40, 60],
                  ["60%", "40%", "20%", "0", "20%", "40%", "60%"])
    ax.set_xlabel("← dissatisfied side          share of group          satisfied side →",
                  fontsize=8.5)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.tick_params(axis="y", length=0)
    ax.grid(axis="x", color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)
    t, n = groups["Completed AI training"], groups["No AI training"]
    fig.suptitle(
        f"AI users who completed training were more often satisfied with AI tools "
        f"({pct(t['satisfied'], t['n']):.0f}% vs {pct(n['satisfied'], n['n']):.0f}%)",
        x=0.02, ha="left", fontsize=12, fontweight="bold", color=TEXT)
    fig.text(0.02, 0.86, "Satisfaction rating (1–5) of employees who use an AI tool, by training "
             "status. Bars are centred on the neutral rating.", fontsize=9, color=TEXT_2)
    handles = [plt.Rectangle((0, 0), 1, 1, color=LIKERT[k]) for k in LIKERT]
    fig.legend(handles, [LIKERT_LABEL[k] for k in LIKERT], loc="lower center", ncol=5,
               frameon=False, fontsize=8.5, bbox_to_anchor=(0.5, 0.02))
    fig.text(0.02, -0.03, "Source: AI Workplace Survey (invented teaching data), 60 unique respondents, "
             "Oct 2026. Excludes non-users, 'unknown' training and missing ratings.",
             fontsize=7.5, color=TEXT_2)
    fig.subplots_adjust(left=0.2, right=0.98, top=0.78, bottom=0.27)
    fig.savefig(path, bbox_inches="tight")
    plt.close(fig)


# --------------------------------------------------------------------------
# Finding 2: main AI tool share
# --------------------------------------------------------------------------
def finding2_data(rows):
    counts = Counter(r["ai_tool"] for r in rows)
    order = [k for k, _ in counts.most_common() if k != "None"] + ["None"]
    return [(("No AI tool" if k == "None" else k), counts[k]) for k in order], len(rows)


def chart_finding2(tools, n, path):
    labels = [t for t, _ in tools][::-1]
    values = [c for _, c in tools][::-1]
    colors = [("#a9a8a2" if lab == "No AI tool" else BLUE) for lab in labels]
    fig, ax = plt.subplots(figsize=(8, 3.8), dpi=200)
    bars = ax.barh(labels, values, color=colors, height=0.6, edgecolor=SURFACE, linewidth=2)
    for b, v in zip(bars, values):
        ax.text(b.get_width() + 0.4, b.get_y() + b.get_height() / 2,
                f"{v}  ({pct(v, n):.0f}%)", va="center", fontsize=9.5, color=TEXT)
    ax.set_xlim(0, max(values) * 1.25)
    ax.set_xlabel("Number of employees (main tool reported)")
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.tick_params(axis="y", length=0)
    ax.grid(axis="x", color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)
    top_name, top_n = tools[0]
    none_n = dict(tools)["No AI tool"]
    fig.suptitle(f"{top_name} is the main AI tool for {pct(top_n, n):.0f}% of employees; "
                 f"{pct(none_n, n):.0f}% use none",
                 x=0.02, ha="left", fontsize=12, fontweight="bold", color=TEXT)
    fig.text(0.02, 0.87, f"Main AI tool reported by each of {n} unique respondents "
             "(spellings such as 'chat gpt' merged).", fontsize=9, color=TEXT_2)
    fig.text(0.02, -0.04, "Source: AI Workplace Survey (invented teaching data), Oct 2026. "
             "One tool per person; people using several tools named only one.",
             fontsize=7.5, color=TEXT_2)
    fig.subplots_adjust(left=0.17, right=0.97, top=0.8, bottom=0.17)
    fig.savefig(path, bbox_inches="tight")
    plt.close(fig)


# --------------------------------------------------------------------------
# Excel workbook (clean data for PivotTables + native charts)
# --------------------------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="1F3A5F")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def write_header(ws, row, values):
    for i, v in enumerate(values, 1):
        c = ws.cell(row=row, column=i, value=v)
        c.fill, c.font = HEADER_FILL, HEADER_FONT


def build_workbook(rows, groups, tools, n):
    wb = openpyxl.Workbook()

    # -- About
    ws = wb.active
    ws.title = "About"
    lines = [
        ("Activity 3 – Visual findings workbook", Font(bold=True, size=14)),
        ("Generated by activity3_visual_findings/visual_findings.py. Do not edit by hand; re-run the script.", None),
        ("", None),
        ("Clean Data", Font(bold=True)),
        (f"{n} unique respondents after the Activity 1-2 cleaning decisions (duplicates removed, labels "
         "standardised, impossible values set to blank, nothing imputed). Names, emails and postal codes "
         "are removed. Blank cell = missing. Use this sheet as the source for your own PivotTable "
         "(see excel_pivot_tutorial.md).", None),
        ("", None),
        ("Finding 1", Font(bold=True)),
        ("Pre-built summary table and Excel chart: satisfaction by training status, AI users only.", None),
        ("", None),
        ("Finding 2", Font(bold=True)),
        ("Pre-built summary table and Excel chart: main AI tool share.", None),
    ]
    for i, (text, font) in enumerate(lines, 1):
        c = ws.cell(row=i, column=1, value=text)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if font:
            c.font = font
    ws.column_dimensions["A"].width = 110

    # -- Clean Data (as an Excel Table so PivotTables pick up the range)
    ws = wb.create_sheet("Clean Data")
    write_header(ws, 1, CLEAN_COLUMNS)
    for r_i, r in enumerate(rows, 2):
        for c_i, col in enumerate(CLEAN_COLUMNS, 1):
            cell = ws.cell(row=r_i, column=c_i, value=r[col])
            if col == "response_date" and r[col]:
                cell.number_format = "yyyy-mm-dd"
    last = f"{get_column_letter(len(CLEAN_COLUMNS))}{len(rows) + 1}"
    table = Table(displayName="CleanData", ref=f"A1:{last}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)
    for i, col in enumerate(CLEAN_COLUMNS, 1):
        ws.column_dimensions[get_column_letter(i)].width = max(12, len(col) + 3)
    ws.freeze_panes = "A2"

    # -- Finding 1: % of each rating, per training group
    ws = wb.create_sheet("Finding 1")
    ws["A1"] = ("AI users who completed training were more often satisfied with AI tools "
                f"({pct(groups['Completed AI training']['satisfied'], groups['Completed AI training']['n']):.0f}% "
                f"vs {pct(groups['No AI training']['satisfied'], groups['No AI training']['n']):.0f}%)")
    ws["A1"].font = Font(bold=True, size=12)
    ws["A2"] = "Share of each satisfaction rating (row %), AI users with a valid rating."
    write_header(ws, 4, ["Training", "1 Very dissatisfied", "2", "3 Neutral", "4",
                         "5 Very satisfied", "n", "% satisfied (4-5)", "Mean rating"])
    for i, (label, g) in enumerate(groups.items(), 5):
        ws.cell(row=i, column=1, value=label)
        for k in range(1, 6):
            c = ws.cell(row=i, column=k + 1, value=g["counts"].get(k, 0) / g["n"])
            c.number_format = "0%"
        ws.cell(row=i, column=7, value=g["n"])
        ws.cell(row=i, column=8, value=g["satisfied"] / g["n"]).number_format = "0%"
        ws.cell(row=i, column=9, value=round(g["mean"], 2))
    ws.column_dimensions["A"].width = 24
    for col in "BCDEFGHI":
        ws.column_dimensions[col].width = 16
    ch = BarChart()
    ch.type, ch.grouping, ch.overlap = "bar", "percentStacked", 100
    ch.title = "Satisfaction rating by training status (AI users)"
    ch.y_axis.title = "Share of group"
    ch.y_axis.number_format = "0%"
    data = Reference(ws, min_col=2, max_col=6, min_row=4, max_row=6)
    cats = Reference(ws, min_col=1, min_row=5, max_row=6)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    for s, k in zip(ch.series, LIKERT):
        s.graphicalProperties.solidFill = LIKERT[k].lstrip("#")
        s.graphicalProperties.line.solidFill = "FFFFFF"
    ch.dataLabels = DataLabelList()
    ch.dataLabels.showVal = True
    ch.legend.position = "b"
    ch.height, ch.width = 8, 20
    ws.add_chart(ch, "A9")

    # -- Finding 2: tool counts
    ws = wb.create_sheet("Finding 2")
    top_name, top_n = tools[0]
    none_n = dict(tools)["No AI tool"]
    ws["A1"] = (f"{top_name} is the main AI tool for {pct(top_n, n):.0f}% of employees; "
                f"{pct(none_n, n):.0f}% use none")
    ws["A1"].font = Font(bold=True, size=12)
    write_header(ws, 3, ["Main AI tool", "Employees", "% of employees"])
    for i, (tool, c) in enumerate(tools, 4):
        ws.cell(row=i, column=1, value=tool)
        ws.cell(row=i, column=2, value=c)
        ws.cell(row=i, column=3, value=c / n).number_format = "0%"
    ws.cell(row=4 + len(tools), column=1, value="Total").font = Font(bold=True)
    ws.cell(row=4 + len(tools), column=2, value=f"=SUM(B4:B{3 + len(tools)})").font = Font(bold=True)
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 16
    ch = BarChart()
    ch.type = "bar"
    ch.title = "Main AI tool reported (employees)"
    ch.x_axis.scaling.orientation = "maxMin"     # largest bar on top
    ch.y_axis.scaling.min = 0
    ch.y_axis.title = "Employees"
    ch.add_data(Reference(ws, min_col=2, min_row=3, max_row=3 + len(tools)), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=4, max_row=3 + len(tools)))
    ch.series[0].graphicalProperties.solidFill = BLUE.lstrip("#")
    ch.dataLabels = DataLabelList()
    ch.dataLabels.showVal = True
    ch.legend = None
    ch.height, ch.width = 8, 16
    ws.add_chart(ch, "E3")

    wb.save(XLSX_OUT)


# --------------------------------------------------------------------------
def main() -> int:
    log("visual_findings.py run log")
    log(f"Run at   : {datetime.now():%Y-%m-%d %H:%M:%S}")
    log(f"Source   : {WORKBOOK.relative_to(REPO_ROOT)} (read-only)")

    p = profile_raw(openpyxl.load_workbook(WORKBOOK, read_only=True)["Raw Data"])
    rows = clean_respondents(p)
    log(f"Clean respondents: {len(rows)} (from {p.n_rows} raw rows)")

    log.section("CANDIDATE SCREENING - relationships considered before choosing")
    users_all = [r for r in rows if r["uses_ai"] == "Yes"]
    pairs = [("hours_per_week_ai", "satisfaction", users_all, "AI users"),
             ("age", "hours_per_week_ai", users_all, "AI users"),
             ("age", "satisfaction", rows, "everyone"),
             ("years_experience", "satisfaction", rows, "everyone")]
    for x, y, pool, who in pairs:
        xy = [(r[x], r[y]) for r in pool if r[x] is not None and r[y] is not None]
        log(f"r({x}, {y}) [{who}] = {pearson(xy):+.2f}  (n={len(xy)})")
    for col, val, pool, who in [("ai_tool", "satisfaction", rows, "everyone"),
                                ("department", "satisfaction", rows, "everyone"),
                                ("job_level", "satisfaction", rows, "everyone"),
                                ("ai_tool", "hours_per_week_ai", users_all, "AI users"),
                                ("department", "hours_per_week_ai", users_all, "AI users")]:
        log(f"mean {val} by {col} [{who}]: " + group_means(pool, col, val))

    log.section("FINDING 1 - satisfaction by training status (AI users)")
    groups, excluded, n_users = finding1_data(rows)
    log(f"AI users: {n_users}; excluded: {excluded['non_users']} non-users, "
        f"{excluded['unknown_training']} with 'unknown' training, "
        f"{excluded['missing_sat']} with missing satisfaction")
    for label, g in groups.items():
        dist = {k: g["counts"].get(k, 0) for k in range(1, 6)}
        log(f"{label:<24} n={g['n']} (of {g['n_group']} in group)  ratings {dist}  "
            f"satisfied(4-5) {g['satisfied']} = {pct(g['satisfied'], g['n']):.0f}%  "
            f"dissatisfied(1-2) {g['dissatisfied']} = {pct(g['dissatisfied'], g['n']):.0f}%  "
            f"mean {g['mean']:.2f}")
    t, nt = groups["Completed AI training"], groups["No AI training"]
    diff = pct(t["satisfied"], t["n"]) - pct(nt["satisfied"], nt["n"])
    log(f"Gap in % satisfied: {diff:.0f} percentage points")
    # what a single respondent is worth in each group (fragility of the gap)
    p_val = fisher_two_sided(t["satisfied"], t["n"] - t["satisfied"],
                             nt["satisfied"], nt["n"] - nt["satisfied"])
    log(f"Fisher exact test, satisfied vs not, trained vs not: p = {p_val:.2f} "
        f"({'not ' if p_val >= 0.05 else ''}significant at 0.05)")
    log(f"One person = {100 / t['n']:.1f} pts (trained) / {100 / nt['n']:.1f} pts (not trained)")
    # Clever Hans checks: is training tangled up with other variables?
    users = [r for r in rows if r["uses_ai"] == "Yes" and r["training_completed"] in ("Yes", "No")]
    for col in ("department", "job_level", "ai_tool"):
        tab = Counter((r[col], r["training_completed"]) for r in users)
        cats = sorted({r[col] for r in users}, key=str)
        log(f"Training by {col}: " + "; ".join(
            f"{c} {tab[(c, 'Yes')]}/{tab[(c, 'Yes')] + tab[(c, 'No')]} trained" for c in cats))
    hrs = {k: [r["hours_per_week_ai"] for r in users if r["training_completed"] == k
               and r["hours_per_week_ai"] is not None] for k in ("Yes", "No")}
    log(f"Mean AI hours/week: trained {sum(hrs['Yes']) / len(hrs['Yes']):.1f} "
        f"(n={len(hrs['Yes'])}) vs not trained {sum(hrs['No']) / len(hrs['No']):.1f} (n={len(hrs['No'])})")
    non_users_sat = [r["satisfaction"] for r in rows if r["uses_ai"] == "No" and r["satisfaction"]]
    log(f"Non-users who still gave a satisfaction rating: {len(non_users_sat)} "
        f"(mean {sum(non_users_sat) / len(non_users_sat):.2f}) - what were they rating?")

    log.section("FINDING 2 - main AI tool share")
    tools, n = finding2_data(rows)
    for tool, c in tools:
        log(f"{tool:<12} {c:>3}  {pct(c, n):5.1f}%")
    log(f"Total {n}")

    FIG_DIR.mkdir(exist_ok=True)
    f1 = FIG_DIR / "finding1_training_satisfaction.png"
    f2 = FIG_DIR / "finding2_tool_share.png"
    chart_finding1(groups, f1)
    chart_finding2(tools, n, f2)
    build_workbook(rows, groups, tools, n)

    log.section("OUTPUTS")
    for path in (f1, f2, XLSX_OUT, LOG):
        log(str(path.relative_to(REPO_ROOT)))
    log.save(LOG)
    return 0


if __name__ == "__main__":
    sys.exit(main())
