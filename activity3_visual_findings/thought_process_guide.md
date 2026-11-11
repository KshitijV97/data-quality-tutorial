# Thought-process guide: from a messy survey to a visual finding

This guide explains *how* we reached the two findings in
[findings.md](findings.md), and *why* each choice was made. Read it as a
reusable method. The numbers are ours, but the steps work for any dataset.

> **Core idea:** a chart is the *last* step. A finding is a sentence that is
> true, matters to someone and has clear limits. The chart is there to make
> that sentence obvious at a glance.

---

## Step 0 – Start from clean, documented data

We did not open a charting tool first. Activities 1–2 had already shown that
the raw sheet would mislead any chart:

| Raw problem | What a chart would have shown |
|---|---|
| 7 duplicate rows | 4 people counted twice and 3 re-entered under new IDs, inflating every bar |
| `chat gpt`, `chatgpt`, `COPILOT ` | Five or six "tools" instead of four, with ChatGPT under-counted |
| Satisfaction `9`, hours `400` | Means pulled far off (one 400 lifts mean hours from 7.0 to 13.4) |
| `'12 hrs'`, `'N/A'` as text | Excel silently drops these from sums and averages |

`common/cleaning.py` applies those decisions once, so every chart uses the
same 60 respondents.

**Intuition:** *garbage in, confident garbage out.* A chart makes bad data look
believable, so clean first.

---

## Step 1 – Know your variables before asking questions

Using the Activity 1 data dictionary, sort each column by **type**, because the
type decides which comparisons and charts are allowed:

| Type | Columns | Allowed summaries | Typical chart |
|---|---|---|---|
| Nominal category | `department`, `ai_tool`, `training_completed`, `uses_ai` | counts, % | bar (sorted) |
| Ordinal category | `job_level`, **`satisfaction` (1–5)** | counts, % per level, median | stacked / diverging bar |
| Numeric | `age`, `years_experience`, `hours_per_week_ai` | mean, median, spread | histogram, box plot, scatter |
| Date | `response_date` | counts over time | line / column |
| Identifier | `respondent_id` | count only | – |

**The trap we avoided:** `satisfaction` looks numeric, but the gap between 1 and 2
may not be the same as the gap between 4 and 5. A bar of **mean** satisfaction
(3.67 vs 2.79) hides *how* people answered. Showing the share of each rating
keeps that detail.

---

## Step 2 – Brainstorm questions that someone would act on

Ask, *"Who would change a decision because of this?"* Our candidate list:

1. Does using AI more make people happier with it? → *tool budget, rollout pace*
2. Does training relate to satisfaction? → *should the organisation fund training?*
3. Which tools do people actually use? → *licences, support, policy*
4. Do departments or job levels differ in satisfaction? → *targeted support*
5. Do older employees use AI less? → *inclusion, training design*

---

## Step 3 – Screen the candidates with quick numbers

Before drawing anything, run a cheap summary for each question (all in
`logs/visual_findings.log`, *CANDIDATE SCREENING*):

| Candidate | Quick result | Verdict |
|---|---|---|
| Hours ↔ satisfaction | r = −0.03 (n = 35) | No relationship. A scatter would show a cloud. Worth one sentence ("more hours ≠ happier"), not a chart |
| Age ↔ hours | r = −0.13 (n = 40) | Weak and noisy. Skip |
| Satisfaction by department | 3.82 (CS, n = 11) … 2.71 (Health Sciences, n = 7) | Interesting, but 7–14 people per group and one department lost 36% of its answers (Activity 2 bias). Too fragile |
| Satisfaction by tool | 3.14 (ChatGPT) … 3.80 (None!) | Non-users rating AI satisfaction is a data-meaning problem, not a finding |
| **Satisfaction by training (AI users)** | **61% vs 37% satisfied; 42% vs 17% dissatisfied** | Clear, actionable, two groups of similar size → **Finding 1** |
| **Tool share** | ChatGPT 42%, none 20% | Simple, robust, gives the context for Finding 1 → **Finding 2** |

**Intuition:** most relationships in real data are weak. Screening stops you
from polishing a chart of noise. A null result (hours vs satisfaction) still
counts as a finding. Say it in a sentence; it doesn't need a chart.

### Selection criteria we used

1. **Relevance**: answers a decision-maker's question.
2. **Signal**: the difference is big enough to see without squinting.
3. **Enough data per group**: we wanted roughly 15 or more per bar. Department splits failed this.
4. **Survives cleaning**: it doesn't depend on a value we set to missing or guessed.
5. **Explainable limits**: we can state honestly what it does *not* show.

---

## Step 4 – Sketch before you code

We sketched both charts on paper first. This forces you to decide what goes on each
axis and what the reader should notice:

```
Finding 1 (sketch)                          Finding 2 (sketch)
TITLE = the finding, in words               TITLE = the finding, in words

            dissatisfied | satisfied        ChatGPT    ██████████████ 25 (42%)
Trained        ▓▓░░[ 3 ]▒▒▒▒████  61%       Copilot    ██████ 11
Not trained ▓▓▓▓▓░░[ 3 ]▒▒▒██     37%       Claude     ████ 8
                 ← 0 →                       Gemini     ██ 4
red = 1-2, grey = 3, blue = 4-5             No AI tool ▒▒▒▒▒▒ 12 (grey, last)
n under each label                          0 ──────────────────── 30
```

Questions the sketch answered:

- *Where does the reader's eye go?* To the side the bars lean towards. Centring
  on "neutral" turns the comparison into left vs right.
- *What is the one number to remember?* 61% vs 37%, so it goes in the title.
- *What could be misread?* Small n, so it is printed on the chart.

---

## Step 5 – Choose the chart form from the data's job

| Data's job | Right form | Forms we rejected, and why |
|---|---|---|
| Compare the *spread* of an ordered scale (Likert) between two groups | **Diverging stacked bar** centred on neutral | *Mean bar*: hides polarisation. *Pie charts*: angles are hard to compare across two pies. *Side-by-side 5-bar clusters*: 10 bars, no single message |
| Compare counts of a nominal category | **Horizontal bar, sorted, from 0** | *Pie*: 5 slices of similar size are hard to rank. *Vertical bars*: long labels crowd. *3D*: distorts lengths |

**Intuition:** pick the form that makes *the comparison in your sentence* the
easiest visual judgement. People judge length along a common baseline far more
accurately than angle, area or colour intensity.

---

## Step 6 – Apply the professional principles

### Do the cognitive work for the reader

- **The title states the finding**, not the variables:
  ✗ *"Satisfaction vs Training"* → ✓ *"AI users who completed training were more
  often satisfied with AI tools (61% vs 37%)"*.
- A **subtitle** says exactly what is plotted and for whom.
- **Direct labels** (percentages inside segments, counts at bar ends) mean the
  reader never has to estimate from gridlines.
- The **"% satisfied (n = …)" summary** sits beside each bar.

### Maintain integrity

- **Axes start at zero.** Finding 2's count axis starts at 0. Finding 1's axis is
  centred on 0, and both arms use the same scale.
- **No cherry-picking the cut-off.** We report both "satisfied 4–5" *and*
  "dissatisfied 1–2" so neither side is hidden.
- **Past tense** ("were more often satisfied") describes this sample. It doesn't
  claim a law.
- The **denominator** (n) is always visible, including what was excluded and why.

### Minimise noise

- No 3D, no gradients, no shadows, no background colour blocks.
- Light grey gridlines; top, right and left spines removed.
- **Colour means something, or it isn't used.** Finding 2 uses one blue for all
  tools (they are all the same kind of thing) and grey for "No AI tool" (a
  different kind of answer). Finding 1 uses a diverging red–grey–blue scale
  because the scale has two poles and a neutral middle.
- One legend, only where colour encodes something (Finding 1). Finding 2 has none.
- Colours come from a colour-blind-checked palette. Labels carry the values,
  so meaning never depends on colour alone.

---

## Step 7 – Stress-test the finding before you write it

Ask these four questions of every chart:

1. **Could it be chance?** Fisher exact test p = 0.19. With 18 vs 19 people, one
   person moves a group by ~5.5 percentage points. *→ Present it as a pattern, not
   proof.*
2. **Could something else explain it? (Clever Hans)** Trained users also use AI
   ~3 h/week more; Engineering is mostly trained. *→ Name the likely confounders.*
3. **Who is missing?** 12 non-users, 3 "unknown" training and 8 missing ratings are
   excluded, and Activity 2 showed the missing ratings are uneven across
   departments. *→ State the exclusions under the chart.*
4. **How far does it travel? (Over-generalisation)** 60 invented people, one
   organisation, one month. *→ Write the boundary sentence.*

---

## Step 8 – Weave the chart into the text

A chart without context is a picture. In `findings.md` each chart is wrapped
the same way:

1. **Lead-in question**: what relationship we looked at, for whom.
2. **The chart**: with its finding-title.
3. **"What it shows"**: 2–3 bullets that walk the reader's eye across the chart,
   using numbers that are *visible on the chart*.
4. **"What it does not show"**: the interpretation risk and the checks behind it.
5. **Boundary sentence**: who, where, when, how many, and what kind of claim
   (association, not cause).

---

## Step 9 – Acknowledge the humans in the numbers (automation bias)

Numbers feel neutral, but every value here passed through people:

- **Respondents** chose a satisfaction score on an unlabelled scale and guessed
  their weekly hours.
- **The survey designer** allowed only one tool per person and let non-users
  answer the satisfaction question.
- **We** chose the cleaning rules: what counts as a duplicate, what counts as
  impossible, and what to drop.

Writing these down is what makes a finding trustworthy. Because the decisions are
in code and logs, another team can check or change them.

---

## Reusable checklist

- [ ] Data cleaned, cleaning decisions written down
- [ ] Variable types known (nominal / ordinal / numeric / date)
- [ ] 3–5 candidate questions, each tied to a decision
- [ ] Quick numbers for each; weak or fragile ones dropped
- [ ] Paper sketch: axes, comparison, the one number to remember
- [ ] Chart form chosen for the comparison (length on a common baseline)
- [ ] Title = the finding; subtitle = what and who; n visible
- [ ] Axes honest (zero or a symmetric centre); no 3D, no decoration
- [ ] Colour only where it encodes meaning; labels carry the values
- [ ] Chance, confounders, missing people and scope checked
- [ ] Text walks through the chart; boundary sentence written
- [ ] Human choices behind the data acknowledged
