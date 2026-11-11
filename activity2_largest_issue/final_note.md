# Activity 2 – The single largest data-quality issue

GNG/ELG 5301 · AI Workplace Survey (`data/`, sheet *Raw Data*)

All numbers below come from `largest_issue.py`; see `logs/largest_issue.log`.

## 1. The largest issue: unusable satisfaction scores

After removing duplicate rows (67 rows → 60 people), we counted how many
respondents each issue makes unusable for analysis:

| # | Issue | Column | Respondents | % of 60 | Role in analysis |
|---|-------|--------|------------:|---------:|------------------|
| 1 | Missing or out-of-scale satisfaction score | `satisfaction` | 10 | 17% | main outcome |
| 2 | Hours not usable as a number (text, 'N/A', > 168) | `hours_per_week_ai` | 7 | 12% | main outcome |
| 3 | Department missing or written inconsistently | `department` | 6 | 10% | grouping |
| 4 | Age missing or impossible | `age` | 4 | 7% | grouping |
| 5 | Non-ISO / text date | `response_date` | 3 | 5% | admin |
| 6 | ai_tool spelled inconsistently | `ai_tool` | 3 | 5% | grouping |
| 7 | training_completed = 'unknown' | `training_completed` | 3 | 5% | grouping |
| 8 | -999 code in years_experience | `years_experience` | 2 | 3% | grouping |

**Missing or out-of-scale satisfaction score** is the largest issue: **10 of 60 respondents (17%)**
have no valid satisfaction score — 9 blanks and 1
value of 9 on a 1–5 scale
(respondent 1057). It is also the survey's main outcome
variable, so it affects every satisfaction result we would report.

Why not the others?

- **Duplicates** (7 of 67 raw rows) distort every count, but
  they are fully fixable: the copies are identical, so removing them loses
  nothing and adds no bias. We did that first.
- **Hours not usable as a number (text, 'N/A', > 168)** (7 people) is the runner-up. Most of it can be
  repaired ('12 hrs' → 12); only the 'N/A' and 400-hour values are truly lost.
- The label issues (department, ai_tool, dates) can be standardised without
  losing any information.

Satisfaction is the one problem that is both large and **cannot be repaired
from the data**: nothing in the file tells us what the missing people would
have answered.

## 2. Cleaning decision

1. Remove duplicate rows first (7 rows; 60 unique respondents remain).
2. Treat the out-of-scale value (9) as missing — it cannot be a 1–5 answer.
3. **Complete-case analysis:** compute satisfaction results only on the
   50 respondents with a valid score. **No imputation** (no filling
   with the mean, median or a guess).
4. Keep the 10 respondents in the dataset for all *other* analyses
   (tool use, hours, training) — they are only excluded where satisfaction is used.
5. Always report the denominator: "n = 50 of 60".

Result: mean satisfaction **3.40** (median 4.0), n = 50.

We rejected mean imputation because it would put 10 invented "average" answers
into the data, shrink the spread and make the results look more certain than
they are.

## 3. Bias and limitation this introduces

Complete-case analysis is only unbiased if the missing answers are **missing
completely at random**. The data suggest they may not be:

| Group | Missing satisfaction (dropped / total) |
|-------|----------------------------------------|
| `department` | Business 4/11; Education 3/17; Computer Science 2/13; Engineering 1/10; Health Sciences 0/7; (blank) 0/2 |
| `job_level` | Manager 3/16; Senior 3/17; Entry 2/13; Mid 2/14 |
| `ai_tool` | Claude 2/8; Gemini 1/4; Copilot 2/11; None 2/12; ChatGPT 3/25 |
| `training_completed` | Yes 5/28; No 5/29; unknown 0/3 |

- **Business** lost 36% of its respondents and
  **Health Sciences** lost 0%, so department-level comparisons
  of satisfaction are skewed toward the departments that answered.
- Dropped respondents are slightly older on average (43.6 vs
  40.0 years).
- If people who were unhappy with AI tools were more likely to skip the
  question (a common pattern), the reported mean of 3.40 is **too high**.
  The true mean could lie anywhere from **3.00** (all 10 would have
  answered 1) to **3.67** (all would have answered 5).
- The sample shrinks from 60 to 50, so estimates are less precise, and
  with only 60 people in total the group differences above are too small to
  test formally. We report them as a warning sign, not as proof.

## 4. Final note (two sentences)

> **Cleaning choice:** After removing 7 duplicate rows, we analysed satisfaction only for the respondents with a valid 1-5 score, treating the 9 blank answers and the 1 out-of-range score as missing instead of imputing them, because nothing in the data tells us what those 10 people would have answered.
>
> **Limitation:** Satisfaction results therefore rest on 50 of 60 respondents (83%) and assume the missing answers are random, which is doubtful because Business lost 36% of its respondents and Health Sciences none, so the true mean could be anywhere from 3.0 to 3.7 rather than the reported 3.4.
