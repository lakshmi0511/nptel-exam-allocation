# NPTEL exam date allocation

A Claude Code skill and a standalone Python script that give every NPTEL course in a
semester one exam date. It follows the operations team's rules and writes a checked
Excel workbook.

Only the **date** is allocated. Every course runs in both sessions (forenoon and
afternoon) on its date, so a student can write 2 exams on one date.

## Rules

Hard rules (always enforced, and checked in code after solving):
1. **Early dates** (e.g. March / September): existing (Rerun) 4-week and 8-week courses only.
2. **Late dates** (e.g. April / October): new 4-week and 8-week courses, plus all 12-week courses.
   Mixed durations such as `8 Weeks (Rerun) + 4 Weeks (New)` go late by default.
3. **No exam before the course ends:** the exam is at least 7 days after the course's
   End date. SET 2 8-week reruns usually end after the early dates, so they go late.

Goals, solved in this priority order:
1. **Domain clashes:** no more than 2 of a domain's courses on one date. Where that's
   impossible, keep the excess as small as possible and push it onto smaller domains.
2. **Big courses** (over 5,000 previous registrations) spread across dates.
3. **Balance:** previous registrations as even as possible across dates in each group,
   then course counts.

Optional: a cap on registrations per early date (e.g. `--early-cap 150000` for 1.5
lakh). Courses that don't fit move to the late dates, and as few as possible are moved.

## Layout

```
SKILL.md                     Instructions Claude follows when the skill runs
scripts/allocate_exams.py    The optimiser (OR-Tools CP-SAT), checks and Excel output
scripts/sheet_json_to_csv.py Converts a Google Sheets connector dump to CSV
references/data-notes.md     Quirks of past course lists and earlier results
```

## Use it as a Claude Code skill

Clone it into your skills folder:

```bash
git clone <repo-url> ~/.claude/skills/nptel-exam-allocation
```

Then ask Claude Code something like *"allocate exam dates for the Jan 2027 courses,
March 20–21 and April 16–25"*. It reads the sheets, shows you the columns it found,
runs the optimiser and reports the result.

## Run the script directly

Requires Python 3.10+ with `ortools`, `pandas` and `openpyxl`:

```bash
uv venv .venv --python 3.11
uv pip install --python .venv/bin/python ortools pandas openpyxl

.venv/bin/python scripts/allocate_exams.py \
  --courses data/course_list.csv --domains data/domain_courses.csv \
  --early-dates 2027-03-20,2027-03-21 \
  --late-dates 2027-04-16,2027-04-17,2027-04-18,2027-04-23,2027-04-24,2027-04-25 \
  --early-cap 150000 \
  --out exam_allocation.xlsx
```

| Option | Default | Meaning |
|---|---|---|
| `--early-dates`, `--late-dates` | (required) | ISO dates, comma-separated |
| `--early-cap` | 0 (none) | Maximum previous registrations per early date |
| `--big` | 5000 | Registrations above which a course counts as big |
| `--mixed-group` | `late` | Where mixed-duration courses go: `late`, `early` or `none` |
| `--min-gap-days` | 7 | Minimum days between course end and exam |
| `--course-cols` | | JSON map when the course-list headers differ |
| `--time` | 60 | Solver time limit per tier, in seconds |

### Input files

**course_list.csv:** `S No`, `NPTEL ID.1` (9-digit NPTEL ID), `Old Course ID`,
`Course Name`, `Duration`, `Type` (New / Rerun), `Previous Registration count`, and
`End date`. New courses with no registrations count as 0.

**domain_courses.csv:** `Discipline name`, `Domain Name`, `core`, `elective`. Every
9-digit number in `core`/`elective` is a course ID, and a course can be in several
domains.

### Output: `exam_allocation.xlsx`

| Sheet | Contents |
|---|---|
| Allocation | One row per course: IDs, name, duration, type, registrations, domains, end date, exam date |
| Date summary | Per date: courses, new courses, total registrations, big courses |
| Domain check | Courses per domain per date, with clashes (more than 2) highlighted |
| Exceptions | Unplaced courses, courses moved late (by the cap or the end date), big courses sharing a date, and domain clashes |

The script also prints each solver tier's status (OPTIMAL / FEASIBLE) and a PASS/FAIL
line for every hard rule.
