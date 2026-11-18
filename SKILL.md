---
name: nptel-exam-allocation
description: Allocate NPTEL exam dates to courses from the course-list and domain-mapping Google Sheets, using an OR-Tools optimiser that keeps March/early dates for existing 4/8-week courses and April/late dates for new and 12-week courses, minimises domain clashes, spreads big courses (over 5,000 registrations) and balances registrations per date, with an optional per-date registration cap. Produces exam_allocation.xlsx (Allocation, Date summary, Domain check, Exceptions). Use this whenever the user wants to schedule, allocate, assign, plan or rebalance NPTEL exam dates, rerun the exam allocation for a new semester, cap registrations per exam date ("1.5 lakh per date"), check domain clashes in an exam schedule, or move courses between exam dates — even if they don't say "skill" or name the script.
---

# NPTEL exam date allocation

Assign every NPTEL course in a semester's offering one exam date, following the
operations team's rules, and hand back a checked Excel workbook plus a short summary.

The bundled `scripts/allocate_exams.py` does the optimisation, the hard-constraint
checks and the workbook. Your job is to get clean inputs to it, run it with the right
settings for this cycle, and report the result honestly.

## The rules (as agreed with the user)

Only an **exam date** is allocated, never a session. Every course runs in both
sessions (Forenoon and Afternoon) on its date, so a student can write 2 exams on one
date. Don't add session columns or session balancing; the user explicitly rejected that.

Hard rules:
1. **Early dates** (e.g. March): only existing (Rerun) 4-week and 8-week courses.
2. **Late dates** (e.g. April): new 4-week and 8-week courses, plus all 12-week courses.
3. Mixed durations such as "8 Weeks (Rerun) + 4 Weeks (New)" total 12 weeks. The user
   decided these go to the late dates (`--mixed-group late`, the default). Mention it
   in the summary. Any other duration or type that fits neither group is left unplaced
   and listed; ask the user where it goes.
4. **No exam before the course ends.** The exam must be at least 7 days after the
   course's End date (`--min-gap-days`). That's what NPTEL's actual schedules show, and
   rule 1 alone breaks it. SET 2 courses start about 4 weeks later, so SET 2 8-week
   reruns end after the early dates. The script sends them to the late dates and lists
   them as exceptions. Always include an `End date` column in course_list.csv. Without
   one the script can't check this and says so; warn the user if that happens.

Optimisation goals, in priority order (the script solves them lexicographically):
1. **Domain clashes:** no more than 2 of a domain's courses on one date. When that's
   impossible, minimise the excess, and push it onto domains with fewer registrations.
2. **Big courses apart:** a big course has Previous Registration above 5,000. Ideally
   one per date. With more big courses than dates, spread them as evenly as possible
   and list the dates with more than one.
3. **Balance:** total Previous Registration as even as possible across dates within
   each group, then course counts.

**Optional cap** (`--early-cap`): a ceiling on registrations per early date, e.g. 1.5
lakh = 150000. Courses that don't fit move to late dates. This deliberately breaks hard
rule 1, so only apply a cap when the user asks for one. The script moves as few courses
as possible, then picks which ones by the domain-clash and balance goals. Always list
the moved courses for the user.

## Workflow

### 1. Get the inputs
The user normally gives two Google Sheets: the course list and the domain mapping.
Save them as `course_list.csv` and `domain_courses.csv` in a fresh `data/` folder.

- First try the CSV export URL
  (`https://docs.google.com/spreadsheets/d/<id>/export?format=csv&gid=<gid>`) with
  `curl -sSL`. Then check the result with `file` or `head`. A private sheet returns an
  HTML "Sign in" page, not CSV.
- If it's private and a Google Sheets connector is available, read it through the
  connector instead (read-only). Use `get_spreadsheet` with
  `fields: ["sheets.properties"]` to find the tab whose `sheetId` equals the URL's
  `gid`. Then call `get_values` on that tab's full range, e.g. `'<tab name>'!A1:AK2000`.
  The result is too large to show inline and gets saved to a file. Convert it with
  `python scripts/sheet_json_to_csv.py <saved file> data/course_list.csv`.
- If neither works, stop and ask the user to share the sheet or download it.

### 2. Profile the data and confirm the column mapping
Before allocating, show the user the columns you found and a short summary. They have
asked for this every time. Cover:
- row counts and the Duration × Type crosstab
- how many courses have an NPTEL ID and a registration count (new courses usually have neither)
- the number of big courses and their split between groups
- any mixed or unclear durations
- how many courses map to at least one domain

Expected columns, which the script checks:
- **Course list:** `S No`, `NPTEL ID.1`, `Old Course ID`, `Course Name`, `Duration`,
  `Type`, `Previous Registration count`. The sheet has two `NPTEL ID` columns; the
  second (`NPTEL ID.1` in pandas) holds the 9-digit IDs, and the first is empty.
  Confirm this in the new data. If the headers differ, pass `--course-cols` with a JSON
  map from the sheet header to `sno`/`nptel_id`/`run_id`/`name`/`duration`/`type`/`reg`.
  Tell the user which column you mapped to which.
- **Domain sheet:** `Discipline name`, `Domain Name`, `core`, `elective`. IDs are
  comma-separated, and `_` or `|` join equivalent IDs; every 9-digit number counts.
  Rows with a blank Domain Name are notes and are skipped. A course can belong to
  several domains. Matching is by NPTEL ID only.

Read `references/data-notes.md` for the quirks seen in past data. Course lists arrive in
different layouts. The July 2026 one was an .xlsx with run codes but no registrations or
NPTEL IDs. In that case, build `course_list.csv` in the layout above from the other
sources (see the July 2026 section of the notes), and save the prep as a small script
beside the outputs so it can be rerun.

### 3. Set up Python and run
The script needs `ortools`, `pandas` and `openpyxl`. If they're missing, create a venv,
e.g. `uv venv .venv --python 3.11 && uv pip install --python .venv/bin/python ortools pandas openpyxl`.
Keep the script outside the downloaded-data folder and run it with absolute paths:

```bash
.venv/bin/python <skill-dir>/scripts/allocate_exams.py \
  --courses data/course_list.csv --domains data/domain_courses.csv \
  --early-dates 2027-03-20,2027-03-21 \
  --late-dates 2027-04-16,2027-04-17,2027-04-18,2027-04-23,2027-04-24,2027-04-25 \
  [--early-cap 150000] [--big 5000] [--mixed-group late|early|none] \
  --out exam_allocation.xlsx
```

Take the dates from the user's request. Check that the weekdays they quoted match: the
script prints each date with its weekday. A typical run takes 10–20 seconds. Each tier
prints OPTIMAL or FEASIBLE. If any tier is only FEASIBLE, rerun with a larger `--time`,
or tell the user the figure may not be the true optimum.

Keep a copy of the script beside the outputs (e.g. as `allocate_exams.py`), so the user
has a rerunnable script with their deliverable. If they ask for the files in Downloads,
copy the script, the workbook and the `data/` CSVs into a folder there.

### 4. Verify, then report
Every check under "Hard-constraint checks" must print PASS. If any FAILs, fix the
cause before reporting. Open the workbook and spot-check a few rows, e.g. that Course
Name holds names and not numbers.

Give the user a short summary:
- a per-date table: courses, previous registrations, big courses
- registration spread (max vs min per date, for each group)
- domain clashes left: excess courses, affected domains, the split between groups, and
  whether the solver proved that's the minimum
- courses moved by the cap, if any, with registrations and new date
- unplaced courses, and the question you need answered about them
- assumptions, e.g. new courses carry 0 registrations and have no domain mapping

Be candid about rules that can't be met. Big courses usually outnumber dates, and early
dates are few, so some domain clashes are structurally unavoidable. Say so plainly, and
don't present an infeasible goal as achieved.

## Outputs
`exam_allocation.xlsx`:
1. **Allocation:** Course ID, Run ID, Course Name, Duration, New/Rerun, Previous
   Registration, Domain(s), Exam Date, Group, Moved (cap)
2. **Date summary:** per date: courses, new courses, total registration, big courses and their names
3. **Domain check:** per domain: courses on each date, with clash cells (more than 2) highlighted
4. **Exceptions:** unplaced courses, courses moved by the cap, big courses sharing a
   date, and courses in a domain clash, each with the reason
