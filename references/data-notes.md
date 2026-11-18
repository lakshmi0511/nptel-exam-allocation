# Data notes from past runs

Observed in the Jan–Apr 2027 cycle (allocation done 2026-10-25). Re-check each of these
against new data; don't assume they still hold.

## Course list (sheet "course list format")
- The header row has unnamed and duplicate columns: two blank-named columns, two
  `current ID`, two `Old course URL`, two `NPTEL ID`. The populated 9-digit ID is in
  the second `NPTEL ID` (pandas: `NPTEL ID.1`).
- `Old Course ID` is the run code (e.g. `noc26-cs107`), not the ID the domain sheet uses.
- New courses (`Type = New`) had no NPTEL ID and no Previous Registration. They are
  treated as 0 registrations with no domain, and get placeholder IDs `NEW-<S No>`.
- A few ID cells hold two IDs separated by a newline (merged courses). Every ID is used
  for domain matching.
- One NPTEL ID (101104063, Engineering Thermodynamics) appeared twice: a 12-week and an
  8-week run. They're treated as separate courses.
- Durations seen: `4 Weeks`, `8 Weeks`, `12 Weeks`, plus 6 mixed ones like
  `8 Weeks (Rerun) + 4 Weeks (New)`. The user put the mixed ones in April.
- Previous Registration ranged from 3 to 82,863 (Introduction to IoT); 52 courses were
  above 5,000.

## Domain mapping
- 83 rows, 77 real domains. Rows with a blank `Domain Name` were notes and two unnamed
  "Environment" rows; they're skipped.
- Domain names repeat across disciplines (e.g. "Engineering Mechanics(Pre-Req)",
  "Power Systems and Power Electronics"). They're labelled `Name [Discipline]`, with
  `#1`/`#2` when the same discipline repeats a name, as with Management "Minor" and
  Mechanical "Robotics".
- `core`/`elective` hold "To Be Developed" and free-text names, which are ignored. Only
  9-digit IDs count.
- Only about 48% of rerun courses matched any domain.

## Outcome of that run (for comparison)
- 333 March courses, 766 April courses. A cap of 150,000 per March date moved 3
  courses to April: Enhancing Soft Skills and Personality, Soft Skill Development, and
  Data Base Management System.
- 44 excess domain clashes (27 in March, 17 in April). The solver proved that's the
  minimum, because 2 March dates can't spread large domains.
- All tiers solved to optimal in about 12 seconds.

## July 2026 cycle (allocation done 2026-10-30)
- The input was `July2026.xlsx` (sheet Main) in Downloads: Course ID (a run code like
  `noc26-ae21`), Course Name, Duration, Type, and Applicable NPTEL Domain (text). It had
  **no registrations and no NPTEL ID**. Type included "Rerun (Upgraded)" for the 4
  mixed-duration courses.
- Built course_list.csv by joining on the run code (see `prepare_inputs.py` in
  `~/Downloads/nptel_exam_alloc_jul2026/`):
  - `Final Course List (July - Dec 2026).xlsx`, sheets New + Rerun, gives `NPTEL ID.1`
    (9-digit), `Check Old ID` (the previous run, e.g. `noc25_ae24`) and `End date`.
  - The Previous Registration comes from the "Enroll&Reg-Jul26 & Domain" tab of the
    Jan 2027 course-list Google Sheet: `Registrations` of the previous run, matched on
    `NOC Course Id`, with `_` and `-` treated as the same. That tab covers every
    semester since 2014.
- The Final Course List also holds the schedule that was actually used (one sheet per
  exam date). Scoring it with the same measures is a useful comparison. Actual: 165
  excess domain clashes. Optimised: 103.
- SET 2 timings: 4-week courses run Aug 17 – Sep 11, 8-week courses Aug 17 – Oct 9. The
  72 SET 2 8-week reruns can't sit the September dates.

## Known issue in the Jan 2027 allocation of 2026-10-25
That run was made before the end-date rule existed. 78 SET 2 8-week reruns, which end
Apr 9, 2027, were given March dates. Rerun that cycle with the current script before
using its output.
