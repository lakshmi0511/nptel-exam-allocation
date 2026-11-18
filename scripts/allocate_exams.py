"""Allocate NPTEL exam dates to courses with OR-Tools CP-SAT.

Only the exam date is allocated: every course runs in both sessions
(Forenoon and Afternoon) on its date, so a student can write up to 2 exams
on one date.

Course groups
-------------
  early dates : existing (Rerun) 4-week and 8-week courses
  late dates  : new 4-week and 8-week courses, plus all 12-week courses
  mixed       : durations like "8 Weeks (Rerun) + 4 Weeks (New)" go to
                --mixed-group (late by default; "none" leaves them unplaced)
  anything else is left unplaced and reported as an exception.
  course end  : if the course list has an "End date" column, every exam must be at
                least --min-gap-days after the course ends.  Early-group courses that
                end too late for every early date (e.g. SET 2 8-week reruns) go to
                the late dates instead and are listed as exceptions.

Method
------
One CP-SAT model covers all dates.  With --early-cap, early-group courses that
don't fit under the cap may move to late dates; the number moved is minimised
first.  Then, lexicographically (solve, freeze the value, next goal):

  1. domain clashes      - per domain and date, courses beyond 2 ("excess").
  2. same, weighted by the domain's total previous registration, so that
     unavoidable clashes land on low-registration domains.
  3. big-course crowding - sum of squares of big courses per date.
  4. registration balance - max-min registrations per date, within each group.
  5. course-count balance - max-min courses per date, within each group.

Example
-------
  python allocate_exams.py --courses course_list.csv --domains domain_courses.csv \
      --early-dates 2027-03-20,2027-03-21 \
      --late-dates 2027-04-16,2027-04-17,2027-04-18,2027-04-23,2027-04-24,2027-04-25 \
      --early-cap 150000 --out exam_allocation.xlsx
"""
import argparse
import datetime as dt
import json
import re
from collections import defaultdict

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from ortools.sat.python import cp_model

PER_DATE = 2  # exams a student can write on one date (one per session)

# Course-list column names (as in the sheet header) -> internal names.
# The NPTEL sheet has two "NPTEL ID" columns; pandas names the second
# "NPTEL ID.1", which is the populated one holding the 9-digit ID used by the
# domain sheet.  Override with --course-cols '{"Sheet header": "internal"}'.
COURSE_COLS = {
    "S No": "sno",
    "NPTEL ID.1": "nptel_id",
    "Old Course ID": "run_id",
    "Course Name": "name",
    "Duration": "duration",
    "Type": "type",
    "Previous Registration count": "reg",
}
OPTIONAL_COURSE_COLS = {"End date": "end"}
DOMAIN_COLS = {"Discipline name": "disc", "Domain Name": "dom", "core": "core", "elective": "elective"}


def parse_dates(s):
    return [dt.date.fromisoformat(x.strip()) for x in s.split(",") if x.strip()]


def group_label(dates):
    return "/".join(dict.fromkeys(d.strftime("%B") for d in dates))


def fmt_date(d):
    return d.strftime("%a %d-%b-%Y")


def load_courses(path, cols, big, mixed_group):
    df = pd.read_csv(path, dtype=str, keep_default_na=False)
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise SystemExit(f"{path} is missing columns: {missing}\nColumns found: {list(df.columns)}")
    opt = {k: v for k, v in OPTIONAL_COURSE_COLS.items() if k in df.columns and v not in cols.values()}
    df = df[list(cols) + list(opt)].rename(columns={**cols, **opt})
    df = df.apply(lambda s: s.str.strip())
    df = df[df["name"] != ""].reset_index(drop=True)
    if "run_id" not in df:
        df["run_id"] = ""
    if "sno" not in df:
        df["sno"] = (df.index + 1).astype(str)
    df["reg_missing"] = df["reg"] == ""
    df["reg"] = pd.to_numeric(df["reg"].str.replace(",", ""), errors="coerce").fillna(0).astype(int)
    df["ids"] = df["nptel_id"].apply(lambda s: re.findall(r"\d{9}", s))
    df["course_id"] = df.apply(lambda r: ", ".join(r.ids) or f"NEW-{r.sno}", axis=1)
    df["weeks"] = df["duration"].str.extract(r"^\s*(\d+)\s*weeks?\s*$", flags=re.I)[0].astype(float)
    df["mixed"] = df["duration"].str.contains(r"\+")
    df["big"] = df["reg"] > big
    if "end" in df:
        df["end"] = pd.to_datetime(df["end"], errors="coerce", format="mixed").dt.date
    else:
        df["end"] = None

    def group(r):
        if r.weeks in (4, 8) and r.type.lower() == "rerun":
            return "early"
        if (r.weeks in (4, 8) and r.type.lower() == "new") or r.weeks == 12:
            return "late"
        if r.mixed and mixed_group != "none":
            return mixed_group
        return None  # unclear -> exception, ask the user

    df["group"] = df.apply(group, axis=1)
    return df


def load_domains(path):
    """Return {domain label: set of NPTEL IDs}. Rows without a Domain Name are notes, not domains."""
    d = pd.read_csv(path, dtype=str, keep_default_na=False)
    missing = [c for c in DOMAIN_COLS if c not in d.columns]
    if missing:
        raise SystemExit(f"{path} is missing columns: {missing}\nColumns found: {list(d.columns)}")
    d = d[list(DOMAIN_COLS)].rename(columns=DOMAIN_COLS).apply(lambda s: s.str.strip())
    d = d[d["dom"] != ""]
    name_counts = d["dom"].value_counts()
    pair_counts = d.groupby(["disc", "dom"]).size()
    seen = defaultdict(int)
    domains, discipline = {}, {}
    for r in d.itertuples():
        label = r.dom if name_counts[r.dom] == 1 else f"{r.dom} [{r.disc}]"
        if pair_counts[(r.disc, r.dom)] > 1:
            seen[(r.disc, r.dom)] += 1
            label += f" #{seen[(r.disc, r.dom)]}"
        # core/elective lists use "," between courses and "_" / "|" between
        # equivalent IDs; every 9-digit ID counts as a member of the domain.
        domains[label] = set(re.findall(r"\d{9}", f"{r.core} {r.elective}"))
        discipline[label] = r.disc
    return domains, discipline


def solve(sub, course_domains, dom_weight, early, late, early_cap, time_limit, earliest):
    """Assign every course in `sub` a date.  Returns {row index: date}.

    `earliest[i]` is the first date course i may be examined on (end date + gap), or None.
    """
    all_dates = early + late
    groups = {"early": range(len(early)), "late": range(len(early), len(all_dates))}
    idx = list(sub.index)
    allowed = {}
    for i in idx:
        g = sub.at[i, "group"]
        ts = list(groups[g]) + (list(groups["late"]) if g == "early" and early_cap else [])
        allowed[i] = [t for t in ts if earliest[i] is None or all_dates[t] >= earliest[i]]
    m = cp_model.CpModel()
    x = {(i, t): m.NewBoolVar(f"x{i}_{t}") for i in idx for t in allowed[i]}

    def on(t, items):
        return [x[i, t] for i in items if (i, t) in x]

    for i in idx:
        m.AddExactlyOne(x[i, t] for t in allowed[i])

    reg = {i: int(sub.at[i, "reg"]) for i in idx}
    total = sum(reg.values())
    reg_date = []
    for t in range(len(all_dates)):
        r = m.NewIntVar(0, total, "")
        m.Add(r == sum(reg[i] * x[i, t] for i in idx if (i, t) in x))
        reg_date.append(r)
    if early_cap:
        for t in groups["early"]:
            m.Add(reg_date[t] <= early_cap)
    early_idx = [i for i in idx if sub.at[i, "group"] == "early"]
    obj_moved = sum(x[i, t] for i in early_idx for t in groups["late"] if (i, t) in x)

    # 1/2. domain clash excess per (domain, date)
    members = defaultdict(list)
    for i in idx:
        for dname in course_domains[i]:
            members[dname].append(i)
    excess, weighted = [], []
    for dname, mem in members.items():
        for t in range(len(all_dates)):
            vs = on(t, mem)
            if len(vs) <= PER_DATE:
                continue
            e = m.NewIntVar(0, len(vs), "")
            m.Add(e >= sum(vs) - PER_DATE)
            excess.append(e)
            weighted.append(dom_weight[dname] * e)

    # 3. big-course crowding
    bigs = [i for i in idx if sub.at[i, "big"]]
    sq_terms = []
    for t in range(len(all_dates)):
        b = m.NewIntVar(0, len(bigs), "")
        m.Add(b == sum(on(t, bigs)))
        sq = m.NewIntVar(0, len(bigs) ** 2, "")
        m.AddMultiplicationEquality(sq, [b, b])
        sq_terms.append(sq)

    # 4/5. registration and course-count balance within each group
    spreads_reg, spreads_cnt = [], []
    for ts in groups.values():
        cnts = []
        for t in ts:
            c = m.NewIntVar(0, len(idx), "")
            m.Add(c == sum(on(t, idx)))
            cnts.append(c)
        for vals, ub, out in ((reg_date[ts.start:ts.stop], total, spreads_reg), (cnts, len(idx), spreads_cnt)):
            mx, mn = m.NewIntVar(0, ub, ""), m.NewIntVar(0, ub, "")
            m.AddMaxEquality(mx, vals)
            m.AddMinEquality(mn, vals)
            out.append(mx - mn)

    solver = cp_model.CpSolver()
    solver.parameters.num_workers = 8
    solver.parameters.max_time_in_seconds = time_limit
    tiers = [
        ("early courses moved to late dates (cap)", obj_moved),
        ("domain clash excess", sum(excess)),
        ("registration-weighted domain clashes", sum(weighted)),
        ("big-course crowding (sum of squares)", sum(sq_terms)),
        ("registration balance (max-min per date, each group)", sum(spreads_reg)),
        ("course-count balance (max-min per date, each group)", sum(spreads_cnt)),
    ]
    xs = list(x.values())
    for name, obj in tiers:
        m.Minimize(obj)
        status = solver.Solve(m)
        if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            raise SystemExit(f"No solution for tier '{name}' ({solver.StatusName(status)}). "
                             "If the early cap is set, check it is achievable.")
        val = int(round(solver.ObjectiveValue()))
        print(f"  {name}: {val}  ({solver.StatusName(status)}, bound {solver.BestObjectiveBound():.0f})")
        m.ClearHints()
        for v in xs:
            m.AddHint(v, solver.Value(v))
        m.Add(obj <= val)

    return {i: all_dates[next(t for t in allowed[i] if solver.Value(x[i, t]))] for i in idx}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--courses", required=True, help="course list CSV")
    ap.add_argument("--domains", required=True, help="domain mapping CSV")
    ap.add_argument("--early-dates", required=True, help="ISO dates for existing 4/8-week courses, comma-separated")
    ap.add_argument("--late-dates", required=True, help="ISO dates for new 4/8-week and all 12-week courses")
    ap.add_argument("--out", default="exam_allocation.xlsx")
    ap.add_argument("--early-cap", type=int, default=0, help="max registrations per early date (0 = no cap)")
    ap.add_argument("--big", type=int, default=5000, help="a course is big if Previous Registration exceeds this")
    ap.add_argument("--mixed-group", choices=["late", "early", "none"], default="late",
                    help="group for mixed durations like '8 Weeks (Rerun) + 4 Weeks (New)'")
    ap.add_argument("--min-gap-days", type=int, default=7,
                    help="exam must be at least this many days after the course End date (if the column exists)")
    ap.add_argument("--course-cols", help="JSON override of the course column mapping")
    ap.add_argument("--time", type=float, default=60, help="CP-SAT seconds per objective tier")
    args = ap.parse_args()

    early, late = parse_dates(args.early_dates), parse_dates(args.late_dates)
    if set(early) & set(late):
        raise SystemExit("Early and late dates overlap")
    labels = {"early": group_label(early), "late": group_label(late)}
    labels["early_ended_late"] = f"{group_label(late)} (course ends after {group_label(early)} dates)"
    all_dates = early + late
    cols = json.loads(args.course_cols) if args.course_cols else COURSE_COLS

    df = load_courses(args.courses, cols, args.big, args.mixed_group)
    gap = dt.timedelta(days=args.min_gap_days)
    df["earliest"] = [e + gap if isinstance(e, dt.date) else None for e in df["end"]]
    # Early-group courses that can't sit any early date (course still running) go late.
    df["ends_late"] = (df.group == "early") & df.earliest.apply(lambda e: e is not None and e > max(early))
    df.loc[df.ends_late, "group"] = "late"
    # Courses that can't sit any date of their group at all are left unplaced.
    last = {"early": max(early), "late": max(late)}
    df["no_date"] = [g is not None and e is not None and e > last[g] for g, e in zip(df.group, df.earliest)]
    df.loc[df.no_date, "group"] = None
    domains, discipline = load_domains(args.domains)
    course_domains = {i: sorted(dn for dn, ids in domains.items() if set(r) & ids) for i, r in df["ids"].items()}
    dom_reg = {dn: int(sum(df.at[i, "reg"] for i in df.index if dn in course_domains[i])) for dn in domains}
    dom_weight = {dn: max(1, v) for dn, v in dom_reg.items()}

    print("Dates: " + ", ".join(f"{fmt_date(d)} ({labels['early'] if d in early else labels['late']})" for d in all_dates))
    print(f"Courses: {len(df)} | early ({labels['early']}): {(df.group == 'early').sum()} "
          f"| late ({labels['late']}): {(df.group == 'late').sum()} | unclassified: {df.group.isna().sum()} "
          f"| mixed durations: {int(df.mixed.sum())} -> {args.mixed_group}")
    print(f"Course end dates: {'used, min gap ' + str(args.min_gap_days) + ' days' if df.end.notna().any() else 'no End date column - not checked'}"
          f" | early courses ending too late for early dates -> late: {int(df.ends_late.sum())}"
          f" | no feasible date: {int(df.no_date.sum())}")
    print(f"Domains: {len(domains)} | courses mapped to >=1 domain: {sum(1 for v in course_domains.values() if v)}")
    print(f"Big courses (>{args.big:,}): {int(df.big.sum())} | early cap: {args.early_cap:,}" if args.early_cap
          else f"Big courses (>{args.big:,}): {int(df.big.sum())} | early cap: none")

    df["date"] = None
    sub = df[df.group.notna()]
    for i, d in solve(sub, course_domains, dom_weight, early, late, args.early_cap, args.time,
                      {i: df.at[i, "earliest"] for i in sub.index}).items():
        df.at[i, "date"] = d
    df["moved"] = (df.group == "early") & df.date.isin(late)
    df["group"] = df.group.where(~df.ends_late, "early_ended_late")

    # ---------------- hard-constraint checks (recomputed from the final table) ----------------
    print("\nHard-constraint checks")
    placed = df[df.date.notna()]
    checks = [
        ("Every classifiable course placed exactly once",
         df[df.group.notna()].date.notna().all() and placed.index.is_unique),
        ("Unclassified courses left unplaced (awaiting decision)", df[df.group.isna()].date.isna().all()),
        (f"Early dates ({labels['early']}) hold only early-group courses",
         (placed[placed.date.isin(early)].group == "early").all()),
        (f"Late dates ({labels['late']}) hold only late-group courses (plus early courses moved by the cap"
         " or ending too late)",
         all(r.group in ("late", "early_ended_late") or r.moved for r in placed.itertuples() if r.date in late)),
        ("Early-group courses moved to late dates only because of the cap", bool(args.early_cap) or not df.moved.any()),
        (f"Every exam at least {args.min_gap_days} days after the course end date",
         all(r.earliest is None or r.date >= r.earliest for r in placed.itertuples())),
    ] + ([(f"Every early date within the cap of {args.early_cap:,}",
           all(placed[placed.date == d].reg.sum() <= args.early_cap for d in early))] if args.early_cap else []) + [
        ("All placements use a listed date", placed.date.isin(all_dates).all()),
    ]
    for name, ok in checks:
        print(f"  [{'PASS' if ok else 'FAIL'}] {name}")

    # ---------------- soft-goal metrics ----------------
    dom_rows, clash_courses, n_clash_cells, n_excess = [], defaultdict(list), 0, 0
    for dn in domains:
        mem = [i for i in placed.index if dn in course_domains[i]]
        row = {"dn": dn, "mem": mem, "dates": {}, "clash": [], "excess": 0}
        for d in all_dates:
            ons = [i for i in mem if df.at[i, "date"] == d]
            row["dates"][d] = len(ons)
            if len(ons) > PER_DATE:
                row["clash"].append(d)
                row["excess"] += len(ons) - PER_DATE
                for i in ons:
                    clash_courses[i].append(f"{dn} ({len(ons)} courses on {fmt_date(d)})")
        n_clash_cells += len(row["clash"])
        n_excess += row["excess"]
        dom_rows.append(row)
    moved = df[df.moved]
    print(f"\nMoved early -> late: {len(moved)} courses, {int(moved.reg.sum()):,} registrations"
          + "".join(f"\n  - {r['name']} ({r.reg:,}) -> {fmt_date(r.date)}" for _, r in moved.iterrows()))
    print(f"Domain clashes: {n_clash_cells} domain-date cells, {n_excess} excess courses, "
          f"{sum(1 for r in dom_rows if r['clash'])} of {len(domains)} domains affected")
    for g, ds in (("early", early), ("late", late)):
        regs = [int(placed[placed.date == d].reg.sum()) for d in ds]
        cnts = [int((placed.date == d).sum()) for d in ds]
        bigs = [int(placed[placed.date == d].big.sum()) for d in ds]
        print(f"{labels[g]}: registrations per date max {max(regs):,} / min {min(regs):,}; "
              f"courses per date {cnts}; big courses per date {bigs}")
    ends_late = df[df.ends_late]
    if len(ends_late):
        print(f"Early-group courses placed late because they end after the early dates: {len(ends_late)} "
              f"({int(ends_late.reg.sum()):,} registrations)")
    for _, r in df[df.group.isna()].iterrows():
        print(f"UNPLACED: {r.course_id} | {r['name']} | {r.duration} | {r.type} | {r.reg:,}")

    # ---------------- Excel output ----------------
    wb = Workbook()
    hdr_font, hdr_fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="305496")
    red = PatternFill("solid", fgColor="F8CBAD")

    def sheet(ws, header, rows, widths):
        ws.append(header)
        for c in ws[1]:
            c.font, c.fill = hdr_font, hdr_fill
            c.alignment = Alignment(wrap_text=True, vertical="top")
        for r in rows:
            ws.append(r)
        for k, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(k)].width = w
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    # Note: with iterrows() use r["name"], not r.name (that is the row index).
    ws = wb.active
    ws.title = "Allocation"
    order = df.assign(_d=df.date.fillna(dt.date(2100, 1, 1))).sort_values(["_d", "reg"], ascending=[True, False])
    sheet(ws, ["Course ID", "Run ID", "Course Name", "Duration", "New/Rerun", "Previous Registration",
               "Domain(s)", "Course End Date", "Exam Date", "Group", f"Moved to {labels['late']} (cap)"],
          [[r.course_id, r.run_id, r["name"], r.duration, r.type, None if r.reg_missing else r.reg,
            "; ".join(course_domains[i]), fmt_date(r.end) if isinstance(r.end, dt.date) else "",
            fmt_date(r.date) if r.date else "NOT PLACED",
            labels.get(r.group, "Unclear - ask"), "Yes" if r.moved else ""] for i, r in order.iterrows()],
          [14, 13, 50, 16, 10, 12, 50, 17, 17, 22, 12])

    ws = wb.create_sheet("Date summary")
    rows = []
    for g, ds in (("early", early), ("late", late)):
        for d in ds:
            gp = placed[placed.date == d]
            rows.append([labels[g], fmt_date(d), len(gp), int((gp.type.str.lower() == "new").sum()), int(gp.reg.sum()),
                         int(gp.big.sum()),
                         "; ".join(f"{r.name} ({r.reg:,})" for r in gp[gp.big].sort_values("reg", ascending=False).itertuples())])
    sheet(ws, ["Group", "Exam Date", "Courses", "of which New", "Total Previous Registration",
               f"Big courses (>{args.big:,})", "Big course names"], rows, [12, 17, 9, 10, 14, 11, 140])
    for r in range(2, len(rows) + 2):
        ws.cell(r, 5).number_format = "#,##0"
        if ws.cell(r, 6).value > 1:
            ws.cell(r, 6).fill = red

    ws = wb.create_sheet("Domain check")
    rows, fills = [], []
    for k, row in enumerate(sorted(dom_rows, key=lambda r: -dom_reg[r["dn"]])):
        rows.append([row["dn"], discipline[row["dn"]], len(domains[row["dn"]]), len(row["mem"]), dom_reg[row["dn"]]]
                    + [row["dates"][d] for d in all_dates] + [row["excess"]])
        fills += [(k + 2, 6 + all_dates.index(d)) for d in row["clash"]]
    sheet(ws, ["Domain", "Discipline", "IDs listed in mapping", "Courses in this offering",
               "Total Previous Registration"] + [fmt_date(d) for d in all_dates] + ["Clash excess"],
          rows, [45, 28, 10, 10, 13] + [10] * len(all_dates) + [9])
    for r, c in fills:
        ws.cell(r, c).fill = red
    ws.cell(len(rows) + 3, 1, f"Clash = more than {PER_DATE} of the domain's courses on one date "
                              "(a student can write one exam per session). Highlighted cells are clashes.").font = Font(italic=True)

    ws = wb.create_sheet("Exceptions")
    rows = []
    for _, r in df[df.group.isna()].iterrows():
        rows.append([r.course_id, r["name"], r.duration, r.type, r.reg, "NOT PLACED",
                     f"Course ends {fmt_date(r.end)}; no listed date is {args.min_gap_days}+ days later" if r.no_date
                     else "Duration/type doesn't clearly fit either date group; decision needed"])
    for i, r in placed.iterrows():
        when = fmt_date(r.date)
        if r.ends_late:
            rows.append([r.course_id, r["name"], r.duration, r.type, r.reg, when,
                         f"Existing {r.duration.lower()} course ends {fmt_date(r.end)}, too late for the "
                         f"{labels['early']} dates (exam needs {args.min_gap_days}+ days after course end); placed in {labels['late']}"])
        if r.moved:
            rows.append([r.course_id, r["name"], r.duration, r.type, r.reg, when,
                         f"Existing {r.duration.lower()} course moved from {labels['early']} to {labels['late']} "
                         f"to keep each {labels['early']} date within {args.early_cap:,} registrations"])
        if r.big:
            others = placed[(placed.date == r.date) & placed.big & (placed.index != i)]
            if len(others):
                rows.append([r.course_id, r["name"], r.duration, r.type, r.reg, when,
                             f"Big course shares its date with {len(others)} other big course(s): "
                             + "; ".join(others["name"])])
        if clash_courses[i]:
            rows.append([r.course_id, r["name"], r.duration, r.type, r.reg, when,
                         "Domain clash: " + "; ".join(clash_courses[i])])
    sheet(ws, ["Course ID", "Course Name", "Duration", "New/Rerun", "Previous Registration", "Exam Date", "Reason"],
          rows, [14, 50, 16, 10, 12, 17, 120])

    wb.save(args.out)
    print(f"\nWrote {args.out}")


if __name__ == "__main__":
    main()
