"""Convert a saved Google Sheets get_values result ({"range": ..., "values": [[...]]}) to CSV.

Use when a sheet isn't public, so the CSV export URL returns a sign-in page and
the data has to be read through the Google Sheets connector instead.  Rows are
padded to equal width so the header and data line up.

Usage: python sheet_json_to_csv.py <saved_result.json|.txt> <out.csv>
"""
import csv
import json
import sys

src, out = sys.argv[1], sys.argv[2]
values = json.load(open(src))["values"]
width = max(len(r) for r in values)
with open(out, "w", newline="") as f:
    w = csv.writer(f)
    for r in values:
        w.writerow(r + [""] * (width - len(r)))
print(f"Wrote {out}: {len(values)} rows (incl. header) x {width} columns")
