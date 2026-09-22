#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fill_np_recompute.py
====================
Bachelor's Thesis - Tim Gotschim | WU Vienna
Supervisor: Dr. Svitlana Vakulenko

Fills the RECOMPUTE placeholders left in the Numerical Precision master
catalogue after the ground-truth correction.

BACKGROUND
----------
A systematic ground-truth error affected the ratio and percentage items of
the NP Master Catalogue: the reference answer had been computed from an
entity that the question text did not actually refer to. The corrected
ground truth was written to `Ground_Truth_Norm` by the correction step,
the superseded value was preserved in `GT_Previous_Value`, and the formula
actually applied was recorded in `GT_Formula_Used`.

Because the two derived columns `New_Deviation_Pct` and `New_Correctness`
depend on the ground truth, they were invalidated by that correction and
set to the sentinel string "RECOMPUTE". This script recomputes them.

The `average` items were never affected and their values are left
untouched, as is every other already-settled row.

CLASSIFICATION RULE (identical to the rule applied to the settled rows)
-----------------------------------------------------------------------
  deviation = |Extracted_Final_Norm - Ground_Truth_Norm| / |Ground_Truth_Norm| * 100

  Not_Attempted : Extracted_Final_Norm is "NULL" or not parseable as a
                  number, i.e. the model stated no final numerical answer
  Correct       : deviation <= 3.0
  Hallucinated  : deviation >  3.0

The +/- 3% tolerance follows the threshold documented in the methodology
chapter and is calibrated to the documented source variation of real-world
physical measurements.

USAGE
-----
    pip install openpyxl
    python3 fill_np_recompute.py \
        --input  FINAL_Numerical_Precision_CLEAN.xlsx \
        --output FINAL_Numerical_Precision_EVALUATED.xlsx

    # report only, write nothing:
    python3 fill_np_recompute.py --input ... --dry-run
"""

import argparse
import statistics
from collections import Counter, defaultdict

from openpyxl import load_workbook

SENTINEL = "RECOMPUTE"
NULL_TOKEN = "NULL"
THRESHOLD = 3.0          # +/- 3 percent
DEV_DECIMALS = 4         # matches the formatting of the settled rows

# Column headers the script relies on (resolved by name, not by index)
COL_EXTRACTED = "Extracted_Final_Norm"
COL_GT = "Ground_Truth_Norm"
COL_DEV = "New_Deviation_Pct"
COL_CORR = "New_Correctness"
COL_MODEL = "Model"
COL_VARIANT = "Prompt Variant"
COL_OPERATION = "Operation"
COL_QID = "Question ID"


def to_number(value):
    """Parse a cell into a float, or return None if it is not a number."""
    if value is None:
        return None
    text = str(value).strip()
    if not text or text.upper().startswith(NULL_TOKEN):
        return None
    try:
        return float(text)
    except ValueError:
        return None


def classify(extracted, ground_truth):
    """
    Return (deviation_or_None, label).

    A row counts as Not_Attempted when the model produced no parseable
    final number, which is how abstentions are encoded in this catalogue.
    """
    ex = to_number(extracted)
    gt = to_number(ground_truth)

    if ex is None or gt is None or gt == 0:
        return None, "Not_Attempted"

    deviation = abs(ex - gt) / abs(gt) * 100.0
    label = "Correct" if deviation <= THRESHOLD else "Hallucinated"
    return deviation, label


def main():
    parser = argparse.ArgumentParser(
        description="Fill RECOMPUTE placeholders in the NP master catalogue"
    )
    parser.add_argument("--input", default="FINAL_Numerical_Precision_CLEAN.xlsx")
    parser.add_argument("--output", default="FINAL_Numerical_Precision_EVALUATED.xlsx")
    parser.add_argument("--sheet", default="NP Master Catalogue")
    parser.add_argument("--dry-run", action="store_true",
                        help="Report the outcome without writing a file")
    args = parser.parse_args()

    wb = load_workbook(args.input)
    ws = wb[args.sheet]

    # Resolve column letters by header name so the script does not depend
    # on a fixed column order.
    headers = {}
    for cell in ws[1]:
        if cell.value is not None:
            headers[str(cell.value).strip()] = cell.column

    required = [COL_EXTRACTED, COL_GT, COL_DEV, COL_CORR]
    missing = [c for c in required if c not in headers]
    if missing:
        raise SystemExit(f"Missing expected column(s): {missing}")

    c_ex = headers[COL_EXTRACTED]
    c_gt = headers[COL_GT]
    c_dev = headers[COL_DEV]
    c_corr = headers[COL_CORR]

    filled = 0
    untouched = 0
    labels = Counter()
    deviations = []
    by_model = defaultdict(Counter)
    by_variant = defaultdict(Counter)
    by_operation = defaultdict(Counter)

    for row in range(2, ws.max_row + 1):
        if ws.cell(row, headers[COL_QID]).value is None:
            continue

        current = ws.cell(row, c_corr).value
        is_placeholder = SENTINEL in str(current).upper()

        if is_placeholder:
            deviation, label = classify(
                ws.cell(row, c_ex).value,
                ws.cell(row, c_gt).value,
            )
            ws.cell(row, c_dev).value = (
                NULL_TOKEN if deviation is None else round(deviation, DEV_DECIMALS)
            )
            ws.cell(row, c_corr).value = label
            filled += 1
        else:
            label = current
            deviation = to_number(ws.cell(row, c_dev).value)
            untouched += 1

        labels[label] += 1
        if label == "Hallucinated" and deviation is not None:
            deviations.append(deviation)

        if COL_MODEL in headers:
            by_model[ws.cell(row, headers[COL_MODEL]).value][label] += 1
        if COL_VARIANT in headers:
            by_variant[ws.cell(row, headers[COL_VARIANT]).value][label] += 1
        if COL_OPERATION in headers:
            by_operation[ws.cell(row, headers[COL_OPERATION]).value][label] += 1

    total = sum(labels.values())

    def pct(n):
        return 100.0 * n / total if total else 0.0

    print("=" * 64)
    print(f"Rows recomputed : {filled}")
    print(f"Rows untouched  : {untouched}")
    print(f"Rows total      : {total}")
    print("-" * 64)
    print(f"Hallucinated : {labels['Hallucinated']:5d}  ({pct(labels['Hallucinated']):.1f}%)")
    print(f"Correct      : {labels['Correct']:5d}  ({pct(labels['Correct']):.1f}%)")
    print(f"Not_Attempted: {labels['Not_Attempted']:5d}  ({pct(labels['Not_Attempted']):.1f}%)")
    if deviations:
        print(f"Median deviation of hallucinated responses: "
              f"{statistics.median(deviations):.1f}%")

    def breakdown(title, mapping):
        print("-" * 64)
        print(title)
        for key, counter in mapping.items():
            n = sum(counter.values())
            if not n:
                continue
            print(f"  {str(key):<22} H {100*counter['Hallucinated']/n:5.1f}  "
                  f"C {100*counter['Correct']/n:5.1f}  "
                  f"NA {100*counter['Not_Attempted']/n:4.1f}   (n={n})")

    breakdown("By model:", by_model)
    breakdown("By prompt variant:", by_variant)
    breakdown("By operation:", by_operation)
    print("=" * 64)

    if args.dry_run:
        print("Dry run: no file written.")
        return

    wb.save(args.output)
    print(f"Written: {args.output}")


if __name__ == "__main__":
    main()
