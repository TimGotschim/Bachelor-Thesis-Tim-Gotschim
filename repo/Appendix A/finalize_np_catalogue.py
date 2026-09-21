#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
finalize_np_catalogue.py
========================
Bachelor's Thesis - Tim Gotschim | WU Vienna

Two final clean-up steps on the evaluated Numerical Precision catalogue:

  1. Complete the `GT_Formula_Used` column. The correction step recorded a
     formula only for the ratio and percentage items, because only those
     required a ground-truth correction. The average items were never
     affected and were left blank. Their ground truth is the mean of the
     three converted entity values, which is recorded here as avg(E1,E2,E3)
     so that the column documents the derivation of every question rather
     than only of the corrected ones.

  2. Remove the `GT_Previous_Value` column. The superseded ground-truth
     values are not referenced anywhere in the thesis and the column is
     therefore dropped from the submitted catalogue.

Both steps are verified before writing: the script recomputes the ground
truth of every average item from the raw entity values and the documented
unit conversion, and refuses to write the file if any item disagrees.
"""

import argparse
from openpyxl import load_workbook

AVERAGE_FORMULA = "avg(E1,E2,E3)"
TOLERANCE_PCT = 0.05   # recomputation must match stored GT to within 0.05%

CONVERSIONS = {
    ("metre", "foot"):                        lambda v: v / 0.3048,
    ("kilometre", "mile"):                    lambda v: v / 1.609344,
    ("square kilometre", "square mile"):      lambda v: v / 2.589988110336,
    ("degree celsius", "degree fahrenheit"):  lambda v: v * 9 / 5 + 32,
    ("kilogram", "pound"):                    lambda v: v * 2.20462262185,
    ("kilometre per hour", "mile per hour"):  lambda v: v / 1.609344,
}


def convert(value, source_unit, target_unit):
    fn = CONVERSIONS.get((str(source_unit).lower(), str(target_unit).lower()))
    return fn(value) if fn else None


def to_number(value):
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--sheet", default="NP Master Catalogue")
    args = ap.parse_args()

    wb = load_workbook(args.input)
    ws = wb[args.sheet]

    headers = {}
    for cell in ws[1]:
        if cell.value is not None:
            headers[str(cell.value).strip()] = cell.column

    c_op = headers["Operation"]
    c_src, c_tgt = headers["Source Unit"], headers["Target Unit"]
    c_e = [headers["Entity1 Value"], headers["Entity2 Value"], headers["Entity3 Value"]]
    c_gt = headers["Ground_Truth_Norm"]
    c_formula = headers["GT_Formula_Used"]
    c_prev = headers.get("GT_Previous_Value")

    # ---- Step 1: verify and fill the average formulas ----------------------
    filled, verified, failures = 0, 0, []

    for row in range(2, ws.max_row + 1):
        if ws.cell(row, headers["Question ID"]).value is None:
            continue
        if str(ws.cell(row, c_op).value).strip().lower() != "average":
            continue

        src = ws.cell(row, c_src).value
        tgt = ws.cell(row, c_tgt).value
        values = [convert(to_number(ws.cell(row, c).value), src, tgt) for c in c_e]
        values = [v for v in values if v is not None]
        stored = to_number(ws.cell(row, c_gt).value)

        if len(values) >= 2 and stored:
            predicted = sum(values) / len(values)
            deviation = abs(predicted - stored) / abs(stored) * 100
            if deviation < TOLERANCE_PCT:
                verified += 1
            else:
                failures.append((ws.cell(row, headers["Question ID"]).value,
                                 round(predicted, 4), stored, round(deviation, 4)))
        else:
            failures.append((ws.cell(row, headers["Question ID"]).value,
                             "not recomputable", stored, None))

        current = ws.cell(row, c_formula).value
        if current is None or not str(current).strip():
            ws.cell(row, c_formula).value = AVERAGE_FORMULA
            filled += 1

    print(f"Average rows verified against raw entity values : {verified}")
    print(f"Average rows given the formula {AVERAGE_FORMULA}    : {filled}")

    if failures:
        print(f"\nABORTED - {len(failures)} average item(s) did not reproduce:")
        for f in failures[:10]:
            print("  ", f)
        raise SystemExit(1)

    # ---- Step 2: drop the GT_Previous_Value column --------------------------
    if c_prev is not None:
        ws.delete_cols(c_prev)
        print(f"Removed column GT_Previous_Value (was column index {c_prev})")
    else:
        print("Column GT_Previous_Value not present, nothing removed")

    wb.save(args.output)
    print(f"Written: {args.output}")


if __name__ == "__main__":
    main()
