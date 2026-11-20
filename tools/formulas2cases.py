#!/usr/bin/env python3
"""Turn a list of formulas (one per line) into a cases JSON file.

Each formula goes into Sheet1!C1 of a workbook with a fixed data set, so the
same list can be fed to the LibreOffice oracle and to the engine.
Lines starting with '#' are comments.
Usage: tools/formulas2cases.py list.txt > tests/cases/name.json
"""
import json, os, sys

# FORMULA_CELL=D5 puts the formula elsewhere (implicit intersection)
TARGET = os.environ.get("FORMULA_CELL", "C1")

DATA = {"A1": 1, "A2": 2, "A3": 3, "A4": "text", "A5": True, "A7": "5", "A8": -1.5,
        "A9": 0, "A10": "Apple", "A11": "banana", "A12": "=1/0",
        "B1": 10, "B2": 20, "B3": 30, "B4": "x", "B5": False, "B6": "  padded  ", "B7": 2.5,
        "B8": "Banana", "B9": 100, "B10": "apple pie", "B11": "", "B12": 7}
# DATASET=2: another data set (sorted numbers, duplicates, other errors)
if os.environ.get("DATASET") == "2":
    DATA = {"A1": -3, "A2": 0, "A3": 0.5, "A4": 2, "A5": 2, "A6": 7, "A7": "7", "A8": 15, "A9": "=NA()", "A10": "apple",
            "A11": "Apple", "A12": "=SQRT(-4)", "B1": "b", "B2": "a", "B3": True, "B4": "", "B5": 3, "B6": "=A6*2",
            "B7": "x y", "B8": " lead", "B9": -0.25, "B10": "=\"\"", "B11": 1E10, "B12": "zeta"}
SHEET2 = {"A1": 10, "A2": 20, "A3": "two", "B1": "=A1*2"}

cases = {}
for line in open(sys.argv[1], encoding="utf-8"):
    f = line.rstrip("\n")
    if not f.strip() or f.startswith("# "):
        continue
    d = dict(DATA)
    d[TARGET] = f if f.startswith("=") else "=" + f
    cases[f] = {"sheets": {"Sheet1": d, "Sheet2": dict(SHEET2)}}
json.dump(cases, sys.stdout, indent=1, ensure_ascii=False)
print()
