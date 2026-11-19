#!/usr/bin/env python3
"""Turn a list of formulas (one per line) into a cases JSON file.

Each formula goes into Sheet1!C1 of a workbook with a fixed data set, so the
same list can be fed to the LibreOffice oracle and to the engine.
Lines starting with '#' are comments.
Usage: tools/formulas2cases.py list.txt > tests/cases/name.json
"""
import json, sys

DATA = {"A1": 1, "A2": 2, "A3": 3, "A4": "text", "A5": True, "A7": "5", "A8": -1.5,
        "A9": 0, "A10": "Apple", "A11": "banana", "A12": "=1/0",
        "B1": 10, "B2": 20, "B3": 30, "B4": "x", "B5": False, "B6": "  padded  ", "B7": 2.5,
        "B8": "Banana", "B9": 100, "B10": "apple pie", "B11": "", "B12": 7}
SHEET2 = {"A1": 10, "A2": 20, "A3": "two", "B1": "=A1*2"}

cases = {}
for line in open(sys.argv[1], encoding="utf-8"):
    f = line.rstrip("\n")
    if not f.strip() or f.startswith("# "):
        continue
    d = dict(DATA)
    d["C1"] = f if f.startswith("=") else "=" + f
    cases[f] = {"sheets": {"Sheet1": d, "Sheet2": dict(SHEET2)}}
json.dump(cases, sys.stdout, indent=1, ensure_ascii=False)
print()
