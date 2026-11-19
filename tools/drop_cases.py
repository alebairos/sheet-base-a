#!/usr/bin/env python3
"""Dev-only: drop named cases from a suite (src list, cases and expected).
Usage: tools/drop_cases.py suite 'formula1' 'formula2' ..."""
import json, sys
suite, keys = sys.argv[1], set(sys.argv[2:])
src = f"tests/src/{suite}.txt"
try:
    lines = [l for l in open(src).read().split("\n") if l not in keys]
    open(src, "w").write("\n".join(lines))
except FileNotFoundError:
    pass
for d in ["cases", "expected"]:
    p = f"tests/{d}/{suite}.json"
    j = json.load(open(p))
    for k in keys:
        j.pop(k, None)
    with open(p, "w") as f:
        json.dump(j, f, indent=1, ensure_ascii=False)
        f.write("\n")
