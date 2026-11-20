#!/usr/bin/env python3
"""Dev-only: mutate formulas inside fixture workbooks (realistic*, workbooks, scenarios).
Usage: python3 tools/mutate_wb.py <seed> <count> > cases.json"""
import json, random, re, sys, glob
random.seed(int(sys.argv[1])); N = int(sys.argv[2])
books = []
for f in sorted(glob.glob("tests/cases/realistic*.json") + glob.glob("tests/cases/workbooks.json") + glob.glob("tests/cases/scenarios.json") + glob.glob("tests/cases/money*.json")):
    books += list(json.load(open(f)).values())
nums = ["0", "1", "-1", "2", "0.5", "10", "\"\"", "\"a\"", "TRUE", "NA()", "1/0", "{1,2}"]
fns1 = ["SUM", "ABS", "LEN", "ISNUMBER", "ISERROR", "NOT", "UPPER", "INT", "COUNT", "N", "TRIM", "ROUND"]
def mutate(f, cells):
    k = random.random()
    if k < 0.3:
        toks = list(re.finditer(r"(?<![A-Za-z$!:])-?\d+(\.\d+)?(?![\d:A-Za-z!])", f))
        if toks: m = random.choice(toks); return f[:m.start()] + random.choice(nums) + f[m.end():]
    if k < 0.6:
        toks = list(re.finditer(r"(?<![A-Za-z!'])\$?[A-Z]\$?\d+(?![\d(])", f))
        if toks and cells: m = random.choice(toks); return f[:m.start()] + random.choice(cells) + f[m.end():]
    if k < 0.8: return "=" + random.choice(fns1) + "(" + f[1:] + ")"
    return f + random.choice(["+1", "&\"x\"", "*{1,2}", "=0", "<\"b\"", "^2", "%", "/0"])
out = {}
while len(out) < N:
    wb = json.loads(json.dumps(random.choice(books)))
    sheets = wb["sheets"]
    fcells = [(s, a) for s in sheets for a, v in sheets[s].items() if isinstance(v, str) and v.startswith("=")]
    if not fcells: continue
    for _ in range(random.randint(1, 3)):
        s, a = random.choice(fcells)
        cells = [c for c in sheets[s] if re.match(r"^[A-Z]\d+$", c)]
        sheets[s][a] = mutate(sheets[s][a], cells)
    out[f"mw{len(out)}"] = wb
print(json.dumps(out))
