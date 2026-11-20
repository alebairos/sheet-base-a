# Aggregates over ranges holding error cells and cells that read the
# aggregate's own cell (circular), optionally inside IFERROR/ISERROR.
# Run: python3 tests/src/cycrange.py 1 300 > tests/cases/cycrange.json
import json, random, sys
random.seed(int(sys.argv[1]) if len(sys.argv) > 1 else 1)
N = int(sys.argv[2]) if len(sys.argv) > 2 else 300
FN = ["SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA", "PRODUCT", "MEDIAN", "STDEV", "AND", "OR",
      "SUMIF({r},\">0\")", "COUNTIF({r},\">0\")", "AVERAGEIF({r},\">0\")", "SUMPRODUCT", "MATCH(3,{r},0)", "MATCH(3,{r})",
      "VLOOKUP(3,{r},1,0)", "INDEX({r},2,1)"]
W = {}
for w in range(N):
    s = {}
    h, wd = random.randint(2, 4), random.randint(1, 3)
    cells = [f"{c}{r}" for c in "BCD"[:wd] for r in range(1, h + 1)]
    for a in cells:
        k = random.random()
        if k < 0.3: s[a] = random.choice([1, 2, 3, 5, "x", True])
        elif k < 0.5: s[a] = random.choice(["=1/0", "=NA()", "=SQRT(-1)", "=\"a\"+1"])
        elif k < 0.75: s[a] = random.choice(["=A1+1", "=A1*2", "=ISERROR(A1)", "=IF(A1>2,1,2)", "=A1&\"\""])
    f = random.choice(FN)
    rg = f"B1:{'BCD'[wd - 1]}{h}"
    if random.random() < 0.3: rg = f"A1:{'BCD'[wd - 1]}{h}"
    call = f.format(r=rg) if "{r}" in f else f"{f}({rg})"
    s["A1"] = "=" + random.choice([call, f"IFERROR({call},0)", f"IFERROR({call},0)", f"ISERROR({call})"])
    s["F1"] = "=A1"
    W[f"cr{w}"] = {"sheets": {"Sheet1": s}}
print(json.dumps(W))
