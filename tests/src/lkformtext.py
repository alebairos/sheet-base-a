# Approximate lookups over columns mixing constant text, text from formulas,
# numbers, formula numbers, blanks and errors (Calc skips only constant text
# at the start of a numeric sorted search).
# Run: python3 tests/src/lkformtext.py <seed> <count> > tests/cases/lkformtext.json; then tools/lo_oracle.py
import json, random, sys
random.seed(int(sys.argv[1]) if len(sys.argv) > 1 else 1)
N = int(sys.argv[2]) if len(sys.argv) > 2 else 200
def cell():
    k = random.random()
    if k < 0.2: return None
    if k < 0.35: return random.choice(["a", "m", "x", "3", "zz"])
    if k < 0.5: return random.choice(['="a"', '="m"', '="x"', '="3"', '=""', '="zz"'])
    if k < 0.8: return random.choice([1, 2, 3, 5, 8, 0.5, -1, 13])
    if k < 0.92: return random.choice(["=2", "=5", "=0.5", "=TRUE", "=1+1"])
    return random.choice(["=1/0", "=NA()"])
W = {}
for t in range(N):
    n = random.randint(2, 9)
    d = {}
    # mostly ascending numbers with a few odd cells, to look like real data
    for r in range(1, n + 1):
        v = cell()
        if v is not None: d[f"A{r}"] = v
        d[f"B{r}"] = r * 10
    qs = [random.choice([0, 0.7, 2, 4, 6, 9, 20, '"m"', '"b"', '"zz"', '"3"', 'TRUE'])]
    for i, q in enumerate(qs):
        d["D1"] = f"=VLOOKUP({q},A1:B{n},2,1)"
        d["D2"] = f"=MATCH({q},A1:A{n},1)"
        d["D3"] = f"=MATCH({q},A1:A{n},-1)"
        d["D4"] = f"=VLOOKUP({q},A1:B{n},2,0)"
    W[f"lft{t}"] = {"sheets": {"Sheet1": d}}
json.dump(W, sys.stdout, indent=0)
