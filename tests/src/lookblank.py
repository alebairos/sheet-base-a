# Sorted columns with blank cells, text and duplicates: approximate VLOOKUP/MATCH/HLOOKUP over cells.
# Run: python3 tests/src/lookblank.py <seed> <count> > tests/cases/lookblank.json
import json, random, sys
random.seed(int(sys.argv[1]) if len(sys.argv) > 1 else 1)
N = int(sys.argv[2]) if len(sys.argv) > 2 else 150
W = {}
for w in range(N):
    s = {}
    n = random.randint(3, 12)
    vals = sorted(random.sample(range(0, 100), n))
    if random.random() < 0.5: vals = [v / 4 for v in vals]
    desc = random.random() < 0.3
    if desc: vals = vals[::-1]
    r = 1
    for v in vals:
        k = random.random()
        if k < 0.2: r += 1                      # a blank cell
        elif k < 0.27: s[f"A{r}"] = random.choice(["x", "", "n/a"]); r += 1
        s[f"A{r}"] = v; s[f"B{r}"] = f"v{r}"; r += 1
    if random.random() < 0.3: s[f"A{r}"] = "text"; s[f"B{r}"] = "t"; r += 1
    for j in range(8):
        q = random.choice([random.randint(-5, 105), random.randint(0, 100) / 4, '""', '"x"', "TRUE", "C20"])
        mt = -1 if desc else 1
        s[f"D{j+1}"] = f"=VLOOKUP({q},A1:B{r},2,TRUE)"
        s[f"E{j+1}"] = f"=MATCH({q},A1:A{r},{mt})"
        s[f"F{j+1}"] = f"=MATCH({q},A1:A{r+3},{mt})"
    W[f"lb{w}"] = {"sheets": {"Sheet1": s}}
print(json.dumps(W))
