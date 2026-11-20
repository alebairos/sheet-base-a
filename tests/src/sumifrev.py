# SUMIF/AVERAGEIF with sum ranges written reversed (bottom-up, right-to-left)
# against criteria ranges of random shapes.
# Run: python3 tests/src/sumifrev.py 1 60 > tests/cases/sumifrev.json
import json, random, sys
random.seed(int(sys.argv[1]) if len(sys.argv) > 1 else 1)
N = int(sys.argv[2]) if len(sys.argv) > 2 else 60
C = "ABCDEFGHIJKL"
W = {}
for w in range(N):
    s = {}
    for r in range(1, 16):
        for c in "FGHIJKL":
            s[f"{c}{r}"] = random.choice([1, 2, 3, 5, 8, 13, 21, 34, 55, 89, "x", True])
    for r in range(1, 7):
        for c in "ABC":
            s[f"{c}{r}"] = random.choice([1, 1, 2, 3])
    for j in range(1, 11):
        h, wd = random.randint(1, 6), random.randint(1, 3)
        r0, c0 = random.randint(1, 7 - h), random.randint(0, 3 - wd)
        crit = f"{C[c0]}{r0}:{C[c0 + wd - 1]}{r0 + h - 1}"
        x, y = random.randint(1, 15), random.randint(1, 15)
        cx, cy = random.randint(5, 11), random.randint(5, 11)
        if random.random() < 0.5: cy = cx
        sr = f"{C[cx]}{x}:{C[cy]}{y}"
        f = random.choice(["SUMIF", "SUMIF", "AVERAGEIF"])
        s[f"N{j}"] = f"=IFERROR({f}({crit},{random.choice(['1', '\">1\"', '2'])},{sr}),-1)"
    W[f"r{w}"] = {"sheets": {"Sheet1": s}}
print(json.dumps(W))
