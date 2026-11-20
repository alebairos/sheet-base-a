# Dev-only: mutate formulas from tests/src/*.txt (numbers, references, wrapping functions).
# Usage: python3 tools/mutate.py <seed> <count> > tests/src/mutateN.txt
import random, re, sys, glob
random.seed(int(sys.argv[1]))
N = int(sys.argv[2])
src = []
for f in glob.glob("tests/src/*.txt"):
    for l in open(f, encoding="utf-8"):
        l = l.rstrip("\n")
        if l.startswith("=") and 6 < len(l) < 90 and "Sheet" not in l: src.append(l)
src = sorted(set(src))
nums = ["0", "1", "-1", "2", "0.5", "3", "10", "1E-3", "TRUE", "\"\"", "\"a\"", "\"3\"", "NA()", "1/0", "A12", "B11", "A6", "{1,2}", "{1;2;3}", "A1:A3", "B1:B3"]
cells = ["A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11", "A12", "B1", "B4", "B5", "B6", "B7", "B11", "B12", "A1:A3", "A1:B3", "A4:A6", "B1:B12", "A10:B12"]
fns1 = ["SUM", "ABS", "LEN", "ISNUMBER", "ISERROR", "NOT", "UPPER", "INT", "COUNT", "N", "ISBLANK", "TRIM"]
fns2 = ["ROUND", "MOD", "IFERROR", "LEFT", "POWER", "MAX", "AND", "CONCATENATE", "MATCH", "COUNTIF"]
out = set()
while len(out) < N:
    f = random.choice(src)
    for _ in range(random.randint(1, 2)):
        k = random.random()
        if k < 0.35:
            toks = list(re.finditer(r"(?<![A-Za-z$])-?\d+(\.\d+)?(?![\d:A-Za-z])", f))
            if toks:
                m = random.choice(toks); f = f[:m.start()] + random.choice(nums) + f[m.end():]
        elif k < 0.65:
            toks = list(re.finditer(r"\$?[A-B]\$?\d+(:\$?[A-B]\$?\d+)?", f))
            if toks:
                m = random.choice(toks); f = f[:m.start()] + random.choice(cells) + f[m.end():]
        elif k < 0.8:
            f = "=" + random.choice(fns1) + "(" + f[1:] + ")"
        elif k < 0.9:
            g = random.choice(fns2); a = random.choice(nums)
            f = "=" + g + "(" + (f[1:] + "," + a if random.random() < 0.5 else a + "," + f[1:]) + ")"
        else:
            f = f + random.choice(["+1", "&\"x\"", "*{1,2}", "=0", "<\"b\"", "^2", "%", "-A12"])
    out.add(f)
print("\n".join(sorted(out)))
