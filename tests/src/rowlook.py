# Approximate lookups along unsorted cell rows mixing numbers, text, booleans,
# blanks and errors (MATCH 1/-1, HLOOKUP sorted), text and numeric queries.
# Run: python3 tests/src/rowlook.py 1 400 > tests/cases/rowlook.json (rowlookb: seed 77)
import json, random, sys
random.seed(int(sys.argv[1]) if len(sys.argv) > 1 else 1)
N = int(sys.argv[2]) if len(sys.argv) > 2 else 300
vals = [-3, 0, 1, 2, 2.5, 5, 10, "a", "b", "c", "m", "x", "apple", "Banana", True, False, None, None, "=NA()", "=1/0"]
qs = ['"a"', '"b"', '"c"', '"m"', '"x"', '"zz"', '""', '0', '1', '2', '4', '11', '-5', 'TRUE']
W = {}
for w in range(N):
    n = random.randint(2, 8)
    s = {}
    for i in range(n):
        v = random.choice(vals[:16] if random.random() < 0.8 else vals)
        if v is not None: s[f"{chr(65 + i)}1"] = v
        s[f"{chr(65 + i)}2"] = i + 1
    last = chr(64 + n)
    for j in range(6):
        q = random.choice(qs)
        f = random.choice([f"=MATCH({q},A1:{last}1,1)", f"=MATCH({q},A1:{last}1,-1)", f"=MATCH({q},A1:{last}1)", f"=HLOOKUP({q},A1:{last}2,2)", f"=HLOOKUP({q},A1:{last}2,2,TRUE)"])
        s[f"A{5 + j}"] = f
    W[f"r{w}"] = {"sheets": {"Sheet1": s}}
print(json.dumps(W))
