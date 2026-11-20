# Approximate (sorted) lookups with wildcard patterns over unsorted columns and
# rows mixing text, numbers, booleans, blanks and errors.
# Run: python3 tests/src/wildlook.py 1 300 > tests/cases/wildlook.json
import json, random, sys
random.seed(int(sys.argv[1]) if len(sys.argv) > 1 else 1)
N = int(sys.argv[2]) if len(sys.argv) > 2 else 300
vals = ["a", "b", "ab", "ba", "apple", "banana", "Bob", "x", "zz", "", 1, 2, 10, -3, True, False, None, None, "=NA()"]
qs = ['"*"', '"?"', '"a*"', '"*a"', '"b?"', '"?b"', '"*p*"', '"??"', '"z*"', '"~*"', '"a?*"']
W = {}
for w in range(N):
    n = random.randint(2, 8)
    s = {}
    horiz = random.random() < 0.3
    for i in range(n):
        v = random.choice(vals if random.random() < 0.3 else vals[:9])
        a = f"{chr(65 + i)}1" if horiz else f"A{i + 1}"
        if v is not None and v != "": s[a] = v
        elif v == "": s[a] = '=""'
    last = f"{chr(64 + n)}1" if horiz else f"A{n}"
    for j in range(5):
        q = random.choice(qs)
        f = random.choice([f"=MATCH({q},A1:{last},1)", f"=MATCH({q},A1:{last},-1)", f"=MATCH({q},A1:{last})"] + ([f"=HLOOKUP({q},A1:{last},1)"] if horiz else [f"=VLOOKUP({q},A1:{last},1)"]))
        s[f"K{j + 1}"] = f
    W[f"w{w}"] = {"sheets": {"Sheet1": s}}
print(json.dumps(W))
