# Random columns of awkward decimals split by blanks, text, booleans and formula cells:
# SUM's last-bit rounding (Calc sums runs of number cells as blocks; formula cells one by one).
# Run: python3 tests/src/sumblocks.py <seed> <count> > tests/cases/sumblocks.json
import json, random, sys
random.seed(int(sys.argv[1]) if len(sys.argv) > 1 else 1)
N = int(sys.argv[2]) if len(sys.argv) > 2 else 200
W = {}
def dec():
    return random.choice([round(random.uniform(-100, 100), 2), round(random.uniform(0, 1000), 3), random.randint(1, 60), 0.1, 0.2, 0.3, 16.99, 33.12, -3.21, 1e-9 * random.randint(1, 9)])
for w in range(N):
    s = {}
    for col in "AB":
        n = random.randint(3, 20)
        for i in range(1, n + 1):
            k = random.random()
            if k < 0.6: s[f"{col}{i}"] = dec()
            elif k < 0.7: pass
            elif k < 0.77: s[f"{col}{i}"] = random.choice(["n/a", "x", "5"])
            elif k < 0.85: s[f"{col}{i}"] = random.choice([True, False])
            else: s[f"{col}{i}"] = random.choice([f"={dec()}", f"={dec()}+{dec()}", "=1/3", "=0.1*3"])
    fs = ["=SUM(A1:A20)", "=MOD(SUM(A1:A20),7)", "=SUM(A1:A20)-ROUND(SUM(A1:A20),2)", "=SUM(A1:B20)", "=SUM(A1:A20,B1:B20)",
          "=SUM(B1:B20)*1E6-ROUND(SUM(B1:B20)*1E6,0)", "=SUM(A1:A20)-SUM(B1:B20)", "=SUM(A1:A10)+SUM(A11:A20)-SUM(A1:A20)"]
    for j, f in enumerate(fs):
        s[f"D{j+1}"] = f
    W[f"sb{w}"] = {"sheets": {"Sheet1": s}}
print(json.dumps(W))
