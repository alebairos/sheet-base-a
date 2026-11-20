# Array formulas over ranges that run past a sheet's data (blank elements count).
# Run: python3 tests/src/pastdata.py <seed> <count> > tests/cases/pastdata.json; then tools/lo_oracle.py
import json, random, sys
random.seed(int(sys.argv[1]) if len(sys.argv) > 1 else 1)
N = int(sys.argv[2]) if len(sys.argv) > 2 else 200
vals = [1, 2, 0, -3, 2.5, "a", "b", "", "x y", True, False, "5", 10, 7]
W = {}
for w in range(N):
    s1 = {}
    nr, nc = random.randint(2, 6), random.randint(1, 3)
    for r in range(1, nr + 1):
        for c in "ABC"[:nc]:
            if random.random() < 0.85: s1[f"{c}{r}"] = random.choice(vals)
    s2 = {"A1": random.choice(vals), "A2": random.choice(vals)}
    for k in range(4):
        c = random.choice("ABC"); r1 = random.randint(1, 4); r2 = r1 + random.randint(1, 12)
        rg = f"{c}{r1}:{c}{r2}"
        rg2 = f"{random.choice('ABC')}{r1}:{random.choice('ABC')}{r2}"
        sh = random.choice(["", "", "Sheet2!"])
        rg = sh + rg
        f = random.choice([
            f'=SUMPRODUCT(--({rg}=""))', f'=SUMPRODUCT(--ISBLANK({rg}))', f'=SUMPRODUCT(--({rg}<1))', f'=SUMPRODUCT(--({rg}=0))',
            f'=SUMPRODUCT(--NOT({rg}))', f'=SUMPRODUCT(LEN({rg}&"x"))', f'=SUMPRODUCT(--ISTEXT({rg}&""))', f'=SUMPRODUCT(--({rg}<>"a"))',
            f'=SUMPRODUCT(({rg}="")*1)', f'=SUMPRODUCT(--({rg}={rg2}))', f'=SUMPRODUCT(--ISBLANK({rg2}))', f'=MATCH(TRUE,INDEX({rg}="",0),0)',
            f'=MATCH(0,INDEX(--({rg}<>""),0),0)', f'=SUMPRODUCT(--(LEN({rg})=0))', f'=SUMPRODUCT(--(TRIM({rg})=""))', f'=SUMPRODUCT(N(+{rg}))',
            f'=COUNT(1/({rg}=""))', f'=SUM(--({rg}=""))', f'=MAX(--({rg}=""))', f'=SUMPRODUCT(--ISNUMBER({rg}))', f'=SUMPRODUCT(--({rg}>=0))',
            f'=SUMPRODUCT(--(UPPER({rg})=""))', f'=SUMPRODUCT(--ISERROR(1/{rg}))', f'=MATCH(1,{rg},1)', f'=MATCH("zz",{rg},1)'])
        s1[f"E{k + 1}"] = f
    W[f"pd{w}"] = {"sheets": {"Sheet1": s1, "Sheet2": s2}}
json.dump(W, sys.stdout)
