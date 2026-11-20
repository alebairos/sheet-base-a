# Everyday money arithmetic through the rounding functions: prices, quantities,
# tax and discount rates, averages and shares, rounded to cents.
# Run: python3 tests/src/money.py 1 300 > tests/cases/money.json
import json, random, sys
random.seed(int(sys.argv[1]) if len(sys.argv) > 1 else 1)
N = int(sys.argv[2]) if len(sys.argv) > 2 else 300
W = {}
rates = [0.07, 0.0825, 0.15, 0.2, 0.05, 0.125, 0.0625, 0.19, 0.035, 0.1]
for w in range(N):
    s = {}
    n = random.randint(3, 10)
    for i in range(1, n + 1):
        s[f"A{i}"] = round(random.choice([random.uniform(0.5, 100), random.uniform(100, 5000), random.randint(1, 99) + random.choice([0.05, 0.15, 0.25, 0.35, 0.45, 0.5, 0.55, 0.65, 0.75, 0.85, 0.95, 0.995, 0.005])]), random.choice([2, 2, 2, 3]))
        s[f"B{i}"] = random.choice([1, 2, 3, 4, 5, 6, 7, 10, 12, 0.5, 1.5, 0.25])
        s[f"C{i}"] = f"=A{i}*B{i}"
    t, d = random.choice(rates), random.choice(rates)
    s["E1"] = t; s["E2"] = d
    fs = [f"=ROUND(SUM(C1:C{n})*(1+E1),2)", f"=ROUND(AVERAGE(A1:A{n}),2)", f"=ROUND(SUMPRODUCT(A1:A{n},B1:B{n})*(1-E2),2)",
          f"=ROUND(A1*(1+E1),2)", f"=ROUND(A2*E1,2)", f"=ROUNDUP(A3*E2,2)", f"=ROUNDDOWN(SUM(C1:C{n})*E1,2)", f"=INT(A1*B1*100)/100",
          f"=ROUND(SUM(C1:C{n})/COUNT(A1:A{n}),2)", f"=ROUND(A1/3,2)*3", f"=ROUND(MEDIAN(A1:A{n}),1)", f"=MOD(ROUND(SUM(C1:C{n}),2),1)",
          f"=ROUND(SUM(C1:C{n})*E1,2)+ROUND(SUM(C1:C{n}),2)", f"=ROUND(A1*(1+E1)*(1-E2),2)", f"=ROUND(STDEV(A1:A{n}),2)", f"=ROUND(A1,-1)",
          f"=ROUND(SUMIF(B1:B{n},\">2\",C1:C{n}),2)", f"=ROUND(AVERAGEIF(B1:B{n},\"<5\",A1:A{n}),2)", f"=ROUND(C1/SUM(C1:C{n})*100,1)", f"=ROUND(A1*1.1,2)=ROUND(A1+A1*0.1,2)",
          f"=ROUND(A1,2)=A1", f"=ROUND(SUM(C1:C{n}),2)-SUM(C1:C{n})", f"=TRUNC(A2*E1*100)", f"=ROUND(A1*B1,0)", f"=ROUND(PRODUCT(A1,B1,1+E1),2)"]
    for j, f in enumerate(fs):
        s[f"F{j+1}"] = f
    W[f"m{w}"] = {"sheets": {"Sheet1": s}}
print(json.dumps(W))
