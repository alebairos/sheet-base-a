# Random "sales report / pricing tiers / contact list / stats" workbooks with everyday formulas
# (input for tests/cases/realistic4.json).
# Run: python3 tests/src/realistic4.py 4242 300 > tests/cases/realistic4.json; then tools/lo_oracle.py for the expected values
import json, random, sys
random.seed(int(sys.argv[1]) if len(sys.argv) > 1 else 4242)
regions = ["North", "south", "East", "west", " North", "North ", "EAST"]
prods = ["Apple", "apple pie", "Banana", "Cherry", "Date", "Elder*", "Fig?", "grape"]
people = ["ana silva", "BRUNO costa", " carla  dias ", "Diego", "eva-maria", "Fábio Ño", "", "x"]
W = {}


def num():
    return random.choice([random.randint(-50, 500), round(random.uniform(0, 1000), 2), round(random.uniform(0, 1), 3),
                          random.randint(1, 9) * 10 ** random.randint(0, 6), 0, "", "n/a", str(random.randint(1, 99)), True])


def sales(w):
    n = random.randint(4, 14)
    s = {"A1": "Region", "B1": "Product", "C1": "Qty", "D1": "Price", "E1": "Total"}
    for i in range(2, n + 2):
        s[f"A{i}"] = random.choice(regions)
        s[f"B{i}"] = random.choice(prods)
        for c in "CD":
            v = num()
            if v != "":
                s[f"{c}{i}"] = v
        s[f"E{i}"] = random.choice([f"=C{i}*D{i}", f"=ROUND(C{i}*D{i},2)", f"=IFERROR(C{i}*D{i},0)", f"=IF(ISNUMBER(C{i}),C{i}*D{i},\"\")",
                                    f"=PRODUCT(C{i}:D{i})", f"=C{i}*D{i}*(1-10%)", f"=IF(C{i}>100,C{i}*D{i}*0.9,C{i}*D{i})"])
    last = n + 1
    r = f"A2:A{last}"
    crits = ['"North"', '"north"', '"*th"', '" North"', '"<>North"', '"N*"', '"?ast"', '"east"', 'A2', '"<>"&A3', '">M"', '"="&B2',
             '"Apple"', '"apple*"', '"Elder~*"', '"Fig~?"', '"*"', '"<>"', '""', '">=b"']
    f = []
    for k in range(random.randint(6, 12)):
        col = random.choice(["A", "B"])
        rr = f"{col}2:{col}{last}"
        crit = random.choice(crits)
        f.append(random.choice([
            f"=SUMIF({rr},{crit},E2:E{last})", f"=COUNTIF({rr},{crit})", f"=AVERAGEIF({rr},{crit},C2:C{last})",
            f"=IFERROR(AVERAGEIF({rr},{crit},E2:E{last}),\"-\")", f"=SUMIF(C2:C{last},\">\"&{random.randint(0, 200)})",
            f"=COUNTIF(D2:D{last},\"<=\"&D{random.randint(2, last)})", f"=SUM(E2:E{last})", f"=AVERAGE(E2:E{last})",
            f"=MAX(C2:C{last})-MIN(C2:C{last})", f"=COUNT(C2:D{last})", f"=COUNTA(A1:E{last})", f"=MEDIAN(D2:D{last})",
            f"=ROUND(STDEV(C2:C{last}),2)", f"=SUMIF({rr},{crit},E2:E{last})/SUM(E2:E{last})",
            f"=INDEX(B2:B{last},MATCH(MAX(E2:E{last}),E2:E{last},0))", f"=VLOOKUP({crit},{col}2:E{last},{random.randint(2, 4)},FALSE)",
            f"=CONCATENATE(\"Total: \",ROUND(SUM(E2:E{last}),1))", f"=\"Avg \"&AVERAGE(C2:C{last})",
            f"=SUM(E2:E{last})*{random.choice(['5%', '0.5%', '150%', '1/3'])}"]))
    for k, x in enumerate(f):
        s[f"G{k + 1}"] = x
    w["Sales"] = s


def tiers(w):
    t = {"A1": 0, "B1": "Bronze", "C1": 0, "A2": 100, "B2": "Silver", "C2": 0.05, "A3": 500, "B3": "Gold", "C3": 0.1,
         "A4": 1000, "B4": "Platinum", "C4": 0.15}
    if random.random() < 0.3:
        t["A5"] = 5000
        t["B5"] = "Diamond"
        t["C5"] = "20%"
    w["Tiers"] = t
    s = {"A1": "Customer", "B1": "Spend", "C1": "Tier", "D1": "Disc", "E1": "Net"}
    n = random.randint(3, 10)
    for i in range(2, n + 2):
        s[f"A{i}"] = random.choice(people)
        v = random.choice([random.randint(-10, 6000), round(random.uniform(0, 2000), 2), 100, 500, 999.999, 1000, "", "abc", "750"])
        if v != "":
            s[f"B{i}"] = v
        s[f"C{i}"] = random.choice([f"=VLOOKUP(B{i},Tiers!$A$1:$C$5,2)", f"=VLOOKUP(B{i},Tiers!$A$1:$C$5,2,TRUE)",
                                    f"=INDEX(Tiers!$B$1:$B$5,MATCH(B{i},Tiers!$A$1:$A$5,1))",
                                    f"=IF(B{i}>=1000,\"Platinum\",IF(B{i}>=500,\"Gold\",IF(B{i}>=100,\"Silver\",\"Bronze\")))",
                                    f"=IFERROR(VLOOKUP(B{i},Tiers!A1:C5,2),\"none\")"])
        s[f"D{i}"] = random.choice([f"=VLOOKUP(B{i},Tiers!$A$1:$C$5,3)", f"=IFERROR(INDEX(Tiers!C1:C5,MATCH(C{i},Tiers!B1:B5,0)),0)",
                                    f"=VLOOKUP(C{i},Tiers!$B$1:$C$5,2,FALSE)"])
        s[f"E{i}"] = random.choice([f"=B{i}*(1-D{i})", f"=ROUND(B{i}*(1-D{i}),2)", f"=ROUNDDOWN(B{i}-B{i}*D{i},0)", f"=IFERROR(B{i}*(1-D{i}),B{i})"])
    last = n + 1
    s["G1"] = f"=SUM(E2:E{last})"
    s["G2"] = f"=COUNTIF(C2:C{last},\"Gold\")"
    s["G3"] = f"=SUMIF(C2:C{last},\"<>Bronze\",E2:E{last})"
    s["G4"] = f"=MATCH(\"Gold\",C2:C{last},0)"
    s["G5"] = f"=AND(E2:E{last}>0)"
    s["G6"] = f"=OR(C2:C{last}=\"Platinum\")"
    s["G7"] = f"=AVERAGEIF(B2:B{last},\">=500\")"
    w["Customers"] = s


def contacts(w):
    s = {"A1": "Name", "B1": "Clean", "C1": "Initials", "D1": "Len", "E1": "Code"}
    n = random.randint(3, 10)
    for i in range(2, n + 2):
        v = random.choice(people + [random.randint(1, 999), 3.5, True])
        if v != "":
            s[f"A{i}"] = v
        s[f"B{i}"] = random.choice([f"=TRIM(A{i})", f"=UPPER(TRIM(A{i}))", f"=LOWER(A{i})", f"=TRIM(LOWER(A{i}))&\"!\""])
        s[f"C{i}"] = random.choice([f"=UPPER(LEFT(B{i},1))&\".\"", f"=LEFT(B{i},1)&MID(B{i},FIND(\" \",B{i})+1,1)" if False else f"=UPPER(LEFT(B{i},2))",
                                    f"=CONCATENATE(LEFT(B{i}),RIGHT(B{i}))", f"=MID(B{i},2,3)", f"=RIGHT(B{i},LEN(B{i})-1)"])
        s[f"D{i}"] = random.choice([f"=LEN(A{i})", f"=LEN(B{i})", f"=LEN(A{i})-LEN(TRIM(A{i}))", f"=IF(LEN(B{i})>5,\"long\",\"short\")"])
        s[f"E{i}"] = random.choice([f"=UPPER(LEFT(B{i},3))&\"-\"&ROW(A{i})" if False else f"=UPPER(LEFT(B{i},3))&\"-\"&{i}",
                                    f"=CONCATENATE(D{i},\"/\",LEN(C{i}))", f"=IF(ISTEXT(A{i}),\"txt\",IF(ISNUMBER(A{i}),\"num\",\"other\"))",
                                    f"=IF(ISBLANK(A{i}),\"(blank)\",A{i}&\"\")", f"=MID(A{i},{random.randint(0, 4)},{random.randint(0, 5)})"])
    last = n + 1
    s["G1"] = f"=COUNTIF(B2:B{last},\"*a*\")"
    s["G2"] = f"=COUNTIF(D2:D{last},\">3\")"
    s["G3"] = f"=MATCH(\"*o*\",B2:B{last},0)"
    s["G4"] = f"=COUNTA(A2:A{last})-COUNT(A2:A{last})"
    s["G5"] = f"=SUM(D2:D{last})/COUNT(D2:D{last})"
    s["G6"] = f"=INDEX(B2:B{last},COUNTA(B2:B{last}))"
    w["Contacts"] = s


def stats(w):
    s = {}
    n = random.randint(4, 15)
    for i in range(1, n + 1):
        v = random.choice([random.randint(-100, 100), round(random.gauss(50, 20), 4), 0.1 * random.randint(1, 30), 1e-9 * random.randint(1, 9), 1e9 + random.randint(0, 9)])
        s[f"A{i}"] = v
        if random.random() < 0.6:
            s[f"B{i}"] = random.choice([random.randint(1, 20), round(random.uniform(-5, 5), 2), "x", True, False])
    r = f"A1:A{n}"
    fs = [f"=AVERAGE({r})", f"=MEDIAN({r})", f"=STDEV({r})", f"=MAX({r})", f"=MIN({r})", f"=SUM({r})", f"=PRODUCT(B1:B{n})",
          f"=ROUND(AVERAGE({r}),{random.randint(-1, 4)})", f"=ROUNDUP(STDEV({r}),{random.randint(0, 3)})", f"=ROUNDDOWN(MEDIAN({r}),1)",
          f"=MOD(SUM({r}),{random.choice([7, 0.5, -3, 2.5])})", f"=INT(AVERAGE({r}))", f"=ABS(MIN({r}))", f"=SQRT(ABS(SUM({r})))",
          f"=POWER(MAX({r}),{random.choice([2, 0.5, -1, 3])})", f"=(A1-AVERAGE({r}))/STDEV({r})", f"=COUNTIF({r},\">\"&AVERAGE({r}))",
          f"=SUMIF({r},\"<0\")", f"=AVERAGEIF(B1:B{n},\">0\",{r})", f"=SUM(A1:B{n})", f"=AVERAGE(A1:B{n})", f"=MEDIAN(A1:B{n})",
          f"=STDEV(A1:B{n})", f"=COUNT(A1:B{n})", f"=MAX(B1:B{n})", f"=SUM({r})/COUNT({r})-AVERAGE({r})", f"=STDEV(B1:B{n})",
          f"=INDEX({r},MATCH(MAX({r}),{r},0))", f"=MATCH(MIN({r}),{r},0)", f"=IF(AVERAGE({r})>MEDIAN({r}),\"right skew\",\"left skew\")",
          f"=ROUND(SUM({r})/{n},2)=ROUND(AVERAGE({r}),2)", f"=A1^2", f"=-A1^2", f"=2^-A2", f"=A1%", f"=SUM({r})&\"\"", f"=AVERAGE({r})&\" avg\""]
    for k, x in enumerate(random.sample(fs, random.randint(8, 16))):
        s[f"D{k + 1}"] = x
    w["Stats"] = s


for k in range(int(sys.argv[2]) if len(sys.argv) > 2 else 200):
    w = {}
    for g in random.sample([sales, tiers, contacts, stats], random.randint(1, 3)):
        g(w)
    W[f"r4_{k}"] = {"sheets": w}
json.dump(W, sys.stdout, indent=1, ensure_ascii=False)
print()
