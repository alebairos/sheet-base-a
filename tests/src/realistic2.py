# Random small "order table" workbooks with typical formulas (input for tests/cases/realistic2.json).
# Run: python3 tests/src/realistic2.py; then tools/lo_oracle.py for the expected values
import json, random
random.seed(2027)
names = ["alice","Bob","carol","Dave","eve","Frank","grace","Heidi","ivan","Judy","mallory","Oscar"]
cats = ["fruit","veg","Dairy","meat","fruit ","bakery"]
W = {}
for w in range(120):
    n = random.randint(3, 10)
    s = {"A1":"Name","B1":"Cat","C1":"Qty","D1":"Price","E1":"Total"}
    for i in range(2, n+2):
        s[f"A{i}"] = random.choice(names)
        s[f"B{i}"] = random.choice(cats)
        q = random.choice([random.randint(0,50), round(random.uniform(-5,50),2), "", "n/a", True, str(random.randint(1,9))])
        if q != "": s[f"C{i}"] = q
        s[f"D{i}"] = random.choice([round(random.uniform(0,20),2), random.randint(1,100), 0.1, 1/3])
        s[f"E{i}"] = random.choice([f"=C{i}*D{i}", f"=ROUND(C{i}*D{i},2)", f"=IF(ISNUMBER(C{i}),C{i}*D{i},0)", f"=IFERROR(C{i}*D{i},\"bad\")"])
    L = n+1
    fs = [f"=SUM(E2:E{L})", f"=AVERAGE(D2:D{L})", f"=ROUND(AVERAGE(E2:E{L}),2)", f"=COUNT(C2:C{L})", f"=COUNTA(A2:A{L})",
          f"=COUNTIF(B2:B{L},\"fruit\")", f"=COUNTIF(B2:B{L},\"fruit*\")", f"=SUMIF(B2:B{L},\"fruit\",E2:E{L})", f"=AVERAGEIF(C2:C{L},\">10\")",
          f"=SUMIF(C2:C{L},\">=\"&D2,D2:D{L})", f"=MAX(E2:E{L})-MIN(E2:E{L})", f"=MEDIAN(C2:C{L})", f"=STDEV(D2:D{L})",
          f"=VLOOKUP(\"{random.choice(names)}\",A2:E{L},5,FALSE)", f"=IFERROR(VLOOKUP(\"{random.choice(names)}\",A2:E{L},4,0),\"none\")",
          f"=INDEX(A2:A{L},MATCH(MAX(D2:D{L}),D2:D{L},0))", f"=MATCH(\"{random.choice(cats)}\",B2:B{L},0)",
          f"=CONCATENATE(UPPER(LEFT(A2,1)),LOWER(MID(A2,2,20)),\" bought \",C2,\" @ \",D2)", f"=A2&\": \"&TEXT(E2,\"0.00\")",
          f"=IF(SUM(E2:E{L})>100,\"high\",IF(SUM(E2:E{L})>50,\"mid\",\"low\"))", f"=AND(COUNT(C2:C{L})>2,MIN(D2:D{L})>0)",
          f"=OR(ISBLANK(C2),ISTEXT(C2))", f"=PRODUCT(D2:D4)", f"=SQRT(SUM(D2:D{L}))", f"=MOD(SUM(C2:C{L}),7)", f"=INT(AVERAGE(D2:D{L}))",
          f"=ROUNDUP(SUM(D2:D{L})/3,1)", f"=ROUNDDOWN(SUM(D2:D{L})/3,-1)", f"=ABS(MIN(C2:C{L}))", f"=POWER(COUNT(C2:C{L}),2)",
          f"=LEN(TRIM(\"  \"&A2&\"   \"&B2))", f"=RIGHT(B2,3)&LEFT(A3,2)", f"=E2/C2", f"=SUM(E2:E{L})/COUNT(E2:E{L})",
          f"=NOT(ISERROR(E2))", f"=Summary!B1*2"]
    for j, f in enumerate(random.sample(fs, 14)):
        s[f"G{j+1}"] = f
    W[f"r{w}"] = {"sheets": {"Sheet1": s, "Summary": {"A1": "Total", "B1": f"=SUM(Sheet1!E2:E{L})", "B2": "=B1/Sheet1!G1"}}}
json.dump(W, open("tests/cases/realistic2.json", "w"), indent=0)
