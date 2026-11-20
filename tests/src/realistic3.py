# Random "gradebook / inventory / budget" workbooks with typical formulas (input for tests/cases/realistic3.json).
# Run: python3 tests/src/realistic3.py; then tools/lo_oracle.py for the expected values
import json, random
random.seed(31337)
names = ["Ana", "bruno", "Carla", "diego", "Eva", "Fabio", "gina", "Hugo", "Iris", "joao", "Kim", "Lia", "Ana"]
items = ["Widget", "gadget", "Bolt", "nut", "Screw", "washer", "Gear", "spring"]
W = {}

def gradebook(w):
    n = random.randint(3, 12)
    s = {"A1": "Student", "B1": "Test1", "C1": "Test2", "D1": "Final", "E1": "Avg", "F1": "Grade"}
    for i in range(2, n + 2):
        s[f"A{i}"] = random.choice(names)
        for c in "BCD":
            v = random.choice([random.randint(0, 100), round(random.uniform(0, 100), 1), "", "abs", True, str(random.randint(40, 99))])
            if v != "":
                s[f"{c}{i}"] = v
        s[f"E{i}"] = random.choice([f"=AVERAGE(B{i}:D{i})", f"=ROUND(AVERAGE(B{i}:D{i}),1)", f"=(B{i}+C{i}+2*D{i})/4",
                                    f"=IFERROR(AVERAGE(B{i}:D{i}),0)", f"=SUM(B{i}:D{i})/3"])
        s[f"F{i}"] = random.choice([f"=VLOOKUP(E{i},Scale!$A$1:$B$5,2)", f"=VLOOKUP(E{i},Scale!$A$1:$B$5,2,TRUE)",
                                    f"=IF(E{i}>=90,\"A\",IF(E{i}>=80,\"B\",IF(E{i}>=70,\"C\",IF(E{i}>=60,\"D\",\"F\"))))",
                                    f"=INDEX(Scale!$B$1:$B$5,MATCH(E{i},Scale!$A$1:$A$5,1))", f"=IFERROR(VLOOKUP(E{i},Scale!A1:B5,2),\"?\")"])
    L = n + 1
    who = random.choice(names)
    fs = [f"=AVERAGE(E2:E{L})", f"=MAX(E2:E{L})", f"=MIN(B2:D{L})", f"=COUNTIF(F2:F{L},\"A\")", f"=COUNTIF(E2:E{L},\">=\"&H1)",
          f"=COUNTIF(B2:B{L},\"<>\")", f"=COUNTIF(B2:B{L},\"\")", f"=COUNTIF(A2:A{L},\"{who[0]}*\")", f"=COUNTIF(A2:A{L},\"?\"&\"{who[1:]}\")",
          f"=SUMIF(A2:A{L},\"{who}\",E2:E{L})", f"=AVERAGEIF(A2:A{L},\"{who}\",D2:D{L})", f"=AVERAGEIF(E2:E{L},\"<60\")",
          f"=MEDIAN(E2:E{L})", f"=STDEV(E2:E{L})", f"=ROUND(STDEV(B2:B{L}),2)", f"=VLOOKUP(\"{who}\",A2:F{L},6,FALSE)",
          f"=INDEX(A2:A{L},MATCH(MAX(E2:E{L}),E2:E{L},0))", f"=MATCH(\"{who}\",A2:A{L},0)", f"=INDEX(A2:F{L},2,5)",
          f"=INDEX(A2:F{L},MATCH(\"{who}\",A2:A{L},0),MATCH(\"Final\",A1:F1,0))", f"=COUNT(B2:D{L})/COUNTA(A2:A{L})",
          f"=COUNTA(B2:D{L})-COUNT(B2:D{L})", f"=SUM(B2:B{L})-SUM(C2:C{L})", f"=IF(COUNTIF(F2:F{L},\"F\")>0,\"some fail\",\"all pass\")",
          f"=ROUNDUP(AVERAGE(D2:D{L}),0)", f"=ROUNDDOWN(AVERAGE(D2:D{L}),-1)", f"=CONCATENATE(A2,\" got \",F2)", f"=A2&\" (\"&ROUND(E2,0)&\")\"",
          f"=UPPER(A3)&\"-\"&LEN(A3)", f"=AND(E2>50,E3>50)", f"=OR(ISBLANK(B2),ISTEXT(B2),ISERROR(E2))", f"=NOT(ISNUMBER(D3))",
          f"=E2-E3", f"=(MAX(E2:E{L})-MIN(E2:E{L}))/2", f"=E2*10%", f"=-E2^2", f"=SQRT(ABS(E2-E3))", f"=MOD(COUNT(B2:D{L}),4)",
          f"=INT(E2/10)*10", f"=POWER(2,COUNTA(A2:A{L}))", f"=PRODUCT(H1,0.5)", f"=MID(A2,2,3)", f"=RIGHT(A2,2)&LEFT(A3,2)",
          f"=TRIM(\"  \"&A2&\"   \"&A3&\" \")", f"=LOWER(A2)=LOWER(A3)", f"=A2=A3", f"=A2<A3", f"=ISERROR(VLOOKUP(\"zed\",A2:F{L},2,0))",
          f"=IFERROR(INDEX(E2:E{L},{n + 3}),\"none\")", f"=SUM(H1:H2)", f"=H2+1", f"=Scale!C1", f"=SUM(Scale!A1:A5)/COUNT(E2:E{L})"]
    s["H1"] = random.choice([60, 70, 75.5, "70"])
    s["H2"] = random.choice(["=H3*2", "=H1/2", "=SUM(H1:H3)", "=H2+1", "=1"])
    for j, f in enumerate(random.sample(fs, 18)):
        s[f"I{j + 1}"] = f
    scale = {"A1": 0, "B1": "F", "A2": 60, "B2": "D", "A3": 70, "B3": "C", "A4": 80, "B4": "B", "A5": 90, "B5": "A",
             "C1": random.choice(["=COUNTA(Sheet1!A:A)", "=Sheet1!I1", "=MAX(A1:A5)", "=VLOOKUP(85,A1:B5,2)"])}
    return {"sheets": {"Sheet1": s, "Scale": scale}}

def inventory(w):
    n = random.randint(3, 10)
    s = {"A1": "Item", "B1": "Stock", "C1": "Cost", "D1": "Value", "E1": "Reorder"}
    for i in range(2, n + 2):
        s[f"A{i}"] = random.choice(items)
        b = random.choice([random.randint(0, 30), random.randint(-3, 3), "", "x", 0])
        if b != "":
            s[f"B{i}"] = b
        s[f"C{i}"] = random.choice([round(random.uniform(0.01, 9.99), 2), random.randint(1, 50), 0, "1.5"])
        s[f"D{i}"] = random.choice([f"=B{i}*C{i}", f"=ROUND(B{i}*C{i},1)", f"=IF(B{i}=\"\",0,B{i}*C{i})", f"=PRODUCT(B{i}:C{i})"])
        s[f"E{i}"] = random.choice([f"=IF(B{i}<5,\"yes\",\"no\")", f"=B{i}<$G$1", f"=IF(AND(ISNUMBER(B{i}),B{i}<$G$1),\"order\",\"\")",
                                    f"=IF(OR(B{i}<=0,B{i}=\"\"),\"out\",IF(B{i}<$G$1,\"low\",\"ok\"))"])
    L = n + 1
    it = random.choice(items)
    s["G1"] = random.choice([5, 10, 3])
    fs = [f"=SUM(D2:D{L})", f"=SUMIF(A2:A{L},\"{it}\",B2:B{L})", f"=SUMIF(B2:B{L},\"<\"&G1)", f"=SUMIF(E2:E{L},\"yes\",D2:D{L})",
          f"=COUNTIF(E2:E{L},TRUE)", f"=COUNTIF(E2:E{L},\"ok\")", f"=COUNTIF(B2:B{L},0)", f"=COUNTIF(B2:B{L},\">0\")", f"=COUNTIF(C2:C{L},\"1.5\")",
          f"=COUNTIF(A2:A{L},\"*{it[1:3]}*\")", f"=AVERAGEIF(A2:A{L},\"{it}\",C2:C{L})", f"=AVERAGEIF(B2:B{L},\">=\"&G1,D2:D{L})",
          f"=VLOOKUP(\"{it}\",A2:D{L},4,0)", f"=VLOOKUP(\"{it.upper()}\",A2:D{L},3,FALSE)", f"=VLOOKUP(\"{it[:2]}*\",A2:D{L},2,FALSE)",
          f"=INDEX(D2:D{L},MATCH(\"{it}\",A2:A{L},0))", f"=MATCH(MAX(D2:D{L}),D2:D{L},0)", f"=MATCH({random.randint(0, 30)},B2:B{L},0)",
          f"=INDEX(A2:E{L},{random.randint(1, n + 1)},{random.randint(1, 6)})", f"=INDEX(B2:B{L},0)", f"=MAX(D2:D{L})/SUM(D2:D{L})",
          f"=ROUND(SUM(D2:D{L})*1.2,2)", f"=ROUND(AVERAGE(C2:C{L}),2)", f"=MEDIAN(B2:B{L})", f"=STDEV(C2:C{L})", f"=MIN(B2:B{L})",
          f"=COUNT(A2:E{L})", f"=COUNTA(A2:E{L})", f"=IF(SUM(B2:B{L})>50,\"big\",\"small\")", f"=SUM(B2:B{L})*G1",
          f"=CONCATENATE(\"Total: \",SUM(D2:D{L}))", f"=\"Items: \"&COUNTA(A2:A{L})&\", value \"&ROUND(SUM(D2:D{L}),2)",
          f"=LEFT(A2,3)&\"-\"&ROW()", f"=LEN(A2)+LEN(A3)", f"=UPPER(LEFT(A2))&LOWER(MID(A2,2,99))", f"=SUM(D2:D{L})/0",
          f"=SQRT(SUM(B2:B{L}))", f"=MOD(SUM(B2:B{L}),G1)", f"=ABS(SUM(B2:B{L})-100)", f"=INT(-SUM(C2:C{L}))", f"=POWER(G1,0.5)",
          f"=ISBLANK(B2)", f"=ISTEXT(B2)", f"=ISNUMBER(C2)", f"=IFERROR(1/MIN(B2:B{L}),\"div\")", f"=Report!A1+1",
          f"=AND(B2:B{L})", f"=OR(B2:B{L}>20)", f"=SUM(B2:B{L}*C2:C{L})", f"=B2:B{L}", f"=SUM(Sheet1!B2:B{L} Sheet1!A3:E3)"]
    for j, f in enumerate(random.sample(fs, 18)):
        s[f"H{j + 1}"] = f
    rep = {"A1": f"=SUM(Sheet1!D2:D{L})", "A2": "=A1*Sheet1!G1", "A3": random.choice(["=A4", "=A3+1", "=Sheet1!H1", "=COUNT(Sheet1!B:B)"]),
           "A4": random.choice(["=A3", "=2", "=Sheet1!H2&\"\""])}
    return {"sheets": {"Sheet1": s, "Report": rep}}

def budget(w):
    n = random.randint(4, 12)
    s = {"A1": "Month", "B1": "Income", "C1": "Spend", "D1": "Net", "E1": "Running"}
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    for i in range(2, n + 2):
        s[f"A{i}"] = months[i - 2]
        s[f"B{i}"] = random.choice([random.randint(1000, 5000), round(random.uniform(1000, 5000), 2), 0])
        s[f"C{i}"] = random.choice([random.randint(500, 6000), round(random.uniform(0, 5000), 2), "=B%d*0.7" % i, "=B%d*%d%%" % (i, random.randint(10, 120))])
        s[f"D{i}"] = f"=B{i}-C{i}"
        s[f"E{i}"] = "=D2" if i == 2 else f"=E{i - 1}+D{i}"
    L = n + 1
    m = random.choice(months[:n])
    fs = [f"=SUM(D2:D{L})", f"=E{L}", f"=E{L}=SUM(D2:D{L})", f"=AVERAGE(B2:B{L})", f"=ROUND(AVERAGE(C2:C{L}),2)", f"=MAX(D2:D{L})",
          f"=MIN(E2:E{L})", f"=COUNTIF(D2:D{L},\"<0\")", f"=SUMIF(D2:D{L},\">0\")", f"=SUMIF(D2:D{L},\"<0\",C2:C{L})",
          f"=AVERAGEIF(D2:D{L},\">0\",B2:B{L})", f"=INDEX(A2:A{L},MATCH(MIN(D2:D{L}),D2:D{L},0))", f"=VLOOKUP(\"{m}\",A2:E{L},4,FALSE)",
          f"=MATCH(\"{m}\",A2:A{L},0)", f"=ROUND(SUM(C2:C{L})/SUM(B2:B{L})*100,1)&\"%\"", f"=IF(E{L}<0,\"deficit\",\"surplus\")",
          f"=MEDIAN(D2:D{L})", f"=ROUND(STDEV(D2:D{L}),2)", f"=ROUNDUP(SUM(C2:C{L})/12,-2)", f"=ROUNDDOWN(SUM(B2:B{L})/7,3)",
          f"=MOD(SUM(B2:B{L}),1)", f"=INT(SUM(C2:C{L})/100)", f"=ABS(MIN(D2:D{L}))", f"=SUM(B2:C{L})", f"=PRODUCT(1+D2/B2,1.05)",
          f"=POWER(1.05,COUNT(B2:B{L}))*B2", f"=SQRT(SUMPRODUCT(D2:D{L},D2:D{L}))", f"=CONCATENATE(A2,\"-\",A{L},\": \",ROUND(E{L},0))",
          f"=LEN(A2&A3)", f"=COUNTIF(A2:A{L},\"J*\")", f"=COUNTIF(A2:A{L},\"???\")", f"=COUNTIF(B2:C{L},\">=\"&B2)",
          f"=AND(D2>0,D3>0,D4>0)", f"=OR(D2:D{L}<0)", f"=NOT(E{L}>0)", f"=IFERROR(C2/B2,\"n/a\")", f"=IF(B2=0,\"none\",C2/B2)",
          f"=B2/C2*100%", f"=-D2^2", f"=2^-1*D2", f"=1-2-3", f"=10/4/2", f"=2^3^2", f"=-2^2", f"=(1+2)*3%", f"=\"10\"+5",
          f"=\"a\"&1+2", f"=1&2+3", f"=1+2&3", f"=1=1=TRUE", f"=\"abc\"<\"abd\"", f"=TRUE>1", f"=\"Z\">5"]
    for j, f in enumerate(random.sample(fs, 18)):
        s[f"G{j + 1}"] = f
    return {"sheets": {"Sheet1": s}}

for w in range(150):
    W[f"g{w}"] = gradebook(w)
    W[f"i{w}"] = inventory(w)
    W[f"b{w}"] = budget(w)
json.dump(W, open("tests/cases/realistic3.json", "w"), indent=0)
