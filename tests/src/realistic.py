# Realistic workbooks (input for tests/cases/realistic.json).
# Run: python3 tests/src/realistic.py > tests/cases/realistic.json
import json
W={}
def wb(name, **sheets): W[name]={"sheets":sheets}
emp={"A1":"Name","B1":"Dept","C1":"Salary","D1":"Years",
 "A2":"Ann","B2":"Sales","C2":52000,"D2":3,"A3":"Bob","B3":"IT","C3":61000,"D3":7,"A4":"Cy","B4":"Sales","C4":48000,"D4":1,
 "A5":"Dee","B5":"HR","C5":45000,"D5":12,"A6":"Eve","B6":"IT","C6":75000.5,"D6":4,"A7":"Fay","B7":"","C7":"n/a","D7":2}
f={"F1":"=SUMIF(B2:B7,\"Sales\",C2:C7)","F2":"=AVERAGEIF(B2:B7,\"IT\",C2:C7)","F3":"=COUNTIF(D2:D7,\">=4\")",
 "F4":"=INDEX(A2:A7,MATCH(MAX(C2:C7),C2:C7,0))","F5":"=VLOOKUP(\"Dee\",A2:D7,3,FALSE)","F6":"=ROUND(AVERAGE(C2:C7),2)",
 "F7":"=MEDIAN(C2:C7)","F8":"=STDEV(C2:C6)","F9":"=IF(D2>5,\"Senior\",IF(D2>2,\"Mid\",\"Junior\"))","F10":"=COUNTA(A2:A7)",
 "F11":"=COUNT(C2:C7)","F12":"=SUM(C2:C7)*1.05","F13":"=MIN(D2:D7)&\"-\"&MAX(D2:D7)","F14":"=COUNTIF(B2:B7,\"\")",
 "F15":"=SUMIF(D2:D7,\">\"&D3,C2:C7)","F16":"=IFERROR(VLOOKUP(\"Zed\",A2:D7,2,FALSE),\"not found\")","F17":"=MATCH(\"IT\",B2:B7,0)",
 "F18":"=INDEX(A1:D7,3,4)","F19":"=AND(C2>50000,D2<5)","F20":"=OR(B7=\"\",ISBLANK(B7))","F21":"=C7*2","F22":"=SUM(C7,1)",
 "F23":"=AVERAGE(C2:C7,1000)","F24":"=UPPER(LEFT(A3,1))&LOWER(MID(A3,2,10))","F25":"=LEN(TRIM(\"  \"&A2&\"  \"))",
 "F26":"=CONCATENATE(A2,\" (\",B2,\")\")","F27":"=ROUNDUP(C6/12,0)","F28":"=ROUNDDOWN(C6/12,-2)","F29":"=MOD(D5,5)","F30":"=INT(C6/1000)",
 "F31":"=ABS(C4-C3)","F32":"=POWER(D3,2)","F33":"=SQRT(C2)","F34":"=PRODUCT(D2:D4)","F35":"=COUNTIF(A2:A7,\"?ve\")",
 "F36":"=SUMIF(A2:A7,\"<>Bob\",D2:D7)","F37":"=AVERAGEIF(D2:D7,\"<5\")","F38":"=NOT(ISNUMBER(C7))","F39":"=ISTEXT(C7)","F40":"=ISERROR(C7*1)",
 "F41":"=VLOOKUP(5,D2:D7,1,TRUE)","F42":"=MATCH(5,{1,3,5,7},1)","F43":"=INDEX(C2:C7,0)","F44":"=SUM(INDEX(C2:D7,0,2))","F45":"=VLOOKUP(\"b*\",A2:C7,3,FALSE)",
 "F46":"=RIGHT(A6,2)&\"!\"","F47":"=MID(\"Spreadsheet\",7,5)","F48":"=C2/D2","F49":"=C2/(D2-3)","F50":"=IF(C7=\"n/a\",0,C7)"}
d=dict(emp); d.update(f); wb("employees",Sheet1=d)
inv={"A1":"Item","B1":"Qty","C1":"Price","D1":"Total","A2":"Pen","B2":10,"C2":1.25,"D2":"=B2*C2","A3":"Pad","B3":3,"C3":4.5,"D3":"=B3*C3",
 "A4":"Ink","B4":0,"C4":12.99,"D4":"=B4*C4","A5":"Bag","B5":"2","C5":19.95,"D5":"=B5*C5","D6":"=SUM(D2:D5)","D7":"=D6*0.0825","D8":"=ROUND(D6+D7,2)",
 "D9":"=IF(D8>100,D8*0.9,D8)","D10":"=TEXT_MISSING","D11":"=D8/COUNT(B2:B5)","D12":"=SUMPRODUCT(B2:B5,C2:C5)","D13":"=MAX(D2:D5)-MIN(D2:D5)",
 "D14":"=Summary!B1*2","E1":"=D6/SUM(B2:B5)","E2":"=AVERAGE(D2:D5)","E3":"=D2/$D$6","E4":"=D3/$D$6","E5":"=TEXT(1,\"0\")","E6":"=ROUND(E3*100,1)&\"%\""}
wb("invoice",Sheet1=inv,Summary={"A1":"Total","B1":"=Sheet1!D8","B2":"=Sheet1!D6>50","B3":"=COUNTIF(Sheet1!B2:B5,\">0\")","B4":"=Sheet1!A2&\"/\"&Sheet1!A5"})
gr={"A1":"Student","B1":"T1","C1":"T2","D1":"T3","E1":"Avg","F1":"Grade",
 "A2":"Al","B2":95,"C2":88,"D2":92,"A3":"Bo","B3":72,"C3":65,"D3":"","A4":"Cal","B4":58,"C4":61,"D4":49,"A5":"Di","B5":"abs","C5":90,"D5":85}
for r in range(2,6):
    gr[f"E{r}"]=f"=AVERAGE(B{r}:D{r})"
    gr[f"F{r}"]=f"=IF(E{r}>=90,\"A\",IF(E{r}>=80,\"B\",IF(E{r}>=70,\"C\",IF(E{r}>=60,\"D\",\"F\"))))"
    gr[f"G{r}"]=f"=VLOOKUP(E{r},$I$2:$J$6,2,TRUE)"
    gr[f"H{r}"]=f"=RANK_X(E{r})"
gr.update({"I2":0,"J2":"F","I3":60,"J3":"D","I4":70,"J4":"C","I5":80,"J5":"B","I6":90,"J6":"A",
 "A7":"=COUNTIF(F2:F5,\"A\")","B7":"=AVERAGE(B2:B5)","C7":"=MAX(E2:E5)","D7":"=INDEX(A2:A5,MATCH(C7,E2:E5,0))","E7":"=COUNTIF(E2:E5,\">=\"&B7)",
 "F7":"=SUM(B2:D5)/COUNT(B2:D5)","G7":"=COUNTBLANK(B2:D5)","H7":"=MATCH(75,E2:E5,1)"})
wb("grades",Sheet1=gr)
import sys; json.dump(W, sys.stdout, indent=1); print()
