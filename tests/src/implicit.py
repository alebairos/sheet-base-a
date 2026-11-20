# Implicit intersection and array workbooks (input for tests/cases/implicit.json).
# Run: python3 tests/src/implicit.py > tests/cases/implicit.json
import json
data={"A1":1,"A2":2,"A3":"x","A4":4,"B1":10,"B2":20,"B3":30,"B4":40,"C1":"a","C2":"b"}
fs=["=A1:A4*2","=ABS(A1:A4)","=A1:A4&\"!\"","=IF(A1:A4>1,\"big\",\"small\")","=LEN(A1:A4)","=-A1:A4","=A1:A4","=B1:B4+A1:A4","=SUM(A1:A4*1)","=ROUND(B1:B4/3,1)",
    "=A1:B1","=A1:B1*2","=Sheet2!A1:A4","=ISNUMBER(A1:A4)","=UPPER(C1:C2)","=A1:A4=2","=NOT(A1:A4)","=MID(C1:C2,1,1)","=INDEX(A1:A4,0)","=A2:A3+0","=SUMPRODUCT(A1:A2*B1:B2)","=SUM(A1:A2*B1:B2)"]
W={}
for i,f in enumerate(fs):
    s=dict(data)
    for r in range(1,7):
        s[f"E{r}"]=f
    s["F2"]=f.replace("A1:A4","A1:A2") if False else f
    W[f"ii{i}"]={"sheets":{"Sheet1":s,"Sheet2":{"A1":5,"A2":6,"A3":7,"A4":8}}}
import sys; json.dump(W, sys.stdout, indent=1)
