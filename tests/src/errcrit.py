# COUNTIF over cells holding different errors (Calc internal errors match their Err:5xx text).
# Run: python3 tests/src/errcrit.py; cases written to tests/cases/errcrit.json
import json
errs = {"A1":"=1/0","A2":"=NA()","A3":"=SQRT(-4)","A4":"=\"a\"+1","A5":"=FOO()","A6":"=INDEX(B1:B2,3)","A7":"=MID(\"a\",0,1)","A8":"=1E200*1E200","A9":"=A9","A10":"=(A1:A2 B5:B6)","A11":"=VLOOKUP(NA(),B1:B2,1,0)","A12":"=MOD(1,0)", "A13":"=10^-400"}
crits = ["*", "*A*", "*a*", "#VALUE!", "#N/A", "#DIV/0!", "Err:502", "*5*", "#*", "<>", "=", "", "<>x", ">A", "<z", "*!"]
W = {}
for i,c in enumerate(crits):
    for a in errs:
        s = dict(errs); s["C1"] = "=COUNTIF(%s,\"%s\")" % (a, c)
        W[f"{c} {a}"] = {"sheets":{"Sheet1":s}}
json.dump(W, open("tests/cases/errcrit.json","w"))
