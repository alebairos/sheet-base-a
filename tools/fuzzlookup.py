#!/usr/bin/env python3
import json, random, sys
# Dev-only: random approximate-lookup workbooks (MATCH 1/-1, VLOOKUP TRUE over unsorted mixed data).
# Usage: python3 tools/fuzzlookup.py <seed> <count> > cases.json; then tools/lo_oracle.py cases.json
random.seed(int(sys.argv[1]))
vals=[1,2,3,5,8,10,-1,0,2.5,"a","b","c","x","apple","Banana","m",True,False,None,None]
qs=['0','1','2','3','4','6','9','11','-5','"a"','"b"','"bb"','"c"','"z"','"m"','"A"','""','TRUE','FALSE']
out={}
for w in range(int(sys.argv[2])):
    s={}
    n=random.randint(3,10)
    for r in range(n):
        v=random.choice(vals)
        if v is not None: s[f"A{r+1}"]=v
    for i in range(12):
        q=random.choice(qs)
        k=random.random()
        if k<0.4: f=f"=MATCH({q},A1:A{n},1)"
        elif k<0.7: f=f"=MATCH({q},A1:A{n},-1)"
        else: f=f"=VLOOKUP({q},A1:A{n},1,TRUE)"
        s[f"C{i+1}"]=f
    out[f"w{w}"]={"sheets":{"Sheet1":s}}
json.dump(out,sys.stdout,indent=1)
