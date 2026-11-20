import json
cases={}
fs=["=SUM(A1:A3)-A1","=SUM(A1:A3)-4503599627370496","=SUM(4503599627370496,1/3,0.3)-4503599627370496","=SUM(0.3,1/3,4503599627370496)-4503599627370496","=(SUM(A1:A3)-A1)*1","=SUM(A1:A3)=A1","=SUM(B1:B3)-B1","=SUM(B1:B3)=B1","=A1+A2+A3-A1","=SUM(A1,A2,A3)-A1","=SUM(A3,A2,A1)-A1","=SUM(A1:A2)-A1","=SUM(A1,A3)-A1","=SUM(A2:A3)","=SUM(C1:C3)-C1"]
for i,f in enumerate(fs):
    cases[f]={"sheets":{"Sheet1":{"A1":"=2^52","A2":"=1/3","A3":"=0.3","B1":4503599627370496,"B2":0.3333333333333333,"B3":0.3,"C1":"=2^52+0.5","C2":"=1/3","C3":"=0.3","D1":f}}}
json.dump(cases,open('tests/cases/sumformcells.json','w'))
