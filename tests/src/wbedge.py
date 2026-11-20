# Workbook-level edge cases: leading '=' signs, a bare '=', lowercase names, quoted sheet names, odd literals.
# Run: python3 tests/src/wbedge.py > tests/cases/wbedge.json
import json
W={}
def wb(n,**s): W[n]={"sheets":s}
wb("w1", Sheet1={"A1":"=","A2":"==1","A3":"= 1 + 2","A4":"=+A3","A5":"'abc","A6":"=A5&\"\"","A7":-0.0,"A8":"=A7&\"\"","A9":1e-320,"A10":"=A9*1",
  "A11":"=sum(a1:a4)","A12":"=sheet2!a1","A13":"=SHEET2!A1","A14":"=NoSheet!A1","A15":"=XFD1","A16":"=XFE1","A17":"=A1048577","A18":"=A0",
  "A19":"= SUM( A3 , A4 )","A20":"=SUM(A3 ,A4)","A21":"=1+\n2","A22":"=\tA3","A23":"=A3 ","A24":"=(A3)","A25":"=((A3))*2","A26":"=A3:A4",
  "B1":"=A1:A2+1","B2":"=ISBLANK(A1)","B3":"=COUNTA(A1:A30)","B4":"=COUNT(A1:A30)","B5":"=1e400*0","B6":1e308,"B7":"=B6*10","B8":"=-B6*10",
  "B9":"=Sheet2!B1+1","B10":"='Sheet2'!A1","B11":"='Sheet 3'!A1","B12":"='It''s'!A1","B13":"=SUM('Sheet 3':'It''s'!A1)","B14":"=Sheet2!A1:B1",
  "B15":"=SUM(Sheet2!A:A)","B16":"=COUNTA(Sheet2!1:1)","B17":"=sUm(1,2)","B18":"=true+true","B19":"=if(true,\"y\")","B20":"=$a$1&\"\"",
  "B21":"=1.","B22":"=.5","B23":"=1.5e1","B24":"=1E+2","B25":"=5E-1%","B26":"=0x10","B27":"=1 2","B28":"=A3 A4","B29":"=(1)(2)","B30":"=\"a\"\"b\""},
  Sheet2={"A1":7,"B1":"=A1*2"}, **{"Sheet 3":{"A1":3}, "It's":{"A1":4}})
wb("w2", Sheet1={"A1":"=B1","B1":"=C1","C1":"=Sheet2!A1","D1":"=SUM(A1:C1)"}, Sheet2={"A1":"=Sheet1!D1"})
wb("w3", Sheet1={"A1":"=IF(B1>0,B1,0)","B1":"=A1+1"})
wb("w4", Sheet1={"A1":"=IF(FALSE,B1,5)","B1":"=A1+1"})
wb("w5", Sheet1={"A1":"=ISERROR(B1)","B1":"=A1"})
wb("w6", Sheet1={"A1":"=ROW()","B2":"=COLUMN()","C3":"=ROWS(A1:B5)*COLUMNS(A1:C1)","A4":"=INDEX(A1:C3,0,0)","A5":"=INDEX(A1:C3,1,0)","A6":"=MATCH(3,C1:C3)","A7":"=VLOOKUP(\"x\",A1:A1,1)"})
wb("w7", Sheet1={"A1":"= ","A2":"==","A3":"=+=1","A4":"= =1","A5":"=\n1","A6":"===1","A7":"=","B7":"=ISBLANK(A7)","A8":"==A6+1","B8":"=A7&\"x\"","B9":"=COUNTA(A1:A8)","B10":"=ISTEXT(A7)"})
# formula cells read by force-array functions are computed as ordinary formulas
wb("fa_cells", Sheet1={"A1":1,"A2":2,"A3":3,"B1":10,"B2":20,"B3":30,
  "D1":"=MAX(A1:A3*B1:B3)","D2":"=A1:A3*2","D3":"=SUM(A1:A3*B1:B3)","D4":"=IF(A1:A3>1,1,0)",
  "E1":"=SUMPRODUCT(D1:D4)","E2":"=MATCH(30,D1:D4,0)","E3":"=SUMPRODUCT(D1:D4*1)","E4":"=VLOOKUP(2,D1:D4,1,0)",
  "E5":"=SUMPRODUCT(MAX(D1:D4*A1:A4))","E6":"=INDEX(D1:D4,0)","E7":"=MATCH(1,INDEX((D1:D4>1)*1,0),0)"})
print(json.dumps(W))
