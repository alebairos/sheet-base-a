# Multi-cell workbooks (input for tests/cases/workbooks.json).
# Run: python3 tests/src/workbooks.py > tests/cases/workbooks.json
import json
W = {}
def wb(name, **sheets):
    W[name] = {"sheets": sheets}

wb("chain", Sheet1={"A1": 1, "A2": "=A1+1", "A3": "=A2*2", "A4": "=A3&\"!\"", "B1": "=A4"})
wb("chain_reverse", Sheet1={"A5": 5, "A4": "=A5+1", "A3": "=A4+1", "A2": "=A3+1", "A1": "=A2+1"})
wb("self_ref", Sheet1={"A1": "=A1+1", "B1": "=A1*2", "C1": "=IFERROR(A1,7)", "D1": "=ISERROR(A1)"})
wb("cycle2", Sheet1={"A1": "=B1", "B1": "=A1", "C1": "=A1+1", "D1": 5})
wb("cycle3", Sheet1={"A1": "=B1+1", "B1": "=C1+1", "C1": "=A1+1", "D1": "=SUM(A1:C1)"})
wb("cycle_range", Sheet1={"A1": 1, "A2": 2, "A3": "=SUM(A1:A3)", "B1": "=A3"})
wb("cycle_iferror", Sheet1={"A1": "=IFERROR(B1,0)", "B1": "=A1+1"})
wb("cross_sheet", Sheet1={"A1": "=Data!A1*2", "A2": "=SUM(Data!A1:A3)", "A3": "='My Sheet'!B2&\"x\""},
   Data={"A1": 5, "A2": "=A1+1", "A3": "=Sheet1!A1"}, **{"My Sheet": {"B2": "hello"}})
wb("cross_cycle", Sheet1={"A1": "=Other!A1"}, Other={"A1": "=Sheet1!A1"})
wb("bool_out", Sheet1={"A1": True, "A2": False, "B1": "=A1", "B2": "=NOT(A1)", "B3": "=A1+0", "B4": "=AND(A1,A2)",
   "B5": "=OR(A1,A2)", "B6": "=A1=TRUE", "B7": "=ISBLANK(A3)", "B8": "=IF(A1,A2,A1)", "B9": "=MAX(A1:A2)",
   "B10": "=SUM(A1:A2)", "B11": "=COUNT(A1:A2)", "B12": "=B1", "B13": "=INDEX(A1:A2,2)", "B14": "=TRUE()",
   "B15": "=IFERROR(A1,0)", "B16": "=MIN(A1,A2)", "B17": "=-A1", "B18": "=A1*A2", "B19": "=A1&A2",
   "B20": "=CONCATENATE(A1)", "B21": "=VLOOKUP(TRUE,A1:A2,1,0)", "B22": "=IF(1,TRUE)", "B23": "=ISNUMBER(A1)",
   "B24": "=A1>A2", "B25": "=ROUND(A1,0)", "B26": "=ABS(A1)", "B27": "=MAX(A1,2)", "B28": "=MEDIAN(A1:A2)"})
wb("blank_refs", Sheet1={"A1": "=B1", "A2": "=A1&\"x\"", "A3": "=ISBLANK(A1)", "A4": "=A1+1", "A5": "=COUNTA(A1)",
   "A6": "=LEN(A1)", "A7": "=IF(TRUE,B1)", "A8": "=A7&\"y\"", "A9": "=INDEX(B1:B3,2)", "A10": "=A9&\"z\"",
   "A11": "=VLOOKUP(1,C1:D1,2,0)", "A12": "=A11&\"w\"", "C1": 1, "A13": "=ISTEXT(A1)", "A14": "=ISNUMBER(A1)",
   "A15": "=COUNT(A1)", "A16": "=A1=\"\"", "A17": "=A1=0", "A18": "=COUNTBLANK(B1:B3)"})
wb("text_vals", Sheet1={"A1": "10", "A2": "abc", "A3": 5, "B1": "=A1+A3", "B2": "=SUM(A1:A3)", "B3": "=A1*1",
   "B4": "=COUNT(A1:A3)", "B5": "=COUNTA(A1:A3)", "B6": "=AVERAGE(A1:A3)", "B7": "=A2+1", "B8": "=MAX(A1:A3)",
   "B9": "=A1&A3", "B10": "=LEN(A1)", "B11": "=ISNUMBER(A1)", "B12": "=A1=10", "B13": "=A1>5", "B14": "=VLOOKUP(10,A1:A3,1,0)",
   "B15": "=MATCH(\"10\",A1:A3,0)", "B16": "=SUMIF(A1:A3,\"10\")", "B17": "=COUNTIF(A1:A3,10)"})
wb("errors_prop", Sheet1={"A1": "=1/0", "A2": "=NA()", "A3": "=SQRT(-1)", "A4": "=\"a\"+1", "A5": "=FOO()",
   "B1": "=A1+A2", "B2": "=A2+A1", "B3": "=SUM(A1:A5)", "B4": "=SUM(A2:A5)", "B5": "=COUNT(A1:A5)", "B6": "=COUNTA(A1:A5)",
   "B7": "=IFERROR(A1,\"x\")", "B8": "=ISERROR(A5)", "B9": "=A3", "B10": "=A4", "B11": "=MAX(A1:A5)", "B12": "=AND(A1:A2)",
   "B13": "=A1&A2", "B14": "=CONCATENATE(A2,A1)", "B15": "=IF(A1,1,2)", "B16": "=VLOOKUP(A2,A1:A5,1,0)",
   "B17": "=COUNTIF(A1:A5,\"#N/A\")", "B18": "=SUMIF(A1:A5,\">0\")", "B19": "=MATCH(1,A1:A5,0)", "B20": "=INDEX(A1:A5,2)",
   "B21": "=ROUND(A1,1)", "B22": "=LEN(A2)", "B23": "=A2=A1", "B24": "=-A1", "B25": "=A1%"})
wb("lookup_table", Sheet1={"A1": "Name", "B1": "Score", "C1": "Grade",
   "A2": "alice", "B2": 91, "C2": "A", "A3": "bob", "B3": 78, "C3": "C", "A4": "carol", "B4": 85, "C4": "B",
   "A5": "dave", "B5": 62, "C5": "D",
   "E1": 0, "F1": "F", "E2": 60, "F2": "D", "E3": 70, "F3": "C", "E4": 80, "F4": "B", "E5": 90, "F5": "A",
   "H1": "=VLOOKUP(\"carol\",A2:C5,2,FALSE)", "H2": "=VLOOKUP(B3,E1:F5,2)", "H3": "=VLOOKUP(B2,E1:F5,2,TRUE)",
   "H4": "=INDEX(A2:A5,MATCH(MAX(B2:B5),B2:B5,0))", "H5": "=MATCH(\"dave\",A1:A5,0)", "H6": "=INDEX(A1:C5,3,3)",
   "H7": "=AVERAGEIF(C2:C5,\"<>D\",B2:B5)", "H8": "=SUMIF(B2:B5,\">=80\")", "H9": "=COUNTIF(A2:A5,\"?o*\")",
   "H10": "=VLOOKUP(\"zed\",A2:C5,2,0)", "H11": "=IFERROR(VLOOKUP(\"zed\",A2:C5,2,0),\"none\")",
   "H12": "=INDEX(B2:B5,MATCH(\"bob\",A2:A5,0))*2", "H13": "=MATCH(85,B2:B5)", "H14": "=VLOOKUP(55,E1:F5,2)",
   "H15": "=ROUND(AVERAGE(B2:B5),1)", "H16": "=MEDIAN(B2:B5)", "H17": "=STDEV(B2:B5)", "H18": "=MAX(B2:B5)-MIN(B2:B5)",
   "H19": "=CONCATENATE(UPPER(LEFT(A2)),MID(A2,2,LEN(A2)))", "H20": "=VLOOKUP(\"B*\",A2:C5,3,0)",
   "H21": "=INDEX(E1:F5,0,2)", "H22": "=MATCH(\"C\",F1:F5,0)", "H23": "=VLOOKUP(-1,E1:F5,2)", "H24": "=PRODUCT(B2:B3)"})
wb("sheet_names", Sheet1={"A1": "=sheet2!A1", "A2": "=SHEET2!A1+1", "A3": "='Sheet 3'!A1", "A4": "=SUM('Sheet 3'!A1:A2)"},
   Sheet2={"A1": 3}, **{"Sheet 3": {"A1": 4, "A2": 6}})
wb("abs_refs", Sheet1={"A1": 2, "B1": 3, "C1": "=$A$1*$B1+A$1", "C2": "=SUM($A$1:$B$1)", "C3": "=Sheet1!$A$1"})
wb("long_chain", Sheet1=dict([("A1", 1)] + [(f"A{i}", f"=A{i-1}+1") for i in range(2, 3001)] + [("B1", "=A3000"), ("B2", "=SUM(A1:A3000)")]))
wb("numbers_out", Sheet1={"A1": "=0.1+0.2", "A2": "=1/3", "A3": "=2/3", "A4": "=10/3", "A5": "=1E-20", "A6": "=1E20/3",
   "A7": "=-1/7", "A8": "=PI()", "A9": "=12345678.9*3", "A10": "=SQRT(2)", "A11": "=100*1.1", "A12": "=2^0.5*2^0.5",
   "A13": "=1-0.9", "A14": "=3*1.1", "A15": "=-0", "A16": "=0*-1", "A17": "=1E15+1", "A18": "=123456789012345678",
   "A19": "=4.35*100", "A20": "=1.1+2.2", "A21": "=MOD(10,3.3)", "A22": "=STDEV(1.1,2.2,3.3)", "A23": "=AVERAGE(0.1,0.2,0.4)"})
wb("whole_cols", Sheet1={"A1": 1, "A2": 2, "A5": 10, "B1": "=SUM(A:A)", "B2": "=COUNT(A:A)", "B3": "=COUNTA(A:A)",
   "B4": "=MAX(A:A)", "B5": "=COUNTIF(A:A,\">1\")", "B6": "=MATCH(10,A:A,0)", "B7": "=VLOOKUP(2,A:A,1,0)", "B8": "=SUM(1:2)",
   "B9": "=AVERAGE(A:A)", "B10": "=SUMIF(A:A,\"<5\")"})
wb("empty_sheet", Sheet1={}, Sheet2={"A1": "=Sheet1!A1", "A2": "=Sheet1!A1&\"\"", "A3": "=ISBLANK(Sheet1!A1)"})
wb("criteria_mix", Sheet1={"A1": 1, "A2": "1", "A3": "a", "A4": "A", "A5": "=\"\"", "A6": True, "A7": 2.5, "A8": "x*y",
   "A9": " a", "A10": "a ", "B1": 1, "B2": 2, "B3": 3, "B4": 4, "B5": 5, "B6": 6, "B7": 7, "B8": 8, "B9": 9, "B10": 10,
   "C1": "=COUNTIF(A1:A10,1)", "C2": "=COUNTIF(A1:A10,\"1\")", "C3": "=COUNTIF(A1:A10,\"a\")", "C4": "=COUNTIF(A1:A10,\"\")",
   "C5": "=COUNTIF(A1:A10,\"<>\")", "C6": "=SUMIF(A1:A10,\"a\",B1:B10)", "C7": "=COUNTIF(A1:A10,\"x~*y\")", "C8": "=COUNTIF(A1:A10,\"*\")",
   "C9": "=COUNTIF(A1:A10,\">=1\")", "C10": "=COUNTIF(A1:A10,\"<>a\")", "C11": "=SUMIF(A1:A10,\"<>1\",B1:B10)", "C12": "=COUNTIF(A1:A10,\"?\")",
   "C13": "=COUNTIF(A1:A10,\"=\")", "C14": "=COUNTIF(A1:A10,\">\"&A7)", "C15": "=AVERAGEIF(A1:A10,\"*a*\",B1:B10)", "C16": "=COUNTIF(A1:A10,2.5)",
   "C17": "=COUNTIF(A1:A10,\"=2.5\")", "C18": "=COUNTIF(A1:A10,\"> \")", "C19": "=COUNTIF(B1:B10,\">\"&\"5\")", "C20": "=COUNTIF(A1:A10,\"<\")"})
print(json.dumps(W, indent=1))
