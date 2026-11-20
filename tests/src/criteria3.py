# Random COUNTIF/SUMIF/AVERAGEIF criteria over mixed cells (input for tests/cases/criteria3.json).
import json, random
random.seed(424242)
vals = [0, 1, 2, -1, 1.5, 10, 100, 0.1, "1", "2", "10", "1.5", "abc", "ABC", "a", "b", "apple", "Apple pie", "banana", " a", "a ",
        "TRUE", "FALSE", True, False, None, None, "=1/0", "=NA()", "=\"\"", "*", "?", "~", "a*b", "x?y", "<5", "=5", "1E1", "$5", "50%", "-", "", "1 1/2", "Jan 1, 2024", "2024-01-01", 45292]
ops = ["", "=", "<>", "<", ">", "<=", ">=", "= ", " ="]
cvals = ["1", "2", "10", "1.5", "0", "-1", "abc", "a*", "*a", "?", "*", "a?", "~*", "~?", "apple*", "TRUE", "FALSE", "", " ", "1E1", "$5", "50%", "a b", "#DIV/0!", "#N/A", "b", "B", "z", "2024-01-01", "1/1/2024"]
W = {}
for w in range(150):
    s = {}
    for i in range(1, 9):
        v = random.choice(vals)
        if v is not None: s[f"A{i}"] = v
        s[f"B{i}"] = random.choice([1, 2, 3, 5, 0.5, "x", True, None, 10]) or 7
    for j in range(1, 7):
        kind = random.random()
        if kind < 0.75:
            c = '"' + random.choice(ops) + random.choice(cvals) + '"'
        elif kind < 0.85:
            c = random.choice(["1", "0", "TRUE", "FALSE", "A1", "A2", "B9", "\"\"", "1.5", "-1"])
        else:
            c = '"' + random.choice(ops) + '"&' + random.choice(["A1", "A3", "1", "TRUE", "B1"])
        f = random.choice(["COUNTIF(A1:A8,%s)", "SUMIF(A1:A8,%s,B1:B8)", "AVERAGEIF(A1:A8,%s,B1:B8)", "SUMIF(A1:A8,%s)", "COUNTIF(A1:B8,%s)"]) % c
        s[f"D{j}"] = "=" + f
    W[f"c{w}"] = {"sheets": {"Sheet1": s}}
json.dump(W, open("tests/cases/criteria3.json", "w"), ensure_ascii=False)
