#!/usr/bin/env python3
"""Dev-only: random columns mixing formula errors, formula/constant numbers,
text, booleans and blanks, summed by SUM (which error SUM reports).
Usage: tools/fuzzsumerr.py <seed> <count> > tests/cases/sumerr.json"""
import json, random, sys
rnd = random.Random(int(sys.argv[1]))
ERR = ['=1/0', '="x"*1', '=NA()', '=FOO()', '=SQRT(-1)', '=INDEX(Z1:Z2,3)']
OTHER = [1, 2.5, '=2+1', '="t"', 'abc', True, False, None, None, '=A1*0']
out = {}
for i in range(int(sys.argv[2])):
    s = {}
    for col in 'CDE':
        for r in range(1, 9):
            v = rnd.choice(ERR) if rnd.random() < 0.3 else rnd.choice(OTHER)
            if v is not None:
                s[f'{col}{r}'] = v
    s['A1'] = 1
    s['B1'] = '=SUM(C1:E8)'
    s['B2'] = '=SUM(C1:C8)'
    s['B3'] = '=SUM(D2:E7,C1:C8)'
    s['B4'] = '=SUM(C3:D8)+1'
    out[f'sumerr{i}'] = {'sheets': {'Sheet1': s}}
json.dump(out, sys.stdout, indent=1)
print()
