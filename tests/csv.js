#!/usr/bin/env node
'use strict';
// CSV import/export used by the UI.
const assert = require('assert');
const { parseCSV, toCSV } = require('../src/csv.js');
const { computeWorkbook } = require('../src/engine.js');

let n = 0;
const t = (f) => { f(); n++; };

t(() => assert.deepStrictEqual(parseCSV('a,1,TRUE\n"x,y",2.5,=B1+B2\n'),
  { A1: 'a', B1: 1, C1: true, A2: 'x,y', B2: 2.5, C2: '=B1+B2' }));
t(() => assert.deepStrictEqual(parseCSV('"he said ""hi""",\n,3'), { A1: 'he said "hi"', B2: 3 }));
t(() => assert.deepStrictEqual(parseCSV('1;2;3\n'), { A1: 1, B1: 2, C1: 3 }));
t(() => assert.deepStrictEqual(parseCSV('"multi\nline",x\r\ny,z\r\n'), { A1: 'multi\nline', B1: 'x', A2: 'y', B2: 'z' }));
t(() => assert.deepStrictEqual(parseCSV('﻿a'), { A1: 'a' }));
t(() => assert.strictEqual(toCSV({ A1: 'a', C1: 1, B2: true, A3: { error: '#N/A' }, B3: 'x,"y"' }),
  'a,,1\n,TRUE,\n#N/A,"x,""y""",\n'));
t(() => assert.strictEqual(toCSV({}), ''));
t(() => {
  // round trip through the engine: formulas computed, values exported
  const cells = parseCSV('1,2,=A1+B1\n');
  const out = computeWorkbook({ sheets: { S: cells } });
  assert.strictEqual(toCSV(out.sheets.S), '1,2,3\n');
});
t(() => {
  // numbers as LibreOffice's CSV export writes them (checked with soffice --convert-to csv)
  const out = computeWorkbook({ sheets: { S: { A1: 1e21, A2: 1e-7, A3: '=1/3', A4: 12345678901234567, A5: 1e16, A6: -2.5e-10, A7: '=0.1+0.2', A8: 1.5e300, A9: 2 ** 60 } } });
  assert.strictEqual(toCSV(out.sheets.S), ['1E+021', '0.0000001', '0.333333333333333', '1.23456789012346E+016', '1E+016',
    '-0.00000000025', '0.3', '1.5E+300', '1.15292150460685E+018'].join('\n') + '\n');
});
console.log(`csv: ${n} passed`);
