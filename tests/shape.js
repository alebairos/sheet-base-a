#!/usr/bin/env node
'use strict';
// Every function, with 3D references, reference lists and arrays in each
// argument position (wrapped in IF/CHOOSE/IFERROR/INDEX/OFFSET, inside SUM,
// SUMPRODUCT or &): the result must be a JSON value of the contract's shape.
const { Worker, isMainThread } = require('worker_threads');
if (isMainThread) {
  const w = new Worker(__filename, { resourceLimits: { stackSizeMb: 256 } });
  w.on('exit', (code) => process.exit(code));
  return;
}
const { computeWorkbook, FUNCS } = require('../src/engine.js');
const refs = ['Sheet1:Sheet3!B2', 'Sheet1:Sheet3!B2:C3', '(A1,A2)', '(A1:A2,B1:B2)', 'Sheet1!A1:Sheet2!B2', '{1,2;3,4}', 'A1:B3'];
const wraps = [(x) => x, (x) => `IF({1,0},${x},0)`, (x) => `IF(1,${x})`, (x) => `CHOOSE({1,2},${x},1)`, (x) => `IFERROR(${x},0)`,
  (x) => `IFERROR({1,2}/0,${x})`, (x) => `INDEX(${x},1,1)`, (x) => `OFFSET(${x},0,0)`];
const outers = [(x) => x, (x) => `SUM(${x})`, (x) => `SUMPRODUCT(${x})`, (x) => `${x}&""`];
let n = 0;
const bad = [];
for (const f of Object.keys(FUNCS)) for (const r of refs) for (const w of wraps) for (const pos of [0, 1, 2]) {
  const args = ['1', '2', '1'].slice(0, pos + 1);
  args[pos] = w(r);
  for (const o of outers) {
    const fm = '=' + o(`${f}(${args.join(',')})`);
    n++;
    let v;
    try {
      v = computeWorkbook({ sheets: { Sheet1: { A1: 1, A2: 2, B2: 5, B3: 'x', E9: fm }, Sheet2: { B2: 7 }, Sheet3: { B2: 'x' } } }).sheets.Sheet1.E9;
    } catch (e) { bad.push(`${fm}: throws ${e.message}`); continue; }
    const ok = v === null || typeof v === 'string' || typeof v === 'boolean' || (typeof v === 'number' && isFinite(v)) ||
      (v !== null && typeof v === 'object' && Object.keys(v).length === 1 && typeof v.error === 'string' && /^#/.test(v.error));
    if (!ok) bad.push(`${fm}: ${JSON.stringify(v).slice(0, 80)}`);
  }
}
bad.slice(0, 20).forEach((b) => console.log('FAIL shape ::', b));
console.log(`shape: ${n - bad.length} passed, ${bad.length} failed`);
process.exit(bad.length ? 1 : 0);
