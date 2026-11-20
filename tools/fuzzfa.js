#!/usr/bin/env node
'use strict';
// Dev-only: random force-array idioms (SUMPRODUCT, MATCH/VLOOKUP tables, INDEX(..,0),
// CHOOSE with array indexes) over functions of ranges, on the standard data set.
// Usage: node tools/fuzzfa.js <seed> <count> > tests/src/faN.txt
const seed0 = parseInt(process.argv[2] || '1', 10);
const count = parseInt(process.argv[3] || '300', 10);
let seed = seed0;
const rnd = () => { seed = (seed + 0x6D2B79F5) | 0; let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
  t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t; return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
const pick = (a) => a[Math.floor(rnd() * a.length)];
const chance = (p) => rnd() < p;
const RNG = ['A1:A3', 'B1:B3', 'A1:A5', 'B1:B5', 'A8:A12', 'B8:B12', 'A10:A11', 'B10:B11', 'A4:A7', 'B4:B7'];
const SC = ['1', '2', '0', '-1', '0.5', '"a"', '""', 'TRUE', 'A1', 'B2', 'A4', 'A6', '"apple"', '10'];
function rng() { return pick(RNG); }
function arr(d) {
  const r = rng();
  if (d <= 0) return r;
  const k = Math.floor(rnd() * 16);
  const a = () => arr(d - 1), s = () => pick(SC);
  switch (k) {
    case 0: return `(${a()}${pick(['>', '<', '=', '<>', '>='])}${s()})`;
    case 1: return `${a()}${pick(['*', '+', '-', '/', '&'])}${chance(0.5) ? s() : a()}`;
    case 2: return `${pick(['LEN', 'ABS', 'INT', 'UPPER', 'TRIM', 'ISNUMBER', 'ISTEXT', 'ISBLANK', 'ISERROR', 'N', 'NOT', 'SQRT', 'LOWER'])}(${a()})`;
    case 3: return `IF(${a()},${pick([a(), s()])}${chance(0.6) ? ',' + pick([a(), s()]) : ''})`;
    case 4: return `IFERROR(${a()},${s()})`;
    case 5: return `${pick(['ROUND', 'MOD', 'LEFT', 'RIGHT', 'POWER', 'ROUNDUP'])}(${a()},${pick(['0', '1', '2', a()])})`;
    case 6: return `MID(${a()},${pick(['1', '2'])},${pick(['1', '3'])})`;
    case 7: return `--(${a()})`;
    case 8: return `COUNTIF(${r},${a()})`;
    case 9: return `INDEX(${a()},0)`;
    case 10: return `CHOOSE({1,2},${a()},${a()})`;
    case 11: return `${pick(['MAX', 'MIN', 'SUM', 'AVERAGE', 'COUNT', 'COUNTA'])}(${a()})`;
    case 12: return `CONCATENATE(${a()},"x")`;
    case 13: return `SUMIF(${r},${a()})`;
    case 14: return `(${a()})*(${a()})`;
    default: return `ROW(${r})`;
  }
}
function top() {
  const d = 1 + Math.floor(rnd() * 3);
  switch (Math.floor(rnd() * 7)) {
    case 0: case 1: case 2: return `SUMPRODUCT(${arr(d)})`;
    case 3: return `SUMPRODUCT(${arr(d)},${arr(d - 1)})`;
    case 4: return `MATCH(${pick(SC)},${arr(d)},${pick(['0', '1', '-1'])})`;
    case 5: return `VLOOKUP(${pick(SC)},CHOOSE({1,2},${arr(d - 1)},${arr(d - 1)}),2,${pick(['0', '1'])})`;
    default: return `INDEX(${arr(d)},${pick(['1', '2', '3'])})`;
  }
}
const out = new Set();
while (out.size < count) out.add('=' + top());
console.log([...out].join('\n'));
