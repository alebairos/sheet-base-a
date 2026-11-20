#!/usr/bin/env node
'use strict';
// Dev-only: random formulas mixing inline arrays with operators, scalar
// functions and aggregates over the standard data set.
// Usage: node tools/fuzzarr.js <seed> <count> > tests/src/fuzzarrN.txt
let seed = parseInt(process.argv[2] || '1', 10);
const count = parseInt(process.argv[3] || '200', 10);
const rnd = () => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed / 0x7fffffff; };
const pick = (a) => a[Math.floor(rnd() * a.length)];
const arrs = ['{1,2,3}', '{1;2;3}', '{0,1}', '{"a","bb",""}', '{2,"x"}', '{TRUE,FALSE}', '{1,2;3,4}', '{-1.5,2.5}', '{"5","1"}', '{0.1,0.2}'];
const cells = ['A1', 'A2', 'A4', 'A5', 'A7', 'A8', 'A9', 'B1', 'B4', 'B11', 'A12', 'A1:A3', 'B1:B3', 'A1:B3'];
const scal = ['1', '2', '0', '"a"', '""', 'TRUE', '-1', '0.5', '"3"'];
const ops = ['+', '-', '*', '/', '^', '&', '=', '<', '>', '<>'];
const F1 = ['ABS', 'INT', 'SQRT', 'LEN', 'UPPER', 'TRIM', 'NOT', 'ISNUMBER', 'ISTEXT', 'ISERROR', 'ISBLANK'];
const F2 = ['ROUND', 'MOD', 'POWER', 'LEFT', 'RIGHT', 'IFERROR', 'ROUNDDOWN'];
const AGG = ['SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'PRODUCT', 'MEDIAN', 'STDEV', 'AND', 'OR', 'SUMPRODUCT'];
function expr(d) {
  const k = rnd();
  if (d <= 0 || k < 0.25) return rnd() < 0.5 ? pick(arrs) : rnd() < 0.5 ? pick(cells) : pick(scal);
  if (k < 0.5) return expr(d - 1) + pick(ops) + expr(d - 1);
  if (k < 0.65) return pick(F1) + '(' + expr(d - 1) + ')';
  if (k < 0.8) return pick(F2) + '(' + expr(d - 1) + ',' + expr(d - 1) + ')';
  if (k < 0.9) return 'IF(' + expr(d - 1) + ',' + expr(d - 1) + ',' + expr(d - 1) + ')';
  return pick(['COUNTIF(A1:B12,', 'SUMIF(A1:A12,', 'MATCH(', 'VLOOKUP(']) + pick(arrs) + (rnd() < 0.5 ? '' : '') +
    { 0: '' }[0] + ')';
}
const seen = new Set();
let guard = 0;
while (seen.size < count && guard++ < count * 20) {
  let e = expr(3);
  e = e.replace(/(MATCH\([^)]*)\)$/, '$1,A1:A3,0)').replace(/(VLOOKUP\([^)]*)\)$/, '$1,A1:B3,2,0)').replace(/((?:COUNTIF|SUMIF)\(A1:[AB]12,[^)]*)\)$/, '$1)');
  const f = pick(AGG) + '(' + e + ')';
  if (/(MATCH|VLOOKUP)\(\{[^}]*\}\)/.test(f)) continue;
  if (!seen.has(f)) seen.add(f);
}
console.log([...seen].join('\n'));
