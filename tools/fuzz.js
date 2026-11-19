#!/usr/bin/env node
'use strict';
// Dev-only: generate random formulas over the standard data set (see tools/formulas2cases.py).
// Usage: node tools/fuzz.js <seed> <count> > tests/src/fuzzN.txt
const seed0 = parseInt(process.argv[2] || '1', 10);
const count = parseInt(process.argv[3] || '200', 10);
let seed = seed0;
const rnd = () => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed / 0x7fffffff; };
const pick = (a) => a[Math.floor(rnd() * a.length)];
const chance = (p) => rnd() < p;

const cells = [];
for (const c of ['A', 'B']) for (let r = 1; r <= 12; r++) cells.push(c + r);
const ranges = ['A1:A3', 'A1:A12', 'B1:B12', 'A1:B12', 'A1:B3', 'B1:B3', 'A8:A9', 'A4:A7', 'B7:B12', 'A1:A11', 'B9:B12', 'Sheet2!A1:A3'];
const nums = ['0', '1', '2', '3', '-1', '0.5', '2.5', '10', '100', '1.005', '1E3', '-2.5', '7', '0.1', '3.14159'];
const strs = ['"abc"', '""', '"5"', '" a b "', '"Apple"', '"x"', '"10"', '"TRUE"', '"a*"', '"?ext"', '">1"', '"<>x"', '"<=2"', '"=5"', '"banana"'];

function leaf() {
  const k = rnd();
  if (k < 0.4) return pick(cells);
  if (k < 0.7) return pick(nums);
  if (k < 0.9) return pick(strs);
  return pick(['TRUE', 'FALSE', 'TRUE()', 'NA()', 'Sheet2!A1', 'Sheet2!B1']);
}

const FN = {
  num1: ['ABS', 'INT', 'SQRT', 'NOT', 'ISBLANK', 'ISNUMBER', 'ISTEXT', 'ISERROR', 'LEN', 'UPPER', 'LOWER', 'TRIM'],
  num2: ['MOD', 'POWER', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'LEFT', 'RIGHT', 'IFERROR'],
  agg: ['SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'PRODUCT', 'MEDIAN', 'STDEV', 'AND', 'OR', 'CONCATENATE'],
};

function expr(d) {
  if (d <= 0) return leaf();
  const k = rnd();
  if (k < 0.15) return leaf();
  if (k < 0.35) {
    const op = pick(['+', '-', '*', '/', '^', '&', '=', '<>', '<', '>', '<=', '>=']);
    return expr(d - 1) + op + expr(d - 1);
  }
  if (k < 0.4) return '-' + expr(d - 1);
  if (k < 0.43) return '(' + expr(d - 1) + ')%';
  if (k < 0.5) return pick(FN.num1) + '(' + expr(d - 1) + ')';
  if (k < 0.6) return pick(FN.num2) + '(' + expr(d - 1) + ',' + expr(d - 1) + ')';
  if (k < 0.72) {
    const n = 1 + Math.floor(rnd() * 3);
    const a = [];
    for (let i = 0; i < n; i++) a.push(chance(0.5) ? pick(ranges) : expr(d - 1));
    return pick(FN.agg) + '(' + a.join(',') + ')';
  }
  if (k < 0.78) return 'IF(' + expr(d - 1) + ',' + expr(d - 1) + (chance(0.7) ? ',' + expr(d - 1) : '') + ')';
  if (k < 0.84) {
    const f = pick(['SUMIF', 'COUNTIF', 'AVERAGEIF']);
    const crit = chance(0.6) ? pick(strs.concat(nums)) : expr(d - 1);
    const extra = f !== 'COUNTIF' && chance(0.5) ? ',' + pick(ranges) : '';
    return f + '(' + pick(ranges) + ',' + crit + extra + ')';
  }
  if (k < 0.88) return 'MID(' + expr(d - 1) + ',' + pick(['1', '2', '0', '-1', '3']) + ',' + pick(['1', '2', '5', '0']) + ')';
  if (k < 0.92) return 'VLOOKUP(' + expr(d - 1) + ',' + pick(['A1:B12', 'A1:B3', 'A10:B11', 'B1:B12']) + ',' + pick(['1', '2', '3']) + (chance(0.7) ? ',' + pick(['0', '1', 'FALSE', 'TRUE']) : '') + ')';
  if (k < 0.96) return 'MATCH(' + expr(d - 1) + ',' + pick(['A1:A12', 'B1:B12', 'A1:A3', 'B1:B3']) + (chance(0.7) ? ',' + pick(['0', '1', '-1']) : '') + ')';
  return 'INDEX(' + pick(['A1:B12', 'A1:A12', 'B1:B3']) + ',' + pick(['1', '2', '3', '0', '13', expr(d - 1)]) + (chance(0.5) ? ',' + pick(['1', '2', '0']) : '') + ')';
}

const out = new Set();
while (out.size < count) out.add(expr(1 + Math.floor(rnd() * 3)));
process.stdout.write([...out].join('\n') + '\n');
