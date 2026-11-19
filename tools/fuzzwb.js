#!/usr/bin/env node
'use strict';
// Dev-only: random multi-cell workbooks (formulas referencing formulas, other sheets, blanks).
// Usage: node tools/fuzzwb.js <seed> <count> > tests/cases/fuzzwbN.json
const seed0 = parseInt(process.argv[2] || '1', 10);
const count = parseInt(process.argv[3] || '50', 10);
let seed = seed0;
const rnd = () => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed / 0x7fffffff; };
const pick = (a) => a[Math.floor(rnd() * a.length)];
const chance = (p) => rnd() < p;

const consts = [0, 1, 2, 3, -4, 0.5, 2.25, 10, 100, -0.1, 7, 'a', 'B', 'text', '5', ' x ', 'apple', true, false, 1e-3, 123456.789];
const FN1 = ['ABS', 'INT', 'SQRT', 'NOT', 'ISBLANK', 'ISNUMBER', 'ISTEXT', 'ISERROR', 'LEN', 'UPPER', 'LOWER', 'TRIM', 'ROUND'];
const AGG = ['SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'PRODUCT', 'MEDIAN', 'STDEV', 'AND', 'OR', 'CONCATENATE'];

function makeBook() {
  const s1 = {}, s2 = {};
  const addrs1 = [];
  for (const c of ['A', 'B', 'C']) for (let r = 1; r <= 6; r++) addrs1.push(c + r);
  const addrs2 = ['A1', 'A2', 'A3', 'B1', 'B2'];
  for (const a of addrs1) if (chance(0.45)) s1[a] = pick(consts);
  for (const a of addrs2) if (chance(0.6)) s2[a] = pick(consts);
  const refs = () => (chance(0.15) ? 'Sheet2!' + pick(addrs2) : pick(addrs1));
  const rng = () => pick(['A1:A6', 'B1:B6', 'C1:C6', 'A1:C6', 'A1:B3', 'B2:C5', 'Sheet2!A1:A3', 'Sheet2!A1:B2', 'A1:A3']);
  const leaf = () => (chance(0.6) ? refs() : pick(['1', '2', '0', '"a"', '""', 'TRUE', '-1', '0.5', '"3"']));
  const expr = (d) => {
    if (d <= 0) return leaf();
    const k = rnd();
    if (k < 0.2) return leaf();
    if (k < 0.45) return expr(d - 1) + pick(['+', '-', '*', '/', '&', '=', '<', '>', '^', '<>']) + expr(d - 1);
    if (k < 0.55) return pick(FN1) + '(' + expr(d - 1) + ')';
    if (k < 0.7) return pick(AGG) + '(' + (chance(0.6) ? rng() : expr(d - 1)) + (chance(0.4) ? ',' + expr(d - 1) : '') + ')';
    if (k < 0.8) return 'IF(' + expr(d - 1) + ',' + expr(d - 1) + ',' + expr(d - 1) + ')';
    if (k < 0.85) return 'IFERROR(' + expr(d - 1) + ',' + expr(d - 1) + ')';
    if (k < 0.9) return pick(['SUMIF', 'COUNTIF', 'AVERAGEIF']) + '(' + rng() + ',' + pick(['">0"', '"a"', '1', '"<>"', '""', '"*"', '">=2"', refs()]) + ')';
    if (k < 0.95) return 'INDEX(' + rng() + ',' + pick(['1', '2', '3']) + ',' + pick(['1', '2']) + ')';
    return pick(['LEFT', 'RIGHT']) + '(' + expr(d - 1) + ',' + pick(['1', '2', '0']) + ')';
  };
  const nForm = 4 + Math.floor(rnd() * 6);
  for (let i = 0; i < nForm; i++) {
    const a = pick(addrs1.filter((x) => !(x in s1)).concat(['D1', 'D2', 'D3', 'D4', 'D5', 'D6']));
    if (a && !(a in s1)) s1[a] = '=' + expr(1 + Math.floor(rnd() * 3));
  }
  if (chance(0.5)) s2.C1 = '=' + expr(2);
  return { sheets: { Sheet1: s1, Sheet2: s2 } };
}

const out = {};
for (let i = 0; i < count; i++) out[`wb${seed0}_${i}`] = makeBook();
process.stdout.write(JSON.stringify(out, null, 1) + '\n');
