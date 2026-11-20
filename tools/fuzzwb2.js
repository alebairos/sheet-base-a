#!/usr/bin/env node
'use strict';
// Dev-only: random workbooks with three sheets (one needing quotes), a 5x10
// grid, all 40 spec functions, chains between sheets and the odd cycle.
// Usage: node tools/fuzzwb2.js <seed> <count> > tests/cases/name.json
const seed0 = parseInt(process.argv[2] || '1', 10);
const count = parseInt(process.argv[3] || '50', 10);
let seed = seed0;
const rnd = () => { // mulberry32
  seed = (seed + 0x6D2B79F5) | 0;
  let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
  t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
  return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
};
const pick = (a) => a[Math.floor(rnd() * a.length)];
const chance = (p) => rnd() < p;

const SHEETS = ['Data', 'My Sheet', 'Calc'];
const q = (s) => (/^[A-Za-z]\w*$/.test(s) ? s : `'${s}'`);
const consts = [0, 1, 2, 3, -4, 0.5, 2.25, 10, 100, -0.1, 7, 12.5, 1e6, 'a', 'B', 'text', '5', ' x ', 'apple', 'Apple', 'pear', 'b c', true, false, 0.001, 42];
const cols = ['A', 'B', 'C', 'D', 'E'];
const addrs = [];
for (const c of cols) for (let r = 1; r <= 10; r++) addrs.push(c + r);

const FN1 = ['ABS', 'INT', 'SQRT', 'NOT', 'ISBLANK', 'ISNUMBER', 'ISTEXT', 'ISERROR', 'LEN', 'UPPER', 'LOWER', 'TRIM'];
const FN2 = ['ROUND', 'ROUNDUP', 'ROUNDDOWN', 'MOD', 'POWER', 'LEFT', 'RIGHT', 'IFERROR'];
const AGG = ['SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'PRODUCT', 'MEDIAN', 'STDEV', 'AND', 'OR', 'CONCATENATE'];

function makeBook() {
  const sh = {};
  for (const s of SHEETS) sh[s] = {};
  // a data block on each sheet: columns A-C mostly constants
  for (const s of SHEETS) for (const a of addrs) if (a[0] <= 'C' && chance(0.55)) sh[s][a] = pick(consts);
  const home = () => pick(SHEETS);
  const ref = (s) => {
    const a = pick(addrs);
    const d = chance(0.2) ? ['$', '$'] : chance(0.1) ? ['$', ''] : ['', ''];
    const t = d[0] + a[0] + d[1] + a.slice(1);
    return chance(0.25) ? q(pick(SHEETS.filter((x) => x !== s))) + '!' + t : t;
  };
  const rng = (s) => {
    const c1 = Math.floor(rnd() * 3), c2 = c1 + Math.floor(rnd() * 2);
    const r1 = 1 + Math.floor(rnd() * 5), r2 = r1 + Math.floor(rnd() * 5);
    const t = cols[c1] + r1 + ':' + cols[c2] + r2;
    return chance(0.25) ? q(pick(SHEETS.filter((x) => x !== s))) + '!' + t : t;
  };
  const col = (s) => { const c = pick(['A', 'B', 'C']); const r1 = 1 + Math.floor(rnd() * 3); return c + r1 + ':' + c + (r1 + 4 + Math.floor(rnd() * 3)); };
  const leaf = (s) => (chance(0.55) ? ref(s) : pick(['1', '2', '0', '3', '"a"', '""', 'TRUE', 'FALSE', '-1', '0.5', '"3"', '"apple"', '10', '2.5']));
  const expr = (s, d) => {
    if (d <= 0) return leaf(s);
    const k = rnd();
    if (k < 0.15) return leaf(s);
    if (k < 0.35) return expr(s, d - 1) + pick(['+', '-', '*', '/', '&', '=', '<', '>', '^', '<>', '>=', '<=']) + expr(s, d - 1);
    if (k < 0.38) return '-' + expr(s, d - 1);
    if (k < 0.45) return pick(FN1) + '(' + expr(s, d - 1) + ')';
    if (k < 0.53) return pick(FN2) + '(' + expr(s, d - 1) + ',' + (chance(0.6) ? pick(['0', '1', '2', '-1', '3']) : expr(s, d - 1)) + ')';
    if (k < 0.66) return pick(AGG) + '(' + (chance(0.65) ? rng(s) : expr(s, d - 1)) + (chance(0.35) ? ',' + expr(s, d - 1) : '') + ')';
    if (k < 0.74) return 'IF(' + expr(s, d - 1) + ',' + expr(s, d - 1) + (chance(0.8) ? ',' + expr(s, d - 1) : '') + ')';
    if (k < 0.8) {
      const f = pick(['SUMIF', 'COUNTIF', 'AVERAGEIF']);
      const crit = pick(['">0"', '"a*"', '1', '"<>"', '""', '"*"', '">=2"', '"<5"', '"apple"', '"?"', '">"&' + leaf(s), ref(s)]);
      return f + '(' + col(s) + ',' + crit + (f !== 'COUNTIF' && chance(0.5) ? ',' + col(s).replace(/[ABC]/g, (c) => pick(['A', 'B', 'C'])) : '') + ')';
    }
    if (k < 0.85) return 'MID(' + expr(s, d - 1) + ',' + pick(['1', '2', '3', '0']) + ',' + pick(['1', '2', '4']) + ')';
    if (k < 0.9) return 'VLOOKUP(' + expr(s, d - 1) + ',' + pick(['A1:C10', 'A1:B10', 'B1:C10', 'A3:C6']) + ',' + pick(['1', '2', '3']) + ',' + pick(['0', 'FALSE', '1', 'TRUE']) + ')';
    if (k < 0.95) return 'MATCH(' + expr(s, d - 1) + ',' + col(s) + ',' + pick(['0', '1', '-1']) + ')';
    if (k < 0.98) return 'INDEX(' + rng(s) + ',' + pick(['1', '2', '3', '0', expr(s, d - 1)]) + (chance(0.6) ? ',' + pick(['1', '2']) : '') + ')';
    return 'CONCATENATE(' + expr(s, d - 1) + ',' + expr(s, d - 1) + ')';
  };
  const nForm = 6 + Math.floor(rnd() * 10);
  for (let i = 0; i < nForm; i++) {
    const s = home();
    const free = addrs.filter((x) => !(x in sh[s]));
    const a = pick(free);
    if (a) sh[s][a] = '=' + expr(s, 1 + Math.floor(rnd() * 3));
  }
  return { sheets: sh };
}

const out = {};
for (let i = 0; i < count; i++) out[`wbx${seed0}_${i}`] = makeBook();
process.stdout.write(JSON.stringify(out, null, 1) + '\n');
