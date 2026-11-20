#!/usr/bin/env node
'use strict';
// Dev-only: random numeric formulas with awkward decimals (precision, rounding, output digits).
// Usage: node tools/fuzzprec.js <seed> <count> > tests/src/precN.txt
const seed0 = parseInt(process.argv[2] || '1', 10);
const count = parseInt(process.argv[3] || '300', 10);
let seed = seed0;
const rnd = () => { seed = (seed + 0x6D2B79F5) | 0; let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
  t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t; return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
const pick = (a) => a[Math.floor(rnd() * a.length)];
const chance = (p) => rnd() < p;
function num() {
  const k = rnd();
  if (k < 0.25) return pick(['0.1', '0.2', '0.3', '0.7', '1.1', '2.675', '1.005', '0.015', '1.45', '2.5', '-2.5', '0.5', '1E-15', '1E15', '123456789.123', '3', '7', '1E-7']);
  const digs = 1 + Math.floor(rnd() * 6), dec = Math.floor(rnd() * 6);
  let s = String(Math.floor(rnd() * 10 ** digs));
  if (dec) s += '.' + String(Math.floor(rnd() * 10 ** dec)).padStart(dec, '0');
  if (chance(0.2)) s = '-' + s;
  if (chance(0.1)) s += 'E' + pick(['-3', '5', '-10', '12', '2']);
  return s;
}
function ex(d) {
  if (d <= 0 || chance(0.25)) return chance(0.15) ? pick(['A1', 'A2', 'A3', 'A8', 'B7', 'B12', 'A9']) : num();
  const k = Math.floor(rnd() * 14);
  const a = () => ex(d - 1);
  switch (k) {
    case 0: return `${a()}${pick(['+', '-', '*', '/'])}${a()}`;
    case 1: return `(${a()}${pick(['+', '-'])}${a()})${pick(['*', '/'])}${a()}`;
    case 2: return `ROUND(${a()},${pick(['0', '1', '2', '3', '-1', '-2', '10', '14', '15'])})`;
    case 3: return `${pick(['ROUNDUP', 'ROUNDDOWN'])}(${a()},${pick(['0', '1', '2', '-1', '5'])})`;
    case 4: return `MOD(${a()},${a()})`;
    case 5: return `${pick(['SUM', 'AVERAGE', 'PRODUCT', 'MEDIAN', 'STDEV', 'MIN', 'MAX'])}(${a()},${a()},${a()}${chance(0.4) ? ',' + a() : ''})`;
    case 6: return `POWER(${a()},${pick(['2', '0.5', '3', '-1', '1/3', '1.5'])})`;
    case 7: return `SQRT(ABS(${a()}))`;
    case 8: return `INT(${a()})`;
    case 9: return `${a()}^${pick(['2', '3', '0.5', '-2', '10'])}`;
    case 10: return `${a()}%`;
    case 11: return `(${a()})&""`;
    case 12: return `${a()}*${a()}-${a()}`;
    default: return `SUM(${a()},-${a()})`;
  }
}
const out = new Set();
while (out.size < count) out.add('=' + ex(1 + Math.floor(rnd() * 3)));
console.log([...out].join('\n'));
