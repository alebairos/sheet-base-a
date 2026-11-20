#!/usr/bin/env node
'use strict';
// Dev-only: small random workbooks whose formula cells reference each other
// (reference cycles), mixing computed errors, resets (IFERROR, IS..., COUNT)
// and ranges, to compare how cycle members are reported with LibreOffice.
// Usage: node tools/fuzzcyc.js <seed> <count> > cases.json
const seed0 = parseInt(process.argv[2] || '1', 10);
const count = parseInt(process.argv[3] || '50', 10);
let seed = seed0;
const rnd = () => {
  seed = (seed + 0x6D2B79F5) | 0;
  let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
  t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
  return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
};
const pick = (a) => a[Math.floor(rnd() * a.length)];
const CELLS = ['A1', 'A2', 'A3', 'B1', 'B2', 'B3', 'C1', 'C2'];
const ERRS = ['1/0', 'NA()', 'SQRT(-1)', '"x"+1', 'MATCH(9,D1:D2,0)'];
function leaf(d) {
  const r = rnd();
  if (r < 0.45) return pick(CELLS);
  if (r < 0.6) return pick(ERRS);
  if (r < 0.75) return String(Math.floor(rnd() * 5));
  if (r < 0.85) return pick(['A1:A3', 'B1:B3', 'A1:C1', 'A2:C2']);
  return pick(['D1', 'D2', '"t"', 'TRUE']);
}
function expr(d) {
  if (d <= 0 || rnd() < 0.25) { const l = leaf(d); return /:/.test(l) ? `SUM(${l})` : l; }
  const a = () => expr(d - 1);
  switch (Math.floor(rnd() * 16)) {
    case 0: case 1: case 2: return `${a()}${pick(['+', '-', '*', '&', '<'])}${a()}`;
    case 3: return `IFERROR(${a()},${a()})`;
    case 4: return `IF(${a()},${a()},${a()})`;
    case 5: return `ISERROR(${a()})`;
    case 6: return `ISNUMBER(${a()})`;
    case 7: return `COUNT(${pick(['A1:A3', 'B1:C2', a()])},${a()})`;
    case 8: return `SUM(${pick(['A1:A3', 'B1:C2', a()])},${a()})`;
    case 9: return `MAX(${pick(['A1:A3', 'B1:C2'])},${a()})`;
    case 10: return `ABS(${a()})`;
    case 11: return `VLOOKUP(${a()},A1:B3,2,0)`;
    case 12: return `INDEX(${pick(['A1:A3', 'B1:B3'])},${1 + Math.floor(rnd() * 3)})`;
    case 13: return `N(${a()})`;
    case 14: return `CONCATENATE(${a()},${a()})`;
    default: return `ROUND(${a()},0)`;
  }
}
const out = {};
for (let k = 0; k < count; k++) {
  const sh = { D1: 1, D2: 'z' };
  const n = 2 + Math.floor(rnd() * 5);
  const cells = CELLS.slice().sort(() => rnd() - 0.5).slice(0, n);
  for (const c of cells) sh[c] = '=' + expr(1 + Math.floor(rnd() * 3));
  for (const c of CELLS) if (!(c in sh) && rnd() < 0.5) sh[c] = Math.floor(rnd() * 4);
  out[`cyc${seed0}_${k}`] = { sheets: { Sheet1: sh } };
}
process.stdout.write(JSON.stringify(out, null, 1) + '\n');
