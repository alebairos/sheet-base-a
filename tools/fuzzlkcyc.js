#!/usr/bin/env node
'use strict';
// Dev-only: random one-sheet workbooks where lookups (VLOOKUP, HLOOKUP,
// MATCH, all match types) search a range holding their own cell or cells
// that read them, to compare which cells Calc reads during a search.
// Usage: node tools/fuzzlkcyc.js <seed> <count> > cases.json
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
const CONSTS = [0, 1, 2, 3, 5, 8, 10, -1, 2.5, 'a', 'b', 'm', 'z', 'Apple', true, false];
const QUERY = ['0', '1', '2', '3', '4', '6', '9', '11', '-5', '2.5', '"a"', '"c"', '"z"', '"zz"', 'TRUE', '1/0', 'NA()'];
function book() {
  const sh = {};
  const sorted = rnd() < 0.5;
  // data in A1:A8 and B1:B8 (sorted ascending, descending or random)
  let nums = [];
  for (let i = 0; i < 8; i++) nums.push(rnd() < 0.7 ? Math.floor(rnd() * 12) : pick(CONSTS));
  if (sorted) {
    const ns = nums.filter((x) => typeof x === 'number').sort((a, b) => a - b);
    nums = rnd() < 0.5 ? ns : ns.reverse();
  }
  nums.forEach((x, i) => { if (rnd() < 0.85) sh['A' + (i + 1)] = x; });
  for (let i = 1; i <= 8; i++) if (rnd() < 0.7) sh['B' + i] = pick(CONSTS);
  for (let i = 1; i <= 3; i++) sh['D' + i] = pick(CONSTS);
  const look = () => {
    const q = rnd() < 0.2 ? pick(['C1', 'C2', 'A' + (1 + Math.floor(rnd() * 8))]) : pick(QUERY);
    const k = rnd();
    const r1 = 1 + Math.floor(rnd() * 3), r2 = 5 + Math.floor(rnd() * 4);
    if (k < 0.4) return `MATCH(${q},A${r1}:A${r2},${pick([0, 1, -1, 1])})`;
    if (k < 0.8) return `VLOOKUP(${q},A${r1}:B${r2},${pick([1, 2])},${pick([0, 1, 'TRUE'])})`;
    return `HLOOKUP(${q},A${r1}:D${r1},${pick([1, 2])},${pick([0, 1])})`;
  };
  // formula cells inside the searched column (and outside it reading them)
  const n = 1 + Math.floor(rnd() * 3);
  for (let i = 0; i < n; i++) sh['A' + (1 + Math.floor(rnd() * 8))] = '=' + look();
  if (rnd() < 0.6) sh['A' + (1 + Math.floor(rnd() * 8))] = '=' + pick(['C1', 'C2', 'C1+1', 'B2*2', 'D1']);
  sh.C1 = '=' + look();
  sh.C2 = '=' + (rnd() < 0.5 ? look() : pick(['A3', 'A5+1', 'C1', 'SUM(A1:A3)', 'A' + (1 + Math.floor(rnd() * 8))]));
  if (rnd() < 0.5) sh['D' + (1 + Math.floor(rnd() * 3))] = '=' + look();
  return { sheets: { Sheet1: sh } };
}
const out = {};
for (let i = 0; i < count; i++) out[`lkc${seed0}_${i}`] = book();
console.log(JSON.stringify(out, null, 1));
