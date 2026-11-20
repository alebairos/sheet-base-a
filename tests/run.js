#!/usr/bin/env node
'use strict';
// Runs every tests/cases/<suite>.json against tests/expected/<suite>.json.
// Expected values were computed by LibreOffice (tools/lo_oracle.py).
const fs = require('fs');
const path = require('path');
const { Worker, isMainThread } = require('worker_threads');
const { computeWorkbook } = require('../src/engine.js');

// run on a thread with the large stack bin/engine uses
if (isMainThread) {
  const w = new Worker(__filename, { argv: process.argv.slice(2), resourceLimits: { stackSizeMb: 256 } });
  w.on('exit', (code) => process.exit(code));
  return;
}

const dir = __dirname;
const only = process.argv[2];
let pass = 0, fail = 0;
const fails = [];

function same(a, b) {
  if (typeof a === 'number' && typeof b === 'number') {
    if (a === b) return true;
    return Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a), Math.abs(b));
  }
  return JSON.stringify(a) === JSON.stringify(b);
}

for (const f of fs.readdirSync(path.join(dir, 'cases')).sort()) {
  if (!f.endsWith('.json')) continue;
  const suite = f.slice(0, -5);
  if (only && suite !== only) continue;
  const cases = JSON.parse(fs.readFileSync(path.join(dir, 'cases', f), 'utf8'));
  const expFile = path.join(dir, 'expected', f);
  if (!fs.existsSync(expFile)) { console.error(`missing expected for ${suite}`); fail++; continue; }
  const expected = JSON.parse(fs.readFileSync(expFile, 'utf8'));
  for (const [name, wb] of Object.entries(cases)) {
    const exp = expected[name];
    if (!exp) { fails.push(`${suite} :: ${name}: no expected value`); fail++; continue; }
    let got;
    try { got = JSON.parse(JSON.stringify(computeWorkbook(wb))); } catch (e) { got = { crash: String(e.stack) }; }
    let ok = true;
    const diffs = [];
    for (const [sn, cells] of Object.entries(exp.sheets)) {
      for (const [a, v] of Object.entries(cells)) {
        const g = got.sheets && got.sheets[sn] ? got.sheets[sn][a] : undefined;
        if (!same(g, v)) { ok = false; diffs.push(`${sn}!${a}: got ${JSON.stringify(g)} want ${JSON.stringify(v)} (${JSON.stringify(wb.sheets[sn][a])})`); }
      }
    }
    if (got.crash) { ok = false; diffs.push(got.crash); }
    if (ok) pass++; else { fail++; fails.push(`${suite} :: ${name}\n    ` + diffs.join('\n    ')); }
  }
}
for (const f of fails) console.log('FAIL ' + f);
console.log(`${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
