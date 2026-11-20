#!/usr/bin/env node
'use strict';
// CLI contract checks: file input, stdin input, --cell.
const { execFileSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');
const assert = require('assert');

const bin = path.join(__dirname, '..', 'bin', 'engine');
const wb = { sheets: { Sheet1: { A1: 1, A2: 'text', A3: true, B1: '=SUM(A1:A2)', H1: '=A1/0', H2: '=A3' }, Sheet2: {} } };
const run = (args, input) => execFileSync(process.execPath, [bin, ...args], { input }).toString();

let n = 0;
const out = JSON.parse(run(['-'], JSON.stringify(wb)));
assert.deepStrictEqual(out, { sheets: { Sheet1: { A1: 1, A2: 'text', A3: true, B1: 1, H1: { error: '#DIV/0!' }, H2: true }, Sheet2: {} } }); n++;

const tmp = path.join(os.tmpdir(), `engine-cli-${process.pid}.json`);
fs.writeFileSync(tmp, JSON.stringify(wb));
try {
  assert.deepStrictEqual(JSON.parse(run([tmp])), out); n++;
} finally { fs.unlinkSync(tmp); }

assert.strictEqual(run(['-', '--cell', 'Sheet1!H1'], JSON.stringify(wb)).trim(), '{"error":"#DIV/0!"}'); n++;
assert.strictEqual(run(['-', '--cell', 'Sheet1!B1'], JSON.stringify(wb)).trim(), '1'); n++;
assert.strictEqual(run(['-', '--cell', 'Sheet1!A2'], JSON.stringify(wb)).trim(), '"text"'); n++;
assert.strictEqual(run(['-', '--cell', 'Sheet1!H2'], JSON.stringify(wb)).trim(), 'true'); n++;
assert.strictEqual(run(['-', '--cell', 'Sheet1!A1'], JSON.stringify({ sheets: { Sheet1: { a1: '=2' } } })).trim(), '2'); n++;
assert.strictEqual(run(['-', '--cell', "'My Sheet'!$B$1"], JSON.stringify({ sheets: { 'My Sheet': { B1: '=1+1' } } })).trim(), '2'); n++;
// integer-like sheet names keep their place in the workbook (3D references, default sheet)
const numbered = '{"sheets": {"Main": {"A1": 1, "B1": "=SUM(Main:\'2024\'!A1)"}, "2024": {"A1": 5}, "Last": {"A1": 100}}}';
assert.strictEqual(run(['-', '--cell', 'Main!B1'], numbered).trim(), '6'); n++;
assert.strictEqual(run(['-', '--cell', 'A1'], numbered).trim(), '1'); n++;
console.log(`cli: ${n} passed`);
