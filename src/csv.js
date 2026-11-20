/*
 * CSV <-> sheet cells, shared by the UI and the tests.
 * parseCSV(text) -> { A1: value, ... }  (numbers, booleans, formulas, text)
 * toCSV(cells, values?) -> text  (values: computed cells to export instead of raw)
 */
(function (root) {
  'use strict';
  const eng = typeof module !== 'undefined' && module.exports ? require('./engine.js') : root.SheetEngine;
  const fmtNum = eng && eng.fmtNum;

  function colName(c) {
    let s = '';
    c += 1;
    while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
    return s;
  }
  function colIndex(s) {
    let n = 0;
    for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }

  function parseRows(text) {
    const rows = [];
    let row = [], field = '', i = 0, quoted = false, any = false;
    if (text.charCodeAt(0) === 0xfeff) i = 1;
    const delim = guessDelimiter(text);
    for (; i < text.length; i++) {
      const ch = text[i];
      if (quoted) {
        if (ch === '"') {
          if (text[i + 1] === '"') { field += '"'; i++; } else quoted = false;
        } else field += ch;
        continue;
      }
      if (ch === '"') { quoted = true; any = true; continue; }
      if (ch === delim) { row.push(field); field = ''; any = true; continue; }
      if (ch === '\r') continue;
      if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; any = false; continue; }
      field += ch; any = true;
    }
    if (any || field !== '' || row.length) { row.push(field); rows.push(row); }
    return rows;
  }

  function guessDelimiter(text) {
    const line = text.split('\n', 1)[0];
    const count = (d) => line.split(d).length - 1;
    if (count(',') === 0 && count(';') > 0) return ';';
    if (count(',') === 0 && count('\t') > 0) return '\t';
    return ',';
  }

  function typed(s) {
    if (s === '') return null;
    if (s.charAt(0) === '=' && s.length > 1) return s;
    if (/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(s.trim())) return Number(s);
    const u = s.toUpperCase();
    if (u === 'TRUE') return true;
    if (u === 'FALSE') return false;
    return s;
  }

  function parseCSV(text) {
    const cells = {};
    parseRows(text).forEach((row, r) => {
      row.forEach((f, c) => {
        const v = typed(f);
        if (v !== null) cells[colName(c) + (r + 1)] = v;
      });
    });
    return cells;
  }

  function cellToText(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'object' && v.error) return v.error;
    // numbers as LibreOffice writes them: 15 significant digits, 1E+021
    if (typeof v === 'number') return fmtNum ? fmtNum(v) : String(v);
    return String(v);
  }

  function quote(s) {
    return /[",\n\r]/.test(s) || /^\s|\s$/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }

  function toCSV(cells) {
    let maxR = -1, maxC = -1;
    const pos = {};
    for (const a of Object.keys(cells)) {
      const m = /^([A-Z]+)(\d+)$/.exec(a);
      if (!m) continue;
      const r = parseInt(m[2], 10) - 1, c = colIndex(m[1]);
      pos[a] = [r, c];
      if (r > maxR) maxR = r;
      if (c > maxC) maxC = c;
    }
    const grid = [];
    for (let r = 0; r <= maxR; r++) grid.push(new Array(maxC + 1).fill(''));
    for (const [a, [r, c]] of Object.entries(pos)) grid[r][c] = quote(cellToText(cells[a]));
    return grid.map((row) => row.join(',')).join('\n') + (grid.length ? '\n' : '');
  }

  const api = { parseCSV, toCSV, parseRows, cellToText, typed };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.SheetCSV = api;
})(typeof window !== 'undefined' ? window : globalThis);
