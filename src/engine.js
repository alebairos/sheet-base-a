/*
 * Spreadsheet engine: computes a JSON workbook the way LibreOffice Calc does.
 * Works in Node (module.exports) and in the browser (window.SheetEngine).
 *
 * Value model (scalars):
 *   number | string | boolean | null (blank, only via references) | FErr
 * Operands may also be Range or Matrix objects.
 */
(function (root) {
  'use strict';

  // ---------------------------------------------------------------- errors
  class FErr {
    constructor(code, fromRef) { this.code = code; this.fromRef = !!fromRef; }
  }
  const E = {};
  const EREF = {};
  for (const [k, code] of Object.entries({
    DIV0: '#DIV/0!', VALUE: '#VALUE!', REF: '#REF!', NAME: '#NAME?', NA: '#N/A', NUM: '#NUM!', NULL: '#NULL!',
  })) { E[k] = new FErr(code, false); EREF[code] = new FErr(code, true); }
  const errByCode = (code) => Object.values(E).find((e) => e.code === code) || new FErr(code);
  const isErr = (v) => v instanceof FErr;
  const asRefErr = (v) => (isErr(v) && !v.fromRef ? (EREF[v.code] || new FErr(v.code, true)) : v);
  const asComputed = (v) => (isErr(v) && v.fromRef ? errByCode(v.code) : v);

  // ---------------------------------------------------------------- helpers
  const MAXROW = 1048576, MAXCOL = 16384;
  const E48 = Math.pow(2, -48);

  function approxEqual(a, b) {
    if (a === b) return true;
    if (a === 0 || b === 0) return false;
    const d = Math.abs(a - b);
    return d < Math.abs(a) * E48 && d < Math.abs(b) * E48;
  }
  function approxAdd(a, b) {
    if (((a < 0 && b > 0) || (b < 0 && a > 0)) && approxEqual(a, -b)) return 0;
    return a + b;
  }
  function approxSub(a, b) {
    if (((a < 0 && b < 0) || (a > 0 && b > 0)) && approxEqual(a, b)) return 0;
    return a - b;
  }
  function approxValue(x) {
    if (!isFinite(x) || x === 0) return x;
    return Number(x.toPrecision(15));
  }
  const approxFloor = (x) => Math.floor(approxValue(x));
  const approxCeil = (x) => Math.ceil(approxValue(x));

  function colName(c) { // 0-based
    let s = '';
    c += 1;
    while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
    return s;
  }
  function colIndex(s) {
    let n = 0;
    for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }
  function parseAddr(a) {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(a);
    if (!m) return null;
    const r = parseInt(m[2], 10) - 1, c = colIndex(m[1]);
    if (r < 0 || r >= MAXROW || c >= MAXCOL) return null;
    return { r, c };
  }
  const addrName = (r, c) => colName(c) + (r + 1);

  // General number format as Calc shows it (used for number -> text).
  function fmtNum(x) {
    if (x === 0 || Object.is(x, -0)) return '0';
    if (!isFinite(x)) return '#NUM!';
    if (Number.isInteger(x) && Math.abs(x) < 1e16) return String(x);
    let s = x.toExponential(14); // 15 significant digits
    let [mant, exp] = s.split('e');
    let e = parseInt(exp, 10);
    if (e <= -15 || e >= 15) {
      mant = mant.indexOf('.') >= 0 ? mant.replace(/0+$/, '').replace(/\.$/, '') : mant;
      const sign = e < 0 ? '-' : '+';
      return mant + 'E' + sign + String(Math.abs(e)).padStart(3, '0');
    }
    const v = Number(x.toPrecision(15));
    // fixed notation with up to 15 significant digits
    let out;
    if (e >= 0) out = v.toFixed(Math.max(0, 14 - e));
    else out = v.toFixed(Math.min(100, 14 - e));
    if (out.indexOf('.') >= 0) out = out.replace(/0+$/, '').replace(/\.$/, '');
    return out;
  }

  // Text -> number, as Calc's unambiguous conversion (en-US).
  const NUMRE = /^([+-]?)(\$?)([+-]?)(\d{1,3}(?:,\d{3})+|\d*)(\.\d*)?(?:[eE]([+-]?\d+))?(%?)$/;
  function textToNum(s) {
    const t = s.trim();
    if (t === '') return null;
    const u = t.toUpperCase();
    if (u === 'TRUE') return 1;
    if (u === 'FALSE') return 0;
    const m = NUMRE.exec(t);
    if (!m) return null;
    if (m[1] && m[3]) return null;
    const intp = m[4].replace(/,/g, '');
    const frac = m[5] || '';
    if (intp === '' && frac.length <= 1) return null;
    if (m[6] !== undefined && m[2]) return null;
    let v = Number((intp || '0') + (frac === '.' ? '' : frac) + (m[6] !== undefined ? 'e' + m[6] : ''));
    if (!isFinite(v)) return null;
    if (m[7]) v /= 100;
    if (m[1] === '-' || m[3] === '-') v = -v;
    return v;
  }

  const collator = (typeof Intl !== 'undefined') ? new Intl.Collator('en', { sensitivity: 'accent' }) : null;
  function strCmp(a, b) {
    if (collator) { const r = collator.compare(a, b); return r < 0 ? -1 : r > 0 ? 1 : 0; }
    a = a.toLowerCase(); b = b.toLowerCase();
    return a < b ? -1 : a > b ? 1 : 0;
  }
  const strEq = (a, b) => a.toLowerCase() === b.toLowerCase() || strCmp(a, b) === 0;

  // ---------------------------------------------------------------- operands
  class Range {
    constructor(sheet, r1, c1, r2, c2) {
      this.sheet = sheet;
      this.r1 = Math.min(r1, r2); this.r2 = Math.max(r1, r2);
      this.c1 = Math.min(c1, c2); this.c2 = Math.max(c1, c2);
    }
    get rows() { return this.r2 - this.r1 + 1; }
    get cols() { return this.c2 - this.c1 + 1; }
  }
  class Matrix {
    constructor(rows) { this.data = rows; }
    get rows() { return this.data.length; }
    get cols() { return this.data[0].length; }
  }
  const MISSING = { missing: true };
  // result of calling an unknown function with arguments: #NAME?, but
  // IFERROR/ISERROR turn it into #VALUE! instead of catching it
  const UNKNOWN_FN = new FErr('#NAME?', false);
  // value of a formula whose result is an empty cell (e.g. =B1 with B1 empty):
  // shown as 0, but compares like an empty cell and is neither number nor text
  const EMPTYR = { emptyResult: true };

  // ---------------------------------------------------------------- tokenizer
  const ERRLITS = ['#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#N/A', '#NUM!', '#NULL!'];

  function tokenize(src) {
    const toks = [];
    let i = 0;
    const n = src.length;
    const refRe = /^(\$?[A-Za-z]{1,3}\$?\d+)(?::(\$?[A-Za-z]{1,3}\$?\d+))?/;
    const colRangeRe = /^(\$?[A-Za-z]{1,3}):(\$?[A-Za-z]{1,3})(?![\w(])/;
    const rowRangeRe = /^(\$?\d+):(\$?\d+)(?![\w(.])/;
    while (i < n) {
      const ch = src[i];
      if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') { i++; continue; }
      if (ch === '"') {
        let s = '';
        i++;
        for (;;) {
          if (i >= n) { const e = new SyntaxError('unterminated string'); e.code = '#NAME?'; throw e; }
          if (src[i] === '"') { if (src[i + 1] === '"') { s += '"'; i += 2; continue; } i++; break; }
          s += src[i++];
        }
        toks.push({ t: 'str', v: s });
        continue;
      }
      if (ch === '#') {
        const lit = ERRLITS.find((e) => src.substr(i, e.length).toUpperCase() === e);
        if (!lit) throw new SyntaxError('bad error literal');
        toks.push({ t: 'err', v: lit });
        i += lit.length;
        continue;
      }
      // sheet-qualified reference
      let sheet = null;
      let j = i;
      if (ch === "'") {
        let s = '';
        j++;
        for (;;) {
          if (j >= n) throw new SyntaxError('unterminated sheet name');
          if (src[j] === "'") { if (src[j + 1] === "'") { s += "'"; j += 2; continue; } j++; break; }
          s += src[j++];
        }
        if (src[j] !== '!') throw new SyntaxError('expected !');
        sheet = s; j++;
      } else {
        const m = /^([A-Za-z_\u00C0-\uFFFF][\w.\u00C0-\uFFFF]*)!/.exec(src.slice(i));
        if (m) { sheet = m[1]; j = i + m[0].length; }
      }
      const rest = src.slice(j);
      if (sheet !== null || /^\$?[A-Za-z]/.test(rest) || /^\$?\d/.test(rest)) {
        let m = refRe.exec(rest);
        if (m && !/^[\w(]/.test(rest.slice(m[0].length)) && !(sheet === null && /^[.]/.test(rest.slice(m[0].length)))) {
          const a = parseAddr(m[1]);
          const b = m[2] ? parseAddr(m[2]) : null;
          if (a && (!m[2] || b)) {
            toks.push({ t: 'ref', sheet, a, b, text: src.slice(i, j + m[0].length) });
            i = j + m[0].length;
            continue;
          }
        }
        m = colRangeRe.exec(rest);
        if (m) {
          const c1 = colIndex(m[1].replace('$', '')), c2 = colIndex(m[2].replace('$', ''));
          toks.push({ t: 'ref', sheet, a: { r: 0, c: c1 }, b: { r: MAXROW - 1, c: c2 }, text: src.slice(i, j + m[0].length) });
          i = j + m[0].length;
          continue;
        }
        if (sheet !== null || ch === '$') {
          m = rowRangeRe.exec(rest);
          if (m) {
            const r1 = parseInt(m[1].replace('$', ''), 10) - 1, r2 = parseInt(m[2].replace('$', ''), 10) - 1;
            toks.push({ t: 'ref', sheet, a: { r: r1, c: 0 }, b: { r: r2, c: MAXCOL - 1 }, text: src.slice(i, j + m[0].length) });
            i = j + m[0].length;
            continue;
          }
          if (sheet !== null) throw new SyntaxError('bad reference');
        }
      }
      if (/[0-9.]/.test(ch)) {
        const rm = rowRangeRe.exec(src.slice(i));
        if (rm) {
          const r1 = parseInt(rm[1].replace('$', ''), 10) - 1, r2 = parseInt(rm[2].replace('$', ''), 10) - 1;
          toks.push({ t: 'ref', sheet: null, a: { r: r1, c: 0 }, b: { r: r2, c: MAXCOL - 1 }, text: rm[0] });
          i += rm[0].length;
          continue;
        }
        const m = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/.exec(src.slice(i));
        if (!m) throw new SyntaxError('bad number');
        toks.push({ t: 'num', v: Number(m[0]) });
        i += m[0].length;
        continue;
      }
      if (/[A-Za-z_\u00C0-\uFFFF]/.test(ch)) {
        const m = /^[A-Za-z_\u00C0-\uFFFF][\w.\u00C0-\uFFFF]*/.exec(src.slice(i));
        toks.push({ t: 'name', v: m[0] });
        i += m[0].length;
        continue;
      }
      const two = src.substr(i, 2);
      if (two === '<=' || two === '>=' || two === '<>') { toks.push({ t: 'op', v: two }); i += 2; continue; }
      if ('+-*/^&=<>%(),;{}:!'.indexOf(ch) >= 0) { toks.push({ t: 'op', v: ch }); i++; continue; }
      throw new SyntaxError('unexpected character ' + ch);
    }
    toks.push({ t: 'eof' });
    return toks;
  }

  // ---------------------------------------------------------------- parser
  // Precedence (low -> high): comparison, &, + -, * /, ^, unary +-, %, :
  function parse(src) {
    const toks = tokenize(src);
    let p = 0;
    const peek = () => toks[p];
    const isOp = (v) => toks[p].t === 'op' && toks[p].v === v;
    // a missing ')' at the end is auto-closed, as Calc does
    const expect = (v) => {
      if (v === ')' && toks[p].t === 'eof') return;
      if (!isOp(v)) throw new SyntaxError('expected ' + v);
      p++;
    };

    function comparison() {
      let a = concat();
      for (;;) {
        const t = peek();
        if (t.t === 'op' && ['=', '<>', '<', '>', '<=', '>='].indexOf(t.v) >= 0) {
          p++; a = { t: 'bin', op: t.v, a, b: concat() };
        } else return a;
      }
    }
    function concat() {
      let a = addsub();
      while (isOp('&')) { p++; a = { t: 'bin', op: '&', a, b: addsub() }; }
      return a;
    }
    function addsub() {
      let a = muldiv();
      while (isOp('+') || isOp('-')) { const op = toks[p++].v; a = { t: 'bin', op, a, b: muldiv() }; }
      return a;
    }
    function muldiv() {
      let a = pow();
      while (isOp('*') || isOp('/')) { const op = toks[p++].v; a = { t: 'bin', op, a, b: pow() }; }
      return a;
    }
    function pow() {
      let a = unary();
      while (isOp('^')) { p++; a = { t: 'bin', op: '^', a, b: unary() }; }
      return a;
    }
    function unary() {
      if (isOp('-')) { p++; return { t: 'neg', a: unary() }; }
      if (isOp('+')) { p++; return unary(); }
      return postfix();
    }
    function postfix() {
      let a = rangeop();
      while (isOp('%')) { p++; a = { t: 'pct', a }; }
      return a;
    }
    function rangeop() {
      let a = primary();
      while (isOp(':')) { p++; a = { t: 'rangeop', a, b: primary() }; }
      return a;
    }
    function args() {
      const list = [];
      if (isOp(')')) { p++; return list; }
      for (;;) {
        if (isOp(',')) { list.push({ t: 'missing' }); p++; continue; }
        if (isOp(')')) { list.push({ t: 'missing' }); p++; return list; }
        if (toks[p].t === 'eof') return list;
        list.push(comparison());
        if (isOp(',')) { p++; if (isOp(')')) { list.push({ t: 'missing' }); p++; return list; } continue; }
        expect(')');
        return list;
      }
    }
    function arrayConst() {
      const rows = [[]];
      for (;;) {
        let neg = false;
        while (isOp('-') || isOp('+')) { if (toks[p].v === '-') neg = !neg; p++; }
        const t = toks[p++];
        let v;
        if (t.t === 'num') v = neg ? -t.v : t.v;
        else if (t.t === 'str' && !neg) v = t.v;
        else if (t.t === 'err' && !neg) v = errByCode(t.v);
        else if (t.t === 'name' && !neg && /^(TRUE|FALSE)$/i.test(t.v)) v = t.v.toUpperCase() === 'TRUE';
        else throw new SyntaxError('bad array constant');
        rows[rows.length - 1].push(v);
        if (isOp(',')) { p++; continue; }
        if (isOp(';')) { p++; rows.push([]); continue; }
        expect('}');
        break;
      }
      const w = rows[0].length;
      if (rows.some((r) => r.length !== w)) throw new SyntaxError('ragged array');
      return { t: 'arr', v: new Matrix(rows) };
    }
    function primary() {
      const t = toks[p++];
      switch (t.t) {
        case 'num': return { t: 'num', v: t.v };
        case 'str': return { t: 'str', v: t.v };
        case 'err': return { t: 'errlit', v: errByCode(t.v) };
        case 'ref': return { t: 'ref', sheet: t.sheet, a: t.a, b: t.b };
        case 'name': {
          if (isOp('(')) { p++; return { t: 'fn', name: t.v.toUpperCase().replace(/^_XLFN\./, ''), args: args() }; }
          const u = t.v.toUpperCase();
          if (u === 'TRUE') return { t: 'bool', v: true };
          if (u === 'FALSE') return { t: 'bool', v: false };
          return { t: 'name', v: t.v };
        }
        case 'op':
          if (t.v === '(') { const e = comparison(); expect(')'); return { t: 'paren', a: e }; }
          if (t.v === '{') return arrayConst();
          if (t.v === '-') return { t: 'neg', a: unary() };
          if (t.v === '+') return unary();
          break;
      }
      const e = new SyntaxError('unexpected token');
      if (t.t === 'eof') e.code = '#N/A';
      throw e;
    }
    const ast = comparison();
    if (peek().t !== 'eof') throw new SyntaxError('trailing input');
    return ast;
  }

  // ---------------------------------------------------------------- workbook
  class Workbook {
    constructor(json) {
      this.sheets = [];
      this.byName = new Map();
      const sheets = (json && json.sheets) || {};
      for (const name of Object.keys(sheets)) {
        const sh = { name, cells: new Map(), maxR: -1, maxC: -1 };
        this.sheets.push(sh);
        this.byName.set(name.toUpperCase(), sh);
      }
      for (const sh of this.sheets) {
        const cells = sheets[sh.name] || {};
        for (const [a, raw] of Object.entries(cells)) {
          const pa = parseAddr(a);
          if (!pa) continue;
          this.setCell(sh, pa.r, pa.c, raw, a);
        }
      }
    }
    setCell(sh, r, c, raw, addr) {
      const key = r * MAXCOL + c;
      // an empty text constant is an empty cell (as after an xlsx round trip)
      if (raw === null || raw === undefined || raw === '') { sh.cells.delete(key); return; }
      const cell = { r, c, raw, addr: addr || addrName(r, c), state: 0, value: null, ast: null };
      if (typeof raw === 'string' && raw.charAt(0) === '=' && raw.length > 1) {
        cell.formula = true;
        try { cell.ast = parse(raw.slice(1)); } catch (e) { cell.parseError = e.code || '#VALUE!'; }
      } else if (raw !== null && typeof raw === 'object') {
        cell.value = raw.error ? errByCode(raw.error) : null;
        cell.state = 2;
      } else {
        cell.value = raw;
        cell.state = 2;
      }
      sh.cells.set(key, cell);
      if (r > sh.maxR) sh.maxR = r;
      if (c > sh.maxC) sh.maxC = c;
    }
    sheet(name) { return this.byName.get(String(name).toUpperCase()) || null; }
  }

  // ---------------------------------------------------------------- evaluator
  class Engine {
    constructor(json) {
      this.wb = new Workbook(json);
    }

    // value of a cell (evaluating formulas), null for blank
    cellValue(sh, r, c) {
      const cell = sh.cells.get(r * MAXCOL + c);
      if (!cell) return null;
      if (cell.state === 2) return cell.value;
      if (cell.state === 1) { cell.circular = true; return E.VALUE; }
      this.evalCell(sh, cell);
      return cell.value;
    }

    evalCell(sh, cell) {
      cell.state = 1;
      let v;
      if (cell.parseError || !cell.ast) v = cell.parseError ? errByCode(cell.parseError) : null;
      else {
        const ctx = { sheet: sh, r: cell.r, c: cell.c };
        try {
          v = this.finalValue(this.ev(cell.ast, ctx), ctx);
        } catch (e) {
          if (e instanceof RangeError) v = E.VALUE; // stack overflow etc.
          else throw e;
        }
      }
      if (cell.circular) v = E.VALUE;
      if (isErr(v)) v = asComputed(v);
      cell.value = v;
      cell.state = 2;
    }


    finalValue(v, ctx) {
      if (v instanceof Range) v = this.intersect(v, ctx);
      else if (v instanceof Matrix) v = v.data[0][0];
      if (v === null) return EMPTYR;
      if (v === MISSING) return 0;
      if (typeof v === 'number') {
        if (!isFinite(v)) return E.NUM;
        return v;
      }
      return v;
    }

    // implicit intersection of a range with the formula cell position
    intersect(rg, ctx) {
      if (rg.rows === 1 && rg.cols === 1) return asRefErr(this.cellValue(rg.sheet, rg.r1, rg.c1));
      if (rg.cols === 1 && ctx && ctx.sheet === rg.sheet && ctx.r >= rg.r1 && ctx.r <= rg.r2) {
        return asRefErr(this.cellValue(rg.sheet, ctx.r, rg.c1));
      }
      if (rg.rows === 1 && ctx && ctx.sheet === rg.sheet && ctx.c >= rg.c1 && ctx.c <= rg.c2) {
        return asRefErr(this.cellValue(rg.sheet, rg.r1, ctx.c));
      }
      // ranges from another sheet: intersect by position as well
      if (rg.cols === 1 && ctx && ctx.sheet !== rg.sheet && ctx.r >= rg.r1 && ctx.r <= rg.r2) {
        return asRefErr(this.cellValue(rg.sheet, ctx.r, rg.c1));
      }
      if (rg.rows === 1 && ctx && ctx.sheet !== rg.sheet && ctx.c >= rg.c1 && ctx.c <= rg.c2) {
        return asRefErr(this.cellValue(rg.sheet, rg.r1, ctx.c));
      }
      return E.VALUE;
    }

    // scalar value of an operand
    scalar(v, ctx) {
      if (v instanceof Range) return this.intersect(v, ctx);
      if (v instanceof Matrix) return v.data[0][0];
      return v;
    }

    sv(node, ctx) { return this.scalar(this.ev(node, ctx), ctx); }

    // iterate a range column-major (as Calc does); fn(value, r, c) may return false to stop
    eachCell(rg, fn, includeBlank) {
      const sh = rg.sheet;
      const r2 = Math.min(rg.r2, sh.maxR), c2 = Math.min(rg.c2, sh.maxC);
      const area = (r2 - rg.r1 + 1) * (c2 - rg.c1 + 1);
      if (!includeBlank && area > 4 * sh.cells.size + 64) {
        const list = [];
        for (const cell of sh.cells.values()) {
          if (cell.r >= rg.r1 && cell.r <= r2 && cell.c >= rg.c1 && cell.c <= c2) list.push(cell);
        }
        list.sort((x, y) => x.c - y.c || x.r - y.r);
        for (const cell of list) {
          const v = this.cellValue(sh, cell.r, cell.c);
          if (fn(v === EMPTYR ? 0 : v, cell.r, cell.c) === false) return;
        }
        return;
      }
      for (let c = rg.c1; c <= c2; c++) {
        for (let r = rg.r1; r <= r2; r++) {
          let v = this.cellValue(sh, r, c);
          if (v === EMPTYR) v = 0;
          if (v === null && !includeBlank) continue;
          if (fn(v, r, c) === false) return;
        }
      }
    }

    // cells of a range/matrix as a 2D array (clamped to used area for ranges)
    grid(v) {
      if (v instanceof Matrix) return v.data;
      if (v instanceof Range) {
        const r2 = Math.min(v.r2, Math.max(v.sheet.maxR, v.r1)), c2 = Math.min(v.c2, Math.max(v.sheet.maxC, v.c1));
        const out = [];
        for (let r = v.r1; r <= r2; r++) {
          const row = [];
          for (let c = v.c1; c <= c2; c++) { const x = this.cellValue(v.sheet, r, c); row.push(x === EMPTYR ? 0 : x); }
          out.push(row);
        }
        return out;
      }
      return [[v]];
    }

    resolveSheet(name, ctx) {
      if (name === null || name === undefined) return ctx.sheet;
      return this.wb.sheet(name);
    }

    ev(node, ctx) {
      switch (node.t) {
        case 'num': return node.v;
        case 'str': return node.v;
        case 'bool': return node.v;
        case 'errlit': return node.v;
        case 'arr': return node.v;
        case 'missing': return MISSING;
        case 'paren': return this.ev(node.a, ctx);
        case 'name': return E.NAME;
        case 'ref': {
          const sh = this.resolveSheet(node.sheet, ctx);
          if (!sh) return E.NAME;
          const b = node.b || node.a;
          return new Range(sh, node.a.r, node.a.c, b.r, b.c);
        }
        case 'rangeop': {
          const a = this.ev(node.a, ctx), b = this.ev(node.b, ctx);
          if (isErr(a)) return a;
          if (isErr(b)) return b;
          if (!(a instanceof Range) || !(b instanceof Range) || a.sheet !== b.sheet) return E.VALUE;
          return new Range(a.sheet, Math.min(a.r1, b.r1), Math.min(a.c1, b.c1), Math.max(a.r2, b.r2), Math.max(a.c2, b.c2));
        }
        case 'neg': {
          const v = this.sv(node.a, ctx);
          const n = this.toNum(v);
          if (isErr(n)) return n;
          return n === 0 ? 0 : -n;
        }
        case 'pct': {
          const v = this.sv(node.a, ctx);
          const n = this.toNum(v);
          if (isErr(n)) return n;
          return n / 100;
        }
        case 'bin': return this.binop(node, ctx);
        case 'fn': return this.call(node, ctx);
      }
      return E.VALUE;
    }

    // ---- coercions
    toNum(v) {
      if (typeof v === 'number') return v;
      if (typeof v === 'boolean') return v ? 1 : 0;
      if (v === null || v === MISSING || v === EMPTYR) return 0;
      if (isErr(v)) return v;
      if (typeof v === 'string') {
        const n = textToNum(v);
        return n === null ? E.VALUE : n;
      }
      return E.VALUE;
    }
    toStr(v) {
      if (typeof v === 'string') return v;
      if (typeof v === 'number') return fmtNum(v);
      if (typeof v === 'boolean') return v ? '1' : '0';
      if (v === null || v === MISSING) return '';
      if (v === EMPTYR) return '0';
      if (isErr(v)) return v;
      return E.VALUE;
    }
    toBool(v) {
      if (typeof v === 'boolean') return v;
      if (typeof v === 'number') return v !== 0;
      if (v === null || v === MISSING || v === EMPTYR) return false;
      if (isErr(v)) return v;
      if (typeof v === 'string') {
        const n = textToNum(v);
        if (n !== null) return n !== 0;
        const u = v.toUpperCase();
        if (u === 'TRUE') return true;
        if (u === 'FALSE') return false;
        return E.VALUE;
      }
      return E.VALUE;
    }

    // ---- operators
    binop(node, ctx) {
      const a = this.sv(node.a, ctx), b = this.sv(node.b, ctx);
      // errors produced by computation (left to right) win over
      // reference/conversion errors, which are checked right operand first.
      if (isErr(a) && !a.fromRef) return a;
      if (isErr(b) && !b.fromRef) return b;
      const op = node.op;
      if (op === '&') {
        const sb = this.toStr(b); if (isErr(sb)) return sb;
        const sa = this.toStr(a); if (isErr(sa)) return sa;
        return sa + sb;
      }
      if (op === '=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=') {
        if (isErr(b)) return b;
        if (isErr(a)) return a;
        const c = compare(a, b);
        switch (op) {
          case '=': return c === 0;
          case '<>': return c !== 0;
          case '<': return c < 0;
          case '>': return c > 0;
          case '<=': return c <= 0;
          default: return c >= 0;
        }
      }
      const y = this.toNum(b); if (isErr(y)) return y;
      const x = this.toNum(a); if (isErr(x)) return x;
      return arith(op, x, y);
    }

    // ---- function calls
    call(node, ctx) {
      const f = FUNCS[node.name];
      if (!f) return node.args.length ? UNKNOWN_FN : E.NAME;
      const n = node.args.length;
      if (n < f.min || (f.max !== undefined && n > f.max)) return E.VALUE;
      return f.fn.call(this, node.args, ctx);
    }

    // Evaluate fixed scalar arguments with Calc's error order:
    // computed errors left to right, then conversion right to left.
    // kinds: 'n' number, 's' string, 'b' bool, 'v' any scalar
    args(nodes, ctx, kinds, defaults) {
      const vals = [];
      for (let i = 0; i < kinds.length; i++) {
        if (i >= nodes.length) { vals.push(undefined); continue; }
        vals.push(this.sv(nodes[i], ctx));
      }
      for (const v of vals) if (isErr(v) && !v.fromRef) return v;
      const out = new Array(kinds.length);
      for (let i = kinds.length - 1; i >= 0; i--) {
        let v = vals[i];
        if (v === undefined || v === MISSING) {
          if (defaults && defaults[i] !== undefined) { out[i] = defaults[i]; continue; }
          if (v === undefined) { out[i] = undefined; continue; }
          v = null;
        }
        let r;
        switch (kinds[i]) {
          case 'n': r = this.toNum(v); break;
          case 's': r = this.toStr(v); break;
          case 'b': r = this.toBool(v); break;
          default: r = v;
        }
        if (isErr(r)) return r;
        out[i] = r;
      }
      return out;
    }

    // Collect numbers for aggregate functions.
    // opts.countText: numeric strings given directly count (COUNT); otherwise #VALUE!
    // opts.skipErrors: errors are ignored (COUNT)
    collect(nodes, ctx, opts) {
      const nums = [];
      let allBool = true;
      for (const node of nodes) {
        if (node.t === 'missing') continue;
        const v = this.ev(node, ctx);
        if (v instanceof Range) {
          let err = null;
          this.eachCell(v, (x) => {
            if (typeof x === 'number') { nums.push(x); allBool = false; } else if (typeof x === 'boolean') nums.push(x ? 1 : 0);
            else if (isErr(x) && !opts.skipErrors) { err = x; return false; }
            return true;
          });
          if (err) return { err };
        } else if (v instanceof Matrix) {
          for (const row of v.data) for (const x of row) {
            if (typeof x === 'number') { nums.push(x); allBool = false; } else if (typeof x === 'boolean') nums.push(x ? 1 : 0);
            else if (isErr(x) && !opts.skipErrors) return { err: x };
          }
        } else if (typeof v === 'number') { nums.push(v); allBool = false; } else if (typeof v === 'boolean') nums.push(v ? 1 : 0);
        else if (isErr(v)) { if (!opts.skipErrors) return { err: v }; } else if (typeof v === 'string') {
          const n = textToNum(v);
          if (opts.skipErrors) { if (n !== null) { nums.push(n); allBool = false; } } else if (n === null || !opts.countText) return { err: E.VALUE };
          else { nums.push(n); allBool = false; }
        }
      }
      return { nums, allBool: allBool && nums.length > 0 };
    }
  }

  function arith(op, x, y) {
    let r;
    switch (op) {
      case '+': r = approxAdd(x, y); break;
      case '-': r = approxSub(x, y); break;
      case '*': r = x * y; break;
      case '/': if (y === 0) return E.DIV0; r = x / y; break;
      case '^': return power(x, y);
    }
    if (!isFinite(r)) return E.NUM;
    return r;
  }

  function power(x, y) {
    if (x === 0 && y < 0) return E.NUM;
    if (x === 0 && y === 0) return 1;
    let r = Math.pow(x, y);
    if (isNaN(r) && x < 0) {
      // odd roots of negative numbers
      const inv = 1 / y;
      if (Math.abs(inv - Math.round(inv)) < 1e-10 && Math.round(inv) % 2 !== 0) r = -Math.pow(-x, y);
      else return E.NUM;
    }
    if (!isFinite(r) || isNaN(r)) return E.NUM;
    return r;
  }

  // compare two non-error scalars as Calc's comparison operators do
  function compare(a, b) {
    if (a === EMPTYR) a = null;
    if (b === EMPTYR) b = null;
    if (a === null || a === MISSING) a = typeof b === 'string' ? '' : 0;
    if (b === null || b === MISSING) b = typeof a === 'string' ? '' : 0;
    if (typeof a === 'boolean') a = a ? 1 : 0;
    if (typeof b === 'boolean') b = b ? 1 : 0;
    const na = typeof a === 'number', nb = typeof b === 'number';
    if (na && nb) return approxEqual(a, b) ? 0 : a < b ? -1 : 1;
    if (na) return -1;
    if (nb) return 1;
    return strCmp(a, b);
  }

  // ---------------------------------------------------------------- criteria (SUMIF etc.)
  function wildcardRe(pat) {
    let s = '';
    for (let i = 0; i < pat.length; i++) {
      const ch = pat[i];
      if (ch === '~' && i + 1 < pat.length && '*?~'.indexOf(pat[i + 1]) >= 0) { s += '\\' + pat[++i]; continue; }
      if (ch === '*') s += '[\\s\\S]*';
      else if (ch === '?') s += '[\\s\\S]';
      else s += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
    return new RegExp('^' + s + '$', 'i');
  }
  const hasWild = (s) => /[*?~]/.test(s);

  function cellText(v) {
    if (typeof v === 'string') return v;
    if (typeof v === 'number') return fmtNum(v);
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (isErr(v)) return v.code;
    return '';
  }

  function makeCriterion(c) {
    if (isErr(c)) return c;
    if (c === null || c === MISSING) c = 0;
    if (typeof c === 'boolean') c = c ? 1 : 0;
    if (typeof c === 'number') {
      return (v) => (typeof v === 'number' || typeof v === 'boolean') && approxEqual(+v, c);
    }
    let s = String(c), op = '=', explicitEq = false;
    for (const o of ['<=', '>=', '<>', '<', '>', '=']) {
      if (s.startsWith(o)) { op = o; s = s.slice(o.length); explicitEq = o === '='; break; }
    }
    const num = textToNum(s);
    if (num !== null) {
      const numMatch = (v) => (typeof v === 'number' || typeof v === 'boolean');
      switch (op) {
        case '=': return (v) => (numMatch(v) ? approxEqual(+v, num) : typeof v === 'string' && strEq(v, s));
        case '<>': return (v) => !(numMatch(v) ? approxEqual(+v, num) : typeof v === 'string' && strEq(v, s));
        case '<': return (v) => numMatch(v) && +v < num && !approxEqual(+v, num);
        case '>': return (v) => numMatch(v) && +v > num && !approxEqual(+v, num);
        case '<=': return (v) => numMatch(v) && (+v < num || approxEqual(+v, num));
        case '>=': return (v) => numMatch(v) && (+v > num || approxEqual(+v, num));
      }
    }
    if (op === '=' || op === '<>') {
      let match;
      if (s === '') match = op === '=' && !explicitEq ? (v) => v === null || v === '' : (v) => v === null;
      else if (hasWild(s)) { const re = wildcardRe(s); match = (v) => v !== null && re.test(cellText(v)); }
      else match = (v) => v !== null && strEq(cellText(v), s);
      if (op === '=') return match;
      if (s === '') return (v) => v !== null;
      return (v) => !match(v);
    }
    return (v) => {
      if (typeof v !== 'string') return false;
      const r = strCmp(v, s);
      switch (op) {
        case '<': return r < 0;
        case '>': return r > 0;
        case '<=': return r <= 0;
        default: return r >= 0;
      }
    };
  }

  // ---------------------------------------------------------------- lookup helpers
  function lookupEq(v, x) {
    if (x === null) return false;
    if (typeof v === 'boolean') v = v ? 1 : 0;
    if (typeof x === 'boolean') x = x ? 1 : 0;
    if (typeof v === 'number') return typeof x === 'number' && approxEqual(v, x);
    if (typeof v === 'string') {
      if (typeof x !== 'string') return false;
      if (hasWild(v)) return wildcardRe(v).test(x);
      return strEq(v, x);
    }
    return false;
  }
  // -1/0/1 comparing cell x to lookup v; null if different type
  function lookupCmp(x, v) {
    if (x === null || isErr(x)) return null;
    if (typeof v === 'boolean') v = v ? 1 : 0;
    if (typeof x === 'boolean') x = x ? 1 : 0;
    if (typeof v === 'number') {
      if (typeof x !== 'number') return null;
      return approxEqual(x, v) ? 0 : x < v ? -1 : 1;
    }
    if (typeof x !== 'string') return null;
    return strCmp(x, v);
  }
  // position (0-based) in a 1-D list, or -1
  function lookupPos(list, v, mode) {
    if (mode === 0) {
      for (let i = 0; i < list.length; i++) if (lookupEq(v, list[i])) return i;
      return -1;
    }
    let best = -1;
    if (mode > 0) {
      for (let i = 0; i < list.length; i++) {
        const c = lookupCmp(list[i], v);
        if (c === null) continue;
        if (c <= 0) best = i; else break;
      }
    } else {
      for (let i = 0; i < list.length; i++) {
        const c = lookupCmp(list[i], v);
        if (c === null) continue;
        if (c >= 0) best = i; else break;
      }
    }
    return best;
  }

  function stripBlank(v) { return v === null ? 0 : v; }

  // ---------------------------------------------------------------- functions
  const FUNCS = {};
  function def(name, min, max, fn) { FUNCS[name] = { min, max, fn }; }

  function aggregate(name, fn, opts, min) {
    def(name, min === undefined ? 0 : min, undefined, function (nodes, ctx) {
      const r = this.collect(nodes, ctx, opts || {});
      if (r.err) return r.err;
      return fn(r.nums, r.allBool);
    });
  }

  const kahan = (nums) => {
    let s = 0, comp = 0;
    for (const x of nums) {
      const t = s + x;
      if (Math.abs(s) >= Math.abs(x)) comp += (s - t) + x; else comp += (x - t) + s;
      s = t;
    }
    return approxAdd(s, comp);
  };
  const finite = (x) => (isFinite(x) ? x : E.NUM);

  aggregate('SUM', (n) => finite(kahan(n)));
  aggregate('AVERAGE', (n) => (n.length ? finite(kahan(n) / n.length) : E.DIV0));
  aggregate('MIN', (n, b) => (n.length ? (b ? !!Math.min(...n) : Math.min(...n)) : 0), {}, 1);
  aggregate('MAX', (n, b) => (n.length ? (b ? !!Math.max(...n) : Math.max(...n)) : 0), {}, 1);
  aggregate('PRODUCT', (n) => (n.length ? finite(n.reduce((a, b) => a * b, 1)) : 0));
  aggregate('COUNT', (n) => n.length, { skipErrors: true, countText: true });
  aggregate('MEDIAN', (n) => {
    if (!n.length) return E.VALUE;
    const s = n.slice().sort((a, b) => a - b);
    const m = s.length >> 1;
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
  }, {}, 1);
  aggregate('STDEV', (n) => {
    if (n.length < 2) return E.DIV0;
    const mean = kahan(n) / n.length;
    const ss = kahan(n.map((x) => (x - mean) * (x - mean)));
    return Math.sqrt(ss / (n.length - 1));
  });

  def('COUNTA', 0, undefined, function (nodes, ctx) {
    let n = 0;
    for (const node of nodes) {
      if (node.t === 'missing') continue;
      const v = this.ev(node, ctx);
      if (v instanceof Range) this.eachCell(v, () => { n++; });
      else if (v instanceof Matrix) n += v.rows * v.cols;
      else if (v !== null) n++;
    }
    return n;
  });

  def('ROUND', 1, 2, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n', 'n'], [undefined, 0]); if (isErr(a)) return a;
    return roundTo(a[0], a[1], 0);
  });
  def('ROUNDUP', 1, 2, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n', 'n'], [undefined, 0]); if (isErr(a)) return a;
    return roundTo(a[0], a[1], 1);
  });
  def('ROUNDDOWN', 1, 2, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n', 'n'], [undefined, 0]); if (isErr(a)) return a;
    return roundTo(a[0], a[1], -1);
  });
  function roundTo(x, d, mode) {
    d = Math.trunc(d);
    if (x === 0) return 0;
    if (d > 20) d = 20;
    if (d < -308) return 0;
    const m = Math.pow(10, Math.abs(d));
    let v = d >= 0 ? x * m : x / m;
    if (!isFinite(v)) return x;
    v = approxValue(v);
    const s = v < 0 ? -1 : 1;
    let a = Math.abs(v);
    if (mode === 0) a = Math.floor(a + 0.5);
    else if (mode > 0) a = Math.ceil(a);
    else a = Math.floor(a);
    a *= s;
    const r = d >= 0 ? a / m : a * m;
    return r === 0 ? 0 : r;
  }
  def('INT', 1, 1, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n']); if (isErr(a)) return a;
    return approxFloor(a[0]);
  });
  def('MOD', 2, 2, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n', 'n']); if (isErr(a)) return a;
    const [x, y] = a;
    if (y === 0) return E.DIV0;
    const q = x / y;
    if (!isFinite(q)) return E.NUM;
    let r = approxSub(x, approxFloor(q) * y);
    if (r !== 0 && (r < 0) !== (y < 0)) r += y;
    return r;
  });
  def('ABS', 1, 1, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n']); if (isErr(a)) return a;
    return Math.abs(a[0]);
  });
  def('POWER', 2, 2, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n', 'n']); if (isErr(a)) return a;
    return power(a[0], a[1]);
  });
  def('SQRT', 1, 1, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n']); if (isErr(a)) return a;
    if (a[0] < 0) return E.VALUE;
    return Math.sqrt(a[0]);
  });

  // ---- conditional aggregates
  function condArgs(self, nodes, ctx) {
    const rg = self.ev(nodes[0], ctx);
    if (isErr(rg)) return rg;
    if (!(rg instanceof Range)) return E.VALUE;
    let crit = self.sv(nodes[1], ctx);
    if (crit === MISSING) crit = null;
    const match = makeCriterion(crit);
    if (isErr(match)) return match;
    let sum = rg;
    if (nodes.length > 2 && nodes[2].t !== 'missing') {
      const s = self.ev(nodes[2], ctx);
      if (isErr(s)) return s;
      if (!(s instanceof Range)) return E.VALUE;
      sum = new Range(s.sheet, s.r1, s.c1, s.r1 + rg.rows - 1, s.c1 + rg.cols - 1);
    }
    return { rg, match, sum };
  }
  function condCollect(self, a) {
    const nums = [];
    let err = null;
    const blankMatches = a.match(null);
    const visit = (v, r, c) => {
      if (!a.match(v)) return true;
      let sv = self.cellValue(a.sum.sheet, a.sum.r1 + (r - a.rg.r1), a.sum.c1 + (c - a.rg.c1));
      if (sv === EMPTYR) sv = 0;
      if (typeof sv === 'number') nums.push(sv);
      else if (typeof sv === 'boolean') nums.push(sv ? 1 : 0);
      else if (isErr(sv)) { err = sv; return false; }
      return true;
    };
    self.eachCell(a.rg, visit, blankMatches && a.sum !== a.rg);
    return err ? { err } : { nums };
  }
  def('SUMIF', 2, 3, function (nodes, ctx) {
    const a = condArgs(this, nodes, ctx); if (isErr(a)) return a;
    const r = condCollect(this, a); if (r.err) return r.err;
    return kahan(r.nums);
  });
  def('AVERAGEIF', 2, 3, function (nodes, ctx) {
    const a = condArgs(this, nodes, ctx); if (isErr(a)) return a;
    const r = condCollect(this, a); if (r.err) return r.err;
    if (!r.nums.length) return E.DIV0;
    return kahan(r.nums) / r.nums.length;
  });
  def('COUNTIF', 2, 2, function (nodes, ctx) {
    const a = condArgs(this, nodes, ctx); if (isErr(a)) return a;
    let n = 0;
    const rg = a.rg, sh = rg.sheet;
    if (a.match(null)) {
      // blanks match: count over the whole area, including cells beyond the used range
      const total = rg.rows * rg.cols;
      let nonBlank = 0, nonBlankMatch = 0;
      this.eachCell(rg, (v) => { nonBlank++; if (a.match(v)) nonBlankMatch++; });
      return total - nonBlank + nonBlankMatch;
    }
    this.eachCell(rg, (v) => { if (a.match(v)) n++; });
    void sh;
    return n;
  });

  // ---- logic
  def('IF', 1, 3, function (nodes, ctx) {
    const c = this.sv(nodes[0], ctx);
    if (isErr(c)) return c;
    const b = this.toBool(c);
    if (isErr(b)) return b;
    if (b) {
      if (nodes.length < 2) return true;
      const v = this.ev(nodes[1], ctx);
      return v === MISSING ? 0 : v;
    }
    if (nodes.length < 3) return false;
    const v = this.ev(nodes[2], ctx);
    return v === MISSING ? 0 : v;
  });
  def('IFERROR', 2, 2, function (nodes, ctx) {
    const v = this.ev(nodes[0], ctx);
    const s = this.scalar(v, ctx);
    if (s === UNKNOWN_FN) return E.VALUE;
    if (isErr(s)) { const alt = this.ev(nodes[1], ctx); return alt === MISSING ? 0 : alt; }
    return v instanceof Range || v instanceof Matrix ? stripBlank(s) : stripBlank(v === MISSING ? 0 : v);
  });
  function logical(name, init, combine) {
    def(name, 1, undefined, function (nodes, ctx) {
      let acc = init, any = false;
      for (const node of nodes) {
        if (node.t === 'missing') continue;
        const v = this.ev(node, ctx);
        if (v instanceof Range) {
          let err = null;
          this.eachCell(v, (x) => {
            if (isErr(x)) { err = x; return false; }
            if (typeof x === 'number' || typeof x === 'boolean') { acc = combine(acc, !!+x); any = true; }
            return true;
          });
          if (err) return err;
        } else if (v instanceof Matrix) {
          for (const row of v.data) for (const x of row) {
            if (isErr(x)) return x;
            if (typeof x === 'number' || typeof x === 'boolean') { acc = combine(acc, !!+x); any = true; }
          }
        } else if (isErr(v)) return v;
        else if (typeof v === 'number' || typeof v === 'boolean') { acc = combine(acc, !!+v); any = true; } else if (typeof v === 'string') return E.VALUE;
      }
      return any ? acc : E.VALUE;
    });
  }
  logical('AND', true, (a, b) => a && b);
  logical('OR', false, (a, b) => a || b);
  def('NOT', 1, 1, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['b']); if (isErr(a)) return a;
    return !a[0];
  });
  def('TRUE', 0, 0, () => true);
  def('FALSE', 0, 0, () => false);
  def('NA', 0, 0, () => E.NA);
  def('PI', 0, 0, () => Math.PI);
  def('COUNTBLANK', 1, 1, function (nodes, ctx) {
    const rg = this.ev(nodes[0], ctx);
    if (isErr(rg)) return rg;
    if (!(rg instanceof Range)) return E.VALUE;
    let nonBlank = 0;
    this.eachCell(rg, (v) => { if (v !== '') nonBlank++; });
    return rg.rows * rg.cols - nonBlank;
  });
  def('ISBLANK', 1, 1, function (nodes, ctx) { return this.sv(nodes[0], ctx) === null; });
  def('ISNUMBER', 1, 1, function (nodes, ctx) {
    const v = this.sv(nodes[0], ctx);
    return typeof v === 'number' || typeof v === 'boolean';
  });
  def('ISTEXT', 1, 1, function (nodes, ctx) { return typeof this.sv(nodes[0], ctx) === 'string'; });
  def('ISERROR', 1, 1, function (nodes, ctx) {
    const v = this.sv(nodes[0], ctx);
    return v === UNKNOWN_FN ? E.VALUE : isErr(v);
  });

  // ---- text
  def('CONCATENATE', 0, undefined, function (nodes, ctx) {
    const vals = nodes.map((n) => this.sv(n, ctx));
    for (const v of vals) if (isErr(v) && !v.fromRef) return v;
    let s = '';
    for (const v of vals) {
      const t = this.toStr(v);
      if (isErr(t)) return t;
      s += t;
    }
    return s;
  });
  def('LEFT', 1, 2, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s', 'n'], [undefined, 1]); if (isErr(a)) return a;
    const n = Math.trunc(a[1]);
    if (n < 0) return E.VALUE;
    return a[0].slice(0, n);
  });
  def('RIGHT', 1, 2, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s', 'n'], [undefined, 1]); if (isErr(a)) return a;
    const n = Math.trunc(a[1]);
    if (n < 0) return E.VALUE;
    return n === 0 ? '' : a[0].slice(Math.max(0, a[0].length - n));
  });
  def('MID', 3, 3, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s', 'n', 'n']); if (isErr(a)) return a;
    const st = Math.trunc(a[1]), n = Math.trunc(a[2]);
    if (st < 1 || n < 0) return E.VALUE;
    return a[0].substr(st - 1, n);
  });
  def('LEN', 1, 1, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s']); if (isErr(a)) return a;
    return a[0].length;
  });
  def('UPPER', 1, 1, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s']); if (isErr(a)) return a;
    return a[0].toUpperCase();
  });
  def('LOWER', 1, 1, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s']); if (isErr(a)) return a;
    return a[0].toLowerCase();
  });
  def('TRIM', 1, 1, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s']); if (isErr(a)) return a;
    return a[0].replace(/^ +| +$/g, '').replace(/ {2,}/g, ' ');
  });

  // ---- lookup
  def('VLOOKUP', 3, 4, function (nodes, ctx) {
    const v = this.sv(nodes[0], ctx);
    const tbl = this.ev(nodes[1], ctx);
    const rest = this.args(nodes.slice(2), ctx, ['n', 'b'], [undefined, true]);
    if (isErr(v) && !v.fromRef) return v;
    if (isErr(tbl)) return tbl;
    if (isErr(rest)) return rest;
    if (isErr(v)) return v;
    const col = Math.trunc(rest[0]);
    const sorted = nodes.length > 3 && nodes[3].t === 'missing' ? false : rest[1];
    if (!(tbl instanceof Range) && !(tbl instanceof Matrix)) return E.VALUE;
    if (col < 1) return E.VALUE;
    if (col > tbl.cols) return E.VALUE;
    if (v === null) return E.NA;
    const g = this.grid(tbl);
    const first = g.map((row) => row[0]);
    const i = lookupPos(first, v, sorted ? 1 : 0);
    if (i < 0) return E.NA;
    let out = tbl instanceof Range ? this.cellValue(tbl.sheet, tbl.r1 + i, tbl.c1 + col - 1) : g[i][col - 1];
    if (out === EMPTYR) out = 0;
    if (typeof out === 'boolean') return out ? 1 : 0;
    return out === null ? 0 : out;
  });
  def('MATCH', 2, 3, function (nodes, ctx) {
    const v = this.sv(nodes[0], ctx);
    const rg = this.ev(nodes[1], ctx);
    const rest = this.args(nodes.slice(2), ctx, ['n'], [1]);
    if (isErr(v) && !v.fromRef) return v;
    if (isErr(rg)) return rg;
    if (isErr(rest)) return rest;
    if (isErr(v)) return v;
    if (!(rg instanceof Range) && !(rg instanceof Matrix)) return E.NA;
    if (rg.rows !== 1 && rg.cols !== 1) return E.VALUE;
    if (v === null) return E.NA;
    const g = this.grid(rg);
    const list = rg.cols === 1 ? g.map((row) => row[0]) : g[0];
    const mode = rest[0] > 0 ? 1 : rest[0] < 0 ? -1 : 0;
    const i = lookupPos(list, v, mode);
    return i < 0 ? E.NA : i + 1;
  });
  def('INDEX', 1, 4, function (nodes, ctx) {
    const ref = this.ev(nodes[0], ctx);
    if (isErr(ref)) return ref;
    const a = this.args(nodes.slice(1), ctx, ['n', 'n', 'n'], [0, undefined, 1]);
    if (isErr(a)) return a;
    let row = Math.trunc(a[0]);
    let col = a[1] === undefined ? undefined : Math.trunc(a[1]);
    let tgt = ref;
    if (!(tgt instanceof Range) && !(tgt instanceof Matrix)) tgt = new Matrix([[tgt]]);
    if (col === undefined) {
      if (tgt.rows === 1 && tgt.cols > 1) { col = row; row = 0; } else col = 0;
    }
    if (row < 0 || col < 0) return E.VALUE;
    if (row > tgt.rows || col > tgt.cols) return E.REF;
    if (tgt instanceof Matrix) {
      if (row === 0 && col === 0) return tgt;
      if (row === 0) return new Matrix(tgt.data.map((r) => [r[col - 1]]));
      if (col === 0) return new Matrix([tgt.data[row - 1]]);
      return tgt.data[row - 1][col - 1];
    }
    const r1 = row === 0 ? tgt.r1 : tgt.r1 + row - 1;
    const r2 = row === 0 ? tgt.r2 : r1;
    const c1 = col === 0 ? tgt.c1 : tgt.c1 + col - 1;
    const c2 = col === 0 ? tgt.c2 : c1;
    return new Range(tgt.sheet, r1, c1, r2, c2);
  });

  // ---------------------------------------------------------------- public API
  function outValue(v) {
    if (v === EMPTYR || v === null) return 0;
    if (isErr(v)) return { error: v.code };
    if (typeof v === 'number') {
      if (!isFinite(v)) return { error: '#NUM!' };
      const r = Number(v.toPrecision(15));
      return r === 0 ? 0 : r;
    }
    return v;
  }

  function computeWorkbook(json) {
    const eng = new Engine(json);
    // evaluate in dependency-friendly order to keep recursion shallow
    for (const sh of eng.wb.sheets) {
      const order = topoOrder(eng, sh);
      for (const cell of order) eng.cellValue(sh, cell.r, cell.c);
    }
    const out = { sheets: {} };
    for (const sh of eng.wb.sheets) {
      const cells = {};
      const src = (json.sheets && json.sheets[sh.name]) || {};
      for (const a of Object.keys(src)) {
        const raw = src[a];
        if (typeof raw === 'string' && raw.charAt(0) === '=' && raw.length > 1) {
          const pa = parseAddr(a);
          cells[a] = pa ? outValue(eng.cellValue(sh, pa.r, pa.c)) : { error: '#REF!' };
        } else cells[a] = raw;
      }
      out.sheets[sh.name] = cells;
    }
    return out;
  }

  // Order formula cells so that static precedents come first (iterative DFS).
  function topoOrder(eng, sh) {
    const res = [];
    const seen = new Set();
    const precedents = (cell) => {
      const out = [];
      const walk = (node) => {
        if (!node || typeof node !== 'object') return;
        if (node.t === 'ref') {
          const s = node.sheet === null ? sh : eng.wb.sheet(node.sheet);
          if (!s || s !== sh) return;
          const b = node.b || node.a;
          const r1 = Math.min(node.a.r, b.r), r2 = Math.min(Math.max(node.a.r, b.r), s.maxR);
          const c1 = Math.min(node.a.c, b.c), c2 = Math.min(Math.max(node.a.c, b.c), s.maxC);
          if ((r2 - r1 + 1) * (c2 - c1 + 1) > 4 * s.cells.size) {
            for (const c of s.cells.values()) if (c.formula && c.r >= r1 && c.r <= r2 && c.c >= c1 && c.c <= c2) out.push(c);
          } else {
            for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
              const x = s.cells.get(r * MAXCOL + c);
              if (x && x.formula) out.push(x);
            }
          }
          return;
        }
        for (const k of ['a', 'b']) if (node[k] && node[k].t) walk(node[k]);
        if (node.args) node.args.forEach(walk);
      };
      walk(cell.ast);
      return out;
    };
    for (const start of sh.cells.values()) {
      if (!start.formula || seen.has(start)) continue;
      const stack = [[start, null]];
      seen.add(start);
      while (stack.length) {
        const top = stack[stack.length - 1];
        if (top[1] === null) top[1] = precedents(top[0]);
        const next = top[1].pop();
        if (next === undefined) { res.push(top[0]); stack.pop(); continue; }
        if (!seen.has(next)) { seen.add(next); stack.push([next, null]); }
      }
    }
    return res;
  }

  const api = {
    computeWorkbook, parse, tokenize, fmtNum, textToNum, parseAddr, addrName, colName, colIndex, Engine, FUNCS,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.SheetEngine = api;
})(typeof window !== 'undefined' ? window : globalThis);
