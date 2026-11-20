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
  // a lookup whose lookup value is #N/A gives an internal Calc error that
  // files show as #N/A but that ISNA does not recognise
  E.LOOKNA = new FErr('Err:LOOKNA', false); EREF['Err:LOOKNA'] = new FErr('Err:LOOKNA', true);
  const errByCode = (code) => Object.values(E).find((e) => e.code === code) || new FErr(code);
  const errText = (code) => (code[0] === '#' ? code : '#N/A');
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
    if (!isFinite(d) || d > Math.abs(a) * E48 || d > Math.abs(b) * E48) return false;
    // representable integers are never "approximately" equal (rtl::math::approxEqual)
    const repInt = (x) => Number.isInteger(x) && Math.abs(x) < 9007199254740992;
    return !(repInt(d) && repInt(a) && repInt(b));
  }
  function approxAdd(a, b) {
    if (((a < 0 && b > 0) || (b < 0 && a > 0)) && approxEqual(a, -b)) return 0;
    return a + b;
  }
  function approxSub(a, b) {
    if (((a < 0 && b < 0) || (a > 0 && b > 0)) && approxEqual(a, b)) return 0;
    return a - b;
  }
  // rtl::math::approxValue: round to 15 significant digits, except values
  // above 2^41, integers and values with at most 11 fraction bits
  const f64 = new DataView(new ArrayBuffer(8));
  function fracBits(a) {
    f64.setFloat64(0, a);
    const hi = f64.getUint32(0), lo = f64.getUint32(4);
    const e = ((hi >>> 20) & 0x7ff) - 1023;
    if (e >= 52) return 0;
    const fhi = hi & 0xfffff;
    let tz; // trailing zeros of the 52-bit fraction
    if (lo) tz = 31 - Math.clz32(lo & -lo);
    else if (fhi) tz = 32 + 31 - Math.clz32(fhi & -fhi);
    else tz = 52;
    const least = fhi || lo ? tz + 1 : 53;
    return Math.max(53 - least - e, 0);
  }
  function approxValue(x) {
    if (!isFinite(x) || x === 0 || x > 2199023255552) return x;
    const a = Math.abs(x);
    if ((Number.isInteger(a) && a < 9007199254740992) || fracBits(a) <= 11) return x;
    return Number(x.toPrecision(15));
  }
  const approxFloor = (x) => Math.floor(approxValue(x));

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
    if (Number.isInteger(x) && Math.abs(x) < 9007199254740992) return String(x); // exact below 2^53
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

  // Text -> number, as Calc's number scanner (en-US): plain numbers with
  // thousands separators, exponents, %, $, signs (leading, trailing or as
  // parentheses), times (h:mm[:ss], am/pm), dates (m/d/y, y-m-d, y/m/d) and
  // mixed fractions ("1 1/2").
  const NUMRE = /^(\d{1,3}(?:,\d{3})+|\d*)(\.\d*)?(?:[eE]([+-]?\d+))?$/;
  function plainNum(t) {
    const m = NUMRE.exec(t);
    if (!m) return null;
    const intp = m[1].replace(/,/g, '');
    const frac = m[2] || '';
    if (intp === '' && frac.length <= 1) return null;
    const v = Number((intp || '0') + (frac === '.' ? '' : frac) + (m[3] !== undefined ? 'e' + m[3] : ''));
    return isFinite(v) ? { v, exp: m[3] !== undefined } : null;
  }
  function daySerial(y, m, d) {
    if (m < 1 || m > 12 || d < 1) return null;
    const dim = [31, (y % 4 === 0 && (y % 100 !== 0 || y % 400 === 0)) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1];
    if (d > dim) return null;
    return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
  }
  function fullYear(y, len) {
    if (len > 2) return y;
    return y < 50 ? 2000 + y : 1900 + y;
  }
  function timeVal(t) {
    let m = /^(\d+):(\d*)(?::(\d+(?:\.\d*)?))?(?:\s*([ap])\.?m\.?)?$/i.exec(t);
    if (m) {
      let h = +m[1];
      const mi = m[2] === '' ? 0 : +m[2], sec = m[3] ? +m[3] : 0;
      if (m[2] === '' && m[3] !== undefined) return null;
      if (mi > 59 || sec >= 60) return null;
      if (m[4]) {
        if (h > 12) return null;
        h = h % 12 + (m[4].toLowerCase() === 'p' ? 12 : 0);
      }
      return (h * 3600 + mi * 60 + sec) / 86400;
    }
    m = /^(\d+):(\d+\.\d*)$/.exec(t); // mm:ss.f
    if (m) {
      if (+m[2] >= 60) return null;
      return (+m[1] * 60 + +m[2]) / 86400;
    }
    return null;
  }
  function dateVal(t) {
    let m = /^(\d{1,2})\/(\d{1,2})(?:\/(\d{1,4}))?$/.exec(t);
    if (m) {
      const y = m[3] !== undefined ? fullYear(+m[3], m[3].length) : new Date().getFullYear();
      return daySerial(y, +m[1], +m[2]);
    }
    m = /^(\d{1,4})-(\d{1,2})-(\d{1,2})$/.exec(t);
    if (m) return daySerial(fullYear(+m[1], m[1].length), +m[2], +m[3]);
    return null;
  }
  function dateTimeVal(t) {
    const tv = timeVal(t);
    if (tv !== null) return tv;
    const dv = dateVal(t);
    if (dv !== null) return dv;
    const m = /^(\S+?)(?:\s+|T)(\S.*)$/.exec(t);
    if (m) {
      const d = dateVal(m[1]), tm = timeVal(m[2]);
      if (d !== null && tm !== null) return d + tm;
    }
    return null;
  }
  function textToNum(s) {
    let t = s.replace(/^ +| +$/g, '');
    if (t === '') return null;
    const u = t.toUpperCase();
    if (u === 'TRUE') return 1;
    if (u === 'FALSE') return 0;
    // mixed fraction: [sign] int num/den
    let m = /^([+-]?)(\d+) +(\d+)\/(\d+)$/.exec(t);
    if (m) {
      if (+m[4] === 0) return null;
      const v = +m[2] + +m[3] / +m[4];
      return m[1] === '-' ? -v : v;
    }
    m = /^([+-]?) *(.*)$/.exec(t);
    if (/[:/]|^\d+-\d/.test(m[2])) {
      const dv = dateTimeVal(m[2]);
      if (dv !== null) return m[1] === '-' ? -dv : dv;
      return null;
    }
    let neg = false, signs = 0, cur = 0, pct = 0, lp = 0, rp = 0;
    // peel leading and trailing decorations; % only as the outermost suffix
    for (;;) {
      t = t.replace(/^ +/, '');
      const c = t[0];
      if (c === '-' || c === '+') { signs++; if (c === '-') neg = true; t = t.slice(1); continue; }
      if (c === '$') { cur++; t = t.slice(1); continue; }
      if (c === '(') { lp++; t = t.slice(1); continue; }
      break;
    }
    for (let first = true; ; first = false) {
      t = t.replace(/ +$/, '');
      const c = t[t.length - 1];
      if (c === '%' && first) { pct++; t = t.slice(0, -1); continue; }
      if (c === '-' || c === '+') { signs++; if (c === '-') neg = true; t = t.slice(0, -1); continue; }
      if (c === '$') { cur++; t = t.slice(0, -1); continue; }
      if (c === ')') { rp++; t = t.slice(0, -1); continue; }
      break;
    }
    if (signs > 1 || cur > 1 || lp > 1 || lp !== rp || (cur && pct)) return null;
    if (lp && signs) return null;
    const pn = plainNum(t);
    if (!pn) return null;
    let v = pn.v;
    if (pct && pn.exp) return null;
    if (pct) v /= 100;
    if (neg || lp) v = -v;
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
  // a union of references, (A1:A2,B1): only aggregate functions take it
  class RefList {
    constructor(ranges) { this.ranges = ranges; }
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
    // whitespace between two references is the intersection operator
    const pushTok = (t) => {
      if (t.t === 'ref' && i > 0 && /\s/.test(src[i - 1]) && toks.length && toks[toks.length - 1].t === 'ref') toks.push({ t: 'op', v: 'isect' });
      toks.push(t);
    };
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
      // sheet-qualified reference (sheet2: a 3D range Sheet1:Sheet3!A1)
      let sheet = null, sheet2 = null;
      let j = i;
      {
        const nm = "(?:'((?:[^']|'')+)'|([A-Za-z_\u00C0-\uFFFF][\\w.\u00C0-\uFFFF]*))";
        const m3 = new RegExp('^' + nm + ':' + nm + '!').exec(src.slice(i));
        if (m3) {
          const un = (q, b) => (q !== undefined ? q.replace(/''/g, "'") : b);
          sheet = un(m3[1], m3[2]); sheet2 = un(m3[3], m3[4]);
          const rest3 = src.slice(i + m3[0].length);
          if (!/^\$?[A-Za-z]{1,3}\$?\d/.test(rest3)) { sheet = null; sheet2 = null; } else j = i + m3[0].length;
        }
      }
      if (sheet2 !== null) { /* already parsed */ } else
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
            pushTok({ t: 'ref', sheet, sheet2, a, b, text: src.slice(i, j + m[0].length) });
            i = j + m[0].length;
            continue;
          }
        }
        m = colRangeRe.exec(rest);
        if (m) {
          const c1 = colIndex(m[1].replace('$', '')), c2 = colIndex(m[2].replace('$', ''));
          pushTok({ t: 'ref', sheet, a: { r: 0, c: c1 }, b: { r: MAXROW - 1, c: c2 }, text: src.slice(i, j + m[0].length) });
          i = j + m[0].length;
          continue;
        }
        if (sheet !== null || ch === '$') {
          m = rowRangeRe.exec(rest);
          if (m) {
            const r1 = parseInt(m[1].replace('$', ''), 10) - 1, r2 = parseInt(m[2].replace('$', ''), 10) - 1;
            pushTok({ t: 'ref', sheet, a: { r: r1, c: 0 }, b: { r: r2, c: MAXCOL - 1 }, text: src.slice(i, j + m[0].length) });
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
          pushTok({ t: 'ref', sheet: null, a: { r: r1, c: 0 }, b: { r: r2, c: MAXCOL - 1 }, text: rm[0] });
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
      let a = rangeop1();
      while (isOp('isect')) { p++; a = { t: 'isect', a, b: rangeop1() }; }
      return a;
    }
    function rangeop1() {
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
        if (toks[p].t === 'eof') { if (list.length) throw new SyntaxError('unterminated argument list'); return list; }
        list.push(comparison());
        if (isOp(',')) {
          p++;
          if (isOp(')')) { list.push({ t: 'missing' }); p++; return list; }
          if (toks[p].t === 'eof') throw new SyntaxError('unterminated argument list');
          continue;
        }
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
        case 'ref': return t.sheet2 ? { t: 'ref3d', sheet: t.sheet, sheet2: t.sheet2, a: t.a, b: t.b } : { t: 'ref', sheet: t.sheet, a: t.a, b: t.b };
        case 'name': {
          if (isOp('(')) { p++; return { t: 'fn', name: t.v.toUpperCase().replace(/^_XLFN\./, ''), args: args() }; }
          const u = t.v.toUpperCase();
          if (u === 'TRUE') return { t: 'bool', v: true };
          if (u === 'FALSE') return { t: 'bool', v: false };
          return { t: 'name', v: t.v };
        }
        case 'op':
          if (t.v === '(') {
            const e = comparison();
            if (isOp(',')) { // reference union (A1:A2,B1)
              const items = [e];
              while (isOp(',')) { p++; items.push(comparison()); }
              expect(')');
              return { t: 'union', args: items };
            }
            expect(')');
            return { t: 'paren', a: e };
          }
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
      // A cell that is being computed reads as an error (Calc's Err:522, which
      // round-trips as #VALUE!); the error then propagates through the cycle.
      if (cell.state === 1) return EREF['#VALUE!'];
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
      if (isErr(v)) v = asComputed(v);
      cell.value = v;
      cell.state = 2;
    }


    finalValue(v, ctx) {
      if (v instanceof Range) v = this.intersect(v, ctx);
      else if (v instanceof Matrix) v = v.data[0][0];
      else if (v instanceof RefList) v = E.VALUE;
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
      return EREF['#VALUE!'];
    }

    // scalar value of an operand
    scalar(v, ctx) {
      if (v instanceof Range) return this.intersect(v, ctx);
      if (v instanceof Matrix) return v.data[0][0];
      if (v instanceof RefList) return E.VALUE;
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

    // ---- array evaluation (SUMPRODUCT's arguments): ranges become matrices,
    // operators and scalar functions apply element by element
    aev(node, ctx) {
      const v = this.aev1(node, ctx);
      return v instanceof Range ? this.rangeMatrix(v) : v;
    }
    rangeMatrix(v) {
      const sh = v.sheet;
      const r2 = Math.min(v.r2, Math.max(v.r1, sh.maxR + 1)), c2 = Math.min(v.c2, Math.max(v.c1, sh.maxC + 1));
      const rows = [];
      for (let r = v.r1; r <= r2; r++) {
        const row = [];
        for (let c = v.c1; c <= c2; c++) { const x = this.cellValue(sh, r, c); row.push(x === EMPTYR ? null : x); }
        rows.push(row);
      }
      return new Matrix(rows);
    }
    aev1(node, ctx) {
      const lit = (x) => (x === MISSING ? { t: 'missing' } : { t: 'lit', v: x });
      switch (node.t) {
        case 'paren': return this.aev(node.a, ctx);
        case 'neg': case 'pct': {
          const a = this.aev(node.a, ctx);
          if (!(a instanceof Matrix)) return this.ev(node, ctx);
          return this.amap([a], (x) => this.ev({ t: node.t, a: lit(x) }, ctx));
        }
        case 'bin': {
          const a = this.aev(node.a, ctx), b = this.aev(node.b, ctx);
          if (!(a instanceof Matrix) && !(b instanceof Matrix)) return asComputed(this.binop({ op: node.op, a: lit(a), b: lit(b) }, ctx));
          return this.amap([a, b], (x, y) => asComputed(this.binop({ op: node.op, a: lit(x), b: lit(y) }, ctx)));
        }
        case 'fn': {
          if (!ELEMENTWISE.has(node.name)) return this.ev(node, ctx);
          const args = node.args.map((n) => (n.t === 'missing' ? MISSING : this.aev(n, ctx)));
          if (!args.some((m) => m instanceof Matrix)) return this.ev(node, ctx);
          return this.amap(args, (...xs) => asComputed(this.call({ t: 'fn', name: node.name, args: xs.map(lit) }, ctx)));
        }
        default: return this.ev(node, ctx);
      }
    }
    // a "force array" parameter: references stay references, expressions
    // are evaluated as arrays
    fev(node, ctx) {
      let n = node;
      while (n.t === 'paren') n = n.a;
      if (n.t === 'ref' || n.t === 'rangeop') return this.ev(n, ctx);
      return this.aev(n, ctx);
    }
    // element-wise map; a single row/column is repeated, otherwise the
    // result is as large as the smallest operand
    amap(args, f) {
      let R = Infinity, C = Infinity;
      for (const m of args) if (m instanceof Matrix) { if (m.rows > 1) R = Math.min(R, m.rows); if (m.cols > 1) C = Math.min(C, m.cols); }
      if (R === Infinity) R = 1;
      if (C === Infinity) C = 1;
      const at = (m, r, c) => {
        if (!(m instanceof Matrix)) return m;
        const rr = m.rows === 1 ? 0 : r, cc = m.cols === 1 ? 0 : c;
        return rr < m.rows && cc < m.cols ? m.data[rr][cc] : E.NA;
      };
      const out = [];
      for (let r = 0; r < R; r++) {
        const row = [];
        for (let c = 0; c < C; c++) {
          const xs = args.map((m) => at(m, r, c));
          row.push(f(...xs));
        }
        out.push(row);
      }
      return new Matrix(out);
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
        case 'lit': return node.v;
        case 'paren': return this.ev(node.a, ctx);
        case 'name': return E.NAME;
        case 'ref': {
          const sh = this.resolveSheet(node.sheet, ctx);
          if (!sh) {
            // 'Sheet1:Sheet3'!A1 (one quoted name) is a 3D reference
            const k = node.sheet ? node.sheet.indexOf(':') : -1;
            if (k > 0) return this.ev({ t: 'ref3d', sheet: node.sheet.slice(0, k), sheet2: node.sheet.slice(k + 1), a: node.a, b: node.b }, ctx);
            return E.NAME;
          }
          const b = node.b || node.a;
          return new Range(sh, node.a.r, node.a.c, b.r, b.c);
        }
        case 'ref3d': {
          const s1 = this.wb.sheet(node.sheet), s2 = this.wb.sheet(node.sheet2);
          if (!s1 || !s2) return E.NAME;
          let i1 = this.wb.sheets.indexOf(s1), i2 = this.wb.sheets.indexOf(s2);
          if (i1 > i2) [i1, i2] = [i2, i1];
          const b = node.b || node.a;
          const ranges = [];
          for (let k = i1; k <= i2; k++) ranges.push(new Range(this.wb.sheets[k], node.a.r, node.a.c, b.r, b.c));
          return ranges.length === 1 ? ranges[0] : new RefList(ranges);
        }
        case 'isect': {
          const a = this.ev(node.a, ctx), b = this.ev(node.b, ctx);
          if (isErr(a)) return a;
          if (isErr(b)) return b;
          if (!(a instanceof Range) || !(b instanceof Range)) return E.VALUE;
          if (a.sheet !== b.sheet) return E.NULL;
          const r1 = Math.max(a.r1, b.r1), r2 = Math.min(a.r2, b.r2), c1 = Math.max(a.c1, b.c1), c2 = Math.min(a.c2, b.c2);
          if (r1 > r2 || c1 > c2) return E.NULL;
          return new Range(a.sheet, r1, c1, r2, c2);
        }
        case 'union': {
          const ranges = [];
          for (const n of node.args) {
            const v = this.ev(n, ctx);
            if (isErr(v)) return v;
            if (v instanceof Range) ranges.push(v);
            else if (v instanceof RefList) ranges.push(...v.ranges);
            else return E.VALUE;
          }
          return new RefList(ranges);
        }
        case 'rangeop': {
          const a = this.ev(node.a, ctx), b = this.ev(node.b, ctx);
          if (isErr(a)) return a;
          if (isErr(b)) return b;
          if (!(a instanceof Range) || !(b instanceof Range) || a.sheet !== b.sheet) return E.VALUE;
          return new Range(a.sheet, Math.min(a.r1, b.r1), Math.min(a.c1, b.c1), Math.max(a.r2, b.r2), Math.max(a.c2, b.c2));
        }
        case 'neg': {
          let depth = 1, inner = node.a;
          while (inner.t === 'neg' || inner.t === 'paren') { if (inner.t === 'neg') depth++; inner = inner.a; }
          const v = this.sv(inner, ctx);
          const n = this.toNum(v);
          if (isErr(n)) return asComputed(n);
          const r = depth % 2 ? -n : n;
          // Calc keeps the boolean number format through negation
          // (but not for a plain cell reference)
          if (typeof v === 'boolean' && inner.t !== 'ref' && (r === 0 || r === 1)) return r === 1;
          return r === 0 ? 0 : r;
        }
        case 'pct': {
          const v = this.sv(node.a, ctx);
          const n = this.toNum(v);
          if (isErr(n)) return asComputed(n);
          return n / 100;
        }
        case 'bin': return asComputed(this.binop(node, ctx));
        case 'fn': { const r = this.call(node, ctx); return r === UNKNOWN_FN ? r : asComputed(r); }
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
      if (!f) {
        if (!node.args.length) return E.NAME;
        // the arguments are still computed: a computed error among them wins
        for (const a of node.args) {
          if (a.t === 'missing') continue;
          const v = this.ev(a, ctx);
          if (isErr(v) && !v.fromRef && v !== UNKNOWN_FN) return v;
        }
        return UNKNOWN_FN;
      }
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
          // an absent argument takes its default; an empty one (f(x,)) is 0/""/FALSE in Calc
          if (defaults && defaults[i] !== undefined && v === undefined) { out[i] = defaults[i]; continue; }
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
      const vals = [];
      for (const node of nodes) {
        if (node.t === 'missing') { vals.push(0); continue; } // an empty argument is 0
        const v = this.ev(node, ctx);
        if (isErr(v) && !v.fromRef && !opts.skipErrors) return { err: v };
        if (v instanceof RefList) vals.push(...v.ranges); else vals.push(v);
      }
      // Calc pops parameters from the stack: the last one is examined first
      // (MEDIAN gathers its parameters into an array first: first one first)
      // for SUM's rounding: Calc's pop order, each item a number or the
      // numbers of one column of a range
      const parts = [];
      for (let j = 0; j < vals.length; j++) {
        const k = opts.forward ? j : vals.length - 1 - j;
        const v = vals[k];
        const n0 = nums.length;
        if (v instanceof Matrix) parts.push(nums); // placeholder, fixed below
        if (v instanceof Range) {
          let err = null, col = -1, cur = null;
          this.eachCell(v, (x, r, c) => {
            if (c !== col) { col = c; cur = []; parts.push(cur); }
            if (typeof x === 'number') { nums.push(x); cur.push(x); } else if (typeof x === 'boolean') { nums.push(x ? 1 : 0); cur.push(x ? 1 : 0); }
            else if (isErr(x) && !opts.skipErrors) { err = x; return false; }
            return true;
          });
          if (err) return { err };
        } else if (v instanceof Matrix) {
          for (const row of v.data) for (const x of row) {
            if (typeof x === 'number') { nums.push(x); } else if (typeof x === 'boolean') nums.push(x ? 1 : 0);
            else if (isErr(x) && !opts.skipErrors) return { err: x };
          }
        } else if (typeof v === 'number') { nums.push(v); } else if (typeof v === 'boolean') { nums.push(v ? 1 : 0); } else if (v === EMPTYR) { nums.push(0); } else if (isErr(v)) { if (!opts.skipErrors) return { err: v }; } else if (typeof v === 'string') {
          const n = textToNum(v);
          if (opts.countText) { if (n !== null) { nums.push(n); } } else return { err: E.VALUE };
        }
        if (parts[parts.length - 1] === nums) parts[parts.length - 1] = nums.slice(n0);
        else if (nums.length > n0 && !(v instanceof Range)) parts.push(nums[n0]);
      }
      if (!opts.forward) nums.reverse();
      // MIN/MAX take the number format of the first value of the first
      // parameter: a boolean cell there makes the result a boolean.
      let boolFmt = false;
      if (vals.length && vals[0] instanceof Range) {
        this.eachCell(vals[0], (x) => {
          if (typeof x === 'number' || typeof x === 'boolean') { boolFmt = typeof x === 'boolean'; return false; }
          return true;
        });
      }
      return { nums, parts, boolFmt: boolFmt && nums.length > 0 };
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
      if (ch === '~' && i + 1 < pat.length) { s += pat[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); continue; } // ~ escapes any character
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
  // exact: an inline/computed array is searched with exact number equality
  function lookupEq(v, x, exact) {
    if (x === null) return false;
    if (typeof v === 'boolean') v = v ? 1 : 0;
    if (typeof v === 'number') {
      if (typeof x === 'boolean') x = x ? 1 : 0;
      return typeof x === 'number' && (exact ? v === x : approxEqual(v, x));
    }
    if (typeof v === 'string') {
      // text is compared with the cell's displayed text (numbers too)
      if (typeof x === 'number') x = fmtNum(x);
      else if (typeof x === 'boolean') x = x ? 'TRUE' : 'FALSE';
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
  // Calc's lookup in "sorted" mode over cells (ScQueryCellIterator::
  // FindEqualOrSortedLastInRange): numbers sort before text; a numeric query
  // skips text cells, a text query may skip one leading text cell; a binary
  // search (columns only) positions the start, then cells are walked while
  // they satisfy the condition, and an exact hit extends over equal entries.
  function sortedLookup(list0, q, asc, binary) {
    // an error cell is a non-text cell that never satisfies the condition
    const ERRC = { err: true };
    const list = list0.map((x) => x === null ? null : isErr(x) ? ERRC : x === EMPTYR ? '' : typeof x === 'boolean' ? +x : x);
    const n = list.length, byStr = typeof q === 'string';
    const isStr = (x) => typeof x === 'string';
    const cmp = (c) => (byStr ? strCmp(c, q) : c === ERRC ? (approxEqual(0, q) ? 0 : 0 < q ? -1 : 1) : (approxEqual(c, q) ? 0 : c < q ? -1 : 1));
    const firstStringIgnore = byStr, allStringIgnore = !byStr;
    let row = 0;
    if (firstStringIgnore && isStr(list[0])) {
      const t = cmp(list[0]);
      if ((asc && t > 0) || (!asc && t < 0)) row++;
    }
    while (row < n && list[row] === null) row++;
    if (allStringIgnore) while (row < n && (list[row] === null || isStr(list[row]))) row++;
    let start = row;
    if (binary && row < n) {
      const idx = [];
      for (let i = row; i < n; i++) if (list[i] !== null) idx.push(i);
      let lo = 0, hi = idx.length - 1, last = 0, found = null, done = false, broken = false;
      let lastV = asc ? -Infinity : Infinity, lastS = asc ? '' : '￿';
      if (isStr(list[idx[0]])) lastS = list[idx[0]]; else lastV = list[idx[0]] === ERRC ? 0 : list[idx[0]];
      while (lo <= hi && !done) {
        const mid = (lo + hi) >> 1, c = list[idx[mid]];
        let r = 0;
        if (!isStr(c) && !byStr) {
          r = cmp(c);
          const cv = c === ERRC ? 0 : c;
          if (r < 0 && asc) { if (lastV <= cv) { lastV = cv; last = mid; } else broken = true; }
          else if (r > 0 && !asc) { if (lastV >= cv) { lastV = cv; last = mid; } else broken = true; }
        } else if (isStr(c) && byStr) {
          r = cmp(c);
          if (r < 0 && asc) { if (strCmp(lastS, c) <= 0) { lastS = c; last = mid; } else broken = true; }
          else if (r > 0 && !asc) { if (strCmp(lastS, c) >= 0) { lastS = c; last = mid; } else broken = true; }
        } else if (!isStr(c)) { r = -1; if (asc) last = mid; } else { r = 1; if (!asc) last = mid; }
        if (broken) break;
        if (r < 0) { if (asc) lo = mid + 1; else if (mid > 0) hi = mid - 1; else done = true; }
        else if (r > 0) { if (asc) { if (mid > 0) hi = mid - 1; else done = true; } else lo = mid + 1; }
        else { found = mid; last = mid; if (mid > 0) hi = mid - 1; else done = true; }
      }
      start = broken ? idx[0] : found !== null ? idx[found] : idx[last];
    }
    if (!asc) {
      // descending data cannot start below the value
      for (let k = row; k < n; k++) {
        const c = list[k];
        if (c === null || c === ERRC || isStr(c) !== byStr) continue;
        if (cmp(c) < 0) return -1;
        break;
      }
    }
    let foundRow = -1, fsi = firstStringIgnore, eq = false;
    for (let r = start; r < n; r++) {
      const c = list[r];
      if (c === null || (allStringIgnore && isStr(c))) continue;
      if (isStr(c) === byStr && c !== ERRC) {
        const t = cmp(c);
        if (asc ? t <= 0 : t >= 0) { foundRow = r; fsi = false; if (t === 0) { eq = true; break; } continue; }
      }
      if (fsi && isStr(c)) { fsi = false; continue; }
      break;
    }
    if (eq) {
      for (let k = foundRow + 1; k < n; k++) {
        const c = list[k];
        if (c === null || (allStringIgnore && isStr(c))) continue;
        if (isStr(c) === byStr && c !== ERRC && cmp(c) === 0) foundRow = k; else break;
      }
    }
    return foundRow;
  }
  // position (0-based) in a 1-D list, or -1.
  // mode 0: exact (wildcards); 1: largest <= v; -1: smallest >= v.
  // ranges: 'col' (a cell column: Calc's binary search), 'row' (a cell row:
  // Calc's sequential search); anything else (inline arrays) uses a simple
  // approximation. On sorted data all agree with Calc.
  function lookupPos(list, v, mode, ranges) {
    if (typeof v === 'boolean') v = v ? 1 : 0;
    if (mode === 0 || (typeof v === 'string' && hasWild(v))) {
      for (let i = 0; i < list.length; i++) if (lookupEq(v, list[i], !ranges)) return i;
      return -1;
    }
    if (ranges === 'col' || ranges === 'row') return sortedLookup(list, v, mode > 0, ranges === 'col');
    if (mode > 0) {
      let best = -1;
      for (let i = 0; i < list.length; i++) {
        const c = lookupCmp(list[i], v);
        if (c === null) continue;
        if (c === 0) return i;
        if (c < 0) best = i;
        else if (best >= 0 || typeof v === 'number') break;
      }
      return best;
    }
    const idx = [];
    for (let i = 0; i < list.length; i++) if (lookupCmp(list[i], v) !== null) idx.push(i);
    // data that starts below the value cannot be in descending order
    if (!idx.length || lookupCmp(list[idx[0]], v) < 0) return -1;
    let lo = 0, hi = idx.length - 1, best = -1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const c = lookupCmp(list[idx[mid]], v);
      if (c === 0) return idx[mid];
      if (c > 0) { best = idx[mid]; lo = mid + 1; } else hi = mid - 1;
    }
    return best;
  }


  // ---------------------------------------------------------------- functions
  const FUNCS = {};
  function def(name, min, max, fn) { FUNCS[name] = { min, max, fn }; }

  function aggregate(name, fn, opts, min) {
    def(name, min === undefined ? 0 : min, undefined, function (nodes, ctx) {
      const r = this.collect(nodes, ctx, opts || {});
      if (r.err) return r.err;
      return fn(r.nums, r.boolFmt, r.parts);
    });
  }

  // LibreOffice's KahanSum (sc/inc/kahan.hxx): Neumaier summation that holds
  // back the latest term and adds it last, snapping to 0 when it cancels.
  const neumaier = (st, v) => {
    const t = st.sum + v;
    if (Math.abs(st.sum) >= Math.abs(v)) st.err += (st.sum - t) + v; else st.err += (v - t) + st.sum;
    st.sum = t;
  };
  const kahanAdd = (st, x) => {
    if (x === 0) return;
    if (st.mem === 0) { st.mem = x; return; }
    neumaier(st, st.mem);
    st.mem = x;
  };
  const kahanGet = (st) => {
    const total = st.sum + st.err;
    if (st.mem === 0) return total;
    if (((st.mem < 0 && total > 0) || (total < 0 && st.mem > 0)) && approxEqual(st.mem, -total)) return 0;
    const t = { sum: st.sum, err: st.err };
    neumaier(t, st.mem);
    return t.sum + t.err;
  };
  const kahan = (nums) => {
    const st = { sum: 0, err: 0, mem: 0 };
    for (const x of nums) kahanAdd(st, x);
    return kahanGet(st);
  };
  // A range's numbers are summed like Calc's sc::op::sumArray (SSE2): blocks of
  // 8 go to 4 accumulators of 2 lanes with Neumaier compensation, combined
  // pairwise; the remaining values are added one by one.
  const sumArray = (xs) => {
    const n = xs.length;
    let i = 0;
    const st = { sum: 0, err: 0, mem: 0 };
    if (n > 7) {
      const S = [];
      for (let j = 0; j < 8; j++) S.push({ sum: 0, err: 0 });
      for (; i + 7 < n; i += 8) for (let j = 0; j < 8; j++) neumaier(S[j], xs[i + j]);
      const comb = (a, b) => {
        for (let l = 0; l < 2; l++) { neumaier(S[a + l], S[b + l].sum); neumaier(S[a + l], S[b + l].err); }
      };
      comb(0, 2); comb(4, 6); comb(0, 4);
      st.sum = S[0].sum; st.err = S[0].err;
      neumaier(st, S[1].sum); neumaier(st, S[1].err);
    }
    for (; i < n; i++) kahanAdd(st, xs[i]);
    return kahanGet(st);
  };
  const finite = (x) => (isFinite(x) ? x : E.NUM);

  // SUM/AVERAGE: parameters are popped last to first; a range is summed on
  // its own (first cell first) and its total added as one term
  const sumParts = (parts) => kahan(parts.map((p) => (Array.isArray(p) ? sumArray(p) : p)));
  aggregate('SUM', (n, b, parts) => finite(sumParts(parts)));
  aggregate('AVERAGE', (n, b, parts) => (n.length ? finite(sumParts(parts) / n.length) : E.DIV0));
  const asFmt = (x, b) => (b && (x === 0 || x === 1) ? x === 1 : x);
  aggregate('MIN', (n, b) => (n.length ? asFmt(Math.min(...n), b) : 0), {}, 1);
  aggregate('MAX', (n, b) => (n.length ? asFmt(Math.max(...n), b) : 0), {}, 1);
  aggregate('PRODUCT', (n) => (n.length ? finite(n.reduce((a, b) => a * b, 1)) : 0));
  aggregate('COUNT', (n) => n.length, { skipErrors: true, countText: true });
  aggregate('MEDIAN', (n) => {
    if (!n.length) return E.VALUE;
    const s = n.slice().sort((a, b) => a - b);
    const m = s.length >> 1;
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
  }, { forward: true }, 1);
  aggregate('STDEV', (n) => {
    if (n.length < 2) return E.DIV0;
    const mean = kahan(n) / n.length;
    const ss = kahan(n.map((x) => (x - mean) * (x - mean)));
    return Math.sqrt(ss / (n.length - 1));
  });

  const variance = (n, pop) => {
    const mean = kahan(n) / n.length;
    return kahan(n.map((x) => (x - mean) * (x - mean))) / (n.length - (pop ? 0 : 1));
  };
  aggregate('VAR', (n) => (n.length < 2 ? E.DIV0 : variance(n, false)));
  aggregate('VARP', (n) => (n.length < 1 ? E.DIV0 : variance(n, true)));
  aggregate('STDEVP', (n) => (n.length < 1 ? E.DIV0 : Math.sqrt(variance(n, true))));
  const kth = (largest) => function (nodes, ctx) {
    if (nodes.length !== 2) return E.VALUE;
    const r = this.collect([nodes[0]], ctx, {});
    if (r.err) return r.err;
    const a = this.args([nodes[1]], ctx, ['v']); if (isErr(a)) return a;
    if (typeof a[0] === 'string') return E.VALUE;
    const k = Math.ceil(approxValue(+a[0]));
    if (k < 1 || k > r.nums.length) return E.VALUE;
    const s = r.nums.slice().sort((x, y) => (largest ? y - x : x - y));
    return s[k - 1];
  };
  def('LARGE', 2, 2, kth(true));
  def('SMALL', 2, 2, kth(false));

  def('COUNTA', 0, undefined, function (nodes, ctx) {
    let n = 0;
    for (const node of nodes) {
      if (node.t === 'missing') { n++; continue; } // an empty argument is 0
      const v = this.ev(node, ctx);
      if (v instanceof RefList) v.ranges.forEach((rg) => this.eachCell(rg, () => { n++; }));
      else if (v instanceof Range) this.eachCell(v, () => { n++; });
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
    let x = a[0];
    // with a digits argument, Calc first rounds a positive non-integer to 12
    // significant digits (tdf#124286)
    if (nodes.length > 1 && x > 0 && Math.trunc(a[1]) < 12 && x % 1 !== 0) {
      const t = Math.floor(approxValue(Math.log10(x))) + 1 - 12;
      x = Math.round(Math.pow(10, -t) * x) * Math.pow(10, t);
    }
    return roundTo(x, a[1], 1);
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
    const vals = nodes.map((n) => this.sv(n, ctx));
    for (const v of vals) if (isErr(v) && !v.fromRef) return v;
    const y = this.toNum(vals[1]); if (isErr(y)) return y;
    if (y === 0) return E.DIV0;
    const x = this.toNum(vals[0]); if (isErr(x)) return x;
    const q = x / y;
    if (!isFinite(q)) return E.NUM;
    const r = approxSub(x, approxFloor(q) * y);
    // a result that rounding pushed outside [0, y) is an invalid argument
    if ((y > 0 && r >= 0 && r < y) || (y < 0 && r <= 0 && r > y)) return r;
    return E.VALUE;
  });
  def('ABS', 1, 1, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n']); if (isErr(a)) return a;
    return Math.abs(a[0]);
  });
  def('POWER', 2, 2, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n', 'n']); if (isErr(a)) return a;
    return power(a[0], a[1]);
  });
  const math1 = (name, fn) => def(name, 1, 1, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n']); if (isErr(a)) return a;
    const r = fn(a[0]);
    return typeof r === 'number' && !isFinite(r) ? E.NUM : r;
  });
  math1('EXP', (x) => Math.exp(x));
  math1('LN', (x) => (x <= 0 ? E.VALUE : Math.log(x)));
  math1('LOG10', (x) => (x <= 0 ? E.VALUE : Math.log10(x)));
  math1('FACT', (x) => {
    x = Math.floor(x);
    if (x < 0) return E.VALUE;
    let f = 1;
    for (let i = 2; i <= x && isFinite(f); i++) f *= i;
    return isFinite(f) ? f : E.VALUE;
  });
  const parity = (odd) => function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['v']); if (isErr(a)) return a;
    if (a[0] === null || a[0] === EMPTYR) return E.VALUE;
    const x = this.toNum(a[0]); if (isErr(x)) return x;
    return Math.abs(Math.trunc(x)) % 2 === (odd ? 1 : 0);
  };
  def('ISODD', 1, 1, parity(true));
  def('ISEVEN', 1, 1, parity(false));
  math1('EVEN', (x) => { const c = Math.ceil(Math.abs(approxValue(x)) / 2) * 2; return x < 0 ? -c : c; });
  math1('ODD', (x) => { let c = Math.ceil(Math.abs(approxValue(x))); if (c % 2 === 0) c++; return x < 0 ? -c : c; });
  def('LOG', 1, 2, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n', 'n'], [undefined, 10]); if (isErr(a)) return a;
    if (a[0] <= 0 || a[1] <= 0 || a[1] === 1) return E.VALUE;
    return Math.log(a[0]) / Math.log(a[1]);
  });
  def('QUOTIENT', 2, 2, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n', 'n']); if (isErr(a)) return a;
    if (a[1] === 0) return E.VALUE;
    return Math.trunc(approxValue(a[0] / a[1]));
  });
  const toMultiple = (up) => function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n', 'n']); if (isErr(a)) return a;
    const [x, m] = a;
    if (m === 0) return up ? 0 : E.VALUE;
    if (x === 0) return 0;
    if (x > 0 && m < 0) return E.VALUE;
    const q = approxValue(x / m);
    return approxValue((up ? Math.ceil(q) : Math.floor(q)) * m);
  };
  def('CEILING', 2, 2, toMultiple(true));
  def('FLOOR', 2, 2, toMultiple(false));
  def('RAND', 0, 0, () => Math.random());
  def('RANDBETWEEN', 2, 2, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n', 'n']); if (isErr(a)) return a;
    const lo = Math.ceil(a[0]), hi = Math.floor(a[1]);
    if (lo > hi) return E.VALUE;
    return lo + Math.floor(Math.random() * (hi - lo + 1));
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
    if (!(rg instanceof Range) && !(rg instanceof Matrix)) return E.VALUE;
    let crit = self.sv(nodes[1], ctx);
    if (crit === MISSING) crit = null;
    const match = makeCriterion(crit);
    if (isErr(match)) return match;
    let sum = rg;
    if (nodes.length > 2) {
      if (nodes[2].t === 'missing') return E.VALUE;
      const s = self.ev(nodes[2], ctx);
      if (isErr(s)) return s;
      if (rg instanceof Matrix) {
        // an inline array: elements beyond the sum array are skipped
        if (!(s instanceof Matrix)) return E.VALUE;
        return { rg, match, sum: s };
      }
      if (!(s instanceof Range)) return E.VALUE;
      sum = new Range(s.sheet, s.r1, s.c1, s.r1 + rg.rows - 1, s.c1 + rg.cols - 1);
    }
    return { rg, match, sum };
  }
  function condCollect(self, a) {
    const nums = [];
    let err = null;
    if (a.rg instanceof Matrix) {
      a.rg.data.forEach((row, r) => row.forEach((v, c) => {
        if (err || !a.match(v)) return;
        const sv = a.sum.data[r] ? a.sum.data[r][c] : undefined;
        if (typeof sv === 'number') nums.push(sv);
        else if (typeof sv === 'boolean') nums.push(sv ? 1 : 0);
        else if (isErr(sv)) err = sv;
      }));
      return err ? { err } : { nums };
    }
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
    const rg = a.rg;
    if (rg instanceof Matrix) {
      for (const row of rg.data) for (const v of row) if (a.match(v)) n++;
      return n;
    }
    if (a.match(null)) {
      // blanks match: count over the whole area, including cells beyond the used range
      const total = rg.rows * rg.cols;
      let nonBlank = 0, nonBlankMatch = 0;
      this.eachCell(rg, (v) => { nonBlank++; if (a.match(v)) nonBlankMatch++; });
      return total - nonBlank + nonBlankMatch;
    }
    this.eachCell(rg, (v) => { if (a.match(v)) n++; });
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
    if (nodes[0].t === 'missing') return E.NA;
    const v = this.ev(nodes[0], ctx);
    const s = this.scalar(v, ctx);
    if (s === UNKNOWN_FN) return E.VALUE;
    if (isErr(s)) { const alt = this.ev(nodes[1], ctx); return alt === MISSING ? 0 : alt; }
    return v === MISSING ? 0 : v;
  });
  function logical(name, init, combine) {
    def(name, 1, undefined, function (nodes, ctx) {
      let acc = init, any = false;
      const vals = [];
      for (const node of nodes) {
        if (node.t === 'missing') { vals.push(false); continue; } // an empty argument is FALSE
        const v = this.ev(node, ctx);
        if (isErr(v) && !v.fromRef) return v;
        vals.push(v);
      }
      // parameters are examined last to first, as Calc pops them
      for (let k = vals.length - 1; k >= 0; k--) {
        const v = vals[k];
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
        else if (typeof v === 'number' || typeof v === 'boolean') { acc = combine(acc, !!+v); any = true; } else if (v === EMPTYR) { acc = combine(acc, false); any = true; } else if (typeof v === 'string') return E.VALUE;
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
  // Calc's LEN/LEFT/RIGHT/MID count Unicode code points, not UTF-16 units
  const codePoints = (s) => Array.from(s);
  const cpCount = (s) => codePoints(s).length;
  def('LEFT', 1, 2, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s', 'n'], [undefined, 1]); if (isErr(a)) return a;
    const n = Math.trunc(a[1]);
    if (a[1] < 0) return E.VALUE; // checked before truncation
    return codePoints(a[0]).slice(0, n).join('');
  });
  def('RIGHT', 1, 2, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s', 'n'], [undefined, 1]); if (isErr(a)) return a;
    const n = Math.trunc(a[1]);
    if (a[1] < 0) return E.VALUE;
    const cps = codePoints(a[0]);
    return n === 0 ? '' : cps.slice(Math.max(0, cps.length - n)).join('');
  });
  def('MID', 3, 3, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s', 'n', 'n']); if (isErr(a)) return a;
    const st = Math.trunc(a[1]), n = Math.trunc(a[2]);
    if (st < 1 || a[2] < 0) return E.VALUE;
    return codePoints(a[0]).slice(st - 1, st - 1 + n).join('');
  });
  def('LEN', 1, 1, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s']); if (isErr(a)) return a;
    return cpCount(a[0]);
  });
  def('UPPER', 1, 1, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s']); if (isErr(a)) return a;
    return a[0].replace(/ß/g, 'ẞ').toUpperCase(); // ICU keeps sharp s one letter
  });
  def('LOWER', 1, 1, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s']); if (isErr(a)) return a;
    return a[0].replace(/[^\u0130]+/g, (m) => m.toLowerCase()); // ICU keeps dotted capital I
  });
  def('TRIM', 1, 1, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s']); if (isErr(a)) return a;
    return a[0].replace(/^ +| +$/g, '').replace(/ {2,}/g, ' ');
  });

  // ---- a few common extras beyond the spec's 40
  def('PROPER', 1, 1, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s']); if (isErr(a)) return a;
    return a[0].toLowerCase().replace(/(^|[^A-Za-z\u00C0-\uFFFF])([a-z\u00C0-\uFFFF])/g, (m, p, c) => p + c.toUpperCase());
  });
  def('EXACT', 2, 2, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s', 's']); if (isErr(a)) return a;
    return a[0] === a[1];
  });
  def('VALUE', 1, 1, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['v']); if (isErr(a)) return a;
    const v = a[0];
    if (typeof v === 'number' || typeof v === 'boolean') return v;
    if (v === null || v === EMPTYR) return E.VALUE;
    const n = textToNum(String(v));
    return n === null ? E.VALUE : n;
  });
  def('FIND', 2, 3, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s', 's', 'n'], [undefined, undefined, 1]); if (isErr(a)) return a;
    const st = Math.trunc(a[2]);
    if (st < 1 || st > a[1].length + 1 || a[0] === '') return E.VALUE;
    const i = a[1].indexOf(a[0], st - 1);
    // the start is in UTF-16 units, the result in code points (as Calc does)
    return i < 0 ? E.VALUE : cpCount(a[1].slice(0, i)) + 1;
  });
  def('SEARCH', 2, 3, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s', 's', 'n'], [undefined, undefined, 1]); if (isErr(a)) return a;
    const st = Math.trunc(a[2]);
    if (st < 1 || st > a[1].length + 1) return E.VALUE;
    const re = new RegExp(wildcardRe(a[0]).source.slice(1, -1), 'iu');
    const m = re.exec(a[1].slice(st - 1));
    return m ? cpCount(a[1].slice(0, m.index + st - 1)) + 1 : E.VALUE;
  });

  def('ISNA', 1, 1, function (nodes, ctx) { const v = this.sv(nodes[0], ctx); return isErr(v) && v.code === '#N/A'; });
  def('ISERR', 1, 1, function (nodes, ctx) { const v = this.sv(nodes[0], ctx); return isErr(v) && v.code !== '#N/A'; });
  def('ISLOGICAL', 1, 1, function (nodes, ctx) { return typeof this.sv(nodes[0], ctx) === 'boolean'; });
  def('ISNONTEXT', 1, 1, function (nodes, ctx) { return typeof this.sv(nodes[0], ctx) !== 'string'; });
  def('N', 1, 1, function (nodes, ctx) {
    const v = this.sv(nodes[0], ctx);
    if (isErr(v)) return v;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    return 0;
  });
  def('T', 1, 1, function (nodes, ctx) {
    const v = this.sv(nodes[0], ctx);
    if (isErr(v)) return v;
    return typeof v === 'string' ? v : '';
  });
  def('SIGN', 1, 1, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n']); if (isErr(a)) return a;
    return a[0] > 0 ? 1 : a[0] < 0 ? -1 : 0;
  });
  def('TRUNC', 1, 2, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n', 'n'], [undefined, 0]); if (isErr(a)) return a;
    return roundTo(a[0], a[1], -1);
  });
  // TEXT's General format: large numbers that need more than 15 digits are
  // shown as d.dddddddddddddd E+nn
  function fmtGeneral(x) {
    if (typeof x !== 'number' || !isFinite(x)) return fmtNum(x);
    const a = Math.abs(x);
    if (a >= 1e15 && !(a < 1e16 && Number(a.toPrecision(15)) === a)) {
      const [m, e] = x.toExponential(14).split('e');
      const n = parseInt(e, 10);
      return m + 'E' + (n < 0 ? '-' : '+') + String(Math.abs(n)).padStart(2, '0');
    }
    return fmtNum(x);
  }
  // TEXT(value, format): number formats without dates (sections, 0 # ?,
  // thousands separators and scaling, %, exponents, quoted literals, @).
  function splitSections(fmt) {
    const out = []; let cur = '';
    for (let i = 0; i < fmt.length; i++) {
      const ch = fmt[i];
      if (ch === '"') { const j = fmt.indexOf('"', i + 1); const e = j < 0 ? fmt.length : j; cur += fmt.slice(i, e + 1); i = e; continue; }
      if (ch === '\\') { cur += fmt.slice(i, i + 2); i++; continue; }
      if (ch === ';') { out.push(cur); cur = ''; continue; }
      cur += ch;
    }
    out.push(cur);
    return out;
  }
  function formatSection(x, sec) {
    // tokens: literals and the number pattern
    const toks = [];
    let i = 0;
    while (i < sec.length) {
      const ch = sec[i];
      if (ch === '"') { const j = sec.indexOf('"', i + 1); const e = j < 0 ? sec.length : j; toks.push({ lit: sec.slice(i + 1, e) }); i = e + 1; continue; }
      if (ch === '\\') { toks.push({ lit: sec[i + 1] || '' }); i += 2; continue; }
      if (ch === '[') { const j = sec.indexOf(']', i); i = j < 0 ? sec.length : j + 1; continue; }
      if (ch === '_') { toks.push({ lit: ' ' }); i += 2; continue; }
      if (ch === '*') { i += 2; continue; }
      if ('0#?.,'.indexOf(ch) >= 0) { toks.push({ p: ch }); i++; continue; }
      if ((ch === 'E' || ch === 'e') && (sec[i + 1] === '+' || sec[i + 1] === '-')) { toks.push({ e: sec[i + 1] }); i += 2; continue; }
      if (ch === '%') { toks.push({ pct: true }); i++; continue; }
      toks.push({ lit: ch }); i++;
    }
    const pat = toks.filter((t) => t.p);
    if (toks.some((t) => t.pct)) x *= 100 ** toks.filter((t) => t.pct).length;
    if (!pat.length) return toks.map((t) => t.lit !== undefined ? t.lit : t.pct ? '%' : '').join('');
    // integer/decimal/exponent placeholder counts
    const ei = toks.findIndex((t) => t.e);
    let intP = '', decP = '', expP = '', seenDot = false, comma = false, trailingCommas = 0;
    toks.forEach((t, k) => {
      if (!t.p) return;
      if (ei >= 0 && k > ei) { if (t.p !== ',' && t.p !== '.') expP += t.p; return; }
      if (t.p === '.') { seenDot = true; return; }
      if (t.p === ',') {
        // a comma after the last digit placeholder scales by 1000
        const rest = toks.slice(k + 1, ei >= 0 ? ei : toks.length).some((u) => u.p && u.p !== ',' && u.p !== '.');
        if (!rest) trailingCommas++; else if (!seenDot) comma = true;
        return;
      }
      if (seenDot) decP += t.p; else intP += t.p;
    });
    x /= 1000 ** trailingCommas;
    let mant = Math.abs(x), exp = 0;
    if (ei >= 0 && mant !== 0) {
      const ip = Math.max(1, intP.length);
      exp = Math.floor(Math.log10(mant)) - (ip - 1);
      mant = mant / 10 ** exp;
      if (Number(mant.toFixed(decP.length)) >= 10 ** ip) { mant /= 10; exp++; }
    }
    const m10 = mant * 10 ** decP.length;
    const scaled = Math.round(Number.isInteger(m10) && m10 < 9007199254740992 ? m10 : approxValue(m10));
    let str = decP.length ? (scaled / 10 ** decP.length).toFixed(decP.length) : String(scaled);
    if (!isFinite(scaled) || scaled > 1e21) str = mant.toFixed(decP.length);
    let [ip, dp = ''] = str.split('.');
    // decimals: trailing optional digits
    let dd = dp.split('');
    for (let k = decP.length - 1; k >= 0 && dd[k] === '0' && decP[k] !== '0'; k--) dd[k] = decP[k] === '?' ? ' ' : '';
    dp = dd.join('');
    // integer part
    const minInt = (intP.match(/0/g) || []).length;
    if (ip === '0' && minInt === 0) ip = '';
    while (ip.length < minInt) ip = '0' + ip;
    if (comma) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    let body = ip + (seenDot && (dp !== '' || /0|\?/.test(decP) || decP === '') && !(dp === '' && decP !== '') ? '.' + dp : '');
    if (ei >= 0) {
      const es = String(Math.abs(exp)).padStart((expP.match(/0/g) || []).length, '0');
      body += 'E' + (exp < 0 ? '-' : toks[ei].e === '+' ? '+' : '') + es;
    }
    // place literals around the number (before first / after last placeholder)
    const first = toks.findIndex((t) => t.p), last = toks.length - 1 - [...toks].reverse().findIndex((t) => t.p || t.e);
    const pre = toks.slice(0, first).map((t) => t.lit !== undefined ? t.lit : t.pct ? '%' : '').join('');
    const mid = toks.slice(first, last + 1).filter((t) => t.lit !== undefined || t.pct).map((t) => t.lit !== undefined ? t.lit : '%').join('');
    const post = toks.slice(last + 1).map((t) => t.lit !== undefined ? t.lit : t.pct ? '%' : '').join('');
    return pre + body + mid + post;
  }
  function textFormat(v, fmt) {
    const secs = splitSections(fmt);
    if (typeof v === 'string') {
      const ts = secs.length >= 4 ? secs[3] : secs.find((x) => x.indexOf('@') >= 0);
      return ts === undefined ? v : ts.replace(/"([^"]*)"/g, '$1').replace(/@/g, v);
    }
    if (fmt === '') return '';
    if (/^general$/i.test(fmt.trim())) return fmtGeneral(v);
    let sec = secs[0], x = v;
    if (v < 0 && secs.length >= 2 && secs[1] !== '') { sec = secs[1]; x = -v; }
    else if (v === 0 && secs.length >= 3) sec = secs[2];
    if (/general/i.test(sec)) return sec.replace(/general/i, fmtGeneral(x)).replace(/"/g, '');
    const out = formatSection(x, sec);
    return (x < 0 && /[0#?]/.test(sec) && out.replace(/[^1-9]/g, '') !== '' ? '-' : '') + out;
  }
  def('TEXT', 2, 2, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['v', 's']); if (isErr(a)) return a;
    let v = a[0];
    if (v === null) v = 0;
    if (v === EMPTYR) v = '';
    if (typeof v === 'boolean') v = v ? 1 : 0;
    if (typeof v === 'string') { const n = textToNum(v); if (n !== null && v !== '') v = n; }
    return textFormat(v, a[1]);
  });
  def('REPT', 2, 2, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s', 'n']); if (isErr(a)) return a;
    const n = Math.trunc(a[1]);
    if (a[1] < 0 || a[0].length * n > 32767) return E.VALUE;
    return a[0].repeat(n);
  });
  def('SUBSTITUTE', 3, 4, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s', 's', 's', 'n']); if (isErr(a)) return a;
    const [t, from, to, nth] = a;
    if (from === '') return t;
    if (nth === undefined) return t.split(from).join(to);
    const k = Math.trunc(nth);
    if (k < 1) return E.VALUE;
    let i = -1;
    for (let j = 0; j < k; j++) { i = t.indexOf(from, i + 1); if (i < 0) return t; }
    return t.slice(0, i) + to + t.slice(i + from.length);
  });
  def('CHAR', 1, 1, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n']); if (isErr(a)) return a;
    const n = Math.trunc(a[0]);
    if (n < 1 || n > 255) return E.VALUE;
    // Calc decodes the byte in the system encoding (UTF-8): a lone byte >= 128 is invalid
    return n < 128 ? String.fromCharCode(n) : '\uFFFD';
  });
  def('CODE', 1, 1, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s']); if (isErr(a)) return a;
    if (a[0] === '') return E.VALUE;
    // Calc converts the first character to the system encoding (UTF-8) and gives its first byte
    const u = a[0].charCodeAt(0);
    if (u < 0x80) return u;
    if (u >= 0xD800 && u <= 0xDFFF) return 0;
    return new TextEncoder().encode(a[0][0])[0];
  });
  def('CHOOSE', 2, undefined, function (nodes, ctx) {
    const a = this.args(nodes.slice(0, 1), ctx, ['n']); if (isErr(a)) return a;
    const i = Math.trunc(a[0]);
    if (i < 1 || i >= nodes.length) return E.VALUE;
    const v = this.ev(nodes[i], ctx);
    return v === MISSING ? 0 : v;
  });
  def('ROWS', 1, 1, function (nodes, ctx) {
    const v = this.ev(nodes[0], ctx);
    if (isErr(v)) return v;
    return v instanceof Range || v instanceof Matrix ? v.rows : E.VALUE;
  });
  def('COLUMNS', 1, 1, function (nodes, ctx) {
    const v = this.ev(nodes[0], ctx);
    if (isErr(v)) return v;
    return v instanceof Range || v instanceof Matrix ? v.cols : E.VALUE;
  });
  const ELEMENTWISE = new Set(['IF', 'IFERROR', 'ISNUMBER', 'ISTEXT', 'ISBLANK', 'ISERROR', 'ISNA', 'ISERR', 'ISLOGICAL',
    'ISNONTEXT', 'LEN', 'ABS', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'INT', 'TRUNC', 'SIGN', 'NOT', 'UPPER', 'LOWER', 'LEFT',
    'RIGHT', 'MID', 'TRIM', 'VALUE', 'MOD', 'POWER', 'SQRT', 'EXACT', 'EXP', 'LN', 'SEARCH', 'FIND', 'SUBSTITUTE',
    'PROPER', 'REPT', 'CODE', 'CHAR', 'ISODD', 'ISEVEN', 'N', 'T']);
  def('SUMPRODUCT', 1, undefined, function (nodes, ctx) {
    const grids = [];
    for (const n of nodes) {
      if (n.t === 'missing') { grids.push([[0]]); continue; }
      const v = this.aev(n, ctx);
      if (isErr(v)) return v;
      grids.push(v instanceof Matrix ? v.data : [[v]]);
    }
    const R = grids[0].length, C = grids[0][0].length;
    if (grids.some((g) => g.length !== R || g[0].length !== C)) return E.VALUE;
    let sum = 0;
    for (let r = 0; r < R; r++) for (let c = 0; c < C; c++) {
      let p = 1;
      for (const g of grids) {
        const x = g[r][c];
        if (isErr(x)) return x;
        p *= typeof x === 'number' ? x : typeof x === 'boolean' ? +x : 0;
      }
      sum += p;
    }
    return sum;
  });
  // COUNTIFS / SUMIFS / AVERAGEIFS: all criteria ranges must have the same shape
  function multiCond(self, nodes, ctx, first) {
    const pairs = [];
    for (let i = first; i + 1 < nodes.length; i += 2) {
      const rg = self.ev(nodes[i], ctx);
      if (isErr(rg)) return rg;
      if (!(rg instanceof Range)) return E.VALUE;
      let crit = self.sv(nodes[i + 1], ctx);
      if (crit === MISSING) crit = null;
      const m = makeCriterion(crit);
      if (isErr(m)) return m;
      pairs.push([rg, m]);
    }
    const R = pairs[0][0].rows, C = pairs[0][0].cols;
    if (pairs.some(([rg]) => rg.rows !== R || rg.cols !== C)) return E.VALUE;
    const hits = [];
    const r2 = Math.min(R - 1, Math.max(...pairs.map(([rg]) => Math.max(rg.sheet.maxR - rg.r1, 0))));
    const c2 = Math.min(C - 1, Math.max(...pairs.map(([rg]) => Math.max(rg.sheet.maxC - rg.c1, 0))));
    for (let c = 0; c <= c2; c++) for (let r = 0; r <= r2; r++) {
      if (pairs.every(([rg, m]) => { let v = self.cellValue(rg.sheet, rg.r1 + r, rg.c1 + c); if (v && v.emptyResult) v = 0; return m(v); })) hits.push([r, c]);
    }
    const blanksBeyond = pairs.every(([, m]) => m(null)) ? R * C - (r2 + 1) * (c2 + 1) : 0;
    return { hits, blanksBeyond };
  }
  def('COUNTIFS', 2, undefined, function (nodes, ctx) {
    if (nodes.length % 2) return E.VALUE;
    const h = multiCond(this, nodes, ctx, 0); if (isErr(h)) return h;
    return h.hits.length + h.blanksBeyond;
  });
  function sumIfs(self, nodes, ctx) {
    if (nodes.length % 2 === 0) return E.VALUE;
    const sr = self.ev(nodes[0], ctx);
    if (isErr(sr)) return sr;
    if (!(sr instanceof Range)) return E.VALUE;
    const h = multiCond(self, nodes, ctx, 1); if (isErr(h)) return h;
    const first = self.ev(nodes[1], ctx);
    if (first.rows !== sr.rows || first.cols !== sr.cols) return E.VALUE;
    const nums = [];
    for (const [r, c] of h.hits) {
      const v = self.cellValue(sr.sheet, sr.r1 + r, sr.c1 + c);
      if (isErr(v)) return v;
      if (typeof v === 'number') nums.push(v); else if (typeof v === 'boolean') nums.push(+v);
    }
    return nums;
  }
  def('SUMIFS', 3, undefined, function (nodes, ctx) {
    const n = sumIfs(this, nodes, ctx); if (isErr(n)) return n;
    return kahan(n);
  });
  def('AVERAGEIFS', 3, undefined, function (nodes, ctx) {
    const n = sumIfs(this, nodes, ctx); if (isErr(n)) return n;
    return n.length ? kahan(n) / n.length : E.DIV0;
  });
  def('HLOOKUP', 3, 4, function (nodes, ctx) {
    const v = this.sv(nodes[0], ctx);
    const tbl = this.fev(nodes[1], ctx);
    const rest = this.args(nodes.slice(2), ctx, ['n', 'b'], [undefined, true]);
    if (isErr(v) && !v.fromRef) return v.code === '#N/A' ? E.LOOKNA : v;
    if (isErr(tbl)) return tbl;
    if (isErr(rest)) return rest;
    if (isErr(v)) return v;
    const row = Math.trunc(rest[0]);
    const sorted = nodes.length > 3 && nodes[3].t === 'missing' ? false : rest[1];
    if (!(tbl instanceof Range) && !(tbl instanceof Matrix)) return E.VALUE;
    if (row < 1 || row > tbl.rows) return E.VALUE;
    if (v === null) return E.NA;
    const g = this.grid(tbl);
    const i = lookupPos(g[0], v, sorted ? 1 : 0, tbl instanceof Range ? 'row' : null);
    if (i < 0) return E.NA;
    let out = tbl instanceof Range ? this.cellValue(tbl.sheet, tbl.r1 + row - 1, tbl.c1 + i) : g[row - 1][i];
    if (out === EMPTYR) out = 0;
    if (typeof out === 'boolean') return out ? 1 : 0;
    return out === null ? 0 : out;
  });

  // ---- lookup
  def('VLOOKUP', 3, 4, function (nodes, ctx) {
    const v = this.sv(nodes[0], ctx);
    const tbl = this.fev(nodes[1], ctx);
    const rest = this.args(nodes.slice(2), ctx, ['n', 'b'], [undefined, true]);
    if (isErr(v) && !v.fromRef) return v.code === '#N/A' ? E.LOOKNA : v;
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
    const i = lookupPos(first, v, sorted ? 1 : 0, tbl instanceof Range ? 'col' : null);
    if (i < 0) return E.NA;
    let out = tbl instanceof Range ? this.cellValue(tbl.sheet, tbl.r1 + i, tbl.c1 + col - 1) : g[i][col - 1];
    if (out === EMPTYR) out = 0;
    if (typeof out === 'boolean') return out ? 1 : 0;
    return out === null ? 0 : out;
  });
  def('MATCH', 2, 3, function (nodes, ctx) {
    let v = this.sv(nodes[0], ctx);
    if (v === MISSING) v = 0;
    const rg = this.fev(nodes[1], ctx);
    const rest = this.args(nodes.slice(2), ctx, ['n'], [1]);
    if (isErr(v) && !v.fromRef) return v.code === '#N/A' ? E.LOOKNA : v;
    if (isErr(rg)) return rg;
    if (isErr(rest)) return rest;
    if (isErr(v)) return v;
    if (!(rg instanceof Range) && !(rg instanceof Matrix)) return E.NA;
    if (rg.rows !== 1 && rg.cols !== 1) return E.VALUE;
    if (v === null) return E.NA;
    const g = this.grid(rg);
    const list = rg.cols === 1 ? g.map((row) => row[0]) : g[0];
    const mode = rest[0] > 0 ? 1 : rest[0] < 0 ? -1 : 0;
    const i = lookupPos(list, v, mode, rg instanceof Range ? (rg.cols === 1 ? 'col' : 'row') : null);
    return i < 0 ? E.NA : i + 1;
  });
  def('INDEX', 1, 4, function (nodes, ctx) {
    const ref = this.ev(nodes[0], ctx);
    if (isErr(ref)) return ref;
    const a = this.args(nodes.slice(1), ctx, ['n', 'n', 'n'], [0, undefined, 1]);
    if (isErr(a)) return a;
    let row = Math.trunc(a[0]);
    let col = a[1] === undefined ? undefined : Math.trunc(a[1]);
    if (Math.trunc(a[2]) !== 1) return a[2] < 1 ? E.VALUE : E.REF; // a single area
    let tgt = ref;
    const scalar = !(tgt instanceof Range) && !(tgt instanceof Matrix);
    if (scalar) { if (row < 0 || (col !== undefined && col < 0)) return E.VALUE; return E.REF; } // not a reference
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
    if (isErr(v)) return { error: errText(v.code) };
    if (typeof v === 'number') {
      if (!isFinite(v)) return { error: '#NUM!' };
      // integers below 2^53 are written exactly, other numbers with 15 digits
      if (Number.isInteger(v) && Math.abs(v) < 9007199254740992) return v === 0 ? 0 : v;
      const r = Number(v.toPrecision(15));
      return r === 0 ? 0 : r;
    }
    return v;
  }

  function computeWorkbook(json) {
    const eng = new Engine(json);
    recalcAll(eng);
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

  // Recalculate every formula cell.
  // Cells that are not on (or downstream of) a reference cycle get the same
  // value in any order, so they are computed in dependency order, which keeps
  // recursion shallow even for long chains. Cells on or after a cycle depend on
  // where the cycle is entered, so they are computed the way Calc does it:
  // sheet by sheet, column by column, row by row, recursing into precedents.
  function recalcAll(eng) {
    const formulas = [];
    eng.wb.sheets.forEach((sh, si) => {
      for (const cell of sh.cells.values()) if (cell.formula) { cell.sheet = sh; cell.si = si; formulas.push(cell); }
    });
    const deps = new Map();
    for (const cell of formulas) deps.set(cell, staticPrecedents(eng, cell));

    // Tarjan's strongly connected components, iterative.
    let index = 0;
    const idx = new Map(), low = new Map(), onStack = new Set(), stack = [], order = [], cyclic = new Set();
    for (const root of formulas) {
      if (idx.has(root)) continue;
      const work = [[root, 0]];
      idx.set(root, index); low.set(root, index); index++; stack.push(root); onStack.add(root);
      while (work.length) {
        const top = work[work.length - 1];
        const [v, i] = top;
        const ds = deps.get(v);
        if (i < ds.length) {
          top[1]++;
          const w = ds[i];
          if (!idx.has(w)) {
            idx.set(w, index); low.set(w, index); index++; stack.push(w); onStack.add(w);
            work.push([w, 0]);
          } else if (onStack.has(w)) low.set(v, Math.min(low.get(v), idx.get(w)));
          continue;
        }
        work.pop();
        if (work.length) { const u = work[work.length - 1][0]; low.set(u, Math.min(low.get(u), low.get(v))); }
        if (low.get(v) === idx.get(v)) {
          const comp = [];
          let w;
          do { w = stack.pop(); onStack.delete(w); comp.push(w); } while (w !== v);
          if (comp.length > 1 || deps.get(v).indexOf(v) >= 0) comp.forEach((c) => cyclic.add(c));
          order.push(...comp);
        }
      }
    }
    // taint everything downstream of a cycle
    const tainted = new Set(cyclic);
    if (cyclic.size) {
      const users = new Map();
      for (const [c, ds] of deps) for (const d of ds) { if (!users.has(d)) users.set(d, []); users.get(d).push(c); }
      const queue = [...cyclic];
      while (queue.length) {
        const c = queue.pop();
        for (const u of users.get(c) || []) if (!tainted.has(u)) { tainted.add(u); queue.push(u); }
      }
    }
    for (const cell of order) if (!tainted.has(cell)) eng.cellValue(cell.sheet, cell.r, cell.c);
    if (tainted.size) {
      const rest = [...tainted].sort((a, b) => a.si - b.si || a.c - b.c || a.r - b.r);
      for (const cell of rest) eng.cellValue(cell.sheet, cell.r, cell.c);
    }
  }

  // Formula cells a formula may read, from its references (an over-approximation).
  function staticPrecedents(eng, cell) {
    const out = [];
    const seen = new Set();
    const add = (x) => { if (x && x.formula && !seen.has(x)) { seen.add(x); out.push(x); } };
    const walk = (node) => {
      if (!node || typeof node !== 'object') return;
      if (node.t === 'ref3d') {
        const s1 = eng.wb.sheet(node.sheet), s2 = eng.wb.sheet(node.sheet2);
        if (!s1 || !s2) return;
        const i1 = eng.wb.sheets.indexOf(s1), i2 = eng.wb.sheets.indexOf(s2);
        for (let k = Math.min(i1, i2); k <= Math.max(i1, i2); k++) walk({ t: 'ref', sheet: eng.wb.sheets[k].name, a: node.a, b: node.b });
        return;
      }
      if (node.t === 'ref') {
        const s = node.sheet === null ? cell.sheet : eng.wb.sheet(node.sheet);
        if (!s && node.sheet && node.sheet.indexOf(':') > 0) {
          const k = node.sheet.indexOf(':');
          walk({ t: 'ref3d', sheet: node.sheet.slice(0, k), sheet2: node.sheet.slice(k + 1), a: node.a, b: node.b });
          return;
        }
        if (!s) return;
        const b = node.b || node.a;
        const r1 = Math.min(node.a.r, b.r), r2 = Math.min(Math.max(node.a.r, b.r), s.maxR);
        const c1 = Math.min(node.a.c, b.c), c2 = Math.min(Math.max(node.a.c, b.c), s.maxC);
        if (r2 < r1 || c2 < c1) return;
        // formula cells of the sheet by column, sorted by row
        if (!s.fcols) {
          s.fcols = new Map();
          for (const x of s.cells.values()) if (x.formula) { if (!s.fcols.has(x.c)) s.fcols.set(x.c, []); s.fcols.get(x.c).push(x); }
          for (const list of s.fcols.values()) list.sort((p, q) => p.r - q.r);
        }
        const cols = c2 - c1 + 1 > s.fcols.size ? [...s.fcols.keys()].filter((c) => c >= c1 && c <= c2).sort((p, q) => p - q)
          : Array.from({ length: c2 - c1 + 1 }, (_, k) => c1 + k);
        for (const c of cols) {
          const list = s.fcols.get(c);
          if (!list) continue;
          let lo = 0, hi = list.length;
          while (lo < hi) { const m = (lo + hi) >> 1; if (list[m].r < r1) lo = m + 1; else hi = m; }
          for (let k = lo; k < list.length && list[k].r <= r2; k++) add(list[k]);
        }
        return;
      }
      if (node.a && node.a.t) walk(node.a);
      if (node.b && node.b.t) walk(node.b);
      if (node.args) node.args.forEach(walk);
    };
    walk(cell.ast);
    return out;
  }

  const api = {
    computeWorkbook, parse, tokenize, fmtNum, textToNum, parseAddr, addrName, colName, colIndex, Engine, FUNCS,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.SheetEngine = api;
})(typeof window !== 'undefined' ? window : globalThis);
