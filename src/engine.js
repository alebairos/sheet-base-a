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
  // Calc's internal errors without an Excel name: files show them as #VALUE!,
  // but text criteria (COUNTIF(A1,"Err:502")) see their own text
  for (const [k, code] of Object.entries({
    IARG: 'Err:502', PARAM: 'Err:504', PAREN: 'Err:508', NOOP: 'Err:509', NOVAR: 'Err:510', NOARG: 'Err:511',
    CIRC: 'Err:522', STROVF: 'Err:513', OVF: 'Err:514', CODEOVF: 'Err:512',
  })) { E[k] = new FErr(code, false); EREF[code] = new FErr(code, true); }
  // a lookup given a computed error as its lookup value gives an internal
  // error: shown with that error's code (#N/A as Err:LOOKNA), but read as
  // #VALUE! once a later function cleared the pending error
  const ELOOK = {};
  const lookErr = (v) => {
    if (v.code === '#N/A' || v === E.LOOKNA) return E.LOOKNA;
    if (v.look) return v;
    if (!ELOOK[v.code]) { ELOOK[v.code] = new FErr(v.code, false); ELOOK[v.code].look = true; }
    return ELOOK[v.code];
  };
  const errByCode = (code) => Object.values(E).find((e) => e.code === code) || new FErr(code);
  // (files show the lookup error and Err:513, string overflow, as #N/A)
  const errText = (code) => (code[0] === '#' ? code : code === 'Err:LOOKNA' || code === 'Err:MISS' || code === 'Err:512' || code === 'Err:513' || code === 'Err:514' ? '#N/A' : '#VALUE!');
  const isErr = (v) => v instanceof FErr;
  const asRefErr = (v) => (isErr(v) && !v.fromRef ? (EREF[v.code] || new FErr(v.code, true)) : v);
  // an error whose pending state was cleared (IFERROR, ISNUMBER, ... ran
  // later) is an ordinary operand; Calc's internal lookup error then reads as #VALUE!
  // (an error token on Calc's stack: popped as an operand it replaces the
  // error a conversion of the other operand set)
  const ETOK = {};
  // an error constant in a formula (=#N/A): an error token; a lookup given
  // one as its lookup value gives #VALUE!
  const ELIT = {};
  const litErr = (code) => {
    if (!ELIT[code]) { ELIT[code] = new FErr(code, true); ELIT[code].tok = true; ELIT[code].lit = true; }
    return ELIT[code];
  };
  // a computed error kept as a token on Calc's stack (see Engine.ev0)
  const ECONT = {};
  const tokOf = (v) => {
    if (!ECONT[v.code]) { const e = ECONT[v.code] = new FErr(v.code, true); e.tok = true; e.lit = true; e.cont = true; }
    return ECONT[v.code];
  };
  // the error functions in a formula (Calc's IsErrFunc), marked and counted once
  const countErrFuncs = (ast) => {
    if (!ast || typeof ast !== 'object') return 0;
    if (ast.efCount !== undefined) return ast.efCount;
    let n = 0;
    const walk = (x) => {
      if (!x || typeof x !== 'object') return;
      if (Array.isArray(x)) { for (const y of x) walk(y); return; }
      if (x.t === 'fn' && ERR_FUNCS.has(x.name)) {
        x.ef = true; n++;
        if (x.args) Object.defineProperty(x.args, 'ef', { value: x, enumerable: false, configurable: true });
      }
      for (const k in x) if (k !== 'ef' && k !== 'efCount') walk(x[k]);
    };
    walk(ast);
    Object.defineProperty(ast, 'efCount', { value: n, enumerable: false });
    return n;
  };
  const ERR_FUNCS = new Set(['IFERROR', 'ISNUMBER', 'ISTEXT', 'ISBLANK', 'ISERROR', 'ISERR', 'ISNA', 'ISLOGICAL',
    'ISNONTEXT', 'ISREF', 'COUNT', 'COUNTA', 'N']);
  const demote = (v) => {
    if (!isErr(v) || v.fromRef) return v;
    const code = v === E.LOOKNA || v.look ? '#VALUE!' : v.code;
    if (!ETOK[code]) { ETOK[code] = new FErr(code, true); ETOK[code].tok = true; }
    return ETOK[code];
  };
  // the error token of an earlier parameter (popped later) replaces a
  // conversion error met in parameter k
  const tokBefore = (vals, k, e) => {
    // (error tokens overwrite the error as Calc pops them: the leftmost wins)
    for (let j = 0; j < k; j++) if (isErr(vals[j]) && vals[j].tok) return vals[j];
    return e;
  };
  // (Calc's internal lookup error read from a cell keeps its code when a
  // later reset makes it an operand: it stays #N/A in files)
  const LOOKNA_CELL = new FErr('Err:LOOKNA', false);
  const asComputed = (v) => (isErr(v) && v.fromRef ? (v.code === 'Err:LOOKNA' ? LOOKNA_CELL : errByCode(v.code)) : v);

  // ---------------------------------------------------------------- helpers
  const MAXROW = 1048576, MAXCOL = 16384;
  const AMISS = { amiss: true }; // an array element beyond a smaller array (see amap)
  const E48 = Math.pow(2, -48);

  function approxEqual(a, b) {
    if (a === b) return true;
    if (a === 0 || b === 0) return false;
    const d = Math.abs(a - b);
    if (!isFinite(d)) return false;
    // (beyond 2^-48 but within 2^-44, values that round to the same 15
    // significant digits still count as equal: 1223082542420 = 1223082542420+18*2^-12)
    if (d > Math.abs(a) * E48 || d > Math.abs(b) * E48) {
      if (d > Math.abs(a) * E48 * 16 || d > Math.abs(b) * E48 * 16) return false;
      if (round15fp(a) !== round15fp(b)) return false;
    }
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
  // x rounded to 15 significant digits by scaling with a power of ten in floating point
  function round15fp(x) {
    const a = Math.abs(x);
    const ne = 14 - Math.floor(Math.log10(a));
    const f = Math.pow(10, Math.abs(ne));
    let v = ne < 0 ? a / f : a * f;
    if (!isFinite(v)) return x;
    v = Math.round(v);
    v = ne < 0 ? v * f : v / f;
    if (!isFinite(v)) return x;
    return x < 0 ? -v : v;
  }
  function approxValue(x) {
    if (!isFinite(x) || x === 0 || x > 2199023255552) return x;
    const a = Math.abs(x);
    if ((Number.isInteger(a) && a < 9007199254740992) || fracBits(a) <= 11) return x;
    // 15 significant digits as rtl_math_approxValue gets them: scaled by a
    // power of ten in floating point, so the scaling's own rounding counts
    // (ROUND(5805.4999999999945,0) is 5806)
    const ne = 14 - Math.floor(Math.log10(a));
    const f = Math.pow(10, Math.abs(ne));
    let v = ne < 0 ? a / f : a * f;
    if (!isFinite(v)) return x;
    v = Math.round(v);
    v = ne < 0 ? v * f : v / f;
    if (!isFinite(v)) return x;
    return x < 0 ? -v : v;
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

  // x with 15 significant digits as Calc writes it: the shortest decimal
  // that round-trips, rounded half up (so 0.19565217391304349, shortest
  // 0.1956521739130435, becomes 0.195652173913044). Returns toExponential(14) form.
  function exp15(x) {
    const [m, ex] = Math.abs(x).toExponential().split('e');
    let e = parseInt(ex, 10);
    let d = m.replace('.', '');
    if (d.length > 15) {
      const up = d.charCodeAt(15) >= 53; // '5'
      d = d.slice(0, 15);
      if (up) {
        const a = d.split('');
        let k = 14;
        while (k >= 0) { if (a[k] === '9') { a[k] = '0'; k--; } else { a[k] = String.fromCharCode(a[k].charCodeAt(0) + 1); break; } }
        if (k < 0) { a.unshift('1'); a.pop(); e++; }
        d = a.join('');
      }
    } else d = d.padEnd(15, '0');
    return (x < 0 ? '-' : '') + d[0] + '.' + d.slice(1) + 'e' + (e < 0 ? '-' : '+') + Math.abs(e);
  }
  const round15 = (x) => Number(exp15(x));

  // General number format as Calc shows it (used for number -> text).
  function fmtNum(x) {
    if (x === 0 || Object.is(x, -0)) return '0';
    if (!isFinite(x)) return '#NUM!';
    if (Number.isInteger(x) && Math.abs(x) < 9007199254740992) return String(x); // exact below 2^53
    // 15 significant digits (all of them when 15 would round past the largest double)
    let s = isFinite(Number(exp15(x))) ? exp15(x) : x.toExponential();
    let [mant, exp] = s.split('e');
    let e = parseInt(exp, 10);
    // fixed or scientific follows the magnitude before rounding
    // (999999999999999.9 is written 1000000000000000)
    const eRaw = parseInt(x.toExponential().split('e')[1], 10);
    if (eRaw <= -15 || eRaw >= 15) {
      mant = mant.indexOf('.') >= 0 ? mant.replace(/0+$/, '').replace(/\.$/, '') : mant;
      const sign = e < 0 ? '-' : '+';
      return mant + 'E' + sign + String(Math.abs(e)).padStart(3, '0');
    }
    const v = round15(x);
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
    // an exponent overflow reads as the largest number (whatever the sign, as Calc)
    return isFinite(v) ? { v, exp: m[3] !== undefined } : { v: Number.MAX_VALUE, exp: true, overflow: true };
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
  // dates with an English month name, as Calc's en-US input accepts them:
  // "Jan 1, 2024", "January 1 2024", "Sep. 5, 99", "Jan 2024", "Jan 31"
  // (this year), "1-Jan-2024", "Jan/1/2024"
  const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  const FULL_MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september',
    'october', 'november', 'december'];
  function monthIndex(w) {
    w = w.toLowerCase().replace(/\.$/, '');
    let i = MONTHS.indexOf(w);
    if (i < 0) i = FULL_MONTHS.indexOf(w);
    if (i < 0 && w === 'sept') i = 8;
    return i < 0 ? null : i + 1;
  }
  function monthDateVal(t) {
    let m = /^([A-Za-z]+\.?)(?: *([-/]) *| +)(\d{1,4})(?:(?: *, +| *,? +| *\/ *)(\d{1,4}))?$/.exec(t);
    if (m) {
      const mo = monthIndex(m[1]);
      if (mo === null) return null;
      if (m[2] === '-' && m[4] !== undefined) return null; // "Jan-1-2024"
      if (m[4] !== undefined) {
        if (m[3].length > 2) return null;
        return daySerial(fullYear(+m[4], m[4].length), mo, +m[3]);
      }
      // one number: a day of this year, or a year
      if (m[3].length <= 2 && +m[3] <= 31) return daySerial(new Date().getFullYear(), mo, +m[3]);
      return daySerial(fullYear(+m[3], m[3].length), mo, 1);
    }
    m = /^(\d{1,2})-([A-Za-z]+\.?)-(\d{1,4})$/.exec(t);
    if (m) {
      const mo = monthIndex(m[2]);
      return mo === null ? null : daySerial(fullYear(+m[3], m[3].length), mo, +m[1]);
    }
    return null;
  }
  function dateTimeVal(t) {
    const tv = timeVal(t);
    if (tv !== null) return tv;
    const dv = dateVal(t);
    if (dv !== null) return dv;
    if (/[A-Za-z]{2}/.test(t)) {
      const md = monthDateVal(t);
      if (md !== null) return md;
      const mt = /^(.*?) +(\d+:[\d:.]*(?: *[ap]\.?m\.?)?)$/i.exec(t);
      if (mt) {
        const d = monthDateVal(mt[1]), tm = timeVal(mt[2]);
        if (d !== null && tm !== null) return d + tm;
      }
      return null;
    }
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
    if (/^[A-Za-z]{3}|^\d+-[A-Za-z]/.test(m[2])) return m[1] ? null : dateTimeVal(m[2]);
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
    if (pn.overflow) return pct || cur ? null : v;
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
  // An array IF without an else branch leaves an empty element where the
  // condition is false: skipped by aggregates (like an empty element) but
  // FALSE where it is taken as a value (epm marks those positions)
  const EPATH = { emptyPath: true };
  const epath = (m) => {
    if (!m.data.some((row) => row.includes(EPATH))) return m;
    m.epm = m.data.map((row) => row.map((x) => x === EPATH));
    m.data = m.data.map((row) => row.map((x) => (x === EPATH ? null : x)));
    return m;
  };
  const m00 = (m) => (m.epm && m.epm[0][0] ? EPATH : m.data[0][0]);
  // a union of references, (A1:A2,B1): only aggregate functions take it
  class RefList {
    constructor(ranges) { this.ranges = ranges; }
  }
  // a 3D reference (Sheet1:Sheet3!A1): one range per sheet
  const refList3d = (ranges) => { const l = new RefList(ranges); l.is3d = true; return l; };
  const MISSING = { missing: true };
  // an empty argument that IF, CHOOSE or IFERROR passes on (IF(1,)): Calc's
  // missing token stays on the stack: 0 as a number, "" as text, neither
  // number nor text; as the formula's result 0, but IFERROR's own empty
  // value argument (IFERROR(,5)) passes one on that ends as #N/A
  const MISSV = { missing: true, passed: true };
  const MISSN = { missing: true, passed: true, na: true };
  const isMV = (v) => v === MISSV || v === MISSN;
  // result of calling an unknown function with arguments: #NAME?, but
  // IFERROR/ISERROR turn it into #VALUE! instead of catching it
  const UNKNOWN_FN = new FErr('#NAME?', false);
  // value of a formula whose result is an empty cell (e.g. =B1 with B1 empty):
  // shown as 0, but compares like an empty cell and is neither number nor text
  const EMPTYR = { emptyResult: true };
  // a blank (LBLANK) or empty-result (LEMPTY) cell returned by VLOOKUP/HLOOKUP:
  // 0 as a number, "" as text, counted by COUNT/COUNTA; only the blank one
  // is ISBLANK
  const LBLANK = { lookupBlank: true }, LEMPTY = { lookupEmpty: true };
  const isLB = (v) => v === LBLANK || v === LEMPTY;

  // ---------------------------------------------------------------- tokenizer
  const ERRLITS = ['#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#N/A', '#NUM!', '#NULL!'];

  function tokenize(src) {
    const toks = [];
    let i = 0;
    const n = src.length;
    const fail0 = (msg, code) => { const e = new SyntaxError(msg); e.code = code; throw e; };
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
      // (an inline array holding anything but constants fails the whole
      // formula, before other syntax errors: #N/A)
      if (ch === '{') {
        let j = i + 1, inStr = false;
        while (j < n && (inStr || src[j] !== '}')) { if (src[j] === '"') inStr = !inStr; else if (!inStr && src[j] === '{') break; j++; }
        if (j < n && src[j] === '}') {
          const body = src.slice(i + 1, j);
          const elems = body.match(/("(?:[^"]|"")*"|[^,;"]+)*|[,;]/g) ? body.split(/[,;](?=(?:[^"]*"[^"]*")*[^"]*$)/) : [body];
          const okEl = (e) => /^\s*(?:[+-]?\s*(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?%?|"(?:[^"]|"")*"|TRUE|FALSE|)\s*$/i.test(e);
          if (!elems.every(okEl)) fail0('bad inline array', '#N/A');
        }
      }
      // (an inline array opened inside another fails the whole formula: #N/A)
      if (ch === '{' && toks.some((t) => t.t === 'op' && t.v === '{') && toks.filter((t) => t.t === 'op' && t.v === '{').length > toks.filter((t) => t.t === 'op' && t.v === '}').length) fail0('nested array', '#N/A');
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
        let lit = ERRLITS.find((e) => src.substr(i, e.length).toUpperCase() === e);
        // (an error constant running into more name characters is a name: #N/A!)
        if (lit === '#N/A' && /^[\d!]/.test(src.slice(i + lit.length))) {
          const w = /^#[\w.!?/$]*/.exec(src.slice(i))[0];
          toks.push({ t: 'name', v: w });
          i += w.length;
          continue;
        }
        // #REF! (an invalid reference) and unknown #words read as unknown names
        if ((!lit && !/^#GETTING_DATA/i.test(src.slice(i))) || lit === '#REF!') {
          const w = /^#[\w.]*[!?]?\w*/.exec(src.slice(i))[0];
          toks.push({ t: 'name', v: w });
          i += w.length;
          continue;
        }
        if (!lit) throw new SyntaxError('bad error literal');
        toks.push({ t: 'err', v: lit });
        i += lit.length;
        continue;
      }
      // sheet-qualified reference (sheet2: a 3D range Sheet1:Sheet3!A1)
      let sheet = null, sheet2 = null;
      let j = i;
      {
        const nm = "(?:'((?:[^']|'')+)'|([A-Za-z_\\p{L}\\p{M}][\\w.\\p{L}\\p{N}\\p{M}\u00B7]*))";
        // (the second sheet name may start with a digit: A1:2024!A1)
        const m3 = new RegExp('^' + nm + ':' + nm.replace('[A-Za-z_', '[\\w') + '!', 'u').exec(src.slice(i));
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
        const m = /^([A-Za-z_\p{L}\p{M}][\w.\p{L}\p{N}\p{M}\u00B7]*)!/u.exec(src.slice(i));
        if (m) { sheet = m[1]; j = i + m[0].length; }
      }
      const rest = src.slice(j);
      if (sheet !== null || /^\$?[A-Za-z]/.test(rest) || /^\$?\d/.test(rest)) {
        let m = refRe.exec(rest);
        if (m && !/^[\w($\p{L}\p{N}\p{M}\u00B7]/u.test(rest.slice(m[0].length)) && !/^[.]/.test(rest.slice(m[0].length))) {
          const a = parseAddr(m[1]);
          const b = m[2] ? parseAddr(m[2]) : null;
          if (a && (!m[2] || b)) {
            // (xlsx import: a sheet's cell followed by a lone column, Sheet2!A1:B,
            // is a range along the cell's row)
            const cm = sheet !== null && !m[2] ? /^:(\$?)([A-Za-z]{1,3})(?![\w.$(!:])/.exec(rest.slice(m[0].length)) : null;
            if (cm && colIndex(cm[2]) < MAXCOL) {
              pushTok({ t: 'ref', sheet, sheet2, a, b: { r: a.r, c: colIndex(cm[2]) }, text: src.slice(i, j + m[0].length + cm[0].length) });
              i = j + m[0].length + cm[0].length;
              continue;
            }
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
          if (sheet !== null) {
            // (a sheet name before something that is not a reference: an unknown name)
            const w = /^[\w.$\p{L}\p{N}\p{M}\u00B7]*/u.exec(rest)[0];
            if (!w && j > i) { toks.push({ t: 'name', v: src.slice(i, j) }); i = j; continue; }
            if (!w || /^[(]/.test(rest.slice(w.length)) && /^\$?[A-Za-z]{1,3}\$?\d+$/.test(w)) throw new SyntaxError('bad reference');
            toks.push({ t: 'name', v: src.slice(i, j) + w });
            i = j + w.length;
            continue;
          }
        }
      }
      // ('[' (an external or table reference) or '.' before a letter starts a name)
      if (ch === ']' || (ch === '[' && src.indexOf(']', i) < 0)) fail0('unbalanced bracket', '#N/A');
      if (ch === '[' || (ch === '.' && /^\.[A-Za-z_.]/.test(src.slice(i)))) {
        const w = /^(?:\[[^\]]*\]?)?[\w.$!:\p{L}\p{N}\p{M}\u00B7]*/u.exec(src.slice(i))[0];
        toks.push({ t: 'name', v: w });
        i += w.length;
        continue;
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
        const after = src.slice(i + m[0].length);
        // a number running into letters is one (unknown) name, as in Calc
        // (2x, 1E5E5), unless a signed exponent ended it (1e+5x); an
        // unfinished exponent before anything but the end or a sign is a
        // syntax error (1e*2)
        if (/^[A-Za-z_\p{L}\p{N}\p{M}\u00B7]/u.test(after) && !/[eE][+-]/.test(m[0])) {
          const w = /^[\w.\p{L}\p{N}\p{M}\u00B7]*/u.exec(after)[0];
          const next = after.slice(w.length);
          if (/^[eE]$/.test(w) && next !== '' && !/^[+-]/.test(next)) fail0('bad exponent', 'Err:509');
          toks.push({ t: 'name', v: m[0] + w });
          i += m[0].length + w.length;
          continue;
        }
        // (an exponent running into a '.' when the mantissa had none, or into
        // another signed exponent, is a name too: 1E3.5, 1E-3., .5E3.5,
        // 1.5e-1E-3; 1.5E3.5 is a syntax error)
        if (m[2] && ((after[0] === '.' && !/^\d+\./.test(m[1])) || /^[eE][+-]/.test(after))) {
          // (after the dot: up to the next dot, letters too unless the exponent was signed)
          const w = after[0] === '.' ? (/[+-]/.test(m[2]) ? /^\.\d*(?:[eE][+-]?\d+)?/ : /^\.\w*/).exec(after)[0] : after[0];
          toks.push({ t: 'name', v: m[0] + w });
          i += m[0].length + w.length;
          continue;
        }
        const v = Number(m[0]);
        // a literal beyond the double range is an error for the whole formula
        if (!isFinite(v) || (v === 0 && /[1-9]/.test(m[1]))) fail0('number out of range', 'Err:503');
        toks.push({ t: 'num', v });
        i += m[0].length;
        continue;
      }
      // ('$' is a name character to Calc: a stray one makes an unknown name)
      // (non-ASCII spaces start a name too)
      if (/^[A-Za-z_$\p{L}\p{N}\p{M}\p{Zs}]/u.test(src.slice(i))) {
        const m = /^[A-Za-z_$\p{L}\p{N}\p{M}\p{Zs}][\w.$\p{L}\p{N}\p{M}\u00B7]*/u.exec(src.slice(i));
        toks.push({ t: 'name', v: m[0] });
        i += m[0].length;
        continue;
      }
      const two = src.substr(i, 2);
      if (two === '<=' || two === '>=' || two === '<>') { toks.push({ t: 'op', v: two }); i += 2; continue; }
      if ('+-*/^&=<>%(),;{}:!@'.indexOf(ch) >= 0) { toks.push({ t: 'op', v: ch }); i++; continue; }
      // (a symbol or punctuation outside ASCII fails the whole formula: #N/A;
      // a stray ~ | \ ` ? where an operand belongs, at the end (or ? before a
      // name), is an unknown name)
      if (ch.charCodeAt(0) >= 0x80) fail0('unexpected character ' + ch, '#N/A');
      const prevOp = !toks.length || (toks[toks.length - 1].t === 'op' && !/^[)%}]$/.test(toks[toks.length - 1].v));
      if (prevOp && '~|\\`?'.indexOf(ch) >= 0) {
        const w = ch === '?' ? /^\?[\w.$]*/.exec(src.slice(i))[0] : ch;
        if (ch === '?' || !src.slice(i + 1).trim()) { toks.push({ t: 'name', v: w }); i += w.length; continue; }
      }
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
    // Calc's syntax errors (all shown as #VALUE!): Err:508 for a bracket or
    // separator problem, Err:509 for a missing operator, Err:510 for a missing
    // operand, Err:511 for an unfinished argument list
    const fail = (msg, code) => { const e = new SyntaxError(msg); e.code = code; throw e; };
    const operandNext = () => toks[p].t !== 'op' || toks[p].v === '(' || toks[p].v === '{';
    const expect = (v) => {
      if (v === ')' && toks[p].t === 'eof') return;
      if (!isOp(v)) fail('expected ' + v, operandNext() ? 'Err:509' : 'Err:508');
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
      if (isOp('@')) { p++; return { t: 'at', a: unary() }; } // implicit intersection
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
      // an unqualified reference after a qualified one is on that one's sheet
      // (Sheet2!A1:B2:A3); different sheets make a 3D reference
      let sheet = a.t === 'ref' ? a.sheet : null;
      while (isOp(':')) {
        p++;
        let b = primary();
        if (b.t === 'ref' && b.sheet === null && sheet !== null && !b.b) b = { ...b, sheet };
        if (b.t === 'ref') sheet = b.sheet;
        a = { t: 'rangeop', a, b };
      }
      return a;
    }
    function args() {
      const list = [];
      if (isOp(')')) { p++; return list; }
      for (;;) {
        if (isOp(',')) { list.push({ t: 'missing' }); p++; continue; }
        if (isOp(')')) { list.push({ t: 'missing' }); p++; return list; }
        if (toks[p].t === 'eof') { if (list.length) fail('unterminated argument list', 'Err:511'); return list; }
        list.push(comparison());
        if (isOp(',')) {
          p++;
          if (isOp(')')) { list.push({ t: 'missing' }); p++; return list; }
          if (toks[p].t === 'eof') fail('unterminated argument list', 'Err:511');
          continue;
        }
        expect(')');
        return list;
      }
    }
    function arrayConst() {
      // Calc's inline arrays: numbers (one sign at most), text and TRUE/FALSE
      // (stored as the numbers 1 and 0); an empty element is an empty matrix
      // element, {} one such element; anything else (errors, expressions) is a
      // syntax error shown as #N/A
      const bad = () => fail('bad array constant', '#N/A');
      const rows = [[]];
      for (;;) {
        let sign = 0;
        if (isOp('-') || isOp('+')) { sign = toks[p].v === '-' ? -1 : 1; p++; }
        const t = toks[p];
        let v;
        if (!sign && t.t === 'op' && (t.v === ',' || t.v === ';' || t.v === '}')) v = null;
        else {
          p++;
          if (t.t === 'num') v = sign < 0 ? -t.v : t.v;
          else if (t.t === 'str' && !sign) v = t.v;
          else if (t.t === 'name' && !sign && /^(TRUE|FALSE)$/i.test(t.v) && !isOp('(')) v = t.v.toUpperCase() === 'TRUE' ? 1 : 0;
          else bad();
        }
        rows[rows.length - 1].push(v);
        if (isOp(',')) { p++; continue; }
        if (isOp(';')) { p++; rows.push([]); continue; }
        if (toks[p].t === 'eof' && v !== null) break; // (an array left open after a value is closed, as Calc does)
        if (!isOp('}')) bad();
        p++;
        break;
      }
      const w = rows[0].length;
      if (rows.some((r) => r.length !== w)) bad();
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
          if (isOp('(')) { p++; return { t: 'fn', name: t.v.toUpperCase(), args: args() }; }
          const u = t.v.toUpperCase();
          if (u === 'TRUE') return { t: 'bool', v: true };
          if (u === 'FALSE') return { t: 'bool', v: false };
          return { t: 'name', v: t.v };
        }
        case 'op':
          if (t.v === '(') {
            if (peek().t === 'eof') fail('unterminated (', 'Err:510');
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
          if (t.v === '%') { const e = new SyntaxError('unexpected %'); e.code = '#N/A'; throw e; }
          if (t.v === '-') return { t: 'neg', a: unary() };
          if (t.v === '+') return unary();
          if (t.v === '@') return { t: 'at', a: unary() };
          break;
      }
      const e = new SyntaxError('unexpected token');
      if (t.t === 'eof') e.code = '#N/A';
      else if (t.t === 'op') e.code = t.v === ')' || t.v === ';' ? 'Err:508' : 'Err:510';
      throw e;
    }
    const ast = comparison();
    if (peek().t !== 'eof') fail('trailing input', isOp(')') || isOp(';') ? 'Err:508' : 'Err:509');
    return ast;
  }

  // ---------------------------------------------------------------- workbook
  class Workbook {
    constructor(json) {
      this.sheets = [];
      this.byName = new Map();
      const sheets = (json && json.sheets) || {};
      for (const name of sheets.__order || Object.keys(sheets)) {
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
      sh.fcols = null; // formula-cell index used by the dependency scan
      // an empty text constant is an empty cell (as after an xlsx round trip)
      // so is a bare '=' (an empty formula)
      if (raw === null || raw === undefined || raw === '' || raw === '=') { sh.cells.delete(key); return; }
      const cell = { r, c, raw, addr: addr || addrName(r, c), state: 0, value: null, ast: null };
      if (typeof raw === 'string' && raw.charAt(0) === '=' && raw.length > 1) {
        cell.formula = true;
        // Calc's compiler skips further leading '=' signs (and spaces between them): '==1' is 1
        try { cell.ast = parse(raw.slice(1).replace(/^[\s=]*=/, '')); } catch (e) { cell.parseError = e.code || '#VALUE!'; }
        if (cell.ast && raw.length > 196) { const e = compileLimits(raw.slice(1), cell.ast); if (e) { cell.ast = null; cell.parseError = e; } }
        if (cell.ast) { const e = compileArity(cell.ast); if (e) { cell.ast = null; cell.parseError = e; } }
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
      this.clr = 0; // count of pending-error resets (see CLEARS_ERROR)
      this.errSeq = 0;
    }

    // value of a cell (evaluating formulas), null for blank
    cellValue(sh, r, c) {
      const cell = sh.cells.get(r * MAXCOL + c);
      if (!cell) return null;
      if (cell.state === 2) return cell.value;
      // A cell that is being computed reads as an error (Calc's Err:522, which
      // round-trips as #VALUE!); the error then propagates through the cycle.
      // If the formulas computing between that cell and this read already have
      // a pending error, the read gives the earliest such error instead (Calc
      // reads the running cell while that error is set).
      if (cell.state === 1) {
        const st = this.running;
        let i = st.length - 1;
        while (i >= 0 && st[i].cell !== cell) i--;
        for (let k = i; k >= 0 && k < st.length; k++) {
          const fe = st[k].fe;
          if (fe && fe.clr === (k + 1 < st.length ? st[k + 1].clr : this.clr)) return fe.err;
        }
        return EREF['Err:522'];
      }
      this.evalCell(sh, cell);
      return cell.value;
    }

    evalCell(sh, cell) {
      cell.state = 1;
      if (!this.running) this.running = [];
      // fe: the first error computed in the formula so far (Calc's pending
      // error); clr: the caller's reset count when this cell started
      this.running.push({ cell, fe: null, clr: this.clr });
      let v;
      if (cell.parseError || !cell.ast) v = cell.parseError ? errByCode(cell.parseError) : null;
      else {
        const ctx = { sheet: sh, r: cell.r, c: cell.c };
        const clr = this.clr, pend = this.pend; // each cell has its own pending error (own interpreter)
        const seen = this.errSeen, lce = this.lastChildErr;
        const efTotal = this.efTotal, efDone = this.efDone, efSeen = this.efSeen;
        this.efTotal = countErrFuncs(cell.ast); this.efDone = 0; this.efSeen = new Set();
        const gerr = this.gerr, inElem = this.inElem; this.gerr = null; this.inElem = 0;
        const fa = this.fa, root = this.root; this.fa = 0; this.root = cell.ast;
        this.pend = null;
        try {
          v = this.finalValue(this.ev(cell.ast, ctx), ctx);
          // (a scalar error computed before the pending one ended the formula first)
          const fe = this.running[this.running.length - 1].se;
          const earlier = fe && fe.clr === this.clr && this.pend && fe.seq < this.pend.seq && fe.err.code !== this.pend.err.code && isErr(v) && !v.tok;
          if (this.pend && this.pend.clr === this.clr && !earlier) v = this.pend.err;
          // (the interpreter went on past an error no function cleared: that error)
          else if (this.gerr && this.gerr.clr === this.clr && !(isErr(v) && v.tok)) v = this.gerr.err;
        } catch (e) {
          if (e instanceof RangeError) v = E.VALUE; // stack overflow etc.
          else if (e instanceof JumpAbort) v = e.err;
          else throw e;
        }
        this.fa = fa; this.root = root;
        this.clr = clr;
        this.pend = pend;
        this.errSeen = seen; this.lastChildErr = lce;
        this.efTotal = efTotal; this.efDone = efDone; this.efSeen = efSeen; this.gerr = gerr; this.inElem = inElem;
      }
      this.running.pop();
      // (a cell left with IFERROR's missing value: an error shown as #N/A, not ISNA)
      if (isMV(v)) v = v.na ? new FErr('Err:MISS', true) : EMPTYR;
      if (isErr(v)) v = asComputed(v);
      cell.value = v;
      cell.state = 2;
    }


    finalValue(v, ctx) {
      if (v instanceof Range) v = this.intersect(v, ctx);
      else if (v instanceof Matrix) { const b = v.boolRes; v = m00(v); if (b && (v === 0 || v === 1)) v = v === 1; }
      else if (v instanceof RefList) {
        // a 3D reference as the result: the formula's own sheet's range, if
        // it is among the sheets (inside operations it is an error)
        v = v.is3d ? this.own3d(v, ctx) : E.VALUE;
      }
      if (v === null) return EMPTYR;
      if (v === EPATH) return false;
      if (v === MISSING || isLB(v)) return 0;
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
      // a block: the cell in the formula's row and column, when both are inside
      if (ctx && ctx.r >= rg.r1 && ctx.r <= rg.r2 && ctx.c >= rg.c1 && ctx.c <= rg.c2) {
        if (ctx.sheet === rg.sheet) return EREF['Err:522']; // the formula cell itself
        return asRefErr(this.cellValue(rg.sheet, ctx.r, ctx.c));
      }
      return EREF['#VALUE!'];
    }

    // scalar value of an operand
    scalar(v, ctx) {
      if (v instanceof Range) return this.intersect(v, ctx);
      if (v instanceof Matrix) return m00(v);
      if (v instanceof RefList) return v.is3d ? this.own3d(v, ctx) : E.IARG;
      return v;
    }
    // a 3D reference in a scalar position: its range on the formula's own
    // sheet, intersected with the formula's row or column (also one cell)
    own3d(v, ctx) {
      const own = ctx && v.ranges.find((x) => x.sheet === ctx.sheet);
      // (no cell: a conversion error, not a pending one)
      if (!own) return EREF['#VALUE!'];
      if (own.c1 === own.c2 && ctx.r >= own.r1 && ctx.r <= own.r2) return asRefErr(this.cellValue(own.sheet, ctx.r, own.c1));
      if (own.r1 === own.r2 && ctx.c >= own.c1 && ctx.c <= own.c2) return asRefErr(this.cellValue(own.sheet, own.r1, ctx.c));
      if (ctx.r >= own.r1 && ctx.r <= own.r2 && ctx.c >= own.c1 && ctx.c <= own.c2) return asRefErr(this.cellValue(own.sheet, ctx.r, ctx.c));
      return EREF['#VALUE!'];
    }

    sv(node, ctx) { return this.scalar(this.ev(node, ctx), ctx); }
    // scalar values of several arguments; an error computed before a later
    // argument cleared the pending error becomes an ordinary operand
    svAll(nodes, ctx) {
      const vals = [], clrs = [];
      for (const n of nodes) { vals.push(this.sv(n, ctx)); clrs.push(this.clr); }
      return vals.map((v, i) => (clrs[i] !== this.clr ? demote(v) : v));
    }

    // iterate a range column-major (as Calc does); fn(value, r, c) may return false to stop
    // pre(r, c), given for cells that are not blank, may return true to skip
    // a cell without computing it
    eachCell(rg, fn, includeBlank, pre) {
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
          if (pre && pre(cell.r, cell.c)) continue;
          const v = this.cellValue(sh, cell.r, cell.c);
          if (fn(v === EMPTYR ? 0 : v, cell.r, cell.c) === false) return;
        }
        return;
      }
      for (let c = rg.c1; c <= c2; c++) {
        for (let r = rg.r1; r <= r2; r++) {
          if (pre && sh.cells.has(r * MAXCOL + c) && pre(r, c)) continue;
          let v = this.cellValue(sh, r, c);
          if (v === EMPTYR) v = 0;
          if (v === null && !includeBlank) continue;
          if (fn(v, r, c) === false) return;
        }
      }
    }

    // number of non-blank cells in a range, without computing formulas
    countCells(rg) {
      const sh = rg.sheet;
      const r2 = Math.min(rg.r2, sh.maxR), c2 = Math.min(rg.c2, sh.maxC);
      if (r2 < rg.r1 || c2 < rg.c1) return 0;
      const area = (r2 - rg.r1 + 1) * (c2 - rg.c1 + 1);
      let n = 0;
      if (area > sh.cells.size) {
        for (const cell of sh.cells.values()) if (cell.r >= rg.r1 && cell.r <= r2 && cell.c >= rg.c1 && cell.c <= c2) n++;
      } else {
        for (let c = rg.c1; c <= c2; c++) for (let r = rg.r1; r <= r2; r++) if (sh.cells.has(r * MAXCOL + c)) n++;
      }
      return n;
    }

    // the first column (or row) of a range as grid() gives it; cached for
    // lookups once every cell in it is final (none is being computed)
    firstLine(v, byRow) {
      if (!(v instanceof Range)) { const g = this.grid(v); return byRow ? g[0] : g.map((row) => row[0]); }
      const r2 = Math.min(v.r2, Math.max(v.sheet.maxR, v.r1)), c2 = Math.min(v.c2, Math.max(v.sheet.maxC, v.c1));
      const key = v.sheet.name + '|' + v.r1 + '|' + v.c1 + '|' + (byRow ? 'r' + c2 : 'c' + r2);
      if (!this.lineCache) this.lineCache = new Map();
      const hit = this.lineCache.get(key);
      if (hit) return hit;
      const n = byRow ? c2 - v.c1 + 1 : r2 - v.r1 + 1;
      const pos = (i) => (byRow ? [v.r1, v.c1 + i] : [v.r1 + i, v.c1]);
      const cellAt = (i) => { const [r, c] = pos(i); return v.sheet.cells.get(r * MAXCOL + c); };
      const get = (i) => { const [r, c] = pos(i); const x = this.cellValue(v.sheet, r, c); return x === EMPTYR ? 0 : x; };
      let final = true;
      for (let i = 0; i < n && final; i++) { const cell = cellAt(i); if (cell && cell.state !== 2 && cell.ast) final = false; }
      if (final) {
        const out = [];
        for (let i = 0; i < n; i++) out.push(get(i));
        this.lineCache.set(key, out);
        return out;
      }
      // cells still to compute are read only when the search reaches them
      // (Calc computes only the cells a lookup probes, so a lookup over a
      // range holding its own cell need not be circular)
      const UNREAD = {}, vals = new Array(n).fill(UNREAD);
      const at = (i) => (vals[i] === UNREAD ? (vals[i] = get(i)) : vals[i]);
      // a cell with nothing in it, told without computing anything
      const blankAt = (i) => { const cell = cellAt(i); return !cell || (!cell.ast && !cell.parseError && cell.value === null); };
      return new Proxy(vals, {
        get(t, k) {
          if (k === 'blankAt') return blankAt;
          if (typeof k === 'string' && /^\d+$/.test(k)) return +k < n ? at(+k) : undefined;
          return Reflect.get(t, k);
        },
      });
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
    // sc: ranges are intersected with the formula cell (an ordinary formula
    // whose inline arrays still compute element by element)
    // (a single cell reference is a scalar even in an array context, as
    // Calc's single reference: SUMPRODUCT(B4^2) with text in B4 is #VALUE!)
    aev(node, ctx, sc) {
      let u = node; while (u.t === 'paren') u = u.a; // ((B7) is the same single reference)
      const v = this.aev1(node, ctx, sc);
      // (a single reference a function passes on, IFERROR(A7,0), IF(1,A7), stays
      // one, as does the one cell INDEX or OFFSET gives)
      const single = v instanceof Range && u.t === 'fn' && (v.single || ((u.name === 'INDEX' || u.name === 'OFFSET') && v.rows === 1 && v.cols === 1));
      return this.aevPost(v, ctx, sc || (u.t === 'ref' && !u.b) || single);
    }
    // an error function of the formula has run (counted once: see countErrFuncs)
    // an error met while the formula still has an error function to run:
    // Calc's interpreter goes on with it set (see ev0); null otherwise
    contErr(e) {
      if (!(this.efDone < this.efTotal) || e === UNKNOWN_FN || this.inElem) return null;
      if (!this.gerr || this.gerr.clr !== this.clr) this.gerr = { err: asComputed(e), clr: this.clr };
      return e === E.LOOKNA || e.look ? demote(e) : tokOf(e);
    }
    // a conversion error once Calc's error is set gives that error
    convErr(e) { return this.gerr && this.gerr.clr === this.clr ? tokOf(this.gerr.err) : e; }
    efRan(node) {
      if (this.efSeen && !this.efSeen.has(node)) { this.efSeen.add(node); this.efDone++; }
    }
    aev1(node, ctx, sc) {
      const r = this.aev1x(node, ctx, sc);
      if (node.ef) this.efRan(node);
      return r;
    }
    aevPost(v, ctx, sc) {
      if (v instanceof Range) v = sc ? this.intersect(v, ctx) : this.rangeMatrix(v);
      // a scalar error met among array operations stays pending (Calc's
      // nGlobalError): it is the formula's result unless a reset follows
      return this.notePend(v);
    }
    // does a branch start with a scalar IF (one that jumps) or CHOOSE? (or
    // an IFERROR given an error: then that error)
    jumpFirst(n, ctx) {
      for (;;) {
        if (n.t === 'paren' || n.t === 'neg' || n.t === 'pct') n = n.a;
        else if (n.t === 'bin') n = n.a;
        else if (n.t === 'fn' && (n.name === 'IF' || n.name === 'CHOOSE')) {
          if (!n.args.length || n.args[0].t === 'missing') return false;
          const pend = this.pend, gerr = this.gerr;
          // (a condition that gives an array makes a nested jump matrix, no
          // quirk: an inline array or a range given directly; ranges inside
          // operators or functions are intersected or reduced)
          const c0 = n.args[0].t === 'paren' ? n.args[0].a : n.args[0];
          if (hasRange(c0) && (c0.t === 'ref' || c0.t === 'ref3d' || c0.t === 'rangeop' || c0.t === 'name')) return false;
          if (hasArr(n.args[0])) {
            const c = this.aev(n.args[0], ctx, true);
            this.pend = pend; this.gerr = gerr;
            if (c instanceof Matrix) return false;
          }
          // (an IF without an else branch whose condition is false does not
          // jump; an error the jump gives ends the formula)
          let r = true;
          if (n.name === 'IF' && n.args.length < 3) { const c = this.toBool(this.sv(n.args[0], ctx)); r = !isErr(c) && !!c; }
          if (r) { const v = this.sv(n, ctx); if (isErr(v)) r = { err: v }; }
          this.pend = pend; this.gerr = gerr;
          return r;
        } else if (n.t === 'fn' && n.name === 'IFERROR') {
          // (an IFERROR whose value is an error: that error ends the formula)
          if (!n.args.length || n.args[0].t === 'missing') return false;
          // (a range in its value outside any function: no quirk seen)
          if (n.args[0].t !== 'fn' && hasRange(n.args[0])) return false;
          const pend = this.pend, gerr = this.gerr;
          const v = this.sv(n.args[0], ctx);
          const f = isErr(v) && n.args.length > 1 && n.args[1].t !== 'missing' ? this.sv(n.args[1], ctx) : null;
          this.pend = pend; this.gerr = gerr;
          return isErr(v) ? { err: isErr(f) ? f : null, valErr: v } : false;
        } else if (n.t === 'fn' && !ERR_FUNCS.has(n.name) && n.args.length && n.args[0].t !== 'missing') n = n.args[0];
        else return false;
      }
    }
    // a computed error stays pending, or is a token if Calc goes on past it (contErr)
    // (the table's own error is never computed then: MATCH(1/0,NA()) is #DIV/0!)
    lookPend(v, pend0) { if (isErr(v) && !v.fromRef && !v.tok && !v.lit) this.pend = pend0; }
    notePend(v) {
      if (this.inElem && this.fa) return v; // (an element's error, in a force-array context)
      if (isErr(v) && !v.fromRef && !v.tok) { const t = this.contErr(v); if (t) return t; }
      if (isErr(v) && !v.fromRef && !this.pend) this.pend = { err: v, clr: this.clr, seq: ++this.errSeq };
      return v;
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
    aev1x(node, ctx, sc) {
      const lit = (x) => (x === MISSING ? { t: 'missing' } : { t: 'lit', v: x });
      // In an ordinary formula (sc) an array made by a function ("jm") reaches
      // another function's scalar parameter as its first element only, as in
      // Calc; operators still apply element by element.
      const one = (m) => (sc && m instanceof Matrix && m.jm ? m00(m) : m);
      const mark = (m) => { if (sc && m instanceof Matrix && !m.jm) { const e = m.epm; m = new Matrix(m.data); m.jm = true; if (e) m.epm = e; } return m; };
      // (NOT's array: numbers inside, but as a formula's result shown as a boolean)
      const markBool = (m) => { m = mark(m); if (m instanceof Matrix) m.boolRes = true; return m; };
      const jmOf = (r, ms) => { if (sc) r.jm = ms.filter((m) => m instanceof Matrix).every((m) => m.jm); return r; };
      switch (node.t) {
        case 'paren': {
          // (brackets do not consume an error: no pending error of their own)
          let v = this.aev1(node.a, ctx, sc);
          let u = node.a; while (u.t === 'paren') u = u.a; // (a single cell reference stays a scalar)
          if (v instanceof Range) v = sc || (u.t === 'ref' && !u.b) ? this.intersect(v, ctx) : this.rangeMatrix(v);
          return v;
        }
        case 'neg': case 'pct': {
          const a = this.aev(node.a, ctx, sc);
          if (!(a instanceof Matrix)) return this.ev0(node, ctx);
          return jmOf(this.amap([a], (x) => this.ev({ t: node.t, a: lit(x) }, ctx)), [a]);
        }
        case 'bin': {
          let a, b;
          const op0 = node.op;
          if (sc && (op0 === '=' || op0 === '<>' || op0 === '<' || op0 === '>' || op0 === '<=' || op0 === '>=')) {
            // a comparison with an array takes a range operand as an array
            const note = (v) => this.notePend(v);
            const ra = note(this.aev1(node.a, ctx, sc)), rb = note(this.aev1(node.b, ctx, sc));
            const anyMat = ra instanceof Matrix || rb instanceof Matrix;
            const fix = (v) => (v instanceof Range ? (anyMat && (v.rows > 1 || v.cols > 1) ? this.rangeMatrix(v) : this.intersect(v, ctx)) : v);
            a = fix(ra); b = fix(rb);
          } else { a = this.aev(node.a, ctx, sc); b = this.aev(node.b, ctx, sc); }
          if (!(a instanceof Matrix) && !(b instanceof Matrix)) return asComputed(this.binop({ op: node.op, a: lit(a), b: lit(b) }, ctx));
          // (text minus an array: Calc computes text's value minus each element,
          // so the text's conversion error wins over error elements: ""-{1,2}/0)
          if (node.op === '-' && typeof a === 'string' && b instanceof Matrix) {
            const x = this.toNum(a);
            if (isErr(x)) return jmOf(this.amap([b], () => x), [b]);
          }
          const op = node.op;
          if (op === '^' && !(a instanceof Matrix && b instanceof Matrix)) {
            // Calc's matrix-scalar power: text and error elements are NaN to
            // pow(), so x^0 and 1^y are 1 and anything else is #NUM!
            // (a pending error is the result; an error Calc went on past is a
            // NaN operand like the others: NA()^ISTEXT({1,2}) is {1,1})
            const pre = !node.fnPow && [a, b].find((v) => !(v instanceof Matrix) && isErr(v) && !v.fromRef);
            if (pre) return pre;
            const num = (v) => {
              if (v === null || v === undefined || v === EMPTYR || v === MISSING) return 0;
              const n = this.toNum(v);
              return isErr(n) ? NaN : n;
            };
            return jmOf(this.amap([a, b], (x, y) => {
              const bx = num(x), by = num(y);
              if (by === 0 || bx === 1) return 1;
              if (isNaN(bx) || isNaN(by)) return E.NUM;
              return asComputed(power(bx, by));
            }), [a, b]);
          }
          const cmpOp = op === '=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=';
          // comparing an array with an error gives that error, not an array
          if (cmpOp) { const e = [a, b].find((v) => !(v instanceof Matrix) && isErr(v)); if (e) return e.lit ? errByCode(e.code) : e; }
          // Calc's matrix division checks the divisor first: a zero divisor
          // is #DIV/0! whatever the dividend (an error, or text beside a
          // scalar divisor); an error to the power 0 and 1 to an error are 1
          const bothMat = a instanceof Matrix && b instanceof Matrix;
          const zero = (y) => y === 0 || y === false || y === null || y === EMPTYR || isLB(y);
          // (array and scalar with + - *: Calc computes element op scalar
          // (scalar - element for a scalar first), so the first operand's
          // error, a text's #VALUE! too, wins: SUMPRODUCT(B8:B12*A12) with
          // text in B8:B12 and #DIV/0! in A12 is #VALUE!)
          const arith = !bothMat && (op === '+' || op === '-' || op === '*');
          const convE = (v) => (isErr(v) ? v : typeof v === 'string' ? (isErr(this.toNum(v)) ? E.VALUE : null) : null);
          return jmOf(this.amap([a, b], (x, y) => {
            if (arith) {
              const [m, sv] = a instanceof Matrix ? [x, y] : [y, x];
              const em = convE(m), es = convE(sv);
              if (em && es) return asComputed(op === '-' && !(a instanceof Matrix) ? es : em);
            }
            if (op === '/' && zero(y) && (isErr(x) || (!bothMat && typeof x === 'string'))) return E.DIV0;
            if (op === '^' && bothMat && ((zero(y) && isErr(x)) || ((x === 1 || x === true) && isErr(y)))) return 1;
            return asComputed(this.binop({ op, a: lit(x), b: lit(y) }, ctx));
          }, cmpOp), [a, b]);
        }
        case 'fn': {
          const name = node.name;
          // (too many arguments: the call fails as a whole)
          if (FUNCS[name] && ((FUNCS[name].max !== undefined && node.args.length > FUNCS[name].max) || node.args.length < FUNCS[name].min)) return this.call(node, ctx);
          if (name === 'IF' || name === 'IFERROR' || name === 'CHOOSE') {
            // a scalar condition/value picks a whole branch; an array one picks
            // element by element (IF without an else branch leaves the element empty;
            // CHOOSE's array index is a jump matrix like IF's condition)
            const pend0 = this.pend;
            let first;
            if (name === 'IFERROR' && !sc && node.args[0].t !== 'missing') {
              // (in an array context IFERROR gives a reference without errors back as it is)
              const raw = this.aev1(node.args[0], ctx, sc);
              let u0 = node.args[0]; while (u0.t === 'paren') u0 = u0.a;
              // (a single cell reference stays a scalar: no array IFERROR)
              if (raw instanceof Range && u0.t === 'ref' && !u0.b) first = this.aevPost(raw, ctx, true);
              else if (raw instanceof Range) {
                const m = this.rangeMatrix(raw);
                if (!m.data.some((row) => row.some(isErr))) { this.clr++; return raw; } // (IFERROR ran: no pending error)
                first = this.aevPost(m, ctx, sc);
              } else first = this.aevPost(raw, ctx, sc);
            } else first = node.args[0].t === 'missing' ? MISSING : this.aev(node.args[0], ctx, sc);
            if (!(first instanceof Matrix)) { this.pend = pend0; return mark(this.ev0(node, ctx)); }
            // a branch no element takes leaves no pending error behind
            const used = (k) => first.data.some((row) => row.some((x) => {
              if (name === 'IFERROR') return k === 1 && isErr(x);
              if (isErr(x) || typeof x === 'string') return false;
              if (name === 'CHOOSE') { const n = this.toNum(x); return !isErr(n) && Math.trunc(n) === k; }
              const b = this.toBool(x);
              return !isErr(b) && (k === 1 ? b : !b);
            }));
            const args = [first];
            let skipK = null;
            // (Calc's quirk: a branch that some element takes and that starts
            // with a scalar IF or CHOOSE ends the whole formula: #VALUE!, or
            // #N/A when a value was already on the stack)
            // (not while Calc goes on past an earlier error)
            if (sc && !this.fa && !(this.gerr && this.gerr.clr === this.clr)) {
              // the branch each element takes, column by column (an array
              // IFERROR jumps only for its error elements; an error or text
              // element of a condition takes no branch)
              const seq = [];
              for (let c = 0; c < first.cols; c++) for (let r = 0; r < first.rows; r++) {
                const x = first.data[r][c];
                if (name === 'IFERROR') { if (isErr(x)) seq.push(1); continue; }
                if (isErr(x) || typeof x === 'string') { seq.push(-1); continue; }
                if (name === 'IF') { const t = this.toBool(x); seq.push(isErr(t) ? -1 : t ? 1 : 2); } else { const n = this.toNum(x); seq.push(isErr(n) ? -1 : Math.trunc(n)); }
              }
              const kLast = seq.length ? seq[seq.length - 1] : -1, kFirst = seq.length ? seq[0] : -1;
              const deep = stackDepth(this.root, node) > 0;
              skipK = new Set();
              for (let k = 1; k < node.args.length; k++) {
                const j = used(k) && this.jumpFirst(node.args[k], ctx);
                if (!j) continue;
                const err = j === true ? null : j.err;
                // (an IFERROR given a computed error, taken by two elements in a row:
                // its value's error; a cell's error does not)
                if (j.valErr && !(j.valErr.fromRef && !j.valErr.tok) && seq.some((b, i) => i > 0 && b === k && seq[i - 1] === k)) throw new JumpAbort(asComputed(j.valErr));
                if (!deep) {
                  // (the last element takes the branch, or is an error and takes
                  // none: the jump's error, else #VALUE!)
                  if (kLast === k || kLast < 0) throw new JumpAbort(err ? asComputed(err) : E.VALUE);
                  // (the first does: #N/A)
                  if (kFirst === k) throw new JumpAbort(E.NA);
                } else {
                  // (with a value on the stack: the jump's error if the last
                  // element takes the branch; else #VALUE! when an odd number
                  // of elements take it, #N/A when an even number do and the
                  // first is one of them)
                  if (err && kLast === k) throw new JumpAbort(asComputed(err));
                  if (seq.filter((b) => b === k).length % 2) throw new JumpAbort(E.VALUE);
                  if (kFirst === k) throw new JumpAbort(E.NA);
                }
                // (else the elements that take it are left as Calc's result
                // matrix starts: #N/A)
                skipK.add(k);
              }
            }
            let anyBool = false;
            for (let k = 1; k < node.args.length; k++) {
              const before = this.pend, gBefore = this.gerr;
              // with an array condition a range branch is taken as an array
              this.lastChildErr = false;
              if (skipK && skipK.has(k)) { args.push(E.NA); continue; }
              // (an empty branch is the number 0 element by element)
              let b = node.args[k].t === 'missing' ? 0 : this.aev1(node.args[k], ctx, sc);
              // (a computed boolean in a branch some element takes gives the
              // whole result the boolean format; a cell's boolean does not)
              if (name !== 'IFERROR' && used(k) && node.args[k].t !== 'ref' && !(b instanceof Range) &&
                  (typeof b === 'boolean' || (b instanceof Matrix && (b.boolRes || b.data.some((row) => row.some((x) => typeof x === 'boolean')))))) anyBool = true;
              if (b instanceof Range) b = this.rangeMatrix(b);
              // a branch's error is only the element's, unless a computed
              // error went through a further operation in the branch
              if (isErr(b) && !b.fromRef && !this.pend && this.lastChildErr) this.pend = { err: b, clr: this.clr };
              if (!(isErr(b) && this.lastChildErr)) this.gerr = gBefore;
              // a branch no element takes neither leaves a pending error nor sizes the result
              if (name !== 'IFERROR' ? !used(k) : k === 1 && !used(1)) { this.pend = before; this.gerr = gBefore; b = null; }
              // the condition's size holds where it is above 1; a branch only
              // extends the result along a dimension where the condition is 1
              if (b instanceof Matrix && ((first.rows > 1 && b.rows > first.rows) || (first.cols > 1 && b.cols > first.cols))) {
                const R = first.rows > 1 ? Math.min(b.rows, first.rows) : b.rows, C = first.cols > 1 ? Math.min(b.cols, first.cols) : b.cols;
                b = new Matrix(b.data.slice(0, R).map((row) => row.slice(0, C)));
              }
              args.push(b);
            }
            const noElse = name === 'IF' && node.args.length < 3;
            // as large as the largest argument; a position the chosen branch
            // lacks is #N/A
            const res = mark(epath(this.amap(args, (...xs) => {
              if (xs[0] === AMISS) return E.NA;
              // a text element of an array condition is an error (even numeric text)
              if (name !== 'IFERROR' && typeof xs[0] === 'string') return E.VALUE;
              let pick = 1;
              if (name === 'IF') { const t = this.toBool(xs[0]); pick = isErr(t) ? 0 : t ? 1 : 2; } else if (name === 'CHOOSE') {
                const n = this.toNum(xs[0]);
                pick = isErr(n) ? 0 : Math.trunc(n);
                if (!isErr(n) && !(pick >= 1 && pick < xs.length)) return E.IARG;
              } else pick = isErr(xs[0]) ? 1 : 0;
              if (pick && xs[pick] === AMISS) return E.NA;
              // (an IF with only a condition gives the number 1 where it holds)
              if (name === 'IF' && xs.length === 1 && pick === 1) return 1;
              xs = xs.map((x) => (x === AMISS ? null : x));
              let r = asComputed(this.call({ t: 'fn', name, args: xs.map(lit) }, ctx));
              // (IFERROR's array loses the boolean format)
              if (name === 'IFERROR' && typeof r === 'boolean') r = +r;
              return noElse && r === false && !this.toBool(xs[0]) ? EPATH : r;
            }, true, true)));
            if (anyBool && res instanceof Matrix) res.boolRes = true;
            // (a cell's boolean an element takes is a plain number then: IF({1,2},A5,-1) is 0)
            else if (res instanceof Matrix && res.data.some((row) => row.some((x) => typeof x === 'boolean'))) {
              const m = new Matrix(res.data.map((row) => row.map((x) => (typeof x === 'boolean' ? +x : x))));
              for (const k of Object.keys(res)) if (k !== 'data' && k !== 'rows' && k !== 'cols') m[k] = res[k];
              return m;
            }
            return res;
          }
          const mp = MAPPED_PARAMS[name];
          if (mp) {
            // a criterion / lookup value / index given as an array: one call per element
            const args = mp.map((i) => {
              if (i >= node.args.length || node.args[i].t === 'missing') return null;
              const v = this.aev(node.args[i], ctx, sc);
              return LOOKUPS.has(name) ? v : one(v);
            });
            if (!args.some((m) => m instanceof Matrix)) return mark(this.ev0(node, ctx));
            return mark(this.amap(args, (...xs) => {
              const a2 = node.args.slice();
              // (an error element of a lookup value is an operand, not a pending
              // error: checked after the column/row index, as a cell's error)
              mp.forEach((i, k) => { if (args[k] !== null) a2[i] = lit(LOOKUPS.has(name) && i === 0 ? asRefErr(xs[k]) : xs[k]); });
              const pos = this.amPos;
              const r = this.call({ t: 'fn', name, args: a2 }, ctx);
              if (r instanceof Range) {
                // a reference per element: its cell at the element's position
                const rr = r.rows === 1 ? 0 : pos[0], cc = r.cols === 1 ? 0 : pos[1];
                if (rr >= r.rows || cc >= r.cols) return E.NA;
                const x = this.cellValue(r.sheet, r.r1 + rr, r.c1 + cc);
                return asComputed(x === EMPTYR ? '' : x);
              }
              return asComputed(r instanceof Matrix ? r.data[0][0] : r);
            }));
          }
          if (!ELEMENTWISE.has(name)) return mark(this.ev0(node, ctx));
          // (Calc's add-in functions take their arguments as in an ordinary
          // formula even in an array context: SUMPRODUCT(QUOTIENT(B1:B3,2)) is B1/2)
          if (!sc && ADDIN.has(name)) {
            const fa = this.fa; this.fa = 0;
            try { return this.aev1x(node, ctx, true); } finally { this.fa = fa; }
          }
          if (name === 'POWER' && node.args.length === 2 && node.args.every((n) => n.t !== 'missing')) {
            // POWER is Calc's ^ operator: over arrays it follows ^
            const r = this.aev1({ t: 'bin', op: '^', a: node.args[0], b: node.args[1], fnPow: true }, ctx, sc);
            if (r instanceof Matrix) return mark(r);
            return mark(this.ev0(node, ctx));
          }
          // when another argument is an array, a range argument is taken as
          // an array too (not intersected), except for POWER (Calc's ^)
          const raws = node.args.map((n) => (n.t === 'missing' ? MISSING : this.aev1(n, ctx, sc)));
          // (NOT takes a function-made array whole: Calc's NOT handles arrays itself)
          // (beside an inline array, a function-made array is taken whole too)
          const inl = raws.some((v) => v instanceof Matrix && !v.jm);
          // (the first element taken from a function-made array is an element:
          // its error is an operand like a cell's, not a pending error)
          let collapsed = false;
          const one1 = (v) => {
            if (name === 'NOT' || inl) return v;
            let o = one(v);
            // (the IS functions see an empty-path element as empty text)
            if (o === EPATH && /^IS/.test(name)) o = '';
            if (o !== v) { collapsed = true; return asRefErr(o); }
            return o;
          };
          const anyMat = raws.some((v) => one1(v) instanceof Matrix);
          const args = raws.map((v) => {
            if (v instanceof Range) v = (!sc || (anyMat && name !== 'POWER')) && (v.rows > 1 || v.cols > 1) ? this.rangeMatrix(v) : this.intersect(v, ctx);
            else v = one1(v);
            return this.notePend(v);
          });
          if (!args.some((m) => m instanceof Matrix)) {
            if (collapsed) return asComputed(this.call({ t: 'fn', name, args: args.map(lit) }, ctx));
            return mark(this.ev0(node, ctx));
          }
          if (name === 'NOT' && args.length === 1) {
            // Calc's matrix NOT: an error or text that is not a number gives 0
            return markBool(this.amap(args, (x) => (isErr(x) || (typeof x === 'string' && textToNum(x) === null) ? 0 : asComputed(this.call({ t: 'fn', name, args: [lit(x)] }, ctx)))));
          }
          // as large as the largest argument, #VALUE! where one lacks the position
          // (Calc's MOD pops the divisor first: 0 is #DIV/0! whatever the dividend)
          const div0 = name === 'MOD' && args.length === 2;
          return mark(this.amap(args, (...xs) => {
            if (div0 && isErr(xs[0]) && (xs[1] === 0 || xs[1] === false || xs[1] === null || isLB(xs[1]))) return E.DIV0;
            return asComputed(this.call({ t: 'fn', name, args: xs.map(lit) }, ctx));
          }, name !== 'POWER'));
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
      // (force array: functions inside take array arguments too,
      // MATCH(1,INDEX((A1:A3=2)*(B1:B3=20),0),0))
      this.fa = (this.fa || 0) + 1;
      // (a function that gives a reference keeps it: INDEX(B1:B9,0))
      try { const v = this.aev1(n, ctx); return v instanceof Range ? v : this.aevPost(v, ctx); } finally { this.fa--; }
    }
    // element-wise map; a single row/column is repeated, otherwise the
    // result is as large as the smallest operand
    // (comparisons, big: as large as the largest operand, #VALUE! where an
    // operand has no element)
    // (pass: a missing element reaches f as AMISS instead of making #VALUE!)
    amap(args, f, big, pass) {
      let R = big ? 1 : Infinity, C = big ? 1 : Infinity;
      const pick = big ? Math.max : Math.min;
      for (const m of args) if (m instanceof Matrix) { if (m.rows > 1) R = pick(R, m.rows); if (m.cols > 1) C = pick(C, m.cols); }
      if (R === Infinity) R = 1;
      if (C === Infinity) C = 1;
      const OUT = {};
      const at = (m, r, c) => {
        if (!(m instanceof Matrix)) return m;
        const rr = m.rows === 1 ? 0 : r, cc = m.cols === 1 ? 0 : c;
        if (!(rr < m.rows && cc < m.cols)) return big ? OUT : E.NA;
        return m.epm && m.epm[rr][cc] ? EPATH : m.data[rr][cc];
      };
      const out = [];
      for (let r = 0; r < R; r++) {
        const row = [];
        for (let c = 0; c < C; c++) {
          const xs = args.map((m) => at(m, r, c));
          if (big && xs.includes(OUT)) {
            if (!pass) { row.push(E.VALUE); continue; }
            for (let k = 0; k < xs.length; k++) if (xs[k] === OUT) xs[k] = AMISS;
          }
          this.amPos = [r, c];
          // (an element's error is the element's: not Calc's error, see contErr)
          this.inElem = (this.inElem || 0) + 1;
          try { row.push(f(...xs)); } finally { this.inElem--; }
        }
        out.push(row);
      }
      return new Matrix(out);
    }

    resolveSheet(name, ctx) {
      if (name === null || name === undefined) return ctx.sheet;
      return this.wb.sheet(name);
    }

    // a formula part holding an inline array ({1,2}*2, LEN({"a","bb"})) is
    // computed element by element even outside array functions, as in Calc
    ev(node, ctx) {
      if (this.fa && (node.t === 'bin' || node.t === 'neg' || node.t === 'pct' || node.t === 'fn')) {
        const v = this.aev1(node, ctx, false);
        // (inside an element's computation an error is the element's, not pending)
        return v instanceof Range || this.inElem ? v : this.aevPost(v, ctx, false);
      }
      if ((node.t === 'bin' || node.t === 'neg' || node.t === 'pct' || node.t === 'fn') && hasArr(node)) {
        // (a function that gives a reference, IF(1,A4,{1,2}), still gives one)
        if (node.t === 'fn') { const v = this.aev1(node, ctx, true); if (v instanceof Range) return v; return this.aevPost(v, ctx, true); }
        return this.aev(node, ctx, true);
      }
      return this.ev0(node, ctx);
    }
    // errSeen: a computed error came out of a part of the formula evaluated
    // so far (within the current node); lastChildErr: the node evaluated
    // last had such an error among its parts (it consumed a computed error)
    ev0(node, ctx) {
      const t = node.t;
      if (t !== 'fn' && t !== 'bin' && t !== 'neg' && t !== 'pct') {
        const r = this.ev0x(node, ctx);
        if (isErr(r) && !r.fromRef) this.errSeen = true;
        return r;
      }
      const s = this.errSeen;
      this.errSeen = false;
      const r = this.ev0x(node, ctx);
      if (isErr(r) && !r.fromRef && !r.tok && this.running) {
        const top = this.running[this.running.length - 1];
        if (top && !top.sup && (!top.fe || top.fe.clr !== this.clr)) top.fe = { err: r, clr: this.clr };
        // (the first scalar error, not an array element's, for the pending error's order)
        if (top && !this.inElem && (!top.se || top.se.clr !== this.clr)) top.se = { err: r, clr: this.clr, seq: ++this.errSeq };
      }
      this.lastChildErr = this.errSeen;
      this.errSeen = s || (isErr(r) && !r.fromRef);
      // Calc's interpreter stops at an error unless the formula still has an
      // error function (COUNT, IS..., N, IFERROR) to run; going on, the error
      // is an ordinary token on the stack (MOD(SQRT(-1),0)+COUNT(1) is #DIV/0!)
      if (isErr(r) && !r.fromRef && !r.tok) { const t = this.contErr(r); if (t) return t; }
      return r;
    }
    ev0x(node, ctx) {
      switch (node.t) {
        case 'num': return node.v;
        case 'str': return node.v;
        case 'bool': return node.v;
        case 'errlit': return litErr(node.v.code); // a plain error token, not a pending error
        case 'arr': return node.v;
        case 'missing': return MISSING;
        case 'lit': return node.v;
        case 'paren': return this.ev(node.a, ctx);
        case 'at': {
          const v = this.ev(node.a, ctx);
          if (v instanceof Range) {
            // the cell in the formula's row/column; a range it cannot be
            // intersected with is passed on whole (SUM(@A1:A3) below row 3)
            if (v.rows === 1 && v.cols === 1) return v;
            if (v.cols === 1 && ctx && ctx.r >= v.r1 && ctx.r <= v.r2) return new Range(v.sheet, ctx.r, v.c1, ctx.r, v.c1);
            if (v.rows === 1 && ctx && ctx.c >= v.c1 && ctx.c <= v.c2) return new Range(v.sheet, v.r1, ctx.c, v.r1, ctx.c);
            return v;
          }
          if (v instanceof Matrix) return v.data[0][0];
          if (v instanceof RefList) return E.VALUE;
          return v;
        }
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
          const rg = new Range(sh, node.a.r, node.a.c, b.r, b.c);
          if (!node.b) rg.single = true; // (Calc's single reference: a scalar in an array context too)
          // (a range written bottom-up or right-to-left keeps its corners as
          // written after LibreOffice's xlsx import: intersecting it is empty)
          if (node.b && (node.a.r > b.r || node.a.c > b.c)) { rg.rev = true; rg.wr = [node.a.r, node.a.c, b.r, b.c]; }
          return rg;
        }
        case 'ref3d': {
          const s1 = this.wb.sheet(node.sheet), s2 = this.wb.sheet(node.sheet2);
          if (!s1 || !s2) return E.NAME;
          let i1 = this.wb.sheets.indexOf(s1), i2 = this.wb.sheets.indexOf(s2);
          if (i1 > i2) [i1, i2] = [i2, i1];
          const b = node.b || node.a;
          const ranges = [];
          for (let k = i1; k <= i2; k++) ranges.push(new Range(this.wb.sheets[k], node.a.r, node.a.c, b.r, b.c));
          return ranges.length === 1 ? ranges[0] : refList3d(ranges);
        }
        case 'isect': {
          const a = this.ev(node.a, ctx), b = this.ev(node.b, ctx);
          if (isErr(a)) return a;
          if (isErr(b)) return b;
          if (!(a instanceof Range) || !(b instanceof Range)) return E.VALUE;
          if (a.sheet !== b.sheet || a.rev || b.rev) return E.NULL;
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
          if (a instanceof Range && b instanceof Range && a.sheet === b.sheet) {
            return new Range(a.sheet, Math.min(a.r1, b.r1), Math.min(a.c1, b.c1), Math.max(a.r2, b.r2), Math.max(a.c2, b.c2));
          }
          // across sheets (Sheet1!A1:Sheet2!B2): a 3D reference over the
          // sheets between, in workbook order
          const parts = [];
          for (const x of [a, b]) {
            if (x instanceof Range) parts.push(x);
            else if (x instanceof RefList && x.is3d && node.b.t === 'ref' && node.b.sheet !== null) parts.push(...x.ranges);
            else return E.VALUE;
          }
          const idx = parts.map((x) => this.wb.sheets.indexOf(x.sheet));
          const i1 = Math.min(...idx), i2 = Math.max(...idx);
          const r1 = Math.min(...parts.map((x) => x.r1)), r2 = Math.max(...parts.map((x) => x.r2));
          const c1 = Math.min(...parts.map((x) => x.c1)), c2 = Math.max(...parts.map((x) => x.c2));
          const ranges = [];
          for (let k = i1; k <= i2; k++) ranges.push(new Range(this.wb.sheets[k], r1, c1, r2, c2));
          return refList3d(ranges);
        }
        case 'neg': {
          let depth = 1, inner = node.a;
          while (inner.t === 'neg' || inner.t === 'paren') { if (inner.t === 'neg') depth++; inner = inner.a; }
          this.boolLost = false;
          const raw = this.ev(inner, ctx);
          const lost = this.boolLost;
          const v = this.scalar(raw, ctx);
          const n = this.toNum(v);
          if (isErr(n)) return asComputed(n);
          const r = depth % 2 ? -n : n;
          // Calc keeps the boolean number format through negation
          // (but not for a cell reference, also one a function returns, nor
          // a value IFERROR passes on or INDEX takes from an inline array)
          if (typeof v === 'boolean' && !lost && inner.t !== 'ref' && !(raw instanceof Range) && (r === 0 || r === 1)) return r === 1;
          return r === 0 ? 0 : r;
        }
        case 'pct': {
          const v = this.sv(node.a, ctx);
          const n = this.toNum(v);
          if (isErr(n)) return asComputed(n);
          return n / 100;
        }
        case 'bin': { const r = asComputed(this.binop(node, ctx)); this.boolLost = false; return r; }
        case 'fn': {
          const r = this.call(node, ctx);
          // IF and CHOOSE pass on what their chosen argument left
          if (!PASS_FMT.has(node.name)) this.boolLost = this.retLost;
          this.retLost = false;
          // IF, CHOOSE and IFERROR leave an error constant on the stack as it is
          if (isErr(r) && r.lit && (PASS_FMT.has(node.name) || node.name === 'IFERROR')) return r;
          return r === UNKNOWN_FN ? r : asComputed(r); }
      }
      return E.VALUE;
    }

    // ---- coercions
    toNum(v) {
      if (typeof v === 'number') return v;
      if (typeof v === 'boolean') return v ? 1 : 0;
      if (v === null || v === MISSING || v === EMPTYR || isLB(v) || v === EPATH || isMV(v)) return 0;
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
      if (v === null || v === MISSING || isLB(v) || isMV(v)) return '';
      if (v === EMPTYR) return '0';
      if (v === EPATH) return 'FALSE';
      if (isErr(v)) return v;
      return E.VALUE;
    }
    toBool(v) {
      if (typeof v === 'boolean') return v;
      if (typeof v === 'number') return v !== 0;
      if (v === null || v === MISSING || v === EMPTYR || isLB(v) || v === EPATH || isMV(v)) return false;
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
      // a reference operand is read when the operator runs, after both
      // operands are computed, the right one first (as Calc pops them): this
      // only matters for when a cell inside a reference cycle is computed
      const ra = this.ev(node.a, ctx);
      const clr = this.clr;
      const rb = this.ev(node.b, ctx);
      const b = this.scalar(rb, ctx);
      let a = this.scalar(ra, ctx);
      // a function that clears Calc's pending error (IFERROR, ISNUMBER, ...)
      // ran after the left operand: its error is then an ordinary operand
      if (this.clr !== clr) a = demote(a);
      // errors produced by computation (left to right) win over
      // reference/conversion errors, which are checked right operand first.
      if (isErr(a) && !a.fromRef) return a;
      if (isErr(b) && !b.fromRef) return b;
      // a left error token overwrites whatever error the right operand set
      if (isErr(a) && a.tok) return a;
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
      const y = this.toNum(b);
      if (isErr(y)) return isErr(a) && a.tok ? a : y;
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
      // too few arguments: Err:511; too many: Err:508 for one-parameter functions, else Err:504
      const few = n < f.min;
      if (few && (!n || node.name === 'IF' || node.name === 'IFERROR')) return E.NOARG;
      if (few || (f.max !== undefined && n > f.max)) {
        if (!few && f.max === 1) return E.PAREN; // (compile errors)
        if (!few && (node.name === 'IF' || node.name === 'IFERROR')) return E.PARAM;
        // Err:504 when the function runs: its arguments are computed
        // first, and a computed error among them stops the formula
        // (references are not read)
        for (const a of node.args) {
          if (a.t === 'missing' || a.t === 'ref' || a.t === 'ref3d') continue;
          const v = this.ev(a, ctx);
          if (isErr(v) && !v.fromRef && !v.tok && v !== UNKNOWN_FN) return v;
        }
        return few ? E.NOARG : E.PARAM;
      }
      // a function that resets the pending error (IFERROR's value, IS...,
      // COUNT, N) computes its arguments with no pending error (seen by a
      // cycle read): errors there are values on Calc's stack
      const top = CLEARS_ERROR.has(node.name) && this.running && this.running[this.running.length - 1];
      if (top) { top.fe = null; top.sup = (top.sup || 0) + 1; }
      let r;
      try { r = f.fn.call(this, node.args, ctx); } finally { if (top) top.sup--; }
      // (as Calc's PushDouble: a result that is not finite is #NUM!)
      if (typeof r === 'number' && !isFinite(r)) r = E.NUM;
      if (CLEARS_ERROR.has(node.name)) {
        this.clr++;
        if (node.ef && node.name !== 'IFERROR') this.efRan(node);
      }
      return r;
    }

    // Evaluate fixed scalar arguments with Calc's error order:
    // computed errors left to right, then conversion right to left.
    // kinds: 'n' number, 's' string, 'b' bool, 'v' any scalar
    // checks[i](value): a function's own validity check of parameter i, made
    // as Calc pops it (last parameter first), before earlier ones are converted
    args(nodes, ctx, kinds, defaults, checks) {
      const vals = [], clrs = [];
      for (let i = 0; i < kinds.length; i++) {
        if (i >= nodes.length) { vals.push(undefined); continue; }
        vals.push(this.sv(nodes[i], ctx));
        clrs.push(this.clr);
      }
      for (let i = 0; i < clrs.length; i++) if (clrs[i] !== this.clr) vals[i] = demote(vals[i]);
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
        if (isErr(r)) return tokBefore(vals, i, r);
        if (checks && checks[i] && r !== undefined) { const e = checks[i](r); if (e) return checks.soft ? tokBefore(vals, i, e) : e; }
        out[i] = r;
      }
      return out;
    }

    // Collect numbers for aggregate functions.
    // opts.countText: numeric strings given directly count (COUNT); otherwise #VALUE!
    // opts.skipErrors: errors are ignored (COUNT)
    collect(nodes, ctx, opts) {
      const nums = [];
      const vals = [], fmtVals = []; // (fmtVals: in parameter order, sheets first to last)
      const evs = [];
      for (const node of nodes) {
        if (node.t === 'missing') { evs.push([0, this.clr]); continue; } // an empty argument is 0
        evs.push([this.ev(node, ctx), this.clr]);
      }
      // the first computed error wins, unless a later argument cleared it
      for (let [v, clr] of evs) {
        if (clr !== this.clr) v = demote(v);
        if (isErr(v) && !v.fromRef && !opts.skipErrors) return { err: v };
        // (a 3D reference is one parameter: its sheets are taken first to last)
        if (v instanceof RefList) { vals.push(...(v.is3d && !opts.forward ? v.ranges.slice().reverse() : v.ranges)); fmtVals.push(...v.ranges); } else { vals.push(v); fmtVals.push(v); }
      }
      // Calc pops parameters from the stack: the last one is examined first
      // (MEDIAN gathers its parameters into an array first: first one first)
      // for SUM's rounding: Calc's pop order, each item a number or the
      // numbers of one column of a range
      const parts = [];
      // error tokens popped later overwrite the error (the first one wins)
      const tokOver = (k, e) => {
        if (opts.noTok) return e;
        if (opts.forward) { for (let j = vals.length - 1; j > k; j--) if (isErr(vals[j]) && vals[j].tok) return vals[j]; return e; }
        for (let j = 0; j < k; j++) if (isErr(vals[j]) && vals[j].tok) return vals[j];
        return e;
      };
      for (let j = 0; j < vals.length; j++) {
        const k = opts.forward ? j : vals.length - 1 - j;
        const v = vals[k];
        const n0 = nums.length;
        if (v instanceof Matrix) parts.push(nums); // placeholder, fixed below
        if (v instanceof Range) {
          let err = null, col = -1, cur = null;
          // opts.colLastErr (SUM): Calc sums a column block by block (a run of
          // formula cells, booleans and error constants being one block); the
          // first error of a block stops it, a later block's error replaces
          // it, and the column's error ends the scan.
          // (cells a block's error skips are not computed: no false cycles)
          let prevR = -2, prevF = false, blockErr = false, pcol = -1, stop = false;
          // opts.colCont (AVERAGE, PRODUCT): an error ends its column, the
          // later columns are still computed; the first error is the result
          let colErr = false, ccol = -1, lastCol = false;
          const pre = opts.colCont ? (r, c) => {
            if (c !== ccol) { ccol = c; colErr = false; if (lastCol) stop = true; }
            return colErr || stop;
          } : opts.colLastErr ? (r, c) => {
            if (c !== pcol) { pcol = c; prevR = -2; if (err) stop = true; }
            if (stop) return true;
            const cell = v.sheet.cells.get(r * MAXCOL + c);
            const f = !!cell && (cell.formula || typeof cell.raw === 'boolean' || (cell.raw !== null && typeof cell.raw === 'object'));
            if (!(f && prevF && prevR === r - 1)) blockErr = false;
            prevF = f; prevR = r;
            return blockErr;
          } : null;
          // (SUM: Calc sums a run of number cells as one block, sumArray, and
          // adds formula cells one by one; blank and text cells end a run)
          let blk = null, blkR = -2;
          const rid = {};
          this.eachCell(v, (x, r, c) => {
            if (c !== col) { if (err && !opts.colCont) return false; col = c; cur = []; cur.segs = []; cur.rid = rid; parts.push(cur); blk = null; }
            if (typeof x === 'number' || typeof x === 'boolean') {
              const n = typeof x === 'boolean' ? (x ? 1 : 0) : x;
              nums.push(n); cur.push(n);
              const cell = v.sheet.cells.get(r * MAXCOL + c);
              if (cell && (cell.formula || typeof cell.raw === 'boolean')) { cur.segs.push(n); blk = null; }
              else { if (!blk || blkR !== r - 1) { blk = []; cur.segs.push(blk); } blk.push(n); blkR = r; }
            } else blk = null;
            if (typeof x === 'number' || typeof x === 'boolean') {}
            else if (isErr(x) && !opts.skipErrors) {
              // (a cell being computed, Err:522, does not end its column, but
              // no later column is computed)
              if (opts.colCont) { if (!err) err = x; if (x.code !== 'Err:522') colErr = true; else lastCol = true; return true; }
              err = x; blockErr = true; return !!opts.colLastErr;
            }
            return true;
          }, false, pre);
          // a range's error only marks the result: an error token among the
          // parameters popped later replaces it (SUM(#N/A,A12) is #N/A
          // whatever A12's error), except that a larger range's error ends
          // SUM's scan
          if (err) return { err: (v.rows === 1 && v.cols === 1) || !opts.colLastErr ? tokOver(k, err) : err };
        } else if (v instanceof Matrix) {
          // (MIN/MAX: an array holding an error, other than the first
          // parameter, is passed over whole, as Calc's matrix minimum is NaN)
          if (opts.arrErrSkip && k > 0 && v.data.some((row) => row.some(isErr))) continue;
          // (and an array of empty elements only counts as 0)
          // (an array of only text or empty elements, too)
          if (opts.arrErrSkip && v.data.every((row) => row.every((x) => x === null || typeof x === 'string'))) { nums.push(0); continue; }
          if (opts.kahanErr) {
            // (SUM/AVERAGE: Calc's Kahan sum goes column by column and adds
            // the held-back last non-zero term last, so an error there wins
            // over the first error met)
            let first = null, last = null;
            for (let c = 0; c < v.cols; c++) for (let r = 0; r < v.rows; r++) {
              const x = v.data[r][c];
              if (isErr(x)) { if (!first) first = x; last = x; } else if ((typeof x === 'number' && x !== 0) || x === true) last = null;
            }
            if (first) return { err: last || first };
          }
          for (const row of v.data) for (const x of row) {
            if (typeof x === 'number') { nums.push(x); } else if (typeof x === 'boolean') nums.push(x ? 1 : 0);
            else if (isErr(x) && !opts.skipErrors) return { err: x };
          }
        } else if (typeof v === 'number') { nums.push(v); } else if (typeof v === 'boolean') { nums.push(v ? 1 : 0); } else if (v === EMPTYR || isLB(v) || isMV(v)) { nums.push(0); } else if (isErr(v)) { if (!opts.skipErrors) return { err: tokOver(k, v) }; } else if (typeof v === 'string') {
          const n = textToNum(v);
          if (opts.countText) { if (n !== null) { nums.push(n); } } else return { err: opts.forward ? tokOver(k, opts.textParam ? E.PARAM : E.VALUE) : tokBefore(opts.noTok ? [] : vals, k, opts.textParam ? E.PARAM : E.VALUE) };
        }
        if (parts[parts.length - 1] === nums) parts[parts.length - 1] = nums.slice(n0);
        else if (nums.length > n0 && !(v instanceof Range)) parts.push(nums[n0]);
      }
      if (!opts.forward) nums.reverse();
      // MIN/MAX take the number format of the first number met in parameter
      // order (text and blank cells skipped): a boolean cell makes the result
      // a boolean, a computed or constant value a number.
      let boolFmt = false;
      for (const v of fmtVals) {
        if (!(v instanceof Range)) {
          if (typeof v === 'number' || typeof v === 'boolean' || v instanceof Matrix) break;
          continue;
        }
        let found = false;
        this.eachCell(v, (x) => {
          if (typeof x === 'number' || typeof x === 'boolean') { boolFmt = typeof x === 'boolean'; found = true; return false; }
          return true;
        });
        if (found) break;
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
    // the C library reports underflow (a subnormal or zero result) as a range error
    if (x !== 0 && Math.abs(r) < 2.2250738585072014e-308) return E.NUM;
    return r;
  }

  // compare two non-error scalars as Calc's comparison operators do
  function compare(a, b) {
    if (a === EMPTYR || isLB(a)) a = null;
    if (b === EMPTYR || isLB(b)) b = null;
    if (a === EPATH || isMV(a)) a = 0;
    if (b === EPATH || isMV(b)) b = 0;
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
    if (isErr(v)) return v.code === 'Err:LOOKNA' ? '#N/A' : v.code;
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
      // text is compared with a cell's displayed text (numbers too); an
      // inline array's numbers never equal text
      if (typeof x === 'number') { if (exact) return false; x = fmtNum(x); }
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
    const conv = (x) => (x === null ? null : isErr(x) ? ERRC : x === EMPTYR ? '' : typeof x === 'boolean' ? +x : x);
    // (cells are read as the search reaches them; blanks are told apart
    // without reading)
    const memo = new Map();
    const list = new Proxy([], {
      get(t, k) {
        if (typeof k === 'string' && /^\d+$/.test(k)) {
          if (!memo.has(k)) memo.set(k, +k < list0.length ? conv(list0[k]) : undefined);
          return memo.get(k);
        }
        return k === 'length' ? list0.length : Reflect.get(t, k);
      },
    });
    const isNull = (i) => (list0.blankAt ? list0.blankAt(i) : list[i] === null);
    const n = list.length, byStr = typeof q === 'string';
    const isStr = (x) => typeof x === 'string';
    const cmp = (c) => (byStr ? strCmp(c, q) : c === ERRC ? (approxEqual(0, q) ? 0 : 0 < q ? -1 : 1) : (approxEqual(c, q) ? 0 : c < q ? -1 : 1));
    const firstStringIgnore = byStr, allStringIgnore = !byStr;
    let row = 0, skipped = false;
    if (firstStringIgnore && isStr(list[0])) {
      const t = cmp(list[0]);
      if ((asc && t > 0) || (!asc && t < 0)) { row++; skipped = true; }
    }
    while (row < n && isNull(row)) row++;
    if (allStringIgnore) while (row < n && (list[row] === null || isStr(list[row]))) row++;
    let start = row;
    if (binary && row < n) {
      const idx = [];
      for (let i = row; i < n; i++) if (!isNull(i)) idx.push(i);
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
        else { found = mid; lo = mid + 1; } // an exact hit: Calc goes on searching after it
      }
      start = broken ? idx[0] : found !== null ? idx[found] : idx[last];
    }
    let foundRow = -1, fsi = firstStringIgnore && !skipped, eq = false;
    // (after the binary search (a column) or along a row, an ascending
    // text search walks on past numbers, but a
    // number undoes the match found so far until the next one, and a first
    // text passed over after numbers is no match: MATCH("x",A1:B1) is 2 with
    // -3,"b" and #N/A with "b",-3; MATCH("x",A1:C1) with "a",1,"b" is 3)
    const passNums = byStr && asc;
    let leadNums = false;
    for (let r = start; r < n; r++) {
      const c = list[r];
      if (c === null || (allStringIgnore && isStr(c))) continue;
      if (passNums && typeof c === 'number') { if (foundRow < 0) leadNums = true; foundRow = -1; continue; }
      if (isStr(c) === byStr && c !== ERRC) {
        const t = cmp(c);
        if (asc ? t <= 0 : t >= 0) { foundRow = r; fsi = false; if (t === 0) { eq = true; break; } continue; }
      }
      // (numbers before a first text passed over: no match)
      if (fsi && isStr(c) && passNums && leadNums) { foundRow = -1; break; }
      if (fsi && isStr(c)) { fsi = false; continue; }
      if (fsi && isStr(c)) { fsi = false; continue; }
      break;
    }
    if (eq) {
      for (let k = foundRow + 1; k < n; k++) {
        const c = list[k];
        if (c === null) continue;
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
  // numsBelowText (MATCH over an inline array): in an ascending search numbers
  // are below a text value
  // Calc's MATCH over an array (ScMatch, lcl_CompareMatrix2Query): a binary
  // search where an empty element is below anything, an error above
  // anything, numbers below text; an exact hit goes on over equal elements
  function matBinary(list, v, asc) {
    const byStr = typeof v === 'string';
    const empty = (x) => x === null || x === undefined || x === EMPTYR || isMV(x);
    const cmp = (i) => {
      let x = list[i];
      if (empty(x)) return -1;
      // (an error is a value: below a text query, above a numeric one)
      if (isErr(x)) return 1;
      if (typeof x === 'boolean') x = +x;
      if (typeof x === 'number') {
        if (byStr) return -1;
        return x === v ? 0 : x < v ? -1 : 1;
      }
      if (!byStr) return 1;
      const c = strCmp(x, v);
      return c < 0 ? -1 : c > 0 ? 1 : 0;
    };
    const n = list.length;
    if (!n) return -1;
    let first = 0, last = n - 1, hit = 0;
    for (let len = last - first; len > 0; len = last - first) {
      let mid = first + (len >> 1);
      const c = cmp(mid);
      if (c === 0) {
        const x = list[mid];
        const same = (y) => (empty(x) ? empty(y) : typeof x === 'string' ? y === x : !isErr(y) && typeof y !== 'string' && !empty(y) && +y === +x);
        while (mid < n - 1 && same(list[mid + 1])) mid++;
        return mid;
      }
      if (len === 1) { hit = c < 0 ? (asc ? last : first) : (asc ? first : last); break; }
      if (c < 0) { if (asc) first = mid; else last = mid; } else if (asc) last = mid; else first = mid;
    }
    if (hit === n - 1) {
      const c = cmp(hit);
      if ((asc && c <= 0) || (!asc && c >= 0)) return hit;
    }
    if (hit <= 0) return -1;
    // (a hit before the last element must be of the query's type: text for a
    // text query, not text for a number)
    const x = list[hit - 1];
    return (typeof x === 'string') === byStr ? hit - 1 : -1;
  }
  // Calc's sorted VLOOKUP/HLOOKUP over an array: a walk that keeps the last
  // element not above the query and stops at the first one above it (for a
  // text query one above at the start is passed over); numbers are below
  // any text (errors too), but a hit that is not text is no match; a numeric
  // query passes over text, and an error element it meets is the result
  function arrLookup(list, v) {
    if (typeof v === 'boolean') v = +v;
    const byStr = typeof v === 'string';
    let hit = -1, hitStr = false;
    for (let i = 0; i < list.length; i++) {
      let x = list[i];
      if (isErr(x) && !byStr) return x;
      if (x === null || x === undefined || x === EMPTYR || isMV(x)) x = byStr ? '' : null;
      let c;
      if (byStr) c = typeof x === 'string' ? strCmp(x, v) : -1;
      else if (typeof x === 'string' || x === null) continue;
      else c = +x === v ? 0 : +x < v ? -1 : 1;
      if (c <= 0) { hit = i; hitStr = typeof x === 'string'; } else if (i > 0 || !byStr) break;
    }
    if (byStr && hit >= 0 && !hitStr) return -1;
    return hit;
  }
  // a descending search with a pattern no cell matches: a walk comparing it
  // as plain text, from the first cell not below it while cells stay so
  // (text below it at the start is passed over, anything else ends the walk)
  function wildDescWalk(list, v) {
    let found = -1;
    for (let i = 0; i < list.length; i++) {
      let x = list[i];
      if (x === null || x === undefined) continue;
      if (x === EMPTYR) x = '';
      if (typeof x === 'string' && strCmp(x, v) >= 0) found = i;
      else if (found >= 0 || typeof x !== 'string') break; // (a cell not text before any: no match)
    }
    return found;
  }
  function lookupPos(list, v, mode, ranges, numsBelowText) {
    if (typeof v === 'boolean') v = v ? 1 : 0;
    // an ascending wildcard search over cells: the first cell the pattern
    // matches (and the matching cells right after it); without one, a sorted
    // search comparing the pattern as plain text
    if (mode !== 0 && typeof v === 'string' && hasWild(v) && (ranges === 'col' || ranges === 'row')) {
      // (numbers and booleans before the first text or error cell are passed
      // over; an error cell matches as its code's text, except as the first
      // cell, and an ascending search starts on one only after a number)
      const t0 = list.findIndex((x) => typeof x === 'string' || x === EMPTYR || isErr(x));
      const re = wildcardRe(v);
      const wm = (x, k) => (isErr(x) ? k > 0 && re.test(x.code) : lookupEq(v, x, false));
      const prevNum = (k) => { while (--k >= 0) { const y = list[k]; if (y !== null && y !== undefined) return typeof y === 'number' || typeof y === 'boolean'; } return false; };
      let i = t0 < 0 ? -1 : list.findIndex((x, k) => k >= t0 && wm(x, k) && (!isErr(x) || mode < 0 || prevNum(k)));
      if (i < 0) return mode > 0 ? sortedLookup(list, v, true, ranges === 'col') : wildDescWalk(list, v);
      // (the run of matching cells goes on over blank cells)
      for (let j = i + 1; j < list.length; j++) { const x = list[j]; if (x === null || x === undefined) continue; if (wm(x, j)) i = j; else break; }
      return i;
    }
    // (inline arrays compare a pattern as plain text in a sorted search)
    if (mode === 0 || (typeof v === 'string' && hasWild(v) && ranges)) {
      for (let i = 0; i < list.length; i++) if (lookupEq(v, list[i], !ranges)) return i;
      return -1;
    }
    if (ranges === 'col' || ranges === 'row') return sortedLookup(list, v, mode > 0, ranges === 'col');
    if (numsBelowText) return matBinary(list, v, mode > 0);
    if (mode > 0) {
      let best = -1;
      for (let i = 0; i < list.length; i++) {
        const x = list[i];
        const c = numsBelowText && typeof v === 'string' && (typeof x === 'number' || typeof x === 'boolean') ? -1 : lookupCmp(x, v);
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


  // does the expression hold a reference to more than one cell?
  function hasRange(n) {
    if (!n || typeof n !== 'object') return false;
    if (n.t === 'ref') return !!n.b && (n.a.r !== n.b.r || n.a.c !== n.b.c);
    if (n.t === 'ref3d' || n.t === 'rangeop' || n.t === 'name') return true;
    if (n.t === 'fn') return (n.args || []).some(hasRange);
    return hasRange(n.a) || hasRange(n.b);
  }
  // a jump-matrix branch that starts with a scalar jump ends the formula
  class JumpAbort { constructor(err) { this.err = err; } }
  // how many values Calc's stack holds when the evaluation of target starts
  function stackDepth(root, target) {
    const walk = (n, d) => {
      if (n === target) return d;
      if (!n || typeof n !== 'object') return -1;
      const kids = [];
      if (n.t === 'bin') { kids.push([n.a, 0], [n.b, 1]); }
      else if (n.t === 'paren' || n.t === 'neg' || n.t === 'pct') kids.push([n.a, 0]);
      else if (n.t === 'fn' && n.args) {
        const jump = n.name === 'IF' || n.name === 'CHOOSE' || n.name === 'IFERROR';
        n.args.forEach((a, i) => kids.push([a, jump ? 0 : i]));
      }
      for (const [k, add] of kids) { const r = walk(k, d + add); if (r >= 0) return r; }
      return -1;
    };
    return Math.max(0, walk(root, 0));
  }

  // does the expression contain an inline array constant? (cached on the node)
  function hasArr(node) {
    if (node._arr !== undefined) return node._arr;
    let r = node.t === 'arr';
    // (ROW/COLUMN of a range several rows/columns across give an array)
    if (!r && node.t === 'fn' && (node.name === 'ROW' || node.name === 'COLUMN') && node.args && node.args[0] && node.args[0].t === 'ref' && node.args[0].b) {
      const a = node.args[0].a, b = node.args[0].b;
      r = node.name === 'ROW' ? a.r !== b.r : a.c !== b.c;
    }
    if (!r) for (const k of ['a', 'b']) if (node[k] && typeof node[k] === 'object' && node[k].t && hasArr(node[k])) r = true;
    if (!r && node.args) r = node.args.some((n) => n && n.t && hasArr(n));
    node._arr = r;
    return r;
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
  const popOrder = (parts) => [].concat(...parts.map((p) => (Array.isArray(p) ? p : [p])));
  // (a range's blocks and formula cells go to one Kahan sum, added to the
  // total as its sum, error and held-back term)
  const sumState = (xs) => {
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
    return st;
  };
  const addState = (st, b) => { kahanAdd(st, b.sum); kahanAdd(st, b.err); kahanAdd(st, b.mem); };
  const sumParts = (parts) => {
    const T = { sum: 0, err: 0, mem: 0 };
    let R = null, rid = null;
    const flush = () => { if (R) addState(T, R); R = null; rid = null; };
    for (const p of parts) {
      if (Array.isArray(p) && p.segs) {
        const C = { sum: 0, err: 0, mem: 0 };
        for (const sg of p.segs) { if (Array.isArray(sg)) kahanAdd(C, kahanGet(sumState(sg))); else kahanAdd(C, sg); }
        // (a column holding formula cells ends with its held-back term added
        // plainly to sum + error, not compensated: SUM of =2^52, =1/3, =0.3 is 2^52)
        kahanAdd(T, C.mem === 0 ? C.sum + C.err : approxAdd(C.sum + C.err, C.mem));
      } else { flush(); if (Array.isArray(p)) kahanAdd(T, sumArray(p)); else kahanAdd(T, p); }
    }
    flush();
    return kahanGet(T);
  };
  // SUM scans a range column by column: the first column holding an error
  // reports the last error in that column (Calc's per-column sum)
  aggregate('SUM', (n, b, parts) => finite(sumParts(parts)), { colLastErr: true, kahanErr: true });
  // AVERAGE adds every value to one Kahan sum (no per-column sums)
  aggregate('AVERAGE', (n, b, parts) => (n.length ? finite(kahan(popOrder(parts)) / n.length) : E.DIV0), { colCont: true, kahanErr: true });
  const asFmt = (x, b) => (b && (x === 0 || x === 1) ? x === 1 : x);
  aggregate('MIN', (n, b) => (n.length ? asFmt(Math.min(...n), b) : 0), { textParam: true, arrErrSkip: true }, 1);
  aggregate('MAX', (n, b) => (n.length ? asFmt(Math.max(...n), b) : 0), { textParam: true, arrErrSkip: true }, 1);
  aggregate('PRODUCT', (n, b, parts) => (n.length ? finite(popOrder(parts).reduce((x, y) => x * y, 1)) : 0), { colCont: true });
  aggregate('COUNT', (n) => n.length, { skipErrors: true, countText: true });
  aggregate('MEDIAN', (n) => {
    if (!n.length) return E.VALUE;
    const s = n.slice().sort((a, b) => a - b);
    const m = s.length >> 1;
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
  }, { forward: true, textParam: true }, 1);
  // variance: values in Calc's order (parameters last to first, cells first to last)
  aggregate('STDEV', (n, b, parts) => (n.length < 2 ? E.DIV0 : finite(Math.sqrt(variance(popOrder(parts), false)))), { textParam: true, noTok: true });

  const variance = (n, pop) => {
    const mean = kahan(n) / n.length;
    let ss = 0; // Calc adds the squared deviations plainly (deviations via approxSub)
    for (const x of n) { const d = approxSub(x, mean); ss += d * d; } // rtl::math::approxSub
    return ss / (n.length - (pop ? 0 : 1));
  };
  aggregate('VAR', (n, b, parts) => (n.length < 2 ? E.DIV0 : finite(variance(popOrder(parts), false))), { textParam: true, noTok: true });
  aggregate('VARP', (n, b, parts) => (n.length < 1 ? E.DIV0 : finite(variance(popOrder(parts), true))), { textParam: true, noTok: true });
  aggregate('STDEVP', (n, b, parts) => (n.length < 1 ? E.DIV0 : finite(Math.sqrt(variance(popOrder(parts), true)))), { textParam: true, noTok: true });
  const kth = (largest) => function (nodes, ctx) {
    if (nodes.length !== 2) return E.VALUE;
    const r = this.collect([nodes[0]], ctx, { textParam: true });
    if (r.err) return r.err;
    const a = this.args([nodes[1]], ctx, ['v']); if (isErr(a)) return a;
    if (typeof a[0] === 'string') return E.PARAM;
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
      // cells of a reference are counted without computing them (Calc's
      // COUNTA never reads a formula cell's value: no circular reference)
      if (v instanceof RefList) v.ranges.forEach((rg) => { n += this.countCells(rg); });
      else if (v instanceof Range) n += this.countCells(v);
      else if (v instanceof Matrix) { for (const row of v.data) for (const x of row) if (x !== null) n++; }
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
    // with a digits argument, Calc first cuts a positive non-integer to 12
    // significant digits (tdf#124286: ROUNDUP(363136512.1386,3) is .138)
    if (nodes.length > 1 && x > 0 && Math.trunc(a[1]) < 12 && x % 1 !== 0) {
      const t = Math.floor(approxValue(Math.log10(x))) + 1 - 12;
      // (scaled in two steps so that tiny values do not overflow the factor)
      const t1 = Math.trunc(-t / 2), t2 = -t - t1;
      x = Math.floor(approxValue(x * Math.pow(10, t1) * Math.pow(10, t2))) / Math.pow(10, t1) / Math.pow(10, t2);
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
    if (mode !== 0) v = approxValue(v);
    const s = v < 0 ? -1 : 1;
    let a = Math.abs(v);
    // (rounding to nearest is approxFloor(a + 0.5): ROUND(0.5-10*2^-54,0) is 1)
    if (mode === 0) a = approxFloor(a + 0.5);
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
    const vals = this.svAll(nodes, ctx);
    for (const v of vals) if (isErr(v) && !v.fromRef) return v;
    const y = this.toNum(vals[1]); if (isErr(y)) return y;
    if (y === 0) return E.DIV0;
    const x = this.toNum(vals[0]); if (isErr(x)) return x;
    const q = x / y;
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
  math1('LN', (x) => (x <= 0 ? E.IARG : Math.log(x)));
  math1('LOG10', (x) => (x <= 0 ? E.IARG : Math.log10(x)));
  math1('FACT', (x) => {
    x = Math.floor(x);
    if (x < 0) return E.IARG;
    let f = 1;
    for (let i = 2; i <= x && isFinite(f); i++) f *= i;
    return isFinite(f) ? f : E.VALUE;
  });
  const parity = (odd) => function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['v']); if (isErr(a)) return a;
    if (a[0] === null || a[0] === EMPTYR) return E.PARAM;
    const x = this.toNum(a[0]); if (isErr(x)) return E.PARAM;
    return Math.abs(Math.trunc(x)) % 2 === (odd ? 1 : 0);
  };
  def('ISODD', 1, 1, parity(true));
  def('ISEVEN', 1, 1, parity(false));
  math1('EVEN', (x) => { const c = Math.ceil(Math.abs(approxValue(x)) / 2) * 2; return x < 0 ? -c : c; });
  math1('ODD', (x) => { let c = Math.ceil(Math.abs(approxValue(x))); if (c % 2 === 0) c++; return x < 0 ? -c : c; });
  def('LOG', 1, 2, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n', 'n'], [undefined, 10]); if (isErr(a)) return a;
    if (a[0] <= 0 || a[1] <= 0 || a[1] === 1) return E.IARG;
    return Math.log(a[0]) / Math.log(a[1]);
  });
  def('QUOTIENT', 2, 2, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n', 'n']); if (isErr(a)) return a;
    if (a[1] === 0) return E.IARG;
    const q = Math.trunc(approxValue(a[0] / a[1]));
    return isFinite(q) ? q : E.IARG; // (an add-in: an overflow is an invalid argument)
  });
  const toMultiple = (up) => function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n', 'n']); if (isErr(a)) return a;
    const [x, m] = a;
    if (m === 0) return up ? 0 : E.IARG;
    if (x === 0) return 0;
    if (x > 0 && m < 0) return E.IARG;
    const q = approxValue(x / m);
    return approxValue((up ? Math.ceil(q) : Math.floor(q)) * m);
  };
  def('CEILING', 2, 2, toMultiple(true));
  def('FLOOR', 2, 2, toMultiple(false));
  def('RAND', 0, 0, () => Math.random());
  def('RANDBETWEEN', 2, 2, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n', 'n']); if (isErr(a)) return a;
    const lo = Math.ceil(a[0]), hi = Math.floor(a[1]);
    if (lo > hi) return E.IARG;
    return lo + Math.floor(Math.random() * (hi - lo + 1));
  });

  def('SQRT', 1, 1, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n']); if (isErr(a)) return a;
    if (a[0] < 0) return E.IARG;
    return Math.sqrt(a[0]);
  });

  // ---- conditional aggregates
  // (count: COUNTIF, where an error constant as the range is the result)
  function condArgs(self, nodes, ctx, count) {
    const rg = self.ev(nodes[0], ctx);
    if (isErr(rg) && !rg.tok && !rg.lit) return rg; // (an error constant range: after the criterion)
    let crit = self.sv(nodes[1], ctx);
    const s = nodes.length > 2 && nodes[2].t !== 'missing' ? self.ev(nodes[2], ctx) : undefined;
    // computed errors first (in order); then, as Calc pops the parameters,
    // the sum range (anything but a range is #VALUE!), the criterion, the range
    if (isErr(crit) && !crit.fromRef) return crit;
    if (isErr(s) && !s.fromRef) return s;
    if (nodes.length > 2) {
      if (nodes[2].t === 'missing') return E.PARAM;
      if (isErr(s) || !(s instanceof Range || s instanceof Matrix)) return E.VALUE;
    }
    // (a criterion range that cannot be intersected is popped first: its #VALUE! wins)
    if (isErr(rg)) return count ? (crit === EREF['#VALUE!'] ? crit : rg) : isErr(crit) ? crit : E.VALUE;
    // (a criterion read from an error cell is popped first: its error wins)
    if (!(rg instanceof Range) && !(rg instanceof Matrix)) return isErr(crit) ? crit : E.PARAM;
    if (crit === MISSING || isLB(crit)) crit = null;
    if (crit === EMPTYR || isMV(crit)) crit = 0;
    let match = makeCriterion(crit);
    if (isErr(match)) return match;
    // (over an array a pattern, or "<>" for not empty, never matches a number)
    if (!(rg instanceof Range) && typeof crit === 'string' && (crit === '<>' || (!/^[<>=]/.test(crit) && hasWild(crit)))) {
      const m0 = match;
      match = (v) => typeof v !== 'number' && typeof v !== 'boolean' && m0(v);
    }
    let sum = rg;
    if (nodes.length > 2) {
      if (rg instanceof Matrix) {
        // an inline array: elements beyond the sum array are skipped; a
        // reference sum range is taken from its top left cell at the
        // array's size (SUMIF({1;2},2,B1) is B2)
        if (s instanceof Matrix) return { rg, match, sum: s };
        return { rg, match, sum: new Range(s.sheet, s.r1, s.c1, s.r1 + rg.rows - 1, s.c1 + rg.cols - 1) };
      }
      if (s instanceof Matrix) return { rg, match, sum: s };
      if (!(s instanceof Range)) return E.VALUE;
      // (a sum range written reversed, x then y with x > y along a dimension,
      // starts at y + min(d, n-1-d) there (d = x-y, n the criteria size), as
      // LibreOffice's xlsx import makes it)
      const st = (x, y, lo, n) => (x > y ? y + Math.min(x - y, Math.max(0, n - 1 - (x - y))) : lo);
      const r0 = s.wr ? st(s.wr[0], s.wr[2], s.r1, rg.rows) : s.r1, c0 = s.wr ? st(s.wr[1], s.wr[3], s.c1, rg.cols) : s.c1;
      sum = new Range(s.sheet, r0, c0, r0 + rg.rows - 1, c0 + rg.cols - 1);
    }
    return { rg, match, sum };
  }
  function condCollect(self, a) {
    const nums = [];
    let err = null;
    if (a.rg instanceof Matrix) {
      a.rg.data.forEach((row, r) => row.forEach((v, c) => {
        // (an error element of a computed array matches any criterion)
        if (err || !(isErr(v) || a.match(v))) return;
        let sv = a.sum instanceof Matrix ? (a.sum.data[r] ? a.sum.data[r][c] : undefined)
          : self.cellValue(a.sum.sheet, a.sum.r1 + r, a.sum.c1 + c);
        if (sv === EMPTYR) sv = 0;
        if (typeof sv === 'number') nums.push(sv);
        else if (typeof sv === 'boolean') nums.push(sv ? 1 : 0);
        else if (isErr(sv)) err = sv;
      }));
      return err ? { err } : { nums };
    }
    const blankMatches = a.match(null);
    const visit = (v, r, c) => {
      if (!a.match(v)) return true;
      // (an inline array as the sum range: the element at the same position, if any)
      let sv = a.sum instanceof Matrix ? (a.sum.data[r - a.rg.r1] ? a.sum.data[r - a.rg.r1][c - a.rg.c1] : undefined)
        : self.cellValue(a.sum.sheet, a.sum.r1 + (r - a.rg.r1), a.sum.c1 + (c - a.rg.c1));
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
    const a = condArgs(this, nodes, ctx, true); if (isErr(a)) return a;
    let n = 0;
    const rg = a.rg;
    if (rg instanceof Matrix) {
      for (const row of rg.data) for (const v of row) if (isErr(v) || a.match(v)) n++; // (error elements match any criterion)
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
    // an error condition: with an error function still to run, Calc passes
    // the error on as a token and goes on (MOD(IF(NA(),1,COUNT(1)),0) is #DIV/0!)
    if (isErr(c)) return this.contErr(c) || (c.lit ? errByCode(c.code) : c); // (an error constant condition: now pending)
    const b = this.toBool(c);
    // (a condition that does not convert: Calc passes on the error already set)
    if (isErr(b)) return this.contErr(this.convErr(b)) || b;
    this.boolLost = false;
    if (b) {
      if (nodes.length < 2) return true;
      const v = this.ev(nodes[1], ctx);
      return v === MISSING ? MISSV : v;
    }
    if (nodes.length < 3) return false;
    const v = this.ev(nodes[2], ctx);
    return v === MISSING ? MISSV : v;
  });
  def('IFERROR', 2, 2, function (nodes, ctx) {
    if (nodes[0].t === 'missing') return MISSN;
    const v = this.ev(nodes[0], ctx);
    if (v instanceof RefList && !v.is3d) return v; // a reference list is not an error: passed on
    const s = this.scalar(v, ctx);
    if (nodes.ef) this.efRan(nodes.ef); // (Calc runs IFERROR after its value, before the fallback)
    if (s === UNKNOWN_FN) return E.NOOP;
    if (isErr(s)) {
      this.boolLost = false;
      this.gerr = null; // (Calc clears its error before the fallback)
      // the alternative is computed as an ordinary part of the formula
      const top = this.running && this.running[this.running.length - 1];
      if (top) { top.fe = null; top.sup--; }
      let alt;
      try { alt = this.ev(nodes[1], ctx); } finally { if (top) top.sup++; }
      this.retLost = this.boolLost;
      return alt === MISSING ? MISSV : alt;
    }
    // Calc drops the boolean format of a value passed on (-IFERROR(FALSE,0) is 0)
    this.retLost = true;
    return v === MISSING ? MISSV : v;
  });
  function logical(name, init, combine) {
    def(name, 1, undefined, function (nodes, ctx) {
      let acc = init, any = false;
      const vals = [], clrs = [];
      for (const node of nodes) {
        if (node.t === 'missing') { vals.push(false); clrs.push(undefined); continue; } // an empty argument is FALSE
        const v = this.ev(node, ctx);
        vals.push(v); clrs.push(this.clr);
      }
      for (let k = 0; k < vals.length; k++) if (clrs[k] !== undefined && clrs[k] !== this.clr) vals[k] = demote(vals[k]);
      for (const v of vals) if (isErr(v) && !v.fromRef) return v;
      // parameters are examined last to first, as Calc pops them
      for (let k = vals.length - 1; k >= 0; k--) {
        const v = vals[k];
        if (v instanceof Range || v instanceof RefList) {
          let err = null;
          for (const rg of v instanceof RefList ? v.ranges : [v]) {
            this.eachCell(rg, (x) => {
              if (isErr(x)) { err = x; return false; }
              if (typeof x === 'number' || typeof x === 'boolean') { acc = combine(acc, !!+x); any = true; }
              return true;
            });
            if (err) return err;
          }
        } else if (v instanceof Matrix) {
          // an inline or computed array holding text or an empty element is
          // an invalid argument
          for (const row of v.data) for (const x of row) {
            if (isErr(x)) return x;
            if (typeof x === 'number' || typeof x === 'boolean') { acc = combine(acc, !!+x); any = true; } else if (typeof x === 'string' || x === null) return E.IARG;
          }
        } else if (isErr(v)) return v;
        else if (typeof v === 'number' || typeof v === 'boolean') { acc = combine(acc, !!+v); any = true; } else if (v === EMPTYR || isLB(v) || isMV(v)) { acc = combine(acc, false); any = true; } else if (typeof v === 'string') return E.VALUE;
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
  def('ISBLANK', 1, 1, function (nodes, ctx) { const v = this.sv(nodes[0], ctx); return v === null || v === LBLANK || v === MISSV; });
  def('ISNUMBER', 1, 1, function (nodes, ctx) {
    const v = this.sv(nodes[0], ctx);
    return typeof v === 'number' || typeof v === 'boolean';
  });
  def('ISTEXT', 1, 1, function (nodes, ctx) { return typeof this.sv(nodes[0], ctx) === 'string'; });
  def('ISERROR', 1, 1, function (nodes, ctx) {
    const r = this.ev(nodes[0], ctx);
    if (r instanceof RefList && !r.is3d) return false;
    const v = this.scalar(r, ctx);
    return v === UNKNOWN_FN ? E.NOOP : isErr(v);
  });

  // ---- text
  def('CONCATENATE', 0, undefined, function (nodes, ctx) {
    const vals = this.svAll(nodes, ctx);
    for (const v of vals) if (isErr(v) && !v.fromRef) return v;
    // parameters are taken first to last: the first error is the result,
    // but an error token overwrites it (the last one wins)
    let s = '', err = null;
    for (const v of vals) {
      const t = this.toStr(v);
      if (isErr(t)) { if (!err || t.tok) err = t; continue; }
      if (!err) s += t;
    }
    return err || s;
  });
  // Calc's LEN/LEFT/RIGHT/MID count Unicode code points, not UTF-16 units
  const codePoints = (s) => Array.from(s);
  const cpCount = (s) => codePoints(s).length;
  const INT32 = 2147483647; // text positions and lengths are 32-bit in Calc
  // a negative or too large text length / position is an error as soon as it
  // is popped (before truncation); a start below 1 only after all arguments
  const lenCheck = (x) => (x < 0 || Math.trunc(x) > INT32 ? E.IARG : null);

  def('LEFT', 1, 2, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s', 'n'], [undefined, 1], [null, lenCheck]); if (isErr(a)) return a;
    const n = Math.trunc(a[1]);
    return codePoints(a[0]).slice(0, n).join('');
  });
  def('RIGHT', 1, 2, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s', 'n'], [undefined, 1], [null, lenCheck]); if (isErr(a)) return a;
    const n = Math.trunc(a[1]);
    const cps = codePoints(a[0]);
    return n === 0 ? '' : cps.slice(Math.max(0, cps.length - n)).join('');
  });
  const MID_CHECKS = Object.assign([null, lenCheck, lenCheck], { soft: true });
  def('MID', 3, 3, function (nodes, ctx) {
    // (MID pops its text after the position errors: an error token there wins)
    const a = this.args(nodes, ctx, ['s', 'n', 'n'], undefined, MID_CHECKS); if (isErr(a)) return a;
    const st = Math.trunc(a[1]), n = Math.trunc(a[2]);
    if (st < 1) return E.IARG;
    if (st >= 2 ** 28 || n >= 2 ** 28) return E.STROVF; // beyond the longest string
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
    if (isMV(v)) return 0;
    if (v === null || v === EMPTYR) return E.IARG;
    const n = textToNum(String(v));
    return n === null ? E.IARG : n;
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
    if (st < 1 || st > a[1].length + 1) return E.IARG;
    if (a[0] === '') return E.VALUE; // (Calc's search finds no empty text)
    const re = new RegExp(wildcardRe(a[0]).source.slice(1, -1), 'iu');
    const m = re.exec(a[1].slice(st - 1));
    return m ? cpCount(a[1].slice(0, m.index + st - 1)) + 1 : E.VALUE;
  });

  def('ISNA', 1, 1, function (nodes, ctx) { const v = this.sv(nodes[0], ctx); return isErr(v) && v.code === '#N/A'; });
  def('ISERR', 1, 1, function (nodes, ctx) {
    const r = this.ev(nodes[0], ctx);
    if (r instanceof RefList && !r.is3d) return false;
    const v = this.scalar(r, ctx);
    return isErr(v) && v.code !== '#N/A';
  });
  def('ISLOGICAL', 1, 1, function (nodes, ctx) { return typeof this.sv(nodes[0], ctx) === 'boolean'; });
  def('ISREF', 1, 1, function (nodes, ctx) {
    const v = this.ev(nodes[0], ctx);
    return v instanceof Range || v instanceof RefList;
  });
  def('ISNONTEXT', 1, 1, function (nodes, ctx) { return typeof this.sv(nodes[0], ctx) !== 'string'; });
  def('N', 1, 1, function (nodes, ctx) {
    const v = this.sv(nodes[0], ctx);
    // Calc's internal lookup error (VLOOKUP(1/0,...)) is #VALUE! to N
    if (v === E.LOOKNA || (isErr(v) && v.look)) return E.VALUE;
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
      const [m, e] = exp15(x).split('e');
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
    // (a number given a text format, one with @, is written as in General)
    if (typeof v === 'number' && /@/.test(a[1].replace(/"[^"]*"?|\\./g, ''))) return fmtNum(v);
    return textFormat(v, a[1]);
  });
  def('REPT', 2, 2, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s', 'n']); if (isErr(a)) return a;
    const n = Math.trunc(a[1]);
    if (a[1] < 0) return E.IARG;
    if (a[0].length * n >= 2 ** 28) return E.STROVF; // a string of 2^28 units or more
    return a[0].repeat(n);
  });
  def('SUBSTITUTE', 3, 4, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s', 's', 's', 'n']); if (isErr(a)) return a;
    const [t, from, to, nth] = a;
    if (from === '') return t;
    if (nth === undefined) return t.split(from).join(to);
    const k = Math.trunc(nth);
    if (k < 1) return E.IARG;
    let i = -1;
    for (let j = 0; j < k; j++) { i = t.indexOf(from, i + 1); if (i < 0) return t; }
    return t.slice(0, i) + to + t.slice(i + from.length);
  });
  def('CHAR', 1, 1, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['n']); if (isErr(a)) return a;
    const n = Math.trunc(a[0]);
    if (n < 1 || n > 255) return E.IARG;
    // Calc decodes the byte in the system encoding (UTF-8): a lone byte >= 128 is invalid
    return n < 128 ? String.fromCharCode(n) : '\uFFFD';
  });
  def('CODE', 1, 1, function (nodes, ctx) {
    const a = this.args(nodes, ctx, ['s']); if (isErr(a)) return a;
    if (a[0] === '') return 0;
    // Calc converts the first character to the system encoding (UTF-8) and gives its first byte
    const u = a[0].charCodeAt(0);
    if (u < 0x80) return u;
    if (u >= 0xD800 && u <= 0xDFFF) return 0;
    return new TextEncoder().encode(a[0][0])[0];
  });
  def('CHOOSE', 2, undefined, function (nodes, ctx) {
    const a = this.args(nodes.slice(0, 1), ctx, ['n']); if (isErr(a)) return a.lit ? errByCode(a.code) : a;
    const i = Math.trunc(a[0]);
    if (i < 1 || i >= nodes.length) return E.IARG;
    this.boolLost = false;
    const v = this.ev(nodes[i], ctx);
    return v === MISSING ? MISSV : v;
  });
  const rowCol = (row) => function (nodes, ctx) {
    if (!nodes.length || nodes[0].t === 'missing') return (row ? ctx.r : ctx.c) + 1;
    const v = this.ev(nodes[0], ctx);
    if (isErr(v)) return v;
    if (!(v instanceof Range)) return E.PARAM;
    // several rows / columns give an array of their numbers (a function-made
    // array: aggregates take it whole, a scalar parameter its first element)
    if ((row ? v.rows : v.cols) > 1) {
      const xs = [];
      if (row) for (let r = v.r1; r <= v.r2; r++) xs.push([r + 1]);
      else { xs.push([]); for (let c = v.c1; c <= v.c2; c++) xs[0].push(c + 1); }
      const m = new Matrix(xs);
      m.jm = true;
      return m;
    }
    return (row ? v.r1 : v.c1) + 1;
  };
  def('ROW', 0, 1, rowCol(true));
  def('COLUMN', 0, 1, rowCol(false));
  def('ROWS', 1, 1, function (nodes, ctx) {
    const v = this.ev(nodes[0], ctx);
    if (isErr(v)) return v;
    if (v instanceof RefList && v.is3d) return v.ranges.length * v.ranges[0].rows; // Calc counts each sheet's rows
    return v instanceof Range || v instanceof Matrix ? v.rows : E.PARAM;
  });
  def('COLUMNS', 1, 1, function (nodes, ctx) {
    const v = this.ev(nodes[0], ctx);
    if (isErr(v)) return v;
    if (v instanceof RefList && v.is3d) return v.ranges.length * v.ranges[0].cols;
    return v instanceof Range || v instanceof Matrix ? v.cols : E.PARAM;
  });
  // OFFSET(ref, rows, cols, [height], [width]): a reference moved and resized
  def('OFFSET', 3, 5, function (nodes, ctx) {
    const ref = this.ev(nodes[0], ctx);
    const a = this.args(nodes.slice(1), ctx, ['n', 'n', 'n', 'n']);
    if (isErr(ref) && !ref.lit) return ref;
    if (isErr(a)) return a;
    if (isErr(ref)) return E.VALUE; // an error constant as the reference: after the other arguments' errors
    if (!(ref instanceof Range)) return E.PARAM;
    const dr = Math.trunc(a[0]), dc = Math.trunc(a[1]);
    let h = a[2] === undefined ? ref.rows : Math.trunc(a[2]);
    let w = a[3] === undefined ? ref.cols : Math.trunc(a[3]);
    // a size below 1 or a result outside the sheet is an invalid argument
    if (h < 1 || w < 1) return E.IARG;
    const r1 = ref.r1 + dr, c1 = ref.c1 + dc, r2 = r1 + h - 1, c2 = c1 + w - 1;
    if (r1 < 0 || c1 < 0 || r2 >= MAXROW || c2 >= MAXCOL) return E.IARG;
    return new Range(ref.sheet, r1, c1, r2, c2);
  });
  // functions that reset Calc's pending error (nGlobalError) when they finish
  const PASS_FMT = new Set(['IF', 'CHOOSE']);
  const CLEARS_ERROR = new Set(['IFERROR', 'ISNUMBER', 'ISTEXT', 'ISBLANK', 'ISERROR', 'ISERR', 'ISNA', 'ISLOGICAL',
    'ISNONTEXT', 'ISREF', 'COUNT', 'COUNTA', 'N']);
  // parameters that take one value: an array there gives one result per element
  const LOOKUPS = new Set(['VLOOKUP', 'HLOOKUP', 'MATCH']);
  const MAPPED_PARAMS = { COUNTIF: [1], SUMIF: [1], AVERAGEIF: [1], VLOOKUP: [0, 2], HLOOKUP: [0, 2], MATCH: [0], INDEX: [1, 2],
    LARGE: [1], SMALL: [1] };
  const ADDIN = new Set(['QUOTIENT', 'RANDBETWEEN']);
  const ELEMENTWISE = new Set(['IF', 'IFERROR', 'ISNUMBER', 'ISTEXT', 'ISBLANK', 'ISERROR', 'ISNA', 'ISERR', 'ISLOGICAL',
    'ISNONTEXT', 'LEN', 'ABS', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'INT', 'TRUNC', 'SIGN', 'NOT', 'UPPER', 'LOWER', 'LEFT',
    'RIGHT', 'MID', 'TRIM', 'VALUE', 'MOD', 'POWER', 'SQRT', 'EXACT', 'EXP', 'LN', 'SEARCH', 'FIND', 'SUBSTITUTE',
    'PROPER', 'REPT', 'CODE', 'CHAR', 'ISODD', 'ISEVEN', 'N', 'T', 'CONCATENATE', 'LOG', 'LOG10', 'FACT', 'EVEN', 'ODD',
    'TEXT', 'QUOTIENT', 'CEILING', 'FLOOR', 'RANDBETWEEN']);
  def('TRANSPOSE', 1, 1, function (nodes, ctx) {
    if (nodes[0].t === 'missing') return E.PARAM;
    let v = this.fev(nodes[0], ctx);
    if (v instanceof Range) v = this.rangeMatrix(v);
    // (booleans come out as numbers)
    const nb = (x) => (typeof x === 'boolean' ? +x : x);
    if (!(v instanceof Matrix)) return nb(v);
    const d = v.data;
    return new Matrix(d[0].map((_, c) => d.map((row) => nb(row[c]))));
  });
  def('SUMPRODUCT', 1, undefined, function (nodes, ctx) {
    const args = [];
    for (const n of nodes) {
      if (n.t === 'missing') { args.push({ g: [[0]] }); continue; }
      // (Calc's force-array parameter: functions inside take their
      // arguments as arrays too, SUMPRODUCT(MAX(A1:A3*B1:B3)))
      this.fa = (this.fa || 0) + 1;
      let v;
      try { v = this.aev(n, ctx); } finally { this.fa--; }
      if (isMV(v)) return E.VALUE;
      if (isErr(v)) args.push({ err: v, cell: n.t === 'ref' && !n.b });
      else args.push({ g: v instanceof Matrix ? v.data : [[v]] });
    }
    // Calc pops the arguments last to first as matrices: a single cell's
    // error sets its error if none is set (the element is 0); an error
    // token sets it and moves it into a 1x1 element, clearing it; a size
    // differing from the last argument's is that error, else #VALUE!; a
    // set error is the result (SUMPRODUCT(A12,B1:B3) is A12's error,
    // SUMPRODUCT(#N/A,B1:B3) is #VALUE!, SUMPRODUCT(#N/A,A12) is #N/A)
    let gerr = null;
    const grids = new Array(args.length);
    for (let i = args.length - 1; i >= 0; i--) {
      const a = args[i];
      if (a.err && a.cell) { if (!gerr) gerr = a.err; grids[i] = [[0]]; }
      else if (a.err) { gerr = null; grids[i] = [[a.err]]; }
      else grids[i] = a.g;
      const L = grids[args.length - 1];
      if (grids[i].length !== L.length || grids[i][0].length !== L[0].length) return gerr || E.VALUE;
    }
    if (gerr) return gerr;
    const R = grids[0].length, C = grids[0][0].length;
    const prods = [];
    for (let c = 0; c < C; c++) for (let r = 0; r < R; r++) {
      // (Calc multiplies the arrays in turn: a text element makes the
      // product 0 even beside an error, and of several errors the last wins)
      let p = 1, err = null, text = false;
      for (const g of grids) {
        const x = g[r][c];
        if (isErr(x)) err = x;
        else if (typeof x === 'string') text = true;
        else p *= typeof x === 'number' ? x : typeof x === 'boolean' ? +x : 0;
      }
      if (text) p = 0;
      else if (err) return err;
      prods.push(p);
    }
    return kahan(prods);
  });
  // COUNTIFS / SUMIFS / AVERAGEIFS: all criteria ranges must have the same shape
  function multiCond(self, nodes, ctx, first) {
    const pairs = [];
    for (let i = first; i + 1 < nodes.length; i += 2) {
      const rg = self.ev(nodes[i], ctx);
      if (isErr(rg)) return rg;
      if (!(rg instanceof Range)) return E.VALUE;
      let crit = self.sv(nodes[i + 1], ctx);
      if (crit === MISSING || isLB(crit)) crit = null;
    if (crit === EMPTYR || isMV(crit)) crit = 0;
      const m = makeCriterion(crit);
      if (isErr(m)) return m;
      pairs.push([rg, m]);
    }
    const R = pairs[0][0].rows, C = pairs[0][0].cols;
    if (pairs.some(([rg]) => rg.rows !== R || rg.cols !== C)) return E.IARG;
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
    if (first.rows !== sr.rows || first.cols !== sr.cols) return E.IARG;
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
    let v = this.sv(nodes[0], ctx);
    if (v === MISSING) v = 0; // (an empty lookup value is 0)
    const clr0 = this.clr;
    const pend0 = this.pend;
    const tbl = this.fev(nodes[1], ctx);
    this.lookPend(v, pend0);
    const rest = this.args(nodes.slice(2), ctx, ['n', 'b'], [undefined, true]);
    if (this.clr !== clr0) v = demote(v); // a later argument cleared the pending error
    if (isErr(v) && !v.fromRef) return lookErr(v);
    if (isErr(tbl) && !tbl.lit) return tbl;
    if (isErr(rest)) return rest;
    // an error constant as the table or the lookup value: #VALUE!
    if (isErr(tbl) || (isErr(v) && v.lit)) return E.VALUE;
    const row = Math.trunc(rest[0]);
    const sorted = nodes.length > 3 && nodes[3].t === 'missing' ? false : rest[1];
    if (!(tbl instanceof Range) && !(tbl instanceof Matrix)) return E.VALUE;
    if (row < 1 || row > tbl.rows) return E.IARG;
    if (isErr(v)) return v; // an error read from a cell, after the index check
    if (v === null) v = ''; else if (isLB(v) || v === EMPTYR || isMV(v)) v = 0; // a blank lookup value looks for empty text (a formula's "")
    const list = this.firstLine(tbl, true);
    // (an exact HLOOKUP of a number along an array stops at an error element)
    if (!sorted && tbl instanceof Matrix && (typeof v === 'number' || typeof v === 'boolean')) {
      const k = list.findIndex((x) => isErr(x) || lookupEq(v, x, true));
      if (k >= 0 && isErr(list[k])) return list[k];
    }
    const i = sorted && tbl instanceof Matrix ? arrLookup(list, v) : lookupPos(list, v, sorted ? 1 : 0, tbl instanceof Range ? 'row' : null);
    if (isErr(i)) return i;
    if (i < 0) return E.NA;
    let out = tbl instanceof Range ? this.cellValue(tbl.sheet, tbl.r1 + row - 1, tbl.c1 + i) : tbl.data[row - 1][i];
    if (out === EMPTYR) return LEMPTY;
    if (typeof out === 'boolean') return out ? 1 : 0;
    return out === null ? (tbl instanceof Matrix ? '' : LBLANK) : out; // (an empty array element: empty text)
  });

  // ---- lookup
  def('VLOOKUP', 3, 4, function (nodes, ctx) {
    let v = this.sv(nodes[0], ctx);
    if (v === MISSING) v = 0; // (an empty lookup value is 0)
    const clr0 = this.clr;
    const pend0 = this.pend;
    const tbl = this.fev(nodes[1], ctx);
    this.lookPend(v, pend0);
    const rest = this.args(nodes.slice(2), ctx, ['n', 'b'], [undefined, true]);
    if (this.clr !== clr0) v = demote(v); // a later argument cleared the pending error
    if (isErr(v) && !v.fromRef) return lookErr(v);
    if (isErr(tbl) && !tbl.lit) return tbl;
    if (isErr(rest)) return rest;
    if (isErr(tbl) || (isErr(v) && v.lit)) return E.VALUE;
    const col = Math.trunc(rest[0]);
    const sorted = nodes.length > 3 && nodes[3].t === 'missing' ? false : rest[1];
    if (!(tbl instanceof Range) && !(tbl instanceof Matrix)) return E.VALUE;
    if (col < 1 || col > tbl.cols) return E.IARG;
    if (isErr(v)) return v; // an error read from a cell, after the index check
    if (v === null) v = ''; else if (isLB(v) || v === EMPTYR || isMV(v)) v = 0; // a blank lookup value looks for empty text (a formula's "")
    const first = this.firstLine(tbl, false);
    const i = sorted && tbl instanceof Matrix ? arrLookup(first, v) : lookupPos(first, v, sorted ? 1 : 0, tbl instanceof Range ? 'col' : null);
    if (isErr(i)) return i;
    if (i < 0) return E.NA;
    let out = tbl instanceof Range ? this.cellValue(tbl.sheet, tbl.r1 + i, tbl.c1 + col - 1) : tbl.data[i][col - 1];
    if (out === EMPTYR) return LEMPTY;
    if (typeof out === 'boolean') return out ? 1 : 0;
    return out === null ? (tbl instanceof Matrix ? '' : LBLANK) : out; // (an empty array element: empty text)
  });
  def('MATCH', 2, 3, function (nodes, ctx) {
    let v = this.sv(nodes[0], ctx);
    if (v === MISSING) v = 0;
    const clr0 = this.clr;
    const pend0 = this.pend;
    const rg = this.fev(nodes[1], ctx);
    this.lookPend(v, pend0);
    const rest = this.args(nodes.slice(2), ctx, ['n'], [1]);
    if (this.clr !== clr0) v = demote(v); // a later argument cleared the pending error
    if (isErr(v) && !v.fromRef) return lookErr(v);
    if (isErr(rg) && !rg.lit) return rg;
    if (rg === MISSING) return E.PARAM;
    if (isErr(rest)) return rest;
    if (isErr(rg) || (isErr(v) && v.lit)) return E.VALUE;
    // the match type is truncated: -1, 0 or 1, anything else is a parameter error
    const mode = Math.trunc(rest[0]);
    if (mode < -1 || mode > 1) return E.PARAM;
    // (a lookup array that is neither a reference nor an array: before a cell's error)
    if (!(rg instanceof Range) && !(rg instanceof Matrix) && !(rg instanceof RefList)) return E.VALUE;
    if (!(rg instanceof RefList) && rg.rows !== 1 && rg.cols !== 1) return E.PARAM;
    if (isErr(v)) return v;
    if (rg instanceof RefList) return E.VALUE;
    if (v === null) v = ''; else if (isLB(v) || v === EMPTYR || isMV(v)) v = 0; // a blank lookup value looks for empty text (a formula's "")
    const list = this.firstLine(rg, rg.cols !== 1);
    const i = lookupPos(list, v, mode, rg instanceof Range ? (rg.cols === 1 ? 'col' : 'row') : null, true);
    return i < 0 ? E.NA : i + 1;
  });
  def('INDEX', 1, 4, function (nodes, ctx) {
    const ref = this.ev(nodes[0], ctx);
    if (isErr(ref) && !ref.lit) return ref;
    const a = this.args(nodes.slice(1), ctx, ['n', 'n', 'n'], [0, undefined, 1]);
    if (isErr(a)) return a;
    // Calc reads the row and area as 32-bit, the column as 16-bit integers
    if (Math.trunc(a[0]) > INT32 || Math.trunc(a[1]) > 32767 || Math.trunc(a[2]) > INT32) return E.IARG;
    // an error constant as the reference: after the other parameters and the area
    if (isErr(ref) && (a[0] < 0 || a[1] < 0)) return E.IARG;
    if (isErr(ref)) return Math.trunc(a[2]) !== 1 ? (a[2] < 1 ? E.IARG : E.REF) : ref;
    let row = Math.trunc(a[0]);
    let col = a[1] === undefined ? undefined : Math.trunc(a[1]);
    let tgt = ref;
    if (ref instanceof RefList && ref.is3d) return E.REF;
    if (ref instanceof RefList) {
      // a reference list: the area-th reference
      const k = Math.trunc(a[2]);
      if (k < 1) return E.IARG;
      if (k > ref.ranges.length) return E.REF;
      tgt = ref.ranges[k - 1];
    } else if (Math.trunc(a[2]) !== 1) return a[2] < 1 ? E.IARG : E.REF; // a single area
    const scalar = !(tgt instanceof Range) && !(tgt instanceof Matrix);
    if (scalar) { if (row < 0 || (col !== undefined && col < 0)) return E.IARG; return E.REF; } // not a reference
    if (col === undefined) {
      if (tgt.rows === 1 && tgt.cols > 1) { col = row; row = 0; } else col = 0;
    }
    if (row < 0 || col < 0) return E.IARG;
    if (row > tgt.rows || col > tgt.cols) return E.REF;
    if (tgt instanceof Matrix) {
      this.retLost = true;
      // an empty element comes out as empty text; booleans lose their format
      const el = (x) => (x === null ? '' : typeof x === 'boolean' ? +x : x);
      if (row === 0 && col === 0) return tgt;
      if (row === 0) return new Matrix(tgt.data.map((r) => [el(r[col - 1])]));
      if (col === 0) return new Matrix([tgt.data[row - 1].map(el)]);
      return el(tgt.data[row - 1][col - 1]);
    }
    const r1 = row === 0 ? tgt.r1 : tgt.r1 + row - 1;
    const r2 = row === 0 ? tgt.r2 : r1;
    const c1 = col === 0 ? tgt.c1 : tgt.c1 + col - 1;
    const c2 = col === 0 ? tgt.c2 : c1;
    return new Range(tgt.sheet, r1, c1, r2, c2);
  });

  // ---------------------------------------------------------------- public API
  function outValue(v) {
    if (v === EMPTYR || v === null || isLB(v)) return 0;
    if (isErr(v)) return { error: errText(v.code) };
    if (typeof v === 'number') {
      if (!isFinite(v)) return { error: '#NUM!' };
      // integers below 2^53 are written exactly, other numbers with 15 digits
      if (Number.isInteger(v) && Math.abs(v) < 9007199254740992) return v === 0 ? 0 : v;
      const r = round15(v);
      // near the largest double the rounded value overflows: keep the number
      if (!isFinite(r)) return v;
      return r === 0 ? 0 : r;
    }
    return v;
  }

  // JSON text -> value, like JSON.parse, but an object whose keys JavaScript
  // would reorder (integer-like keys come first: a sheet named "2024") keeps
  // the text's key order in a non-enumerable __order property, so sheets stay
  // in workbook order for 3D references.
  function parseJSONOrdered(text) {
    const v = JSON.parse(text);
    if (!/"\d+"\s*:/.test(text)) return v;
    let i = 0;
    const ws = () => { while (i < text.length && ' \t\n\r'.indexOf(text[i]) >= 0) i++; };
    const str = () => { const st = i; i++; while (text[i] !== '"') i += text[i] === '\\' ? 2 : 1; i++; return JSON.parse(text.slice(st, i)); };
    const val = (into) => {
      ws();
      const ch = text[i];
      if (ch === '{') {
        i++; const keys = []; ws();
        if (text[i] === '}') { i++; return; }
        for (;;) {
          ws(); const k = str(); keys.push(k); ws(); i++; // ':'
          val(into && typeof into[k] === 'object' && into[k] !== null ? into[k] : null);
          ws(); if (text[i] === ',') { i++; continue; } i++; break;
        }
        if (into && keys.some((k, n) => Object.keys(into)[n] !== k)) Object.defineProperty(into, '__order', { value: keys });
        return;
      }
      if (ch === '[') {
        i++; ws(); let n = 0;
        if (text[i] === ']') { i++; return; }
        for (;;) { val(into ? into[n] : null); n++; ws(); if (text[i] === ',') { i++; continue; } i++; break; }
        return;
      }
      if (ch === '"') { str(); return; }
      while (i < text.length && ',]} \t\n\r'.indexOf(text[i]) < 0) i++;
    };
    val(v);
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
        } else cells[a] = raw === '=' ? '' : raw;
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
    // a multi-cell range is one node of the graph (shared by every formula
    // using it), linked to the formula cells inside it
    const rnodes = new Map();
    for (const cell of formulas) deps.set(cell, staticPrecedents(eng, cell, rnodes));
    for (const rn of rnodes.values()) deps.set(rn, rn.deps);

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
    for (const cell of order) if (!cell.isRange && !tainted.has(cell)) eng.cellValue(cell.sheet, cell.r, cell.c);
    if (tainted.size) {
      const rest = [...tainted].filter((c) => !c.isRange).sort((a, b) => a.si - b.si || a.c - b.c || a.r - b.r);
      for (const cell of rest) eng.cellValue(cell.sheet, cell.r, cell.c);
    }
  }

  // formula cells of a sheet inside a rectangle
  function formulaCellsIn(s, r1, c1, r2, c2) {
    // formula cells of the sheet by column, sorted by row
    if (!s.fcols) {
      s.fcols = new Map();
      for (const x of s.cells.values()) if (x.formula) { if (!s.fcols.has(x.c)) s.fcols.set(x.c, []); s.fcols.get(x.c).push(x); }
      for (const list of s.fcols.values()) list.sort((p, q) => p.r - q.r);
    }
    const out = [];
    const cols = c2 - c1 + 1 > s.fcols.size ? [...s.fcols.keys()].filter((c) => c >= c1 && c <= c2).sort((p, q) => p - q)
      : Array.from({ length: c2 - c1 + 1 }, (_, k) => c1 + k);
    for (const c of cols) {
      const list = s.fcols.get(c);
      if (!list) continue;
      let lo = 0, hi = list.length;
      while (lo < hi) { const m = (lo + hi) >> 1; if (list[m].r < r1) lo = m + 1; else hi = m; }
      for (let k = lo; k < list.length && list[k].r <= r2; k++) out.push(list[k]);
    }
    return out;
  }

  // Calc's compiler limits: parentheses nest at most 98 deep (Err:514), a
  // formula holds at most 8191 tokens and a
  // function at most 255 parameters (Err:512); both show as #N/A.
  // Argument counts Calc's compiler rejects, even in a branch never taken:
  // more than one for a function of at most one parameter, none for a
  // function that needs some, and a wrong count for IF and IFERROR.
  const RUNTIME_COUNT = new Set(['COUNTBLANK', 'ROWS', 'COLUMNS', 'TRANSPOSE']);
  function compileArity(ast) {
    const todo = [ast];
    while (todo.length) {
      const node = todo.pop();
      if (!node || typeof node !== 'object') continue;
      if (node.t === 'fn' && FUNCS[node.name]) {
        const f = FUNCS[node.name], n = node.args.length, cond = node.name === 'IF' || node.name === 'IFERROR' || (node.name === 'CHOOSE' && !n);
        // (Calc's fixed one- and no-parameter functions; not the reference
        // and matrix ones, which check their count when they run)
        const fixed = f.max === f.min && f.max <= 1 && !RUNTIME_COUNT.has(node.name);
        if (n < f.min && ((!n && fixed) || cond)) return E.NOARG.code;
        if (n > f.max && (fixed || cond)) return fixed ? E.PAREN.code : E.PARAM.code;
      }
      if (node.args) todo.push(...node.args);
      if (node.a && node.a.t) todo.push(node.a);
      if (node.b && node.b.t) todo.push(node.b);
    }
    return null;
  }
  function compileLimits(src, ast) {
    let d = 0, max = 0, n = 0;
    for (const t of tokenize(src)) {
      if (t.t === 'eof') break;
      n++;
      if (t.t === 'op') { if (t.v === '(') max = Math.max(max, ++d); else if (t.v === ')') d--; }
    }
    if (max > 98) return 'Err:514';
    if (n > 8191) return 'Err:512';
    const todo = [ast];
    while (todo.length) {
      const node = todo.pop();
      if (!node || typeof node !== 'object') continue;
      if (node.args) { if (node.args.length > 255) return 'Err:512'; todo.push(...node.args); }
      if (node.a && node.a.t) todo.push(node.a);
      if (node.b && node.b.t) todo.push(node.b);
    }
    return null;
  }

  // Formula cells a formula may read, from its references (an over-approximation).
  function staticPrecedents(eng, cell, rnodes) {
    const out = [];
    const seen = new Set();
    const add = (x) => { if (x && (x.formula || x.isRange) && !seen.has(x)) { seen.add(x); out.push(x); } };
    // (an explicit stack: formulas can nest deeper than the call stack)
    const todo = [];
    const walk = (node) => { todo.push(node); };
    const visit = (node) => {
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
        if ((r2 > r1 || c2 > c1) && rnodes) {
          const key = eng.wb.sheets.indexOf(s) + ':' + r1 + ':' + c1 + ':' + r2 + ':' + c2;
          let rn = rnodes.get(key);
          if (!rn) { rn = { isRange: true, deps: formulaCellsIn(s, r1, c1, r2, c2) }; rnodes.set(key, rn); }
          add(rn);
        } else formulaCellsIn(s, r1, c1, r2, c2).forEach(add);
        return;
      }
      if (node.a && node.a.t) walk(node.a);
      if (node.b && node.b.t) walk(node.b);
      if (node.args) node.args.forEach(walk);
    };
    walk(cell.ast);
    while (todo.length) visit(todo.pop());
    return out;
  }

  const api = {
    computeWorkbook, parseJSONOrdered, recalcAll, parse, tokenize, fmtNum, textToNum, parseAddr, addrName, colName, colIndex, Engine, FUNCS,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.SheetEngine = api;
})(typeof window !== 'undefined' ? window : globalThis);
