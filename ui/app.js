/* Browser UI: grid, selection, editing, formula bar, open/save (CSV, JSON). */
(function () {
  'use strict';
  const { computeWorkbook, fmtNum, colName, parseAddr } = window.SheetEngine;
  const { parseCSV, toCSV, typed } = window.SheetCSV;

  const ERROR_TEXT = {
    '#DIV/0!': 'Division by zero',
    '#VALUE!': 'Wrong type of argument (or circular reference)',
    '#REF!': 'Invalid reference',
    '#NAME?': 'Unknown name or function',
    '#N/A': 'Value not available',
    '#NUM!': 'Invalid numeric value',
    '#NULL!': 'Empty intersection',
  };

  let book = { sheets: { Sheet1: {} } };
  let current = 'Sheet1';
  let computed = null;
  let sel = { r: 0, c: 0 };
  let editing = null; // { input, td, original }
  let nRows = 100, nCols = 26;

  const $ = (id) => document.getElementById(id);
  const grid = $('grid'), wrap = $('grid-wrap');
  const fInput = $('formula-input'), nameBox = $('name-box'), info = $('cell-info');

  const addr = (r, c) => colName(c) + (r + 1);
  const raw = (r, c) => book.sheets[current][addr(r, c)];
  const value = (r, c) => (computed.sheets[current] || {})[addr(r, c)];

  function status(msg) { $('status').textContent = msg || ''; }

  function recalc() {
    const t0 = performance.now();
    computed = computeWorkbook(book);
    status(`recalculated in ${Math.round(performance.now() - t0)} ms`);
  }

  function display(v) {
    if (v === undefined || v === null) return { text: '', cls: '' };
    if (typeof v === 'number') return { text: fmtNum(v), cls: 'num' };
    if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE', cls: 'bool' };
    if (typeof v === 'object' && v.error) return { text: v.error, cls: 'err' };
    return { text: String(v), cls: '' };
  }

  function fitSize() {
    let mr = 0, mc = 0;
    for (const a of Object.keys(book.sheets[current])) {
      const p = parseAddr(a);
      if (p) { mr = Math.max(mr, p.r); mc = Math.max(mc, p.c); }
    }
    nRows = Math.max(100, mr + 20, sel.r + 20);
    nCols = Math.max(26, mc + 5, sel.c + 5);
  }

  function render() {
    fitSize();
    const parts = ['<colgroup><col style="width:44px">' + '<col style="width:96px">'.repeat(nCols) + '</colgroup>'];
    parts.push('<thead><tr><th class="corner"></th>');
    for (let c = 0; c < nCols; c++) parts.push(`<th data-c="${c}">${colName(c)}</th>`);
    parts.push('</tr></thead><tbody>');
    for (let r = 0; r < nRows; r++) {
      parts.push(`<tr><th data-r="${r}">${r + 1}</th>`);
      for (let c = 0; c < nCols; c++) {
        const d = display(value(r, c));
        parts.push(`<td data-r="${r}" data-c="${c}" class="${d.cls}">${escapeHtml(d.text)}</td>`);
      }
      parts.push('</tr>');
    }
    parts.push('</tbody>');
    grid.innerHTML = parts.join('');
    grid.style.width = (44 + 96 * nCols) + 'px';
    renderTabs();
    select(sel.r, sel.c, false);
  }

  function escapeHtml(s) {
    return s.replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));
  }

  function td(r, c) { return grid.querySelector(`td[data-r="${r}"][data-c="${c}"]`); }

  function select(r, c, scroll = true) {
    r = Math.max(0, r); c = Math.max(0, c);
    if (r >= nRows || c >= nCols) { sel = { r, c }; render(); return; }
    grid.querySelectorAll('.sel, .hl').forEach((el) => el.classList.remove('sel', 'hl'));
    sel = { r, c };
    const cell = td(r, c);
    if (cell) {
      cell.classList.add('sel');
      if (scroll) cell.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }
    const ch = grid.querySelector(`thead th[data-c="${c}"]`); if (ch) ch.classList.add('hl');
    const rh = grid.querySelector(`tbody th[data-r="${r}"]`); if (rh) rh.classList.add('hl');
    nameBox.value = addr(r, c);
    const rv = raw(r, c);
    fInput.value = rv === undefined ? '' : typeof rv === 'boolean' ? (rv ? 'TRUE' : 'FALSE') : String(rv);
    showInfo();
  }

  function showInfo() {
    const v = value(sel.r, sel.c);
    if (v && typeof v === 'object' && v.error) {
      info.textContent = `${v.error} ${ERROR_TEXT[v.error] || ''}`;
      info.className = 'cell-info error';
    } else {
      const rv = raw(sel.r, sel.c);
      info.textContent = typeof rv === 'string' && rv.startsWith('=') ? '= ' + display(v).text : '';
      info.className = 'cell-info';
    }
  }

  function refreshCells() {
    for (const cell of grid.querySelectorAll('td')) {
      const d = display(value(+cell.dataset.r, +cell.dataset.c));
      if (cell.textContent !== d.text) cell.textContent = d.text;
      cell.classList.remove('num', 'bool', 'err');
      if (d.cls) cell.classList.add(d.cls);
    }
    showInfo();
  }

  function commit(r, c, text) {
    const a = addr(r, c);
    const v = typed(text);
    if (v === null) delete book.sheets[current][a]; else book.sheets[current][a] = v;
    recalc();
    refreshCells();
    select(sel.r, sel.c, false);
  }

  // ---- in-cell editing
  function startEdit(initial) {
    if (editing) return;
    const cell = td(sel.r, sel.c);
    if (!cell) return;
    const input = document.createElement('input');
    input.value = initial !== undefined ? initial : fInput.value;
    cell.textContent = '';
    cell.classList.add('editing');
    cell.appendChild(input);
    editing = { input, cell, r: sel.r, c: sel.c };
    input.focus();
    input.setSelectionRange(input.value.length, input.value.length);
    input.addEventListener('input', () => { fInput.value = input.value; });
    fInput.value = input.value;
    input.addEventListener('keydown', (e) => {
      e.stopPropagation(); // keep the grid's own key handling out of the editor
      if (e.key === 'Enter') { e.preventDefault(); endEdit(true); select(sel.r + (e.shiftKey ? -1 : 1), sel.c); }
      else if (e.key === 'Tab') { e.preventDefault(); endEdit(true); select(sel.r, sel.c + (e.shiftKey ? -1 : 1)); }
      else if (e.key === 'Escape') { e.preventDefault(); endEdit(false); }
    });
    input.addEventListener('blur', () => { if (editing && editing.input === input) endEdit(true); });
  }

  function endEdit(save) {
    if (!editing) return;
    const { input, cell, r, c } = editing;
    editing = null;
    cell.classList.remove('editing');
    if (input.parentNode) input.remove();
    if (save) commit(r, c, input.value);
    else { const d = display(value(r, c)); cell.textContent = d.text; select(r, c, false); }
    wrap.focus();
  }

  // ---- events
  grid.addEventListener('mousedown', (e) => {
    const cell = e.target.closest('td');
    if (!cell || cell.classList.contains('editing')) return;
    if (editing) endEdit(true);
    select(+cell.dataset.r, +cell.dataset.c);
    wrap.focus();
    e.preventDefault();
  });
  grid.addEventListener('dblclick', (e) => {
    const cell = e.target.closest('td');
    if (cell) startEdit();
  });

  wrap.addEventListener('keydown', (e) => {
    if (editing) return;
    const moves = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (moves[e.key]) { e.preventDefault(); select(sel.r + moves[e.key][0], sel.c + moves[e.key][1]); return; }
    if (e.key === 'Enter') { e.preventDefault(); if (e.shiftKey) select(sel.r - 1, sel.c); else startEdit(); return; }
    if (e.key === 'F2') { e.preventDefault(); startEdit(); return; }
    if (e.key === 'Tab') { e.preventDefault(); select(sel.r, sel.c + (e.shiftKey ? -1 : 1)); return; }
    if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); commit(sel.r, sel.c, ''); return; }
    if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) { e.preventDefault(); startEdit(e.key); }
  });

  fInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); commit(sel.r, sel.c, fInput.value); select(sel.r + 1, sel.c); wrap.focus(); }
    else if (e.key === 'Escape') { e.preventDefault(); select(sel.r, sel.c, false); wrap.focus(); }
  });
  fInput.addEventListener('focus', () => { if (editing) endEdit(true); });

  nameBox.addEventListener('keydown', (e) => {
    if (e.key !== 'Enter') return;
    e.preventDefault();
    const p = parseAddr(nameBox.value.trim().replace(/\$/g, ''));
    if (p) { select(p.r, p.c); wrap.focus(); } else nameBox.value = addr(sel.r, sel.c);
  });

  // ---- sheets
  function renderTabs() {
    const tabs = $('sheet-tabs');
    tabs.innerHTML = '';
    for (const name of Object.keys(book.sheets)) {
      const b = document.createElement('button');
      b.textContent = name;
      if (name === current) b.className = 'active';
      b.addEventListener('click', () => { if (editing) endEdit(true); current = name; sel = { r: 0, c: 0 }; render(); wrap.focus(); });
      b.addEventListener('dblclick', () => renameSheet(name));
      tabs.appendChild(b);
    }
  }
  function renameSheet(name) {
    const nn = prompt('Sheet name', name);
    if (!nn || nn === name || book.sheets[nn]) return;
    const sheets = {};
    for (const [k, v] of Object.entries(book.sheets)) sheets[k === name ? nn : k] = v;
    book.sheets = sheets;
    if (current === name) current = nn;
    recalc(); render();
  }
  $('btn-add-sheet').addEventListener('click', () => {
    let i = Object.keys(book.sheets).length + 1;
    while (book.sheets['Sheet' + i]) i++;
    book.sheets['Sheet' + i] = {};
    current = 'Sheet' + i;
    sel = { r: 0, c: 0 };
    recalc(); render(); wrap.focus();
  });

  // ---- open / save
  function load(newBook, sheet) {
    book = newBook;
    current = sheet && book.sheets[sheet] ? sheet : Object.keys(book.sheets)[0];
    sel = { r: 0, c: 0 };
    recalc(); render(); wrap.focus();
  }
  $('btn-new').addEventListener('click', () => load({ sheets: { Sheet1: {} } }));
  $('file-open').addEventListener('change', async (e) => {
    const f = e.target.files[0];
    e.target.value = '';
    if (!f) return;
    const text = await f.text();
    try {
      if (/\.json$/i.test(f.name) || /^\s*\{/.test(text)) {
        const data = JSON.parse(text);
        if (!data || typeof data.sheets !== 'object' || !Object.keys(data.sheets).length) throw new Error('no "sheets" object');
        load(data);
      } else {
        const name = f.name.replace(/\.[^.]*$/, '').replace(/[^\w ]/g, '_') || 'Sheet1';
        load({ sheets: { [name]: parseCSV(text) } });
      }
      status(`opened ${f.name}`);
    } catch (err) {
      alert(`Could not open ${f.name}: ${err.message}`);
    }
  });
  function download(name, text, type) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([text], { type }));
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 0);
  }
  $('btn-save-json').addEventListener('click', () => download('workbook.json', JSON.stringify(book, null, 1) + '\n', 'application/json'));
  $('btn-save-csv').addEventListener('click', () => download(current + '.csv', toCSV(computed.sheets[current] || {}), 'text/csv'));

  window.__sheetDebug = () => ({ book, sel, current }); // used by the UI smoke test
  // demo content so the grid is not empty on first load
  book.sheets.Sheet1 = {
    A1: 'Item', B1: 'Qty', C1: 'Price', D1: 'Total',
    A2: 'apples', B2: 3, C2: 0.5, D2: '=B2*C2',
    A3: 'pears', B3: 2, C3: 0.75, D3: '=B3*C3',
    A4: 'plums', B4: 10, C4: 0.2, D4: '=B4*C4',
    A6: 'Sum', D6: '=SUM(D2:D4)', A7: 'Average', D7: '=ROUND(AVERAGE(D2:D4),2)',
  };
  recalc();
  render();
  wrap.focus();
})();
