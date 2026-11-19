// UI smoke test: drives the real app with synthetic events; result goes to document.title.
(function () {
  const $ = (id) => document.getElementById(id);
  const wrap = $('grid-wrap'), fi = $('formula-input');
  const key = (el, k) => el.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true }));
  const cell = (r, c) => document.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
  const click = (r, c) => cell(r, c).dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
  const fails = [];
  const eq = (what, got, want) => { if (got !== want) fails.push(`${what}: got ${JSON.stringify(got)} want ${JSON.stringify(want)}`); };
  try {
    eq('demo total', cell(5, 3).textContent, '5');
    click(0, 5);
    eq('name box', $('name-box').value, 'F1');
    key(wrap, '=');
    const inp = cell(0, 5).querySelector('input');
    inp.value = '=D6*2';
    key(inp, 'Enter');
    eq('edited value', cell(0, 5).textContent, '10');
    eq('moved down', $('name-box').value, 'F2');
    fi.value = '=1/0';
    key(fi, 'Enter');
    eq('error shown', cell(1, 5).textContent, '#DIV/0!');
    key(wrap, 'ArrowUp');
    eq('formula bar formula', fi.value, '=1/0');
    eq('formula bar error', $('cell-info').textContent.startsWith('#DIV/0!'), true);
    key(wrap, 'ArrowUp');
    eq('formula bar F1', fi.value, '=D6*2');
    click(1, 2);
    fi.value = '2';
    key(fi, 'Enter');
    eq('recalc dependents', cell(5, 3).textContent, '9.5');
    eq('recalc chain', cell(0, 5).textContent, '19');
    click(0, 5);
    key(wrap, 'Delete');
    eq('deleted', cell(0, 5).textContent, '');
    key(wrap, 'x');
    cell(0, 5).querySelector('input').value = 'abc';
    key(cell(0, 5).querySelector('input'), 'Escape');
    eq('escape cancels', cell(0, 5).textContent, '');
    $('btn-add-sheet').click();
    eq('tabs', $('sheet-tabs').textContent, 'Sheet1Sheet2');
  } catch (e) {
    fails.push('exception ' + e.message);
  }
  document.title = fails.length ? 'FAIL ' + fails.join('; ') : 'PASS';
})();
