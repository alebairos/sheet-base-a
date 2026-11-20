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
  const finish = () => { document.title = fails.length ? 'FAIL ' + fails.join('; ') : 'PASS'; };
  // open and save go through the real file input and download link
  const until = async (cond) => { for (let i = 0; i < 200 && !cond(); i++) await new Promise((r) => setTimeout(r, 10)); };
  const open = async (name, text) => {
    const dt = new DataTransfer();
    dt.items.add(new File([text], name));
    const fo = $('file-open');
    fo.files = dt.files;
    fo.dispatchEvent(new Event('change'));
    await until(() => $('status').textContent === 'opened ' + name);
  };
  let saved = null;
  URL.createObjectURL = (blob) => { saved = blob; return 'blob:test'; };
  URL.revokeObjectURL = () => {};
  HTMLAnchorElement.prototype.click = function () {};
  const save = async (btn) => { saved = null; $(btn).click(); return saved ? saved.text() : null; };
  (async () => {
    try {
      await open('data.csv', 'name,qty\nx,"2"\ny,=B2*3\n');
      eq('csv tab', $('sheet-tabs').textContent, 'data');
      eq('csv text', cell(1, 0).textContent, 'x');
      eq('csv formula', cell(2, 1).textContent, '6');
      const csv = await save('btn-save-csv');
      eq('saved csv', csv && csv.replace(/\r/g, ''), 'name,qty\nx,2\ny,6\n');
      const json = JSON.parse(await save('btn-save-json'));
      eq('saved json formula', json.sheets.data.B3, '=B2*3');
      await open('book.json', JSON.stringify({ sheets: { S1: { A1: 4, A2: '=A1/0' }, S2: { A1: '=S1!A1+1' } } }));
      eq('json tabs', $('sheet-tabs').textContent, 'S1S2');
      eq('json error', cell(1, 0).textContent, '#DIV/0!');
      const json2 = JSON.parse(await save('btn-save-json'));
      eq('json round trip', JSON.stringify(json2), JSON.stringify({ sheets: { S1: { A1: 4, A2: '=A1/0' }, S2: { A1: '=S1!A1+1' } } }));
    } catch (e) {
      fails.push('exception ' + e.message);
    }
    finish();
  })();
})();
