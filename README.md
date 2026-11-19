# A spreadsheet engine (experiment)

A public demo for an experiment: the same spec built by the same AI harness, once with cooknship and once without.

The only spec is [SPEC.md](SPEC.md).

## Use

Requires Node.js (>= 18). No dependencies, no build step.

- Engine CLI: `bin/engine book.json` or `bin/engine - < book.json`; `bin/engine - --cell Sheet1!H1` prints one cell.
- UI: open `ui/index.html` in a browser (no server needed). Click or arrow keys to select, type or Enter/F2/double-click to edit, Enter/Tab to commit, Esc to cancel, Delete to clear. The formula bar shows the selected cell's formula and, for errors, what the error means. Open reads `.json` workbooks and `.csv` files; Save JSON keeps formulas, Save CSV writes the current sheet's computed values. Double-click a sheet tab to rename it.
- Tests: `./test.sh` runs the whole suite: engine fixtures whose expected values were computed by LibreOffice, CLI checks, CSV checks, and a UI smoke test in headless Chromium (skipped if no Chromium is installed).

## Layout

- `src/engine.js` — tokenizer, parser, evaluator and the 40 functions (runs in Node and in the browser).
- `bin/engine` — the CLI.
- `src/csv.js` — CSV import/export; `ui/` — the browser UI.
- `tests/src/*.txt` — formula lists; `tests/cases/*.json` the workbooks built from them; `tests/expected/*.json` LibreOffice's results.
- `tools/lo_oracle.py`, `tools/probe.sh` — dev-only: regenerate expected values with LibreOffice headless.
