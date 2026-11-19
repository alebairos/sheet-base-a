# A spreadsheet engine (experiment)

A public demo for an experiment: the same spec built by the same AI harness, once with cooknship and once without.

The only spec is [SPEC.md](SPEC.md).

## Use

Requires Node.js (>= 18). No dependencies, no build step.

- Engine CLI: `bin/engine book.json` or `bin/engine - < book.json`; `bin/engine - --cell Sheet1!H1` prints one cell.
- Tests: `./test.sh` runs the whole suite (engine fixtures whose expected values were computed by LibreOffice, plus CLI checks).

## Layout

- `src/engine.js` — tokenizer, parser, evaluator and the 40 functions (runs in Node and in the browser).
- `bin/engine` — the CLI.
- `tests/src/*.txt` — formula lists; `tests/cases/*.json` the workbooks built from them; `tests/expected/*.json` LibreOffice's results.
- `tools/lo_oracle.py`, `tools/probe.sh` — dev-only: regenerate expected values with LibreOffice headless.
