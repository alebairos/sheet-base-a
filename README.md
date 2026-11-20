# A spreadsheet engine (experiment)

A public demo for an experiment: the same spec built by the same AI harness, once with cooknship and once without.

The only spec is [SPEC.md](SPEC.md).

## Use

Requires Node.js (>= 18). No dependencies, no build step.

- Engine CLI: `bin/engine book.json` or `bin/engine - < book.json`; `bin/engine - --cell Sheet1!H1` prints one cell.
- UI: open `ui/index.html` in a browser (no server needed). Click or arrow keys to select, type or Enter/F2/double-click to edit, Enter/Tab to commit, Esc to cancel, Delete to clear. The formula bar shows the selected cell's formula and, for errors, what the error means. Open reads `.json` workbooks and `.csv` files; Save JSON keeps formulas, Save CSV writes the current sheet's computed values. Double-click a sheet tab to rename it.
- Tests: `./test.sh` runs the whole suite: engine fixtures whose expected values were computed by LibreOffice, CLI checks, CSV checks, and a UI smoke test in headless Chromium (skipped if no Chromium is installed).

## Layout

- `src/engine.js` — tokenizer, parser, evaluator and the functions (runs in Node and in the browser): the spec's 40, plus PI, COUNTBLANK, PROPER, EXACT, VALUE, FIND, SEARCH, ISNA, ISERR, ISLOGICAL, ISNONTEXT, N, T, SIGN, TRUNC, REPT, SUBSTITUTE, CHAR, CODE, CHOOSE, ROWS, COLUMNS, SUMPRODUCT, COUNTIFS, SUMIFS, AVERAGEIFS, HLOOKUP, EXP, LN, LOG, LOG10, FACT, ISODD, ISEVEN, EVEN, ODD, QUOTIENT, CEILING, FLOOR, VAR, VARP, STDEVP, LARGE, SMALL, RAND, RANDBETWEEN, TEXT (number formats, no dates or fractions). Unknown functions give `#NAME?` after their arguments are computed (a computed error argument wins), as in Calc.
- `bin/engine` — the CLI.
- `src/csv.js` — CSV import/export; `ui/` — the browser UI.
- `tests/src/*.txt` — formula lists; `tests/cases/*.json` the workbooks built from them; `tests/expected/*.json` LibreOffice's results.
- `tools/lo_oracle.py`, `tools/probe.sh`, `tools/fuzz*.{js,py}` — dev-only: regenerate expected values with LibreOffice headless.

## Semantics notes

- Expected values come from LibreOffice recalculating the workbook and writing it out (the oracle in `tools/lo_oracle.py` round-trips through xlsx). In that output Calc's internal errors without an Excel name — `Err:502` (invalid argument, e.g. `SQRT(-1)`, `MID("abc",0,1)`) and `Err:522` (circular reference) — appear as `#VALUE!`, and the engine reports them the same way.
- Numbers in the output are rounded to 15 significant digits, as Calc stores them in files; integers below 2^53 are written exactly.
- Arithmetic follows Calc's rounding: `+`/`-` snap results that cancel to within 2^-48 to 0 (but never between exact integers), and SUM/AVERAGE use Calc's Kahan summation (parameters taken last to first, each range column summed on its own, blocks of 8 values in SSE2 lanes, the last term added with the same snapping), so `SUM(0.1,0.2,-0.3)` is 0. Inline and computed arrays are searched with exact number equality, cell ranges with approximate equality.
- An empty argument (`SUM(1,)`, `LEFT("abc",)`) is 0 as in Calc; references can be intersected (`A1:B3 B2`, a disjoint intersection is `#NULL!`) and joined (`SUM((A1:A2,B1))`).
- Text functions count Unicode code points; `CHAR`/`CODE` use UTF-8 like Calc on Linux (`CHAR(233)` is U+FFFD).
- Booleans are numbers with a boolean format, as in Calc: `TRUE+1` is 2, `TRUE&""` is `"1"`, and a result is reported as `true`/`false` when Calc would format it as a boolean.
- An empty text constant (`""`) in the input is an empty cell.
- A lookup (MATCH, VLOOKUP, HLOOKUP) whose lookup value is a computed `#N/A` gives Calc's internal lookup error: shown as `#N/A`, but `ISNA` is false and `ISERR` true. A `#N/A` read from a cell stays a plain `#N/A`.
- SUMPRODUCT's arguments and the table argument of MATCH/VLOOKUP/HLOOKUP are evaluated as arrays (`SUMPRODUCT((A1:A9>1)*B1:B9)`); elsewhere a range in a scalar position is intersected with the formula's row or column, as in Calc.

## Known deviations

Approximate-match lookups (`VLOOKUP(..., TRUE)`, `HLOOKUP(..., TRUE)`, `MATCH` with type 1 or -1) agree with LibreOffice on sorted data. On unsorted data the engine models Calc's search (a binary search down a column, then a walk while cells satisfy the condition, numbers sorting before text); random unsorted mixed columns agree with LibreOffice about 97% of the time, rows somewhat less.
SUM over ranges mixing very large and tiny values can still differ from LibreOffice in the last digit (about 1 in 200 such random sums).
When a single function call meets several different errors at once (for example an error argument and an invalid argument), LibreOffice's choice of which error to report depends on interpreter internals; the engine follows the common cases (first error computed, then arguments examined last to first) but not every combination.
