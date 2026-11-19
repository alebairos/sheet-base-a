#!/bin/sh
# Dev helper: tools/probe.sh suite  (reads tests/src/<suite>.txt, regenerates cases + expected)
set -e
cd "$(dirname "$0")/.."
python3 tools/formulas2cases.py "tests/src/$1.txt" > "tests/cases/$1.json"
python3 tools/lo_oracle.py "tests/cases/$1.json" > "tests/expected/$1.json"
node tests/run.js "$1"
