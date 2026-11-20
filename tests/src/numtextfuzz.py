# Random number-like strings converted by arithmetic ("..."+0): signs, currency,
# percent, brackets, separators, exponents, fractions, times, dates.
# Run: python3 tests/src/numtextfuzz.py <seed> <count> > tests/src/numtextfuzzN.txt; then formulas2cases + lo_oracle
import random, sys
random.seed(int(sys.argv[1]) if len(sys.argv) > 1 else 1)
N = int(sys.argv[2]) if len(sys.argv) > 2 else 500
toks = ["1", "12", "123", "1234", "0", "05", "7", "59", "60", "99", "2024", " ", " ", ".", ",", "-", "+", "$", "%", "(", ")", ":", "/", "e", "E", "e+", "e-", " AM", " PM", "am", "1,234", "1.5", ".5", "1/2", "0:30", "12:00", "-", "Jan", " 1,"]
seen = set()
while len(seen) < N:
    s = "".join(random.choice(toks) for _ in range(random.randint(1, 5)))
    if '"' in s or s in seen or "/" in s and not any(c.isdigit() for c in s.split("/")[-1][:1] or "x"):
        pass
    # leave out day-month forms without a year (they depend on the current year)
    import re
    if re.search(r"\d/\d", s) and not re.search(r"\d+/\d+/\d+", s) and not re.search(r"\d +\d+/\d+$", s.strip()):
        continue
    if "Jan" in s and not re.search(r"\d{3,4}", s):
        continue
    seen.add(s)
for s in sorted(seen):
    print(f'="{s}"+0')
