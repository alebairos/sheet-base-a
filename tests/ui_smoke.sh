#!/bin/sh
# Runs tests/ui/smoke.html in headless Chromium/Chrome if one is installed; skips otherwise.
cd "$(dirname "$0")/.."
for b in chromium chromium-browser google-chrome-stable google-chrome; do
  if command -v "$b" >/dev/null 2>&1; then BROWSER=$b; break; fi
done
if [ -z "$BROWSER" ]; then echo "ui: skipped (no headless Chromium found)"; exit 0; fi
PROFILE=$(mktemp -d)
TITLE=$(timeout 90 "$BROWSER" --headless --disable-gpu --no-sandbox --user-data-dir="$PROFILE" \
  --allow-file-access-from-files --dump-dom "file://$PWD/tests/ui/smoke.html" 2>/dev/null \
  | sed -n 's:.*<title>\(.*\)</title>.*:\1:p' | head -n 1)
rm -rf "$PROFILE"
case "$TITLE" in
  PASS) echo "ui: smoke test passed" ;;
  *) echo "ui: smoke test FAILED: $TITLE"; exit 1 ;;
esac
