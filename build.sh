#!/usr/bin/env bash
# Build a store-ready zip for the Microsoft Edge Add-ons store (also valid for the Chrome Web Store).
# Strips the "key" field (stores reject it and assign their own ID) and excludes dev files.
set -euo pipefail
cd "$(dirname "$0")"
VERSION=$(python3 -c "import json;print(json.load(open('manifest.json'))['version'])")
OUT="dist/claude-newtab-$VERSION.zip"
STAGE=$(mktemp -d)
trap 'rm -rf "$STAGE"' EXIT
SW=$(python3 -c "import json;print(json.load(open('manifest.json'))['background']['service_worker'])")
cp manifest.json "$SW" ics.js newtab.html newtab.css newtab.js options.html options.js icon.png "$STAGE/"
python3 - "$STAGE/manifest.json" <<'PY'
import json, sys
p = sys.argv[1]; m = json.load(open(p)); m.pop("key", None)
json.dump(m, open(p, "w"), indent=2)
PY
mkdir -p dist
rm -f "$OUT"
(cd "$STAGE" && zip -qr "$OLDPWD/$OUT" .)
echo "built $OUT"
unzip -l "$OUT" | tail -n +4 | head -n -2 | awk '{print "  " $4}'
