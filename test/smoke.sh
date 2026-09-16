#!/bin/bash
# Headless smoke test of both shells against test/page.html. It cannot reach
# OWA, so it proves only that the seams work: settings are read, the cache is
# written, a palette reaches the theme engine, and a plain tab is left alone.
# Nothing here opens a window on the desktop.
#
# Prints one line per case; exits non-zero if any expectation fails.

set -euo pipefail
cd "$(dirname "$0")/.."
./build.sh >/dev/null

BROWSER=${BROWSER:-chromium}
WORK=$(mktemp -d /tmp/owa-smoke.XXXXXX)
trap 'rm -rf "$WORK"' EXIT
HERE=$(pwd)

run() {  # run <profile> <url> [chromium flags...]
  local prof=$1 url=$2; shift 2
  timeout 60 "$BROWSER" --headless=new --no-sandbox --disable-gpu \
    --user-data-dir="$WORK/$prof" --allow-file-access-from-files \
    --virtual-time-budget=8000 "$@" --dump-dom "$url" 2>/dev/null \
    | grep -o 'RESULT:[^<]*' | head -1 | sed 's/^RESULT://' | sed 's/&quot;/"/g'
}

expect() {  # expect <name> <json> <python expression over d>
  local name=$1 json=$2 cond=$3
  if [ -z "$json" ]; then echo "FAIL $name: no result"; FAILED=1; return; fi
  if python3 -c "import json,sys; d=json.loads(sys.argv[1]); sys.exit(0 if ($cond) else 1)" "$json"; then
    echo "ok   $name"
  else
    echo "FAIL $name: $json"; FAILED=1
  fi
}
FAILED=0

# --- userscript: the flag opts a tab in, the feed themes it ---
sed "s|<body>|<body><script src=\"$HERE/test/gm-shim.js\"></script><script src=\"$HERE/foresight.user.js\"></script>|" \
  test/page.html > "$WORK/userscript.html"
R=$(run us "file://$WORK/userscript.html?omarchy=1")
expect "userscript, flagged tab: reduced and themed" "$R" \
  "d['minimal'] and d['bar'] and d['style'] and d['themeLen']>0 and d['literalsLen']>0 and not d['errors']"
R=$(run us2 "file://$WORK/userscript.html")
expect "userscript, plain tab: untouched" "$R" "not d['style'] and not d['errors']"

# --- extension: a test copy that also matches file:// and picks a palette ---
cp -r extension "$WORK/ext"
python3 - "$WORK/ext" <<'PY'
import json, sys
root = sys.argv[1]
m = json.load(open(f'{root}/manifest.json'))
m['content_scripts'][0]['matches'] = ['<all_urls>']
m['web_accessible_resources'][0]['matches'] = ['<all_urls>']
m['host_permissions'].append('<all_urls>')
json.dump(m, open(f'{root}/manifest.json', 'w'))
# force a bundled palette in the shell itself: writing it to storage from
# the install hook races the shell's read under virtual time
c = open(f'{root}/content.js').read()
assert "const settings = sync || {};" in c
c = c.replace("const settings = sync || {};", "const settings = Object.assign({ themeSource: 'tokyo-night' }, sync || {});")
open(f'{root}/content.js', 'w').write(c)
PY
R=$(run ext "file://$HERE/test/page.html?omarchy=1" --load-extension="$WORK/ext")
expect "extension, flagged tab: reduced, bundled palette applied" "$R" \
  "d['minimal'] and d['bar'] and d['style'] and d['themeLen']>0 and not d['errors']"
R=$(run ext2 "file://$HERE/test/page.html" --load-extension="$WORK/ext")
expect "extension, plain tab: untouched" "$R" "not d['style'] and not d['errors']"

exit $FAILED
