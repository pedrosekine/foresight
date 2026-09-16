#!/bin/bash
# Builds the two installable forms from one source.
#
#   foresight.user.js        header + extension/core.js + userscript/env.js
#   dist/foresight-<v>.zip   extension/ packed for the Chrome Web Store
#
# The version comes from extension/manifest.json and is stamped into the
# userscript header, so bumping it in one place bumps both.

set -euo pipefail
cd "$(dirname "$0")"

VERSION=$(python3 -c 'import json;print(json.load(open("extension/manifest.json"))["version"])')

for f in extension/core.js extension/content.js extension/background.js extension/options.js userscript/env.js; do
  node --check "$f"
done

{
  sed "s/@VERSION@/$VERSION/" userscript/header.js
  echo
  cat extension/core.js
  echo
  cat userscript/env.js
} > foresight.user.js
node --check foresight.user.js

mkdir -p dist
rm -f "dist/foresight-$VERSION.zip"
(cd extension && zip -qr "../dist/foresight-$VERSION.zip" . -x 'icons/icon.svg')

echo "foresight.user.js            $(wc -c < foresight.user.js) bytes, v$VERSION"
echo "dist/foresight-$VERSION.zip  $(wc -c < "dist/foresight-$VERSION.zip") bytes"
