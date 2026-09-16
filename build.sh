#!/bin/bash
# Builds the two installable forms from one source.
#
#   owa-minimal.user.js      header + extension/core.js + userscript/env.js
#   dist/owa-minimal-<v>.zip extension/ packed for the Chrome Web Store
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
} > owa-minimal.user.js
node --check owa-minimal.user.js

mkdir -p dist
rm -f "dist/owa-minimal-$VERSION.zip"
(cd extension && zip -qr "../dist/owa-minimal-$VERSION.zip" . -x 'icons/icon.svg')

echo "owa-minimal.user.js            $(wc -c < owa-minimal.user.js) bytes, v$VERSION"
echo "dist/owa-minimal-$VERSION.zip  $(wc -c < "dist/owa-minimal-$VERSION.zip") bytes"
