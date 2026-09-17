#!/bin/bash
set -e
cd "$(dirname "$0")"
node -e "
const fs=require('fs');
const out=fs.readFileSync('../index.html','utf8');
const app=/<script id=\"app-script\">([\s\S]*?)<\/script>/.exec(out);
fs.writeFileSync('/tmp/pwa-app.js',app[1]);
"
export SHEETJS=$PWD/../lib/xlsx.full.min.js
export OOXMLJS=$PWD/../lib/ooxml.js
export APPJS=/tmp/pwa-app.js
echo "ooxml:"; node ooxml-test.js | tail -1
echo "spreadsheet PASS:"; node spreadsheet-test.js | grep -c "PASS -"
echo "spreadsheet FAIL:"; node spreadsheet-test.js | grep -c "FAIL" || true
echo "fidelity:"; node fidelity-test.js | tail -1
