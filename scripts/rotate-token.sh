#!/usr/bin/env bash
# Replace the bridge token, rebuild the manifest and reinstall it into Office.
# The running bridge still holds the old token: restart it, then Cmd+Q and reopen Office.
set -euo pipefail
cd "$(dirname "$0")/.."
mkdir -p local
chmod 700 local
node -e 'require("fs").writeFileSync("local/.token", require("crypto").randomBytes(32).toString("base64url") + "\n", { mode: 0o600 })'
chmod 600 local/.token
echo "new token written to local/.token"
node scripts/build-manifest.mjs
scripts/install-addin.sh
if pgrep -f "node src/bridge.mjs" >/dev/null; then
  echo "NOTE: the bridge is running with the OLD token. Restart it (Ctrl+C, then npm start)."
fi
