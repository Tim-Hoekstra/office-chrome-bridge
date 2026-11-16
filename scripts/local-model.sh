#!/usr/bin/env bash
# Turn the local model gateway on or off. Works before the first install and on an existing one.
#
#   npm run local-model localhost:11434    # Ollama (LM Studio: localhost:1234)
#   npm run local-model off
#
# The address is saved in local/llm-url, which the bridge and the manifest builder both read, so
# it only has to be given once. If the add-in is already installed, its manifest is rebuilt and
# reinstalled.
set -euo pipefail
cd "$(dirname "$0")/.."
ADDR="${1:-}"
[ -n "$ADDR" ] || { echo "usage: npm run local-model <host:port | url | off>" >&2; exit 1; }
mkdir -p local
chmod 700 local

if [ "$ADDR" = off ]; then
  rm -f local/llm-url
  echo "local model gateway off"
else
  printf '%s\n' "$ADDR" > local/llm-url
  echo "local model gateway set to $ADDR"
fi

if [ -f local/.token ]; then
  node scripts/build-manifest.mjs
  scripts/install-addin.sh
  if pgrep -f "node src/bridge.mjs" >/dev/null; then
    echo "NOTE: restart the bridge (Ctrl+C, then npm start) so it picks up the change."
  fi
else
  echo "Saved. Continue with the install steps in README.md (npm start, npm run manifest, npm run install-addin)."
fi
