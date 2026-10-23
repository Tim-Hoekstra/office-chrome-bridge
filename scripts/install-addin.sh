#!/usr/bin/env bash
# Sideload local/manifest.xml into Word, Excel and PowerPoint for Mac by copying it into
# each app's wef folder as <add-in id>.manifest.xml. Quit and reopen the apps afterwards.
set -euo pipefail
cd "$(dirname "$0")/.."
MANIFEST=local/manifest.xml
[ -f "$MANIFEST" ] || { echo "missing $MANIFEST: run 'npm run manifest' first" >&2; exit 1; }
ID="$(grep -oE '<Id>[^<]+</Id>' "$MANIFEST" | head -1 | sed -E 's#</?Id>##g')"
[ -n "$ID" ] || { echo "could not read <Id> from $MANIFEST" >&2; exit 1; }

for app in Word Excel Powerpoint; do
  container="$HOME/Library/Containers/com.microsoft.$app"
  [ -d "$container" ] || { echo "  skip $app (not installed)"; continue; }
  dir="$container/Data/Documents/wef"
  mkdir -p "$dir"
  cp "$MANIFEST" "$dir/$ID.manifest.xml"
  chmod 600 "$dir/$ID.manifest.xml"
  echo "  installed $dir/$ID.manifest.xml"
done
echo "Quit (Cmd+Q) and reopen Word, then Insert > My Add-ins > Claude (developer/shared add-in)."
