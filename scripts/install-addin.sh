#!/usr/bin/env bash
# Sideload local/manifest.xml into Word, Excel and PowerPoint for Mac by copying it into
# each app's wef folder as <add-in id>.manifest.xml, and drop Office's cached copy of the
# manifest so the new one is read. Run with the Office apps closed.
set -euo pipefail
cd "$(dirname "$0")/.."
MANIFEST=local/manifest.xml
[ -f "$MANIFEST" ] || { echo "missing $MANIFEST: run 'npm run manifest' first" >&2; exit 1; }
ID="$(grep -oE '<Id>[^<]+</Id>' "$MANIFEST" | head -1 | sed -E 's#</?Id>##g')"
[ -n "$ID" ] || { echo "could not read <Id> from $MANIFEST" >&2; exit 1; }

# Office caches manifests by id + version under Wef/.../Manifests and won't re-read ours
# while the version is unchanged, so the cached copy (old URL and token) must go. A running
# app can write it back, hence the check.
running=()
for proc in "Microsoft Word" "Microsoft Excel" "Microsoft PowerPoint"; do
  pgrep -x "$proc" >/dev/null && running+=("$proc")
done
if [ ${#running[@]} -gt 0 ]; then
  echo "Quit these first (Cmd+Q), then re-run: ${running[*]}" >&2
  exit 1
fi

for app in Word Excel Powerpoint; do
  container="$HOME/Library/Containers/com.microsoft.$app"
  [ -d "$container" ] || { echo "  skip $app (not installed)"; continue; }
  dir="$container/Data/Documents/wef"
  mkdir -p "$dir"
  cp "$MANIFEST" "$dir/$ID.manifest.xml"
  chmod 600 "$dir/$ID.manifest.xml"
  echo "  installed $dir/$ID.manifest.xml"
  cache="$container/Data/Library/Application Support/Microsoft/Office/16.0/Wef"
  [ -d "$cache" ] || continue
  find "$cache" -path "*/Manifests/${ID}_*" -type f -print -delete | sed 's/^/  cleared cache /'
done
echo "Open Word, then Insert > My Add-ins > Claude (developer/shared add-in)."
