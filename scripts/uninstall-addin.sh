#!/usr/bin/env bash
# Remove the sideloaded manifest from Word, Excel and PowerPoint for Mac.
set -euo pipefail
cd "$(dirname "$0")/.."
ID="$(grep -oE '<Id>[^<]+</Id>' manifest/template.xml | head -1 | sed -E 's#</?Id>##g')"
for app in Word Excel Powerpoint; do
  f="$HOME/Library/Containers/com.microsoft.$app/Data/Documents/wef/$ID.manifest.xml"
  if [ -f "$f" ]; then rm "$f"; echo "  removed $f"; fi
done
echo "Quit and reopen the Office apps. To also untrust the cert:"
echo "  security delete-certificate -c 'localhost (office-chrome-bridge)' ~/Library/Keychains/login.keychain-db"
