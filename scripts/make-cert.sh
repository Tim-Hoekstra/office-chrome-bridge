#!/usr/bin/env bash
# Generate a self-signed TLS cert for https://localhost (the Word taskpane's WebView
# rejects plain http and untrusted certs) and trust it in the login keychain.
#   scripts/make-cert.sh            # generate (if missing) + trust (prompts for password)
#   scripts/make-cert.sh --force    # regenerate even if a cert exists
set -euo pipefail
cd "$(dirname "$0")/.."
DIR=local/certs
mkdir -p "$DIR"
chmod 700 local "$DIR"

if [ -f "$DIR/localhost.crt" ] && [ "${1:-}" != "--force" ]; then
  echo "cert exists: $DIR/localhost.crt (use --force to regenerate)"
else
  openssl req -x509 -newkey rsa:2048 -nodes -days 397 \
    -keyout "$DIR/localhost.key" -out "$DIR/localhost.crt" \
    -subj "/CN=localhost (office-chrome-bridge)" \
    -addext "subjectAltName=DNS:localhost,IP:127.0.0.1" \
    -addext "basicConstraints=critical,CA:FALSE" \
    -addext "keyUsage=critical,digitalSignature,keyEncipherment" \
    -addext "extendedKeyUsage=serverAuth"
  chmod 600 "$DIR/localhost.key"
fi

echo "Trusting cert for TLS in your login keychain (macOS will ask for your password)..."
security add-trusted-cert -r trustRoot -p ssl -k ~/Library/Keychains/login.keychain-db "$PWD/$DIR/localhost.crt"
echo "done"
