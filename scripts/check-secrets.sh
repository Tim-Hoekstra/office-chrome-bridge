#!/usr/bin/env bash
# Fail (exit 1) if a secret is tracked, staged or anywhere in git history.
# Run before every commit and push:  npm run check-secrets
#
# Checks: files under local/ or key/token/cert files tracked; the current bridge token
# (local/.token) in the index or any commit; PEM private keys; long Bearer tokens.
# no pipefail: grep -q exits early and SIGPIPEs git log, which pipefail would treat as "no match".
# History checks use process substitution for the same reason.
set -u
cd "$(dirname "$0")/.."

fail=0
report() { echo "LEAK: $*" >&2; fail=1; }

# 1. forbidden paths tracked or staged
bad_paths="$(git ls-files --cached | grep -E '^local/|(^|/)\.token$|\.(key|pem|p12|crt)$|(^|/)manifest\.xml$|(^|/)\.env' || true)"
[ -n "$bad_paths" ] && report "tracked secret files:"$'\n'"$bad_paths"

# 2. the current token, in the index or in any commit on any ref
if [ -f local/.token ]; then
  token="$(tr -d '\n' < local/.token)"
  if [ -n "$token" ]; then
    git grep -q --cached -F "$token" -- . && report "current token is in staged/tracked files"
    grep -q -F "$token" < <(git log --all -p) && report "current token is in git history"
  fi
fi

# 3. generic patterns (PEM private keys, long Bearer tokens), in the index and in history
pattern='BEGIN ([A-Z]+ )?PRIVATE KEY|Bearer[ +%20]+[A-Za-z0-9_-]{30,}'
git grep -q --cached -I -E "$pattern" -- . && report "private key or Bearer token in staged/tracked files"
grep -q -E "$pattern" < <(git log --all -p) && report "private key or Bearer token in git history"

if [ "$fail" -eq 0 ]; then
  echo "check-secrets: clean"
else
  echo "check-secrets: FAILED. Do not commit or push. If it is already pushed, rotate (npm run rotate) and rewrite history." >&2
fi
exit "$fail"
