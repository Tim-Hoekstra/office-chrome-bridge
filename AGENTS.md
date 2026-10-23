# Agent instructions: office-chrome-bridge

Localhost-only bridge that exposes Claude Code's Claude in Chrome MCP server
(`claude --claude-in-chrome-mcp`, stdio) to the Claude for Office add-ins as a Streamable HTTP
MCP server at `https://localhost:47823/mcp`. See `README.md` for setup and the threat model.

## Before every commit and push: run the secret check

```bash
npm run check-secrets        # scripts/check-secrets.sh, exits 1 on any finding
```

- Run it **before** `git commit` and again **before** `git push`. Push only if it prints
  `check-secrets: clean` and exits 0. Gate on the exit code (`npm run check-secrets && git push`),
  not on reading the output.
- If it fails, stop and tell the user. If the secret was already pushed, the token must be rotated
  (`npm run rotate`) and history rewritten; rotating alone is not enough.
- Don't hand-roll a replacement check. Two past mistakes this script exists to prevent:
  - `cmd | head -1 && echo LEAK` reports a leak even when nothing matched, because `head` succeeds
    on empty input. A false alarm like that got ignored, and the push went ahead.
  - With `set -o pipefail`, `git log -p | grep -q ...` misses real matches: `grep -q` exits early,
    `git log` dies of SIGPIPE and the pipeline counts as failed. The script uses process
    substitution (`grep -q ... < <(git log --all -p)`) for this reason.

## Secrets: never print, log or commit them

Everything secret lives in `local/` (git-ignored): `.token`, `certs/localhost.key`, `manifest.xml`
(embeds the token), `bridge.log`.

- Never `cat` or echo `local/.token` or `local/manifest.xml`, and don't print anything derived from
  the token. Tool output ends up in chat transcripts. Read it into a shell variable for tests
  (`T=$(cat local/.token)`) and print only status codes or truncated responses.
- `scripts/build-manifest.mjs` and `scripts/rotate-token.sh` deliberately don't print the token.
  Keep it that way.
- Office keeps its own copies of the manifest (with the token) in
  `~/Library/Containers/com.microsoft.*/Data/Library/Application Support/Microsoft/Office/16.0/Wef/.../Manifests/`.
  Don't print those either.

## Testing changes

No dependencies, no test suite. After editing `src/bridge.mjs`:

```bash
node --check src/bridge.mjs
pkill -f "node src/bridge.mjs"; (npm start >> local/bridge.log 2>&1 &)
T=$(cat local/.token)
c(){ curl -s --cacert local/certs/localhost.crt -X POST https://localhost:47823/mcp \
  -H "Origin: https://pivot.claude.ai" -H "Content-Type: application/json" "$@"; }
c -o /dev/null -w '%{http_code}\n' -d '{}'                     # expect 401 (no token)
c -H "Authorization: Bearer $T" -d '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-06-18","capabilities":{},"clientInfo":{"name":"curl","version":"0"}}}'
c -H "Authorization: Bearer $T" -d '{"jsonrpc":"2.0","id":2,"method":"tools/call","params":{"name":"tabs_context_mcp","arguments":{}}}'
```

`local/bridge.log` logs every request (method, path, origin, preflight headers, tool calls). Check
it first when Office reports a connector failure.

## Things that cost time before

- **Office reads the manifest only when it starts.** After `npm run install-addin` or
  `npm run rotate`, Word must be fully quit (Cmd+Q) and reopened. Until then it sends the old
  URL/token (`401`, or `404` for a removed path). Office rewrites its `Wef/.../Manifests` cache
  on every start; there is no need to clear it.
- **The add-in connects from the taskpane on this computer**, not from Anthropic's cloud, so
  `https://localhost` works. It needs the trusted cert and CORS for `https://pivot.claude.ai`.
  The preflight asks for headers such as `x-mcp-client-name` and `x-mcp-client-session-id`; the
  bridge echoes the requested headers. Don't replace that with a fixed list.
- **Several Chrome browsers can be connected to one Claude account**, including other computers
  through Anthropic's relay. The bridge pins the session to this computer's browser
  (`ensureLocalBrowser`). Without it, tabs and Allow prompts appear on the other machine.
- **The token authenticates, not the path.** `/mcp/<anything>` is accepted on purpose so manifests
  from older builds keep working.
- The permission guard refusing `skip_all_permission_checks` is intentional; don't remove it.
- The default tool set is `all`, because the add-in asks the user to approve each MCP tool call.

## Conventions

- Plain Node (>=20) ES modules, no npm dependencies. Shell scripts are bash with `set -euo pipefail`,
  except `check-secrets.sh` (see above). macOS paths for Office and the keychain.
- Generated or secret files go in `local/`, never at the repo root.
- `manifest/template.xml` is Anthropic's add-in manifest. Edit it only via
  `node scripts/build-manifest.mjs --fetch`.
- The repo is private on GitHub (`Tim-Hoekstra/office-chrome-bridge`). Commit and push only when
  the user asks, after `npm run check-secrets` passes.
