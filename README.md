# office-chrome-bridge

Use **Claude in Chrome** from the **Claude for Word** add-in (also Excel and PowerPoint).

The bridge runs `claude --claude-in-chrome-mcp` (Claude Code's built-in Claude in Chrome MCP
server, which speaks stdio) and serves it as a Streamable HTTP MCP endpoint on
`https://localhost:47823/mcp`. The add-in's `mcp_servers` setting points at that URL.

```
Word taskpane (https://pivot.claude.ai, WebKit)
   │  HTTPS + Bearer token, CORS-pinned origin        ← never leaves this computer
   ▼
src/bridge.mjs  127.0.0.1:47823
   │  stdio JSON-RPC
   ▼
claude --claude-in-chrome-mcp ──► Claude in Chrome extension ──► your Chrome
```

The add-in connects to `mcp_servers` URLs **directly from the taskpane on your machine**
(not from Anthropic's cloud), so nothing has to be exposed to the network.

## Requirements

- Microsoft 365 Word, Excel or PowerPoint (desktop)
- A Claude account with access to Claude for Word
- [Claude Code](https://code.claude.com) installed and signed in with the **same** Claude account (`claude` on `PATH`)
- Google Chrome with the [Claude extension](https://claude.ai/chrome), signed in with the same account
- Node.js 20+ and `openssl` (both preinstalled or via Homebrew)

No npm dependencies.

### Platform support

The bridge (`src/bridge.mjs`) and manifest builder are plain Node and run on any OS. The cert and
sideload helper scripts are **macOS only** (login keychain, `~/Library/Containers/.../wef`).
On Windows, do those two steps by hand (untested):

- **Trust the cert:** generate it with `openssl` as in `scripts/make-cert.sh`, then
  `certutil -user -addstore Root local\certs\localhost.crt`
- **Sideload the manifest:** add a string value under
  `HKCU\Software\Microsoft\Office\16.0\WEF\Developer` whose data is the full path to
  `local\manifest.xml`, then restart Office.

## Install

```bash
cd office-chrome-bridge

npm run cert          # 1. self-signed cert for localhost, trusted in your login keychain (asks for password)
npm start             # 2. start the bridge; first run generates local/.token (leave running, Ctrl+C stops)

# in a second terminal:
npm run manifest      # 3. build local/manifest.xml with the bridge URL + token
npm run install-addin # 4. sideload it into Word / Excel / PowerPoint
```

Then:

1. Quit Word completely (**Cmd+Q**) and reopen it.
2. **Insert → My Add-ins**, pick **Claude** under the developer / shared add-ins (not the AppSource entry).
3. If you land on the enterprise connection screen, press **Back** and sign in with your Claude account.
4. Ask: *"open example.com in Chrome"*. The first time, Chrome shows a Claude permission prompt for the
   site; click **Allow**.

Claude in Chrome only sees tabs in its own Chrome tab group, so "list my tabs" is empty until it
has opened one. Drag a tab into the group to share it.

## Configuration

Environment variables for `npm start`:

| Variable | Default | |
|---|---|---|
| `BRIDGE_PORT` | `47823` | Rebuild + reinstall the manifest if you change it |
| `BRIDGE_TOOLS` | `all` | Which tools Office can use: `all`, `read`, or a comma-separated list (see below) |
| `BRIDGE_BROWSER` | this computer's Chrome | Extension `deviceId` to drive (see below) |
| `BRIDGE_EVIDENCE_DIR` | `~/evidence` | The only folder `save_screenshot` writes to and `list_evidence`/`view_evidence` read from |
| `BRIDGE_ORIGINS` | `https://pivot.claude.ai` | Allowed browser origins |
| `BRIDGE_TOKEN` | `local/.token` | Override the bearer token |
| `CLAUDE_CHROME_PERMISSION_MODE` | unset (extension default: ask) | Passed through to Claude in Chrome. `skip_all_permission_checks` is refused |
| `BRIDGE_ALLOW_SKIP_PERMISSIONS` | unset | `1` permits `skip_all_permission_checks` (not recommended) |
| `CLAUDE_BIN` | `claude` | Path to the Claude Code binary |

### Tool exposure

`all` (the default) exposes every Claude in Chrome tool. Claude for Office asks you to approve
each MCP tool call, and the extension asks before acting on a new site.

To narrow it, `read` lets Claude open, read and move between pages but not act on them:
`tabs_context_mcp`, `tabs_create_mcp`, `tabs_close_mcp`, `navigate`, `get_page_text`, `read_page`,
`find`, `read_console_messages`, `read_network_requests`, `resize_window`, `list_connected_browsers`,
`page_screenshot`, `list_evidence`, `view_evidence` (see [Screenshots and evidence](#screenshots-and-evidence)).
That leaves out tools that act on a page as you (`javascript_tool`, `form_input`, `computer`,
`file_upload`, `upload_image`, `shortcuts_execute`, `browser_batch`, `gif_creator`), `save_screenshot`
(writes files), and browser
re-selection (`select_browser`, `switch_browser`). Or pass an explicit list:

```bash
BRIDGE_TOOLS=read npm start
BRIDGE_TOOLS=tabs_context_mcp,tabs_create_mcp,navigate,get_page_text,form_input npm start
```

Each action inside `browser_batch` is checked against the same list, so it can't be used to reach
a tool that isn't allowed.

### Screenshots and evidence

With the default `all`, Claude can click, scroll, type, press keys and fill in forms in Chrome
(`computer`, `form_input`, `find`/`read_page` to locate elements), so it can bring the evidence
onto the screen before capturing it. The bridge adds four tools of its own:

| Tool | Does |
|---|---|
| `page_screenshot` | Captures the current tab (or `tabId`) and returns the image, page URL and title, and a one-time link |
| `save_screenshot` | Same, and saves it as `<evidence folder>/<folder>/<name>-<timestamp>-<random>.jpg`, e.g. `~/evidence/5.1.3/weak-password-2026-10-26T01-02-03Z-3f9a1c.jpg`. Optional `ref` scrolls an element into view first; `region` crops to `[x0, y0, x1, y1]` |
| `list_evidence` | Lists saved screenshots (whole evidence folder or one subfolder) |
| `view_evidence` | Opens a saved screenshot so Claude can check it shows the issue, with a one-time link |

The evidence folder is `~/evidence`, or `BRIDGE_EVIDENCE_DIR`. Use one per engagement:

```bash
BRIDGE_EVIDENCE_DIR=~/engagements/clientx/evidence npm start
```

The one-time link (`https://localhost:47823/shots/<id>.jpg`) lets Claude insert the picture into
the document: taskpane code fetches it and passes the base64 bytes to e.g.
`insertInlinePictureFromBase64`. See [SECURITY.md](SECURITY.md#screenshots-and-the-evidence-folder)
for how links and the folder are protected.

Ask, for example:

> *For each finding heading in this document (e.g. 5.1.3 Weak password), check the issue in
> Chrome, take a screenshot that shows it, save it to the evidence folder under the heading number,
> view it to confirm it shows the issue, and insert it under the heading with the URL as caption.*

### Multiple browsers

Every Chrome signed into the Claude extension with your account is reachable, including other
computers (via Anthropic's relay). Without a choice the session can land on another machine, and
its tabs and Allow prompts then appear there. The bridge therefore pins the session to the one
browser reported as running on this computer before the first browser action, and errors if that is
ambiguous. Set `BRIDGE_BROWSER` to pick one explicitly; IDs are in the error message, or ask
Claude in Word to run `list_connected_browsers`.

## Security

See [SECURITY.md](SECURITY.md) for the controls in the bridge, the threat model, where the token
lives, the evidence folder and how to rotate the token. In short: localhost only, bearer token,
origin pinned to `https://pivot.claude.ai`, Chrome's per-site Allow prompt stays on, files are
written only inside the evidence folder, and you approve every tool call in Office.

## Troubleshooting

`local/bridge.log` (if started with `npm start >> local/bridge.log 2>&1`) or the terminal shows every request.

| Symptom | Cause |
|---|---|
| Claude in Word says it has no browser access, no requests in the log | Word loaded the store add-in, not the sideloaded one. Re-run `npm run install-addin`, Cmd+Q Word, pick the developer add-in |
| Connector shows *Connection failed / Load failed*, only `OPTIONS` in the log | CORS preflight rejected; check `BRIDGE_ORIGINS` |
| *Load failed*, nothing in the log | Cert not trusted (`npm run cert`) or bridge not running |
| Permission prompt never shows / times out | Session is on another computer's Chrome; see [Multiple browsers](#multiple-browsers) |
| `401` in the log | Office is sending an old token. Restart the bridge if you rotated, then Cmd+Q and reopen Office so it reads the new manifest |
| Office keeps calling an old URL or port | It hasn't re-read the manifest: `npm run install-addin`, then Cmd+Q and reopen Office |
| Claude says a tool is "not exposed by bridge" | Blocked by `BRIDGE_TOOLS`; see [Tool exposure](#tool-exposure) |
| Bridge exits with "refusing to start with ... skip_all_permission_checks" | Unset `CLAUDE_CHROME_PERMISSION_MODE` |

## Uninstall

```bash
npm run uninstall-addin                                   # removes the sideloaded manifest
security delete-certificate -c 'localhost (office-chrome-bridge)' ~/Library/Keychains/login.keychain-db
rm -rf local                                              # token, cert, manifest, logs
# ~/evidence (or BRIDGE_EVIDENCE_DIR) holds your saved screenshots; delete it yourself when done
```

## Layout

```
src/bridge.mjs             the bridge (HTTP ⇄ stdio MCP, auth, CORS, tool allowlist, browser pinning, screenshot + evidence tools)
scripts/make-cert.sh       generate + trust the localhost TLS cert
scripts/build-manifest.mjs build local/manifest.xml from manifest/template.xml
scripts/install-addin.sh   sideload into Word / Excel / PowerPoint (macOS)
scripts/uninstall-addin.sh remove the sideloaded manifest
scripts/rotate-token.sh    new token, rebuild + reinstall the manifest
scripts/check-secrets.sh   fail if a token, key or secret file is tracked or in git history
SECURITY.md                security controls, threat model, token and evidence-folder handling
AGENTS.md                  instructions for coding agents (CLAUDE.md points here)
manifest/template.xml      snapshot of Anthropic's add-in manifest (pivot.claude.ai/manifest.xml);
                           refresh with `node scripts/build-manifest.mjs --fetch`
local/                     git-ignored: .token, certs/, manifest.xml, bridge.log
```

## Caveats

- Unofficial. `mcp_servers` is documented by Anthropic for third-party-platform deployments of
  the add-in; it also works with Claude.ai sign-in today, but that may change.
- The sideloaded manifest uses the same add-in ID as the store version. Office updates or a new
  add-in version can override it; rerun `npm run manifest -- --fetch` and `npm run install-addin`.
- The bridge must be running whenever you use the connector.
