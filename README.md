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
`find`, `read_console_messages`, `read_network_requests`, `resize_window`, `list_connected_browsers`.
That leaves out tools that act on a page as you (`javascript_tool`, `form_input`, `computer`,
`file_upload`, `upload_image`, `shortcuts_execute`, `browser_batch`, `gif_creator`) and browser
re-selection (`select_browser`, `switch_browser`). Or pass an explicit list:

```bash
BRIDGE_TOOLS=read npm start
BRIDGE_TOOLS=tabs_context_mcp,tabs_create_mcp,navigate,get_page_text,form_input npm start
```

Each action inside `browser_batch` is checked against the same list, so it can't be used to reach
a tool that isn't allowed.

### Multiple browsers

Every Chrome signed into the Claude extension with your account is reachable, including other
computers (via Anthropic's relay). Without a choice the session can land on another machine, and
its tabs and Allow prompts then appear there. The bridge therefore pins the session to the one
browser reported as running on this computer before the first browser action, and errors if that is
ambiguous. Set `BRIDGE_BROWSER` to pick one explicitly; IDs are in the error message, or ask
Claude in Word to run `list_connected_browsers`.

## Security

Controls in the bridge:

- Binds `127.0.0.1` only and refuses to start on any other address.
- Every request needs the bearer token from `local/.token` (random 256-bit, mode 0600).
- Requests with an `Origin` other than `https://pivot.claude.ai` are rejected, preflight included
  (stops other websites and DNS rebinding from driving your browser).
- Optional tool allowlist (`BRIDGE_TOOLS`), including inside `browser_batch` ([Tool exposure](#tool-exposure)).
- Refuses to run with `CLAUDE_CHROME_PERMISSION_MODE=skip_all_permission_checks`, so the
  extension's per-site **Allow** prompt stays on.
- Pins the session to this computer's Chrome ([Multiple browsers](#multiple-browsers)).

### Threat model

| If this happens | Can someone drive your browser? |
|---|---|
| **Token leaks**, attacker is remote (internet / LAN) | No. Nothing listens off loopback. |
| **Token leaks**, attacker controls a website you visit | No. Browsers always send `Origin`; anything but `https://pivot.claude.ai` gets 403. |
| **Token leaks**, another OS user on this computer | Yes, loopback is shared between local users. Rotate the token. |
| **Malware runs as you** | Yes, but it doesn't need the bridge: it can run `claude --claude-in-chrome-mcp` or use the extension's local socket directly. |
| **TLS private key leaks** | No. The cert is valid only for `localhost` / `127.0.0.1`, is `CA:FALSE` and server-auth only, so it can't mint certs for other sites. Worst case: something local impersonates the bridge while it is stopped to capture the token. Regenerate with `npm run cert -- --force`. |
| **Your Claude account is compromised** | Yes, every Chrome connected to the account, through Anthropic's relay, with or without this bridge. Protect the account and remove browsers you don't recognise in the extension. |
| **Prompt injection**: a document or web page tells Claude to misuse the browser | This is the realistic risk. Limited by approving each tool call in Office, the per-site Allow prompt, `BRIDGE_TOOLS=read` if you want it, and what your Chrome profile is logged into. |

Recommendations:

- Use a **separate Chrome profile** for the extension, logged into only what you need.
- Read what each tool call will do before approving it in Office; use `BRIDGE_TOOLS=read` for sessions that only need to read pages.
- Read Allow prompts; don't approve sites you didn't expect.
- Run the bridge only while you need it.

### Where the token lives

- `local/.token` and `local/manifest.xml` (mode 0600)
- The installed manifests under `~/Library/Containers/com.microsoft.*/Data/Documents/wef/` (mode 0600).
  In the manifest it sits in the URL **fragment**, which browsers never send to a server.
- **Office's manifest cache**
  (`~/Library/Containers/com.microsoft.*/Data/Library/Application Support/Microsoft/Office/16.0/Wef/.../Manifests/<id>_<version>`).
  Office copies the manifest here each time it starts, token included, readable by other local
  users (mode 0644).
- Possibly **the add-in's WebView storage** (`.../Data/Library/WebKit/WebsiteData`), where it keeps
  MCP registrations. Only your macOS user account protects it, the same as the rest of your Office data.

Keep the token out of terminals, logs and chat transcripts you share.

### Rotating the token

`npm run rotate`, restart the bridge, then quit Office (Cmd+Q) and reopen it.

Office only reads the manifest when it starts, so until it's fully restarted it keeps sending the
old token and gets `401`. Old tokens are dead once rotated; the bridge only accepts the current one.

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
```

## Layout

```
src/bridge.mjs             the bridge (HTTP ⇄ stdio MCP, auth, CORS, tool allowlist, browser pinning)
scripts/make-cert.sh       generate + trust the localhost TLS cert
scripts/build-manifest.mjs build local/manifest.xml from manifest/template.xml
scripts/install-addin.sh   sideload into Word / Excel / PowerPoint (macOS)
scripts/uninstall-addin.sh remove the sideloaded manifest
scripts/rotate-token.sh    new token, rebuild + reinstall the manifest
scripts/check-secrets.sh   fail if a token, key or secret file is tracked or in git history
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
