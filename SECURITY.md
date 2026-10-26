# Security

The bridge lets the Claude for Office add-in drive your logged-in Chrome and save screenshots to
disk. This page covers what protects that, what doesn't, and where secrets end up. Setup is in
[README.md](README.md).

## Controls in the bridge

- Binds `127.0.0.1` only and refuses to start on any other address.
- Every `/mcp` request needs the bearer token from `local/.token` (random 256-bit, mode 0600).
- Requests with an `Origin` other than `https://pivot.claude.ai` are rejected, preflight included
  (stops other websites and DNS rebinding from driving your browser).
- Optional tool allowlist (`BRIDGE_TOOLS`), including inside `browser_batch`
  ([Tool exposure](README.md#tool-exposure)).
- Refuses to run with `CLAUDE_CHROME_PERMISSION_MODE=skip_all_permission_checks`, so the
  extension's per-site **Allow** prompt stays on.
- Pins the session to this computer's Chrome ([Multiple browsers](README.md#multiple-browsers)).
- Writes files only inside the evidence folder, only as images, never overwriting
  ([below](#screenshots-and-the-evidence-folder)).

## What Claude can do with the default `all` tools

Everything you can do in that Chrome profile: open pages, click, scroll, type, press keys, fill and
submit forms, run JavaScript on a page, upload files, as you and with your cookies. The add-in asks
you to approve each tool call, and the extension asks before Claude acts on a new site. That is
deliberate for verifying findings in a test (logging in, triggering an error, showing a response),
but it means an approved call does real things on real sites. `BRIDGE_TOOLS=read` (or an explicit
list) removes clicking, typing and file saving when a session only needs to read pages.

## Screenshots and the evidence folder

**Writes** (`save_screenshot`) go only under the evidence folder (`~/evidence`, or
`BRIDGE_EVIDENCE_DIR`):

- The tool takes a relative `folder` (`5.1.3`, `client/5.1.3`) and a `name`. Each part must start
  with a letter, digit or `_` and contain only letters, digits, spaces and `. _ - , ( ) +`, so
  absolute paths, `~`, `..` and dotfiles are refused.
- Each directory is resolved with `realpath` as it is walked or created, and must stay inside the
  root, so a symlink placed in the evidence folder can't lead out of it.
- The bridge picks the extension from the image type (`.jpg` / `.png`): no other file types.
- Files are created with `O_EXCL` (`wx`), mode 0600, named `<name>-<timestamp>-<random>.<ext>`:
  existing files are never overwritten and a symlink planted at the target name isn't followed.
- Directories are created mode 0700.

**Reads**: `list_evidence` uses the same path rules and returns only the names, sizes and dates of
`.jpg`/`.jpeg`/`.png` files. No tool reads a saved file back, and the bridge serves nothing outside
`/mcp`, where every request needs the bearer token.

The evidence itself (screenshots of client systems, possibly with credentials or personal data on
screen) is as sensitive as the rest of your engagement data: keep the folder on an encrypted disk,
use one folder per engagement, and delete it according to your retention rules. The bridge never
deletes evidence.

## Threat model

| If this happens | Can someone drive your browser? |
|---|---|
| **Token leaks**, attacker is remote (internet / LAN) | No. Nothing listens off loopback. |
| **Token leaks**, attacker controls a website you visit | No. Browsers always send `Origin`; anything but `https://pivot.claude.ai` gets 403. |
| **Token leaks**, another OS user on this computer | Yes, loopback is shared between local users. Rotate the token. |
| **Malware runs as you** | Yes, but it doesn't need the bridge: it can run `claude --claude-in-chrome-mcp` or use the extension's local socket directly, and read `~/evidence` itself. |
| **TLS private key leaks** | No. The cert is valid only for `localhost` / `127.0.0.1`, is `CA:FALSE` and server-auth only, so it can't mint certs for other sites. Worst case: something local impersonates the bridge while it is stopped to capture the token. Regenerate with `npm run cert -- --force`. |
| **Your Claude account is compromised** | Yes, every Chrome connected to the account, through Anthropic's relay, with or without this bridge. Protect the account and remove browsers you don't recognise in the extension. |
| **Prompt injection**: a document or web page tells Claude to misuse the browser | This is the realistic risk, and a target you're testing may contain such text. Limited by approving each tool call in Office, the per-site Allow prompt, `BRIDGE_TOOLS=read` if you want it, and what your Chrome profile is logged into. File writes are confined to images in the evidence folder, so injection can't plant or overwrite other files, but it could fill the folder with junk or misleading screenshots: check what you put in the report. |

## Recommendations

- Use a **separate Chrome profile** for the extension, logged into only what the engagement needs.
- Read what each tool call will do before approving it in Office, especially clicks, typing and
  form submissions on the systems you're testing. Use `BRIDGE_TOOLS=read` for sessions that only
  need to read pages.
- Read Allow prompts; don't approve sites you didn't expect.
- Point `BRIDGE_EVIDENCE_DIR` at the engagement's folder, and review the screenshots before they
  go into a report.
- Run the bridge only while you need it.

## Where the token lives

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

## Rotating the token

`npm run rotate`, restart the bridge, then quit Office (Cmd+Q) and reopen it.

Office only reads the manifest when it starts, so until it's fully restarted it keeps sending the
old token and gets `401`. Old tokens are dead once rotated; the bridge only accepts the current one.
