#!/usr/bin/env node
// Localhost-only Streamable HTTP front for `claude --claude-in-chrome-mcp`, so the
// Claude for Word add-in (which calls `mcp_servers[].url` from the taskpane on this
// machine) can use Claude in Chrome. No dependencies.
//
//   node src/bridge.mjs
//
// Env:
//   BRIDGE_PORT      default 47823
//   BRIDGE_HOST      default 127.0.0.1 (never bind a public interface)
//   BRIDGE_CERT/KEY  default local/certs/localhost.{crt,key}
//   BRIDGE_TOKEN     default: contents of local/.token (generated on first run, mode 0600)
//   BRIDGE_ORIGINS   default https://pivot.claude.ai (comma-separated)
//   BRIDGE_TOOLS     `all` (default), `read` (read + navigate only), or a comma-separated list
//   BRIDGE_BROWSER   deviceId to drive; default: the one connected browser on this computer
//                    (several Chromes can be connected to one Claude account, incl. other PCs)
//   BRIDGE_EVIDENCE_DIR  default ~/evidence: the only folder save_screenshot writes to and
//                    list_evidence / view_evidence read from
//   BRIDGE_ALLOW_SKIP_PERMISSIONS=1  permit CLAUDE_CHROME_PERMISSION_MODE=skip_all_permission_checks
//   CLAUDE_BIN       default `claude`
import { spawn } from 'node:child_process'
import { createServer } from 'node:https'
import { randomBytes, timingSafeEqual } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, readdirSync, realpathSync, statSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { basename, dirname, extname, join, relative, resolve, sep } from 'node:path'
import { fileURLToPath } from 'node:url'

const LOCAL = join(dirname(fileURLToPath(import.meta.url)), '..', 'local')
const PORT = Number(process.env.BRIDGE_PORT || 47823)
const HOST = process.env.BRIDGE_HOST || '127.0.0.1'
const CERT = process.env.BRIDGE_CERT || join(LOCAL, 'certs', 'localhost.crt')
const KEY = process.env.BRIDGE_KEY || join(LOCAL, 'certs', 'localhost.key')
const ORIGINS = new Set((process.env.BRIDGE_ORIGINS || 'https://pivot.claude.ai').split(',').map(s => s.trim()).filter(Boolean))
// All tools by default: the add-in asks the user to approve each MCP tool call. `read`
// narrows it to reading pages and moving between them, without tools that act on a page
// as you (javascript_tool, form_input, computer, file_upload, browser_batch, ...).
const READ_TOOLS = ['tabs_context_mcp', 'tabs_create_mcp', 'tabs_close_mcp', 'navigate', 'get_page_text',
  'read_page', 'find', 'read_console_messages', 'read_network_requests', 'resize_window', 'list_connected_browsers',
  'page_screenshot', 'list_evidence', 'view_evidence']
const TOOLS_SPEC = (process.env.BRIDGE_TOOLS || 'all').trim()
const TOOLS = TOOLS_SPEC === 'all' ? null
  : new Set(TOOLS_SPEC === 'read' ? READ_TOOLS : TOOLS_SPEC.split(',').map(s => s.trim()).filter(Boolean))
const CLAUDE_BIN = process.env.CLAUDE_BIN || 'claude'
const BROWSER = process.env.BRIDGE_BROWSER || null
const BROWSER_TOOLS = new Set(['list_connected_browsers', 'select_browser', 'switch_browser'])
const MAX_BODY = 10 * 1024 * 1024
const CALL_TIMEOUT_MS = 5 * 60 * 1000
const SHOT_TTL_MS = 5 * 60 * 1000
const MAX_SHOTS = 20
const EVIDENCE_DIR = resolve((process.env.BRIDGE_EVIDENCE_DIR || '~/evidence').replace(/^~(?=$|\/)/, homedir()))
const MAX_EVIDENCE_READ = 15 * 1024 * 1024

if (!['127.0.0.1', '::1', 'localhost'].includes(HOST)) {
  console.error(`refusing to bind ${HOST}: this bridge drives your logged-in browser and must stay on loopback`)
  process.exit(1)
}
// The extension's per-site Allow prompt is the main guard against prompt-injected browser use
if (process.env.CLAUDE_CHROME_PERMISSION_MODE === 'skip_all_permission_checks' && process.env.BRIDGE_ALLOW_SKIP_PERMISSIONS !== '1') {
  console.error('refusing to start with CLAUDE_CHROME_PERMISSION_MODE=skip_all_permission_checks: anything that reaches ' +
    'Claude in Word (documents, web pages) could then drive your browser unprompted. Set BRIDGE_ALLOW_SKIP_PERMISSIONS=1 to override.')
  process.exit(1)
}

function loadToken() {
  if (process.env.BRIDGE_TOKEN) return process.env.BRIDGE_TOKEN
  const file = join(LOCAL, '.token')
  if (!existsSync(file)) {
    mkdirSync(LOCAL, { recursive: true, mode: 0o700 })
    writeFileSync(file, randomBytes(32).toString('base64url') + '\n', { mode: 0o600 })
  }
  return readFileSync(file, 'utf8').trim()
}
const TOKEN = Buffer.from(loadToken())

const log = (...a) => console.error(new Date().toISOString(), '[bridge]', ...a)

// ---- child: one long-lived Claude in Chrome MCP server ------------------------------
let child = null
let childBuf = ''
let nextId = 1
const pending = new Map()      // internal id -> { resolve, timer }
let initResult = null          // first initialize result, replayed to later clients
let browserSelected = null     // deviceId pinned for the current child

function startChild() {
  child = spawn(CLAUDE_BIN, ['--claude-in-chrome-mcp'], { stdio: ['pipe', 'pipe', 'pipe'] })
  childBuf = ''
  initResult = null
  browserSelected = null
  log('started', CLAUDE_BIN, '--claude-in-chrome-mcp pid', child.pid)
  child.stdout.setEncoding('utf8')
  child.stdout.on('data', chunk => {
    childBuf += chunk
    let i
    while ((i = childBuf.indexOf('\n')) >= 0) {
      const line = childBuf.slice(0, i); childBuf = childBuf.slice(i + 1)
      if (line.trim()) onChildLine(line)
    }
  })
  child.stderr.setEncoding('utf8')
  child.stderr.on('data', d => process.stderr.write('[chrome-mcp] ' + d))
  child.on('exit', (code, sig) => {
    log(`chrome mcp exited code=${code} sig=${sig}; restarting in 1s`)
    for (const [, p] of pending) p.resolve({ error: { code: -32000, message: 'Claude in Chrome MCP server restarted' } })
    pending.clear()
    setTimeout(startChild, 1000)
  })
}

function onChildLine(line) {
  let msg
  try { msg = JSON.parse(line) } catch { return log('non-JSON from child:', line.slice(0, 200)) }
  if (msg.id !== undefined && msg.method) {
    // server -> client request (roots/sampling/elicitation): not supported over this bridge
    return childSend({ jsonrpc: '2.0', id: msg.id, error: { code: -32601, message: 'not supported by bridge' } })
  }
  if (msg.id !== undefined) {
    const p = pending.get(msg.id)
    if (p) { pending.delete(msg.id); clearTimeout(p.timer); p.resolve(msg) }
  }
  // notifications from the child (logging, list_changed) are dropped: no GET stream
}

function childSend(obj) { child.stdin.write(JSON.stringify(obj) + '\n') }

function childRequest(method, params) {
  const id = nextId++
  return new Promise(resolve => {
    const timer = setTimeout(() => { pending.delete(id); resolve({ error: { code: -32001, message: `timed out after ${CALL_TIMEOUT_MS}ms` } }) }, CALL_TIMEOUT_MS)
    pending.set(id, { resolve, timer })
    childSend({ jsonrpc: '2.0', id, method, ...(params !== undefined && { params }) })
  })
}

// Pin the session to this computer's Chrome before the first browser action, otherwise the
// extension may route to another computer's Chrome signed into the same account.
async function ensureLocalBrowser() {
  if (browserSelected) return null
  let deviceId = BROWSER
  if (!deviceId) {
    const r = await childRequest('tools/call', { name: 'list_connected_browsers', arguments: {} })
    let list = []
    try { list = JSON.parse(r.result?.content?.[0]?.text || '[]') } catch {}
    const here = list.filter(b => b.onThisComputer === true)
    const local = here.length ? here : list.filter(b => b.isLocal)
    if (local.length !== 1) {
      return `bridge: expected exactly one Chrome on this computer, found ${local.length} of ${list.length} connected ` +
        `(${list.map(b => `${b.name}/${b.osPlatform}/${b.deviceId}`).join(', ') || 'none'}); set BRIDGE_BROWSER=<deviceId>`
    }
    deviceId = local[0].deviceId
  }
  const s = await childRequest('tools/call', { name: 'select_browser', arguments: { deviceId } })
  if (s.error || s.result?.isError) return `bridge: select_browser ${deviceId} failed: ${JSON.stringify(s.error || s.result.content)}`
  browserSelected = deviceId
  log('pinned browser', deviceId)
  return null
}

// ---- bridge-provided tools: screenshots and the evidence folder ----------------------
// Claude in Word sees an image block but can't copy its bytes into the document, so each
// screenshot is also kept in memory behind a one-time, expiring link that taskpane code can
// fetch and pass to insertInlinePictureFromBase64. The unguessable id is the credential:
// that code has no access to the bearer token.
const shots = new Map()        // id -> { buf, mime, timer }

const TAB_ID = { type: 'number', description: 'Tab to capture; default: the tab Claude in Chrome currently has selected' }
const CAPTURE_PROPS = {
  tabId: TAB_ID,
  ref: { type: 'string', description: 'Element reference (from read_page or find) to scroll into view before capturing' },
  region: { type: 'array', items: { type: 'number' }, minItems: 4, maxItems: 4,
    description: '(x0, y0, x1, y1) viewport region to capture instead of the whole tab, e.g. to crop to the evidence' },
}
const INSERT_HINT = 'To put an image in the document, fetch its one-time link (valid 5 minutes, works once) from ' +
  'Office.js code, base64-encode the bytes and insert them as a picture (e.g. insertInlinePictureFromBase64 in Word).'

const OWN_TOOLS = {
  page_screenshot: {
    run: pageScreenshot,
    def: {
      name: 'page_screenshot',
      description: 'Take a screenshot of the Chrome tab Claude in Chrome is on (or of tabId). Returns the image, the ' +
        'page URL and title, and a one-time link to the image file. ' + INSERT_HINT,
      inputSchema: { type: 'object', properties: CAPTURE_PROPS },
    },
  },
  save_screenshot: {
    run: saveScreenshot,
    def: {
      name: 'save_screenshot',
      description: `Take a screenshot of the Chrome tab and save it in the evidence folder (${EVIDENCE_DIR}), e.g. ` +
        'folder "5.1.3" for finding 5.1.3. Click, scroll or zoom with the computer tool first so the evidence is ' +
        'visible, or pass ref/region. Files are never overwritten: a timestamp and random suffix are added to the name. Returns the image ' +
        '(check it shows the issue), the saved path, the page URL and a one-time link. ' + INSERT_HINT,
      inputSchema: {
        type: 'object',
        properties: {
          folder: { type: 'string', description: 'Subfolder of the evidence folder, e.g. "5.1.3" or "client/5.1.3"; created if missing' },
          name: { type: 'string', description: 'File name without extension, e.g. "weak-password-login" (default "screenshot")' },
          ...CAPTURE_PROPS,
        },
        required: ['folder'],
      },
    },
  },
  list_evidence: {
    run: listEvidence,
    local: true,
    def: {
      name: 'list_evidence',
      description: `List the screenshots saved in the evidence folder (${EVIDENCE_DIR}) or one of its subfolders.`,
      inputSchema: { type: 'object', properties: { folder: { type: 'string', description: 'Subfolder, e.g. "5.1.3"; default: the whole evidence folder' } } },
    },
  },
  view_evidence: {
    run: viewEvidence,
    local: true,
    def: {
      name: 'view_evidence',
      description: 'Open a saved screenshot from the evidence folder to check it, by its path as list_evidence or ' +
        'save_screenshot report it (e.g. "5.1.3/weak-password-2026-10-26T01-02-03Z-3f9a1c.jpg"). Returns the image and a ' +
        'one-time link. ' + INSERT_HINT,
      inputSchema: { type: 'object', properties: { path: { type: 'string', description: 'Path relative to the evidence folder' } }, required: ['path'] },
    },
  },
}

const firstJson = r => { try { return JSON.parse(r.result?.content?.[0]?.text) } catch { return null } }
const text = t => ({ type: 'text', text: t })
const failure = (tool, t) => ({ result: { content: [text(`${tool}: ${t}`)], isError: true } })

function storeShot(buf, mime) {
  if (shots.size >= MAX_SHOTS) dropShot(shots.keys().next().value)
  const id = randomBytes(16).toString('base64url')
  shots.set(id, { buf, mime, timer: setTimeout(() => dropShot(id), SHOT_TTL_MS) })
  return `https://localhost:${PORT}/shots/${id}.${mime === 'image/png' ? 'png' : 'jpg'}`
}
function dropShot(id) { const s = shots.get(id); if (s) { clearTimeout(s.timer); shots.delete(id) } }

// Capture a tab through the child's computer tool; returns { img, tabId, tab } or { problem }
async function capture(args) {
  const ctx = firstJson(await childRequest('tools/call', { name: 'tabs_context_mcp', arguments: {} }))
  const tabId = args?.tabId ?? ctx?.selectedTabId
  if (tabId === undefined) return { problem: failure('capture', 'no tab to capture; open a page first (tabs_create_mcp / navigate)') }
  const call = async a => {
    const r = await childRequest('tools/call', { name: 'computer', arguments: { ...a, tabId } })
    return r.error || r.result?.isError ? { problem: r.error ? r : { result: r.result } } : r
  }
  if (args?.ref) { const s = await call({ action: 'scroll_to', ref: args.ref }); if (s.problem) return s }
  const r = await call(args?.region ? { action: 'zoom', region: args.region } : { action: 'screenshot' })
  if (r.problem) return r
  const img = r.result?.content?.find(b => b.type === 'image')
  if (!img) return { problem: { result: r.result } }
  return { img, tabId, tab: ctx?.availableTabs?.find(t => t.tabId === tabId) }
}
const describeTab = c => `tab ${c.tabId}${c.tab ? `: "${c.tab.title}" (${c.tab.url})` : ''}`

async function pageScreenshot(args) {
  const c = await capture(args)
  if (c.problem) return c.problem
  const link = storeShot(Buffer.from(c.img.data, 'base64'), c.img.mimeType)
  return { result: { content: [c.img, text(`Screenshot of ${describeTab(c)}\nImage file (${c.img.mimeType}, one-time link): ${link}`)] } }
}

// ---- evidence folder: every path stays inside EVIDENCE_DIR -------------------------
// Names are plain segments (no leading dot, no slashes, no `..`) and each directory is
// resolved with realpath as it is walked, so a symlink can't lead outside the root.
const SEGMENT = /^\w[\w .,()+-]*$/
const IMAGE_EXT = { '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.png': 'image/png' }

function evidenceDir(folder, create) {
  const raw = String(folder ?? '').trim()
  const parts = raw.split('/').filter(Boolean)
  if (raw.startsWith('/') || raw.startsWith('~') || parts.some(p => !SEGMENT.test(p))) {
    throw new Error(`invalid folder "${raw}": use relative names like 5.1.3 or client/5.1.3 (no "..", "~", "/" prefix or dotfiles)`)
  }
  if (create) mkdirSync(EVIDENCE_DIR, { recursive: true, mode: 0o700 })
  if (!existsSync(EVIDENCE_DIR)) throw new Error(`evidence folder ${EVIDENCE_DIR} does not exist yet; save a screenshot first`)
  const root = realpathSync(EVIDENCE_DIR)
  let dir = root
  for (const p of parts) {
    const next = join(dir, p)
    if (create && !existsSync(next)) mkdirSync(next, { mode: 0o700 })
    dir = realpathSync(next)
    if (!dir.startsWith(root + sep) || !statSync(dir).isDirectory()) throw new Error(`"${raw}" is not a folder inside the evidence folder`)
  }
  return { root, dir }
}

async function saveScreenshot(args) {
  const name = String(args?.name ?? 'screenshot').trim().replace(/\.(jpe?g|png)$/i, '')
  if (!SEGMENT.test(name)) return failure('save_screenshot', `invalid name "${name}": letters, digits, spaces and . _ - , ( ) + only`)
  if (!args?.folder) return failure('save_screenshot', 'folder is required, e.g. "5.1.3"')
  const c = await capture(args)
  if (c.problem) return c.problem
  const buf = Buffer.from(c.img.data, 'base64')
  const ext = c.img.mimeType === 'image/png' ? 'png' : 'jpg'
  const stamp = new Date().toISOString().replace(/\.\d+Z$/, 'Z').replace(/:/g, '-')
  const { root, dir } = evidenceDir(args.folder, true)
  // timestamp keeps files in capture order; the random suffix makes same-second names unique
  const file = join(dir, `${name}-${stamp}-${randomBytes(3).toString('hex')}.${ext}`)
  writeFileSync(file, buf, { flag: 'wx', mode: 0o600 })   // wx: never overwrite, never follow a planted symlink
  log('saved evidence', relative(root, file), buf.length, 'bytes')
  const link = storeShot(buf, c.img.mimeType)
  return { result: { content: [c.img, text(
    `Saved ${relative(root, file)} (${file}) from ${describeTab(c)} at ${new Date().toISOString()}\n` +
    `Image file (${c.img.mimeType}, one-time link): ${link}`)] } }
}

async function listEvidence(args) {
  const { root, dir } = evidenceDir(args?.folder, false)
  const out = []
  const walk = d => {
    for (const e of readdirSync(d, { withFileTypes: true })) {
      if (out.length >= 500 || e.name.startsWith('.')) continue
      const p = join(d, e.name)
      if (e.isDirectory()) walk(p)
      else if (e.isFile() && IMAGE_EXT[extname(e.name).toLowerCase()]) {
        const st = statSync(p)
        out.push(`${relative(root, p)}  ${st.size} bytes  ${st.mtime.toISOString()}`)
      }
    }
  }
  walk(dir)
  return { result: { content: [text(out.length ? out.join('\n') : `no screenshots in ${relative(root, dir) || EVIDENCE_DIR}`)] } }
}

async function viewEvidence(args) {
  const raw = String(args?.path ?? '').trim()
  const mime = IMAGE_EXT[extname(raw).toLowerCase()]
  if (!mime) return failure('view_evidence', `"${raw}" is not a .jpg or .png file`)
  const { root, dir } = evidenceDir(dirname(raw) === '.' ? '' : dirname(raw), false)
  const name = basename(raw)
  if (!SEGMENT.test(name)) return failure('view_evidence', `invalid file name "${name}"`)
  const file = realpathSync(join(dir, name))
  if (!file.startsWith(root + sep) || !statSync(file).isFile()) return failure('view_evidence', `"${raw}" is not a file inside the evidence folder`)
  if (statSync(file).size > MAX_EVIDENCE_READ) return failure('view_evidence', `"${raw}" is larger than ${MAX_EVIDENCE_READ} bytes`)
  const buf = readFileSync(file)
  const link = storeShot(buf, mime)
  return { result: { content: [{ type: 'image', data: buf.toString('base64'), mimeType: mime },
    text(`${relative(root, file)} (${buf.length} bytes)\nImage file (${mime}, one-time link): ${link}`)] } }
}

async function runOwnTool(name, args) {
  try { return await OWN_TOOLS[name].run(args) } catch (e) {
    return failure(name, e.code === 'ENOENT' ? 'not found' : e.code === 'EEXIST' ? 'file already exists' : e.message)
  }
}

// ---- JSON-RPC handling with tool allowlist ------------------------------------------
async function handle(msg) {
  const isRequest = msg && msg.id !== undefined && msg.id !== null && typeof msg.method === 'string'
  if (!isRequest) {
    // notifications: the child is initialized once; swallow repeats from new clients
    if (msg && msg.method === 'notifications/initialized' && initResult?.forwarded) return null
    if (msg && typeof msg.method === 'string') childSend(msg)
    if (msg && msg.method === 'notifications/initialized' && initResult) initResult.forwarded = true
    return null
  }
  const reply = r => ({ jsonrpc: '2.0', id: msg.id, ...('error' in r ? { error: r.error } : { result: r.result }) })

  if (msg.method === 'initialize') {
    if (initResult) return reply({ result: initResult.result })
    const r = await childRequest('initialize', msg.params)
    if (r.result) initResult = { result: r.result, forwarded: false }
    return reply(r)
  }
  if (msg.method === 'tools/call' && TOOLS) {
    // browser_batch runs other tools by name, so each inner action must be allowed too
    const names = [msg.params?.name]
    if (msg.params?.name === 'browser_batch') {
      const actions = msg.params?.arguments?.actions
      names.push(...(Array.isArray(actions) ? actions.map(a => a?.name) : [undefined]))
    }
    const blocked = names.filter(n => !TOOLS.has(n))
    if (blocked.length) {
      log('blocked tool', blocked.join(','))
      return reply({ error: { code: -32602, message: `tool not exposed by bridge: ${blocked.join(', ')}` } })
    }
  }
  const own = msg.method === 'tools/call' && Object.hasOwn(OWN_TOOLS, msg.params?.name) ? OWN_TOOLS[msg.params.name] : null
  if (msg.method === 'tools/call' && !BROWSER_TOOLS.has(msg.params?.name) && !own?.local) {
    const problem = await ensureLocalBrowser()
    if (problem) return reply({ result: { content: [{ type: 'text', text: problem }], isError: true } })
  }
  if (msg.method === 'tools/call' && msg.params?.name === 'select_browser') browserSelected = msg.params?.arguments?.deviceId || null
  const r = own ? await runOwnTool(msg.params.name, msg.params?.arguments) : await childRequest(msg.method, msg.params)
  if (msg.method === 'tools/list' && r.result?.tools) {
    const tools = [...r.result.tools, ...Object.values(OWN_TOOLS).map(t => t.def)]
    r.result = { ...r.result, tools: TOOLS ? tools.filter(t => TOOLS.has(t.name)) : tools }
  }
  if (msg.method === 'tools/call') log('tools/call', msg.params?.name, r.error ? 'error' : r.result?.isError ? 'isError' : 'ok')
  return reply(r)
}

// ---- HTTP ---------------------------------------------------------------------------
function cors(req, res) {
  const origin = req.headers.origin
  if (origin && ORIGINS.has(origin)) {
    res.setHeader('Access-Control-Allow-Origin', origin)
    res.setHeader('Vary', 'Origin')
    res.setHeader('Access-Control-Allow-Methods', 'POST, GET, DELETE, OPTIONS')
    // echo what the preflight asks for: security comes from the token + origin pin, not header filtering
    res.setHeader('Access-Control-Allow-Headers', req.headers['access-control-request-headers'] || 'authorization, content-type, accept, mcp-protocol-version, mcp-session-id, last-event-id')
    res.setHeader('Access-Control-Expose-Headers', 'mcp-session-id')
    res.setHeader('Access-Control-Max-Age', '600')
    // Chromium Private Network Access (Word on the web); harmless elsewhere
    if (req.headers['access-control-request-private-network']) res.setHeader('Access-Control-Allow-Private-Network', 'true')
  }
}

function authorized(req) {
  const h = req.headers.authorization || ''
  const got = Buffer.from(h.startsWith('Bearer ') ? h.slice(7).trim() : '')
  return got.length === TOKEN.length && timingSafeEqual(got, TOKEN)
}

function send(res, status, body) {
  if (body === undefined) { res.writeHead(status); return res.end() }
  res.writeHead(status, { 'Content-Type': 'application/json' })
  res.end(JSON.stringify(body))
}

const server = createServer({ cert: readFileSync(CERT), key: readFileSync(KEY) }, (req, res) => {
  cors(req, res)
  const origin = req.headers.origin
  log(req.method, req.url.replace(/^\/shots\/[^?]*/, '/shots/<id>'), 'origin=' + (origin || '-'),
    req.method === 'OPTIONS'
      ? `acr-method=${req.headers['access-control-request-method'] || '-'} acr-headers=${req.headers['access-control-request-headers'] || '-'}`
      : `accept=${req.headers.accept || '-'} proto=${req.headers['mcp-protocol-version'] || '-'}`)
  // DNS-rebinding / drive-by protection: browsers always send Origin cross-site
  if (origin && !ORIGINS.has(origin)) { log('rejected origin', origin); return send(res, 403, { error: 'origin not allowed' }) }
  if (req.method === 'OPTIONS') return send(res, 204)
  const path = new URL(req.url, 'https://localhost').pathname
  // page_screenshot links: the one-time id authenticates (taskpane code has no bearer token)
  const shot = req.method === 'GET' && path.match(/^\/shots\/([\w-]{22})\.(?:jpg|png)$/)
  if (shot) {
    const s = shots.get(shot[1])
    if (!s) { log('shot not found or expired'); return send(res, 404, { error: 'not found' }) }
    dropShot(shot[1])
    log('shot served', s.buf.length, 'bytes')
    res.writeHead(200, { 'Content-Type': s.mime, 'Content-Length': s.buf.length, 'Cache-Control': 'no-store' })
    return res.end(s.buf)
  }
  // the token authenticates, not the path; sub-paths keep manifests from older builds working
  if (path !== '/mcp' && !path.startsWith('/mcp/')) return send(res, 404, { error: 'not found' })
  if (!authorized(req)) { log('rejected unauthorized', req.method, origin || '-'); return send(res, 401, { error: 'unauthorized' }) }
  if (req.method === 'GET') return send(res, 405)        // no server-initiated stream
  if (req.method === 'DELETE') return send(res, 204)
  if (req.method !== 'POST') return send(res, 405)

  let size = 0
  const chunks = []
  req.on('data', c => {
    size += c.length
    if (size > MAX_BODY) { send(res, 413, { error: 'body too large' }); req.destroy() } else chunks.push(c)
  })
  req.on('end', async () => {
    if (res.writableEnded) return
    let body
    try { body = JSON.parse(Buffer.concat(chunks).toString('utf8')) } catch {
      return send(res, 400, { jsonrpc: '2.0', id: null, error: { code: -32700, message: 'parse error' } })
    }
    const batch = Array.isArray(body)
    const out = (await Promise.all((batch ? body : [body]).map(handle))).filter(Boolean)
    if (out.length === 0) return send(res, 202)
    send(res, 200, batch ? out : out[0])
  })
})

startChild()
server.listen(PORT, HOST, () => {
  log(`listening on https://${HOST === '127.0.0.1' ? 'localhost' : HOST}:${PORT}/mcp`)
  log(`origins: ${[...ORIGINS].join(', ')}  tools (${TOOLS_SPEC.includes(',') ? 'custom' : TOOLS_SPEC}): ${TOOLS ? [...TOOLS].join(', ') : 'all'}`)
  log(`chrome permission mode: ${process.env.CLAUDE_CHROME_PERMISSION_MODE || 'extension default (ask)'}`)
})
const bye = () => { try { child.removeAllListeners('exit'); child.kill('SIGTERM') } catch {} ; process.exit(0) }
process.on('SIGINT', bye)
process.on('SIGTERM', bye)
