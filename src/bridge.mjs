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
//   BRIDGE_ALLOW_SKIP_PERMISSIONS=1  permit CLAUDE_CHROME_PERMISSION_MODE=skip_all_permission_checks
//   CLAUDE_BIN       default `claude`
import { spawn } from 'node:child_process'
import { createServer } from 'node:https'
import { createHash, randomBytes, timingSafeEqual } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
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
  'read_page', 'find', 'read_console_messages', 'read_network_requests', 'resize_window', 'list_connected_browsers']
const TOOLS_SPEC = (process.env.BRIDGE_TOOLS || 'all').trim()
const TOOLS = TOOLS_SPEC === 'all' ? null
  : new Set(TOOLS_SPEC === 'read' ? READ_TOOLS : TOOLS_SPEC.split(',').map(s => s.trim()).filter(Boolean))
const CLAUDE_BIN = process.env.CLAUDE_BIN || 'claude'
const BROWSER = process.env.BRIDGE_BROWSER || null
const BROWSER_TOOLS = new Set(['list_connected_browsers', 'select_browser', 'switch_browser'])
const MAX_BODY = 10 * 1024 * 1024
const CALL_TIMEOUT_MS = 5 * 60 * 1000

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
// The add-in stores MCP registrations (URL + headers) in its own WebView storage, so a
// rotated token on an unchanged URL keeps sending the old header. A token-derived path
// segment gives every token its own URL; must match scripts/build-manifest.mjs.
const KEY_ID = createHash('sha256').update(TOKEN).digest('hex').slice(0, 12)

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
  if (msg.method === 'tools/call' && !BROWSER_TOOLS.has(msg.params?.name)) {
    const problem = await ensureLocalBrowser()
    if (problem) return reply({ result: { content: [{ type: 'text', text: problem }], isError: true } })
  }
  if (msg.method === 'tools/call' && msg.params?.name === 'select_browser') browserSelected = msg.params?.arguments?.deviceId || null
  const r = await childRequest(msg.method, msg.params)
  if (msg.method === 'tools/list' && TOOLS && r.result?.tools) {
    r.result = { ...r.result, tools: r.result.tools.filter(t => TOOLS.has(t.name)) }
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
  log(req.method, req.url, 'origin=' + (origin || '-'),
    req.method === 'OPTIONS'
      ? `acr-method=${req.headers['access-control-request-method'] || '-'} acr-headers=${req.headers['access-control-request-headers'] || '-'}`
      : `accept=${req.headers.accept || '-'} proto=${req.headers['mcp-protocol-version'] || '-'}`)
  // DNS-rebinding / drive-by protection: browsers always send Origin cross-site
  if (origin && !ORIGINS.has(origin)) { log('rejected origin', origin); return send(res, 403, { error: 'origin not allowed' }) }
  if (req.method === 'OPTIONS') return send(res, 204)
  const path = new URL(req.url, 'https://localhost').pathname
  if (path !== '/mcp' && path !== `/mcp/${KEY_ID}`) {
    if (path.startsWith('/mcp/')) log('stale key id in URL (token was rotated; reopen Office):', path)
    return send(res, 404, { error: 'not found' })
  }
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
