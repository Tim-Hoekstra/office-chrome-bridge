// Local model gateway: lets the Claude for Office add-in talk to a model on this machine.
//
// The add-in's "gateway" connection speaks the Anthropic Messages API (GET /v1/models,
// POST /v1/messages). This module answers those two routes and translates to whichever
// API the local server speaks:
//
//   ollama   Ollama's native API     GET /api/tags     POST /api/chat             (NDJSON stream)
//   openai   OpenAI-compatible API   GET /v1/models    POST /v1/chat/completions  (SSE stream)
//            (LM Studio, llama.cpp server, vLLM, LocalAI, and Ollama's own /v1)
//
// The two are different formats. The format is detected from the server (see detect) unless
// BRIDGE_LLM_FORMAT says which one to use.
//
// Env:
//   BRIDGE_LLM_URL      address of the local server, e.g. localhost:11434 (Ollama) or
//                       localhost:1234 (LM Studio). http:// is assumed. Unset: the gateway is off.
//                       This is the only setting needed; the rest are optional.
//   BRIDGE_LLM_FORMAT   `auto` (default), `ollama` or `openai`
//   BRIDGE_LLM_MODEL    model to use when the add-in asks for one the server doesn't have
//                       (default: the first model the server lists)
//   BRIDGE_LLM_KEY      API key sent upstream as `Authorization: Bearer`, if the server wants one
//   BRIDGE_LLM_NUM_CTX  Ollama only: context window (num_ctx) to ask for. Default: 32768, or
//                       the model's own maximum if that is smaller (see contextFor).
import { randomBytes } from 'node:crypto'

const RAW = (process.env.BRIDGE_LLM_URL || '').trim()
const BASE = (RAW && !/^\w+:\/\//.test(RAW) ? 'http://' + RAW : RAW).replace(/\/+$/, '').replace(/\/v1$/, '')
const FORMAT = (process.env.BRIDGE_LLM_FORMAT || 'auto').trim().toLowerCase()
const DEFAULT_MODEL = (process.env.BRIDGE_LLM_MODEL || '').trim()
const KEY = (process.env.BRIDGE_LLM_KEY || '').trim()
const NUM_CTX = Number(process.env.BRIDGE_LLM_NUM_CTX || 0)
const DEFAULT_NUM_CTX = 32768
const MAX_BODY = 64 * 1024 * 1024       // conversations carry base64 images
const DETECT_TIMEOUT_MS = 5000
const PING_MS = 10000

// The add-in drops every model whose name doesn't contain "claude" from the gateway's list, so
// local models are listed as `claude-local/<name>`. The prefix is stripped again on the way in.
const PREFIX = 'claude-local/'

export const llmEnabled = BASE !== ''
export const llmDescription = () => `${BASE} (format: ${format || FORMAT})`

if (llmEnabled) {
  if (!/^https?:\/\//.test(BASE)) { console.error(`BRIDGE_LLM_URL must be host:port or an http(s) URL, got ${RAW}`); process.exit(1) }
  if (!['auto', 'ollama', 'openai'].includes(FORMAT)) { console.error(`BRIDGE_LLM_FORMAT must be auto, ollama or openai, got ${FORMAT}`); process.exit(1) }
}

const log = (...a) => console.error(new Date().toISOString(), '[llm]', ...a)
const upstreamHeaders = () => ({ 'Content-Type': 'application/json', ...(KEY && { Authorization: `Bearer ${KEY}` }) })

class UpstreamError extends Error {
  constructor(status, message) { super(message); this.status = status }
}

// ---- format detection ---------------------------------------------------------------
// Ollama answers GET /api/version; nothing else does. Ollama is checked first because it
// also serves the OpenAI routes, and its native API is the one that accepts num_ctx.
// A failed detection isn't cached, so the server can be started after the bridge.
let format = FORMAT === 'auto' ? null : FORMAT

async function getJson(path) {
  const r = await fetch(BASE + path, { headers: upstreamHeaders(), signal: AbortSignal.timeout(DETECT_TIMEOUT_MS) })
  if (!r.ok) throw new UpstreamError(r.status, `GET ${path}: HTTP ${r.status}`)
  return r.json()
}

async function detect() {
  if (format) return format
  let reason
  try {
    if (typeof (await getJson('/api/version')).version === 'string') format = 'ollama'
  } catch (e) { reason = e.message }
  if (!format) {
    try {
      if (Array.isArray((await getJson('/v1/models')).data)) format = 'openai'
    } catch (e) { reason = e.message }
  }
  if (!format) throw new UpstreamError(502, `no Ollama or OpenAI-compatible server at ${BASE} (${reason || 'unrecognised response'}). Is it running?`)
  log(`detected ${format} API at ${BASE}`)
  return format
}

async function listModels() {
  const f = await detect()
  if (f === 'ollama') return ((await getJson('/api/tags')).models || []).map(m => m.name).filter(Boolean)
  return ((await getJson('/v1/models')).data || []).map(m => m.id).filter(Boolean)
}

// The add-in also asks for real Claude model names (e.g. a small model for chat titles).
// Anything the server doesn't have goes to the default model.
async function resolveModel(requested) {
  const name = String(requested || '').startsWith(PREFIX) ? requested.slice(PREFIX.length) : String(requested || '')
  const models = await listModels()
  if (models.includes(name)) return name
  const fallback = DEFAULT_MODEL || models[0]
  if (!fallback) throw new UpstreamError(502, `the server at ${BASE} lists no models; pull or load one first`)
  return fallback
}

// Ollama loads models with a small context window unless asked, and silently cuts off whatever
// doesn't fit, which here is the add-in's long system prompt. So every request asks for a
// usable window: 32768 tokens, or the model's maximum if it has less.
const contexts = new Map()     // model -> num_ctx
async function contextFor(model) {
  if (NUM_CTX > 0) return NUM_CTX
  if (!contexts.has(model)) {
    let max
    try {
      const r = await fetch(BASE + '/api/show', { method: 'POST', headers: upstreamHeaders(), body: JSON.stringify({ model }), signal: AbortSignal.timeout(DETECT_TIMEOUT_MS) })
      const info = r.ok ? (await r.json()).model_info || {} : {}
      max = Number(Object.entries(info).find(([k]) => k.endsWith('.context_length'))?.[1])
    } catch {}
    contexts.set(model, max > 0 ? Math.min(max, DEFAULT_NUM_CTX) : DEFAULT_NUM_CTX)
    log(`${model}: num_ctx ${contexts.get(model)}${max > 0 ? ` (model maximum ${max})` : ''}`)
  }
  return contexts.get(model)
}

// ---- Anthropic request -> OpenAI chat messages (the Ollama request is built from these) ----
const blockText = b => b?.type === 'text' ? b.text
  : b?.type === 'document' && b.source?.type === 'text' ? b.source.data
  : b?.type === 'document' ? '[document omitted: the local model bridge only passes text and images]'
  : ''
const imageUrl = b => b.source?.type === 'base64' ? `data:${b.source.media_type};base64,${b.source.data}`
  : b.source?.type === 'url' ? b.source.url : null
const joinText = blocks => blocks.map(blockText).filter(Boolean).join('\n')

function userMessage(blocks) {
  const images = blocks.filter(b => b.type === 'image').map(imageUrl).filter(Boolean)
  const body = joinText(blocks)
  if (!images.length) return body ? { role: 'user', content: body } : null
  return { role: 'user', content: [...(body ? [{ type: 'text', text: body }] : []), ...images.map(url => ({ type: 'image_url', image_url: { url } }))] }
}

function toOpenAIMessages(body) {
  const out = []
  const system = typeof body.system === 'string' ? body.system : Array.isArray(body.system) ? joinText(body.system) : ''
  if (system) out.push({ role: 'system', content: system })
  for (const m of body.messages || []) {
    const blocks = typeof m.content === 'string' ? [{ type: 'text', text: m.content }] : Array.isArray(m.content) ? m.content : []
    if (m.role === 'assistant') {
      const calls = blocks.filter(b => b.type === 'tool_use')
        .map(b => ({ id: b.id, type: 'function', function: { name: b.name, arguments: JSON.stringify(b.input ?? {}) } }))
      const content = joinText(blocks)     // thinking blocks are not sent back
      if (content || calls.length) out.push({ role: 'assistant', content, ...(calls.length && { tool_calls: calls }) })
      continue
    }
    // Tool results must directly follow the assistant turn that asked for them. A tool message
    // can only hold text, so images in a result (screenshots) follow in a user message.
    const resultImages = []
    for (const b of blocks.filter(b => b.type === 'tool_result')) {
      const inner = typeof b.content === 'string' ? [{ type: 'text', text: b.content }] : Array.isArray(b.content) ? b.content : []
      resultImages.push(...inner.filter(i => i.type === 'image'))
      out.push({ role: 'tool', tool_call_id: b.tool_use_id, content: (b.is_error ? 'Error: ' : '') + (joinText(inner) || (inner.length ? '[see image below]' : '')) })
    }
    const user = userMessage([...resultImages, ...blocks.filter(b => b.type !== 'tool_result')])
    if (user) out.push(user)
  }
  return out
}

// Only tools the add-in runs itself are passed on. Anthropic's server tools (web search, code
// execution) have no input_schema and nothing local can run them.
const toTools = body => (body.tools || []).filter(t => t?.name && t.input_schema)
  .map(t => ({ type: 'function', function: { name: t.name, description: t.description || '', parameters: t.input_schema } }))

function openAIRequest(body, model) {
  const tools = toTools(body)
  const choice = body.tool_choice?.type
  return {
    model, stream: true, stream_options: { include_usage: true },
    messages: toOpenAIMessages(body),
    ...(body.max_tokens && { max_tokens: body.max_tokens }),
    ...(body.temperature !== undefined && { temperature: body.temperature }),
    ...(body.top_p !== undefined && { top_p: body.top_p }),
    ...(body.stop_sequences?.length && { stop: body.stop_sequences }),
    ...(tools.length && choice !== 'none' && { tools }),
    ...(tools.length && choice === 'any' && { tool_choice: 'required' }),
    ...(tools.length && choice === 'tool' && { tool_choice: { type: 'function', function: { name: body.tool_choice.name } } }),
  }
}

function ollamaRequest(body, model, numCtx) {
  const names = new Map()     // tool call id -> tool name: Ollama matches results by name
  const messages = toOpenAIMessages(body).map(m => {
    if (m.role === 'assistant') {
      for (const c of m.tool_calls || []) names.set(c.id, c.function.name)
      return { role: 'assistant', content: m.content,
        ...(m.tool_calls && { tool_calls: m.tool_calls.map(c => ({ function: { name: c.function.name, arguments: JSON.parse(c.function.arguments) } })) }) }
    }
    if (m.role === 'tool') return { role: 'tool', content: m.content, tool_name: names.get(m.tool_call_id) }
    if (!Array.isArray(m.content)) return m
    // Ollama takes images as bare base64, so images given by URL are left out
    const images = m.content.filter(p => p.type === 'image_url').map(p => p.image_url.url.match(/^data:[^;]+;base64,(.*)$/s)?.[1]).filter(Boolean)
    return { role: m.role, content: m.content.filter(p => p.type === 'text').map(p => p.text).join('\n'), ...(images.length && { images }) }
  })
  const tools = toTools(body)
  return {
    model, stream: true, messages,
    ...(tools.length && body.tool_choice?.type !== 'none' && { tools }),
    options: {
      ...(body.max_tokens && { num_predict: body.max_tokens }),
      ...(body.temperature !== undefined && { temperature: body.temperature }),
      ...(body.top_p !== undefined && { top_p: body.top_p }),
      ...(body.stop_sequences?.length && { stop: body.stop_sequences }),
      num_ctx: numCtx,
    },
  }
}

// ---- upstream streams -> { text } | { tool: {id,name,args} } | { finish, input, output } ----
async function* lines(stream) {
  const decoder = new TextDecoder()
  let buf = ''
  for await (const chunk of stream) {
    buf += decoder.decode(chunk, { stream: true })
    let i
    while ((i = buf.indexOf('\n')) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (line) yield line }
  }
  if (buf.trim()) yield buf.trim()
}

async function* ollamaEvents(stream) {
  for await (const line of lines(stream)) {
    const j = JSON.parse(line)
    if (j.error) throw new UpstreamError(502, String(j.error))
    if (j.message?.content) yield { text: j.message.content }
    for (const c of j.message?.tool_calls || []) yield { tool: { name: c.function?.name, args: c.function?.arguments ?? {} } }
    if (j.done) yield { finish: j.done_reason, input: j.prompt_eval_count, output: j.eval_count }
  }
}

async function* openAIEvents(stream) {
  const calls = []            // tool calls arrive in fragments, keyed by index
  let finish, usage
  for await (const line of lines(stream)) {
    if (!line.startsWith('data:')) continue
    const data = line.slice(5).trim()
    if (data === '[DONE]') break
    const j = JSON.parse(data)
    if (j.error) throw new UpstreamError(502, j.error.message || JSON.stringify(j.error))
    if (j.usage) usage = j.usage
    const choice = j.choices?.[0]
    if (!choice) continue
    if (choice.delta?.content) yield { text: choice.delta.content }
    for (const c of choice.delta?.tool_calls || []) {
      let slot = calls[c.index ?? 0]
      // some servers reuse index 0 for every call and send each one whole: a new id means a new call
      if (!slot || (c.id && slot.id && c.id !== slot.id)) { slot = { id: c.id, name: '', args: '' }; if (calls[c.index ?? 0]) calls.push(slot); else calls[c.index ?? 0] = slot }
      if (c.function?.name) slot.name = c.function.name
      if (c.function?.arguments) slot.args += c.function.arguments
    }
    if (choice.finish_reason) finish = choice.finish_reason
  }
  for (const c of calls.filter(Boolean)) {
    let args = {}
    try { args = c.args ? JSON.parse(c.args) : {} } catch { log(`tool call ${c.name}: arguments were not valid JSON, sent as {}`) }
    yield { tool: { id: c.id, name: c.name, args } }
  }
  yield { finish, input: usage?.prompt_tokens, output: usage?.completion_tokens }
}

async function upstream(body, signal) {
  const f = await detect()
  const model = await resolveModel(body.model)
  const path = f === 'ollama' ? '/api/chat' : '/v1/chat/completions'
  const r = await fetch(BASE + path, {
    method: 'POST', headers: upstreamHeaders(), signal,
    body: JSON.stringify(f === 'ollama' ? ollamaRequest(body, model, await contextFor(model)) : openAIRequest(body, model)),
  })
  if (!r.ok) {
    const detail = (await r.text().catch(() => '')).slice(0, 2000)
    throw new UpstreamError(r.status, `${f} server ${path} (model ${model}): HTTP ${r.status} ${detail}`)
  }
  return { model, events: f === 'ollama' ? ollamaEvents(r.body) : openAIEvents(r.body) }
}

// ---- events -> Anthropic response -----------------------------------------------------
const newId = p => p + randomBytes(12).toString('hex')

// Collect the whole reply as Anthropic content blocks, reporting each text piece and each
// tool call as it arrives
async function collect(events, onText, onTool) {
  const content = []
  let finish, input = 0, output = 0
  for await (const e of events) {
    if (e.text !== undefined) {
      const last = content.at(-1)
      if (last?.type === 'text') last.text += e.text; else content.push({ type: 'text', text: e.text })
      onText?.(e.text, content.length - 1, last?.type !== 'text')
    } else if (e.tool) {
      content.push({ type: 'tool_use', id: e.tool.id || newId('toolu_local_'), name: e.tool.name, input: e.tool.args })
      onTool?.(content.at(-1), content.length - 1)
    } else {
      finish = e.finish; input = e.input ?? 0; output = e.output ?? 0
    }
  }
  const stop_reason = content.some(b => b.type === 'tool_use') ? 'tool_use' : finish === 'length' ? 'max_tokens' : 'end_turn'
  return { content, stop_reason, usage: { input_tokens: input, output_tokens: output } }
}

const envelope = (requested, extra) => ({
  id: newId('msg_local_'), type: 'message', role: 'assistant', model: requested,
  content: [], stop_reason: null, stop_sequence: null, usage: { input_tokens: 0, output_tokens: 0 }, ...extra,
})

async function messages(req, res, body) {
  const abort = new AbortController()
  res.on('close', () => { if (!res.writableEnded) abort.abort() })
  const { model, events } = await upstream(body, abort.signal)
  log(`messages ${body.model} -> ${model} (${format}) stream=${!!body.stream} tools=${toTools(body).length}`)

  if (!body.stream) {
    return sendJson(res, 200, envelope(body.model, await collect(events)))
  }
  res.writeHead(200, { 'Content-Type': 'text/event-stream', 'Cache-Control': 'no-cache', Connection: 'keep-alive' })
  const emit = (type, data) => res.write(`event: ${type}\ndata: ${JSON.stringify({ type, ...data })}\n\n`)
  const ping = setInterval(() => emit('ping', {}), PING_MS)    // local models can take a while to load
  try {
    emit('message_start', { message: envelope(body.model) })
    // Text is streamed as it arrives. Each tool call is sent whole: the arguments are small,
    // and Ollama only delivers them whole anyway. Blocks must start in index order.
    let open = -1
    const closeText = () => { if (open >= 0) emit('content_block_stop', { index: open }); open = -1 }
    const done = await collect(events, (piece, index, isNew) => {
      if (isNew) { closeText(); emit('content_block_start', { index, content_block: { type: 'text', text: '' } }); open = index }
      emit('content_block_delta', { index, delta: { type: 'text_delta', text: piece } })
    }, (block, index) => {
      closeText()
      emit('content_block_start', { index, content_block: { ...block, input: {} } })
      emit('content_block_delta', { index, delta: { type: 'input_json_delta', partial_json: JSON.stringify(block.input) } })
      emit('content_block_stop', { index })
    })
    closeText()
    emit('message_delta', { delta: { stop_reason: done.stop_reason, stop_sequence: null }, usage: done.usage })
    emit('message_stop', {})
  } catch (e) {
    if (!abort.signal.aborted) { log('stream failed:', e.message); emit('error', { error: { type: 'api_error', message: e.message } }) }
  } finally {
    clearInterval(ping)
    res.end()
  }
}

// ---- HTTP -------------------------------------------------------------------------------
function sendJson(res, status, body) {
  res.writeHead(status, { 'Content-Type': 'application/json' })
  res.end(JSON.stringify(body))
}
const sendError = (res, status, message) => sendJson(res, status, {
  type: 'error',
  error: { type: status === 404 ? 'not_found_error' : status === 400 || status === 413 ? 'invalid_request_error' : 'api_error', message },
})

function readJson(req) {
  return new Promise((resolve, reject) => {
    let size = 0
    const chunks = []
    req.on('data', c => { size += c.length; if (size > MAX_BODY) { reject(new UpstreamError(413, 'body too large')); req.destroy() } else chunks.push(c) })
    req.on('end', () => { try { resolve(JSON.parse(Buffer.concat(chunks).toString('utf8'))) } catch { reject(new UpstreamError(400, 'body is not valid JSON')) } })
    req.on('error', reject)
  })
}

// Routes under /v1. The caller has already checked the origin and the bridge token.
export async function handleLLM(req, res, path) {
  try {
    if (req.method === 'GET' && path === '/v1/models') {
      const data = (await listModels()).map(name => ({ type: 'model', id: PREFIX + name, display_name: name, created_at: '1970-01-01T00:00:00Z' }))
      return sendJson(res, 200, { data, has_more: false, first_id: data[0]?.id ?? null, last_id: data.at(-1)?.id ?? null })
    }
    if (req.method === 'POST' && path === '/v1/messages') return await messages(req, res, await readJson(req))
    if (req.method === 'POST' && path === '/v1/messages/count_tokens') {
      // no local tokenizer: about 4 characters per token is close enough for the add-in's meter
      return sendJson(res, 200, { input_tokens: Math.ceil(JSON.stringify(await readJson(req)).length / 4) })
    }
    return sendError(res, 404, `not served by the local model gateway: ${req.method} ${path}`)
  } catch (e) {
    const refused = e.cause?.code === 'ECONNREFUSED' || e.name === 'TimeoutError'
    const message = refused ? `cannot reach the local model server at ${BASE}. Is it running?` : e.message
    log(`${req.method} ${path} failed:`, message)
    if (res.headersSent) return res.end()
    // A bad request stays a 4xx. Everything else is a 502, including an upstream 401: the add-in
    // would read that as its own bridge token being wrong.
    sendError(res, e instanceof UpstreamError && [400, 404, 413, 429].includes(e.status) ? e.status : 502, message)
  }
}
