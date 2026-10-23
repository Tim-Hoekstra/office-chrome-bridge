#!/usr/bin/env node
// Build local/manifest.xml: the Claude for Office add-in manifest with the bridge attached
// as an `mcp_servers` entry. The entry (it carries the bridge token) goes in the URL
// fragment so it never leaves the machine in a request; allow_1p=1 keeps Claude.ai sign-in.
//
//   node scripts/build-manifest.mjs            # from manifest/template.xml
//   node scripts/build-manifest.mjs --fetch    # refresh the template from pivot.claude.ai first
import { createHash } from 'node:crypto'
import { chmodSync, existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
const TEMPLATE = join(ROOT, 'manifest', 'template.xml')
const LOCAL = join(ROOT, 'local')
const OUT = join(LOCAL, 'manifest.xml')
const CANONICAL = 'https://pivot.claude.ai/manifest.xml'
const PORT = Number(process.env.BRIDGE_PORT || 47823)

// Every URL slot Office reads the taskpane from must carry the same params.
const URL_SLOTS = [/(<SourceLocation\s+DefaultValue=")([^"]+)(")/g, /(id="Taskpane\.Url"\s+DefaultValue=")([^"]+)(")/g]

if (process.argv.includes('--fetch')) {
  const res = await fetch(CANONICAL)
  if (!res.ok) throw new Error(`fetch ${CANONICAL}: ${res.status}`)
  writeFileSync(TEMPLATE, await res.text())
  console.log('refreshed', TEMPLATE)
}

const tokenFile = join(LOCAL, '.token')
if (!existsSync(tokenFile)) {
  console.error('local/.token missing: start the bridge once (npm start) to generate it')
  process.exit(1)
}
const token = readFileSync(tokenFile, 'utf8').trim()
// Token-derived path segment: the add-in caches registrations per URL, so a new token
// needs a new URL to be picked up (see KEY_ID in src/bridge.mjs).
const keyId = createHash('sha256').update(token).digest('hex').slice(0, 12)
const url = `https://localhost:${PORT}/mcp/${keyId}`

const query = new URLSearchParams({ allow_1p: '1' })
const fragment = new URLSearchParams({
  mcp_servers: JSON.stringify([
    { url, label: 'Claude in Chrome', headers: { Authorization: `Bearer ${token}` } },
  ]),
})
const amp = s => s.replaceAll('&', '&amp;')   // XML attribute escaping

let xml = readFileSync(TEMPLATE, 'utf8')
for (const slot of URL_SLOTS) {
  let n = 0
  xml = xml.replace(slot, (_, pre, url, post) => {
    n++
    if (url.includes('#')) throw new Error(`template URL already has a fragment: ${url}`)
    return pre + url + (url.includes('?') ? '&amp;' : '?') + amp(query.toString()) + '#' + amp(fragment.toString()) + post
  })
  if (!n) throw new Error(`template missing URL slot ${slot.source}`)
}

mkdirSync(LOCAL, { recursive: true, mode: 0o700 })
writeFileSync(OUT, xml, { mode: 0o600 })
chmodSync(OUT, 0o600)
const version = xml.match(/<Version>([^<]+)<\/Version>/)?.[1]
console.log(`wrote ${OUT} (add-in ${version}, bridge ${url}). It contains the bridge token: keep it private.`)
