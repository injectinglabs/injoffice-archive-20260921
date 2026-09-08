import type { IncomingMessage, ServerResponse } from 'node:http'

const REQUEST_LIMIT = 48 * 1024
const RESPONSE_LIMIT = 32 * 1024
const TIMEOUT_MS = 15_000
type JsonObject = Record<string, unknown>
type HostOptions = {
  env?: Record<string, string | undefined>
  fetch?: typeof globalThis.fetch
  /** Tests may shorten, but cannot extend, the production deadline. */
  timeoutMs?: number
}

class RequestFailure extends Error {
  readonly status: number
  constructor(status: number, message: string) { super(message); this.status = status }
}

function object(value: unknown): value is JsonObject {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

function plainJson(value: unknown, depth = 0): boolean {
  if (depth > 24) return false
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return true
  if (typeof value === 'number') return Number.isFinite(value)
  if (Array.isArray(value)) return value.every((entry) => plainJson(entry, depth + 1))
  return object(value) && Object.entries(value).every(([key, entry]) =>
    !['__proto__', 'constructor', 'prototype'].includes(key) && plainJson(entry, depth + 1))
}

function loopback(hostname: string): boolean {
  return ['localhost', '127.0.0.1', '[::1]'].includes(hostname)
}

function endpointFrom(env: HostOptions['env']): URL | undefined {
  try {
    const url = new URL(env?.INJOFFICE_AGENT_PROPOSAL_URL ?? '')
    if (url.username || url.password || url.hash) return undefined
    if (url.protocol !== 'https:' && !(url.protocol === 'http:' && loopback(url.hostname))) return undefined
    return url
  } catch { return undefined }
}

function trustedRequest(req: IncomingMessage, requireOrigin: boolean): boolean {
  const host = req.headers.host
  if (!host || /[\s/@?#\\]/.test(host)) return false
  try {
    const expected = new URL(`http://${host}`)
    if (!loopback(expected.hostname) || expected.host !== host.toLowerCase()) return false
    const origin = req.headers.origin
    if (!origin) return !requireOrigin
    // This is deliberately a local HTTP demo host, not a public authenticated proxy.
    return origin === expected.origin
  } catch { return false }
}

function send(res: ServerResponse, status: number, body: JsonObject) {
  res.statusCode = status
  res.setHeader('Content-Type', 'application/json; charset=utf-8')
  res.setHeader('Cache-Control', 'no-store')
  res.setHeader('X-Content-Type-Options', 'nosniff')
  res.end(JSON.stringify(body))
}

function readRequest(req: IncomingMessage, signal: AbortSignal): Promise<unknown> {
  return new Promise((resolve, reject) => {
    let length = 0
    const chunks: Buffer[] = []
    const cleanup = () => {
      req.off('data', onData); req.off('end', onEnd); req.off('error', onError)
      signal.removeEventListener('abort', onAbort)
    }
    const fail = (error: RequestFailure) => { cleanup(); req.resume(); reject(error) }
    const onAbort = () => fail(new RequestFailure(504, 'Proposal request timed out'))
    const onError = () => fail(new RequestFailure(400, 'Could not read request'))
    const onData = (chunk: Buffer | string) => {
      const bytes = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk)
      length += bytes.length
      if (length > REQUEST_LIMIT) { fail(new RequestFailure(413, 'Request is too large')); return }
      chunks.push(bytes)
    }
    const onEnd = () => {
      cleanup()
      try { resolve(JSON.parse(Buffer.concat(chunks).toString('utf8'))) }
      catch { reject(new RequestFailure(400, 'Valid JSON object required')) }
    }
    req.on('data', onData); req.on('end', onEnd); req.on('error', onError)
    signal.addEventListener('abort', onAbort, { once: true })
    if (signal.aborted) onAbort()
  })
}

function validateRequest(body: unknown): JsonObject {
  if (!object(body) || !plainJson(body) || Object.keys(body).some((key) =>
    !['request', 'context', 'capabilities', 'consent'].includes(key))) {
    throw new RequestFailure(400, 'Invalid proposal request')
  }
  if (body.consent !== true) throw new RequestFailure(400, 'Explicit data-sharing consent required')
  if (typeof body.request !== 'string' || !body.request.trim() || body.request.length > 2000 ||
      !object(body.context) || !Array.isArray(body.capabilities) || body.capabilities.length > 64) {
    throw new RequestFailure(400, 'Invalid request, context, or capabilities')
  }
  return { request: body.request, context: body.context, capabilities: body.capabilities }
}

async function readResponse(response: Response): Promise<unknown> {
  if (!response.ok || !/^application\/json(?:\s*;|$)/i.test(response.headers.get('content-type') ?? '')) {
    await response.body?.cancel()
    throw new RequestFailure(502, 'Proposal endpoint returned an invalid response')
  }
  const reader = response.body?.getReader()
  if (!reader) throw new RequestFailure(502, 'Proposal endpoint returned an empty response')
  let size = 0
  const chunks: Uint8Array[] = []
  try {
    for (;;) {
      const { done, value } = await reader.read()
      if (done) break
      size += value.byteLength
      if (size > RESPONSE_LIMIT) throw new RequestFailure(502, 'Proposal response is too large')
      chunks.push(value)
    }
    return JSON.parse(Buffer.concat(chunks).toString('utf8'))
  } finally { await reader.cancel().catch(() => {}) }
}

function validateProposal(value: unknown): JsonObject {
  if (!object(value) || !plainJson(value) || !Array.isArray(value.operations) ||
      value.operations.length < 1 || value.operations.length > 8) {
    throw new RequestFailure(502, 'Proposal endpoint returned invalid operations')
  }
  const operations = value.operations.map((op: unknown) => {
    if (!object(op) || typeof op.name !== 'string' || !/^[a-zA-Z][\w.-]{0,119}$/.test(op.name) ||
        !object(op.input) || (op.operationId !== undefined &&
        (typeof op.operationId !== 'string' || op.operationId.length < 1 || op.operationId.length > 120))) {
      throw new RequestFailure(502, 'Proposal endpoint returned invalid operations')
    }
    // Never relay provider commentary, headers, approval flags, or hidden metadata.
    return { name: op.name, input: op.input, ...(op.operationId === undefined ? {} : { operationId: op.operationId }) }
  })
  return { operations }
}

/** Provider-neutral proposal-only bridge. It never approves, commits, or calls Office tools. */
export function createAgentProposalHandler(options: HostOptions = {}) {
  const env = options.env ?? process.env
  const fetcher = options.fetch ?? globalThis.fetch
  const timeoutMs = Math.max(1, Math.min(TIMEOUT_MS, options.timeoutMs ?? TIMEOUT_MS))
  let busy = false
  return async (req: IncomingMessage, res: ServerResponse): Promise<boolean> => {
    const path = req.url?.split('?')[0]
    if (path !== '/api/agent/proposal-status' && path !== '/api/agent/propose') return false
    if (!trustedRequest(req, path === '/api/agent/propose')) {
      send(res, 403, { error: 'Same-origin loopback demo requests only' }); return true
    }
    const endpoint = endpointFrom(env)
    if (path === '/api/agent/proposal-status') {
      if (req.method !== 'GET') send(res, 405, { error: 'GET required' })
      else send(res, 200, { configured: Boolean(endpoint), ...(endpoint ? { destination: endpoint.hostname } : {}) })
      return true
    }
    if (req.method !== 'POST') { send(res, 405, { error: 'POST required' }); return true }
    if (!/^application\/json(?:\s*;|$)/i.test(req.headers['content-type'] ?? '')) {
      send(res, 415, { error: 'application/json required' }); return true
    }
    if (Number(req.headers['content-length']) > REQUEST_LIMIT) {
      send(res, 413, { error: 'Request is too large' }); return true
    }
    if (!endpoint) { send(res, 503, { error: 'Live proposal endpoint is not configured' }); return true }
    if (busy) { send(res, 429, { error: 'A proposal request is already running' }); return true }
    busy = true
    const controller = new AbortController()
    const onDisconnect = () => controller.abort()
    res.once('close', onDisconnect)
    let timer: ReturnType<typeof setTimeout> | undefined
    try {
      const timeout = new Promise<never>((_, reject) => {
        timer = setTimeout(() => {
          controller.abort()
          reject(new RequestFailure(504, 'Proposal request timed out'))
        }, timeoutMs)
      })
      const operation = async () => {
        const body = validateRequest(await readRequest(req, controller.signal))
        controller.signal.throwIfAborted()
        const token = env.INJOFFICE_AGENT_PROPOSAL_TOKEN
        const response = await fetcher(endpoint, {
          method: 'POST', redirect: 'error', signal: controller.signal,
          headers: { 'Content-Type': 'application/json', ...(token ? { Authorization: `Bearer ${token}` } : {}) },
          body: JSON.stringify(body),
        })
        return validateProposal(await readResponse(response))
      }
      send(res, 200, await Promise.race([operation(), timeout]))
    } catch (error) {
      const failure = error instanceof RequestFailure ? error : new RequestFailure(502, 'Proposal endpoint unavailable or invalid')
      send(res, failure.status, { error: failure.message })
    } finally {
      clearTimeout(timer); controller.abort(); busy = false
      res.off('close', onDisconnect)
    }
    return true
  }
}

export const handleAgentProposalRequest = createAgentProposalHandler()
