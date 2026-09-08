import { EventEmitter } from 'node:events'
import type { IncomingMessage, ServerResponse } from 'node:http'
import { PassThrough } from 'node:stream'
import { describe, expect, it, vi } from 'vitest'
import { createAgentProposalHandler } from '../agentProposalHost'

const body = { request: 'Mark launch ready', context: { workbook: { cell: 'C5' } }, capabilities: [], consent: true }
const operation = { name: 'xlsx.cell.set_value', input: { sheetId: 'sheet1', cell: { row: 4, column: 2 }, value: 'Ready' }, operationId: 'edit-1' }
const response = (value: unknown = { operations: [operation] }) => new Response(JSON.stringify(value), {
  headers: { 'Content-Type': 'application/json' },
})
const environment = { INJOFFICE_AGENT_PROPOSAL_URL: 'https://proposals.example.test/private?secret=hidden', INJOFFICE_AGENT_PROPOSAL_TOKEN: 'private-token' }
type Handler = ReturnType<typeof createAgentProposalHandler>
const mockBody = { request: 'Mark Security as Ready', capabilities: [{ name: 'xlsx.cell.set_value' }],
  context: { constraints: { maxOperations: 1, allowedValues: ['Ready'],
    allowedTargets: [{ sheetId: 'sheet1', row: 4, column: 2, ref: 'C5', workstream: 'Security' }] } } }

function request(handler: Handler, options: {
  method?: string; path?: string; headers?: Record<string, string | undefined>; value?: unknown; raw?: string
} = {}) {
  const req = Object.assign(new PassThrough(), {
    method: options.method ?? 'POST', url: options.path ?? '/api/agent/propose',
    headers: { host: 'localhost:3100', origin: 'http://localhost:3100', 'content-type': 'application/json', ...options.headers },
  }) as unknown as IncomingMessage
  let output = ''
  const headers: Record<string, string> = {}
  const res = Object.assign(new EventEmitter(), {
    statusCode: 200,
    setHeader: (key: string, value: string) => { headers[key] = value },
    end: (value: string) => { output = value },
  }) as unknown as ServerResponse
  const done = handler(req, res).then((handled) => ({ handled, status: res.statusCode, headers, body: output ? JSON.parse(output) : null }))
  ;(req as unknown as PassThrough).end(options.raw ?? JSON.stringify(options.value ?? body))
  return { done, res }
}

describe('local proposal-only host', () => {
  it.each([{}, environment])('serves the built-in mock without consent, credentials, or outward calls', async (env) => {
    const fetcher = vi.fn(() => { throw new Error('Must not contact provider') })
    const handler = createAgentProposalHandler({ env, fetch: fetcher })
    const result = await request(handler, { path: '/api/agent/mock-propose', value: mockBody }).done
    expect(result.status).toBe(200)
    expect(result.body).toEqual({ operations: [{ ...operation, operationId: 'mock-status-4-2' }] })
    expect(JSON.stringify(result)).not.toContain('private-token')
    expect(fetcher).not.toHaveBeenCalled()
  })

  it('does not even read the server token for a mock request', async () => {
    const env = { INJOFFICE_AGENT_PROPOSAL_URL: 'https://provider.example.test',
      get INJOFFICE_AGENT_PROPOSAL_TOKEN(): string { throw new Error('Token must stay unused') } }
    expect((await request(createAgentProposalHandler({ env }), { path: '/api/agent/mock-propose', value: mockBody }).done).status).toBe(200)
  })

  it('keeps mock refusals distinct from malformed input without exposing arbitrary context', async () => {
    const handler = createAgentProposalHandler({ env: {} })
    const result = await request(handler, { path: '/api/agent/mock-propose', value: { ...mockBody, request: 'Mark PRIVATE UNKNOWN as Ready' } }).done
    expect(result.status).toBe(422)
    expect(JSON.stringify(result.body)).not.toContain('PRIVATE UNKNOWN')
    expect((await request(handler, { path: '/api/agent/mock-propose', raw: '{bad' }).done).status).toBe(400)
    expect((await request(handler, { path: '/api/agent/mock-propose', value: { ...mockBody, approved: true } }).done).status).toBe(400)
  })

  it.each([
    [{ method: 'GET' }, 405], [{ headers: { origin: undefined } }, 403],
    [{ headers: { host: 'evil.test:3100' } }, 403], [{ headers: { origin: 'http://localhost:9999' } }, 403],
    [{ headers: { 'content-type': 'text/plain' } }, 415],
    [{ headers: { 'content-length': '50000' } }, 413],
    [{ raw: JSON.stringify({ ...mockBody, extra: 'x'.repeat(49 * 1024) }) }, 413],
  ])('preserves HTTP guards on the mock route', async (options, status) => {
    const fetcher = vi.fn()
    const handler = createAgentProposalHandler({ env: environment, fetch: fetcher })
    expect((await request(handler, { ...options, path: '/api/agent/mock-propose', value: mockBody }).done).status).toBe(status)
    expect(fetcher).not.toHaveBeenCalled()
  })

  it('shares the deadline and concurrency guard with live requests', async () => {
    const fetcher = vi.fn().mockImplementation(() => new Promise(() => {}))
    const handler = createAgentProposalHandler({ env: environment, fetch: fetcher, timeoutMs: 10 })
    const live = request(handler)
    expect((await request(handler, { path: '/api/agent/mock-propose', value: mockBody }).done).status).toBe(429)
    expect((await live.done).status).toBe(504)
    expect((await request(handler, { path: '/api/agent/mock-propose', value: mockBody }).done).status).toBe(200)
    expect(fetcher).toHaveBeenCalledTimes(1)
  })

  it('times out an unfinished mock request body and releases its listeners and slot', async () => {
    const fetcher = vi.fn()
    const handler = createAgentProposalHandler({ env: {}, fetch: fetcher, timeoutMs: 10 })
    const req = Object.assign(new PassThrough(), { method: 'POST', url: '/api/agent/mock-propose',
      headers: { host: 'localhost:3100', origin: 'http://localhost:3100', 'content-type': 'application/json' } })
    let result = ''
    const res = Object.assign(new EventEmitter(), { statusCode: 200, setHeader: () => {}, end: (value: string) => { result = value } })
    await handler(req as unknown as IncomingMessage, res as unknown as ServerResponse)
    expect(res.statusCode).toBe(504)
    expect(JSON.parse(result).error).toContain('timed out')
    expect(req.listenerCount('data')).toBe(0)
    expect(req.listenerCount('end')).toBe(0)
    expect(res.listenerCount('close')).toBe(0)
    expect((await request(handler, { path: '/api/agent/mock-propose', value: mockBody }).done).status).toBe(200)
    expect(fetcher).not.toHaveBeenCalled()
    req.destroy()
  })

  it('passes through unrelated routes', async () => {
    expect((await request(createAgentProposalHandler(), { path: '/other' }).done).handled).toBe(false)
  })

  it('reports only safe configuration metadata without contacting the endpoint', async () => {
    const fetcher = vi.fn()
    const handler = createAgentProposalHandler({ env: environment, fetch: fetcher })
    const result = await request(handler, { path: '/api/agent/proposal-status', method: 'GET', headers: { origin: undefined } }).done
    expect(result.body).toEqual({ configured: true, destination: 'proposals.example.test' })
    expect(JSON.stringify(result)).not.toContain('private')
    expect(fetcher).not.toHaveBeenCalled()
  })

  it.each(['', 'http://remote.example.test', 'file:///secret', 'https://user:pass@example.test', 'https://example.test/#secret'])('fails closed for unconfigured or unsafe endpoint %s', async (url) => {
    const fetcher = vi.fn()
    const handler = createAgentProposalHandler({ env: { INJOFFICE_AGENT_PROPOSAL_URL: url }, fetch: fetcher })
    expect((await request(handler).done).status).toBe(503)
    expect(fetcher).not.toHaveBeenCalled()
  })

  it.each([
    { host: 'evil.test:3100' }, { host: 'localhost.evil.test' }, { host: 'evil@localhost:3100' },
    { origin: 'https://evil.test' }, { origin: 'http://localhost:3101' }, { origin: undefined },
    { host: 'localhost:3100', origin: 'null' },
  ])('rejects untrusted Host/Origin %j', async (headers) => {
    const fetcher = vi.fn()
    const handler = createAgentProposalHandler({ env: environment, fetch: fetcher })
    expect((await request(handler, { headers }).done).status).toBe(403)
    expect(fetcher).not.toHaveBeenCalled()
  })

  it.each([
    [{ ...body, consent: false }, 400], [{ ...body, request: 'x'.repeat(2001) }, 400],
    [{ ...body, request: '' }, 400], [{ ...body, context: [] }, 400],
    [{ ...body, capabilities: new Array(65).fill({}) }, 400],
    [{ ...body, url: 'https://evil.test' }, 400], [{ ...body, context: { text: 'x'.repeat(49 * 1024) } }, 413],
  ])('rejects invalid consent/schema/size before forwarding', async (value, status) => {
    const fetcher = vi.fn()
    const handler = createAgentProposalHandler({ env: environment, fetch: fetcher })
    expect((await request(handler, { value }).done).status).toBe(status)
    expect(fetcher).not.toHaveBeenCalled()
  })

  it('rejects non-JSON content and malformed JSON', async () => {
    const fetcher = vi.fn()
    const handler = createAgentProposalHandler({ env: environment, fetch: fetcher })
    expect((await request(handler, { headers: { 'content-type': 'text/plain' } }).done).status).toBe(415)
    expect((await request(handler, { raw: '{broken' }).done).status).toBe(400)
    expect((await request(handler, { method: 'GET' }).done).status).toBe(405)
    expect(fetcher).not.toHaveBeenCalled()
  })

  it('forwards only the bounded contract and returns only operations', async () => {
    const fetcher = vi.fn().mockResolvedValue(response({ operations: [{ ...operation, approval: true }], commentary: 'secret' }))
    const handler = createAgentProposalHandler({ env: environment, fetch: fetcher })
    const result = await request(handler).done
    expect(result.body).toEqual({ operations: [operation] })
    expect(result.headers['Cache-Control']).toBe('no-store')
    const [url, options] = fetcher.mock.calls[0]
    expect(url.href).toBe(environment.INJOFFICE_AGENT_PROPOSAL_URL)
    expect(options.redirect).toBe('error')
    expect(options.headers.Authorization).toBe('Bearer private-token')
    expect(JSON.parse(options.body)).toEqual({ request: body.request, context: body.context, capabilities: [] })
  })

  it('allows explicitly configured loopback HTTP endpoints', async () => {
    const handler = createAgentProposalHandler({ env: { INJOFFICE_AGENT_PROPOSAL_URL: 'http://127.0.0.1:3999/propose' }, fetch: vi.fn().mockResolvedValue(response()) })
    expect((await request(handler).done).status).toBe(200)
  })

  it.each([
    {}, { operations: [] }, { operations: new Array(9).fill(operation) },
    { operations: [{ name: 'xlsx.set_value', input: [] }] },
    { operations: [{ name: 'x', input: {}, operationId: 1 }] },
    { operations: [{ name: 'x', input: {}, operationId: '' }] },
    { operations: [{ name: 'x'.repeat(121), input: {} }] },
    { operations: [{ name: 'x', input: JSON.parse('{"__proto__":{"polluted":true}}') }] },
  ])('rejects malformed operations', async (value) => {
    const handler = createAgentProposalHandler({ env: environment, fetch: vi.fn().mockResolvedValue(response(value)) })
    expect((await request(handler).done).status).toBe(502)
  })

  it('bounds responses and hides raw upstream errors', async () => {
    for (const upstream of [
      response({ operations: [{ name: 'x', input: { text: 'x'.repeat(33 * 1024) } }] }),
      new Response('secret error', { status: 401 }), new Response('secret HTML', { headers: { 'Content-Type': 'text/html' } }),
      new Response('broken secret JSON', { headers: { 'Content-Type': 'application/json' } }),
    ]) {
      const result = await request(createAgentProposalHandler({ env: environment, fetch: vi.fn().mockResolvedValue(upstream) })).done
      expect(result.status).toBe(502)
      expect(JSON.stringify(result.body)).not.toContain('secret')
    }
    const result = await request(createAgentProposalHandler({ env: environment, fetch: vi.fn().mockRejectedValue(new Error('private-token')) })).done
    expect(JSON.stringify(result.body)).not.toContain('private-token')
  })

  it('limits concurrency, times out, aborts upstream and releases the slot', async () => {
    const fetcher = vi.fn().mockImplementationOnce(() => new Promise(() => {})).mockResolvedValue(response())
    const handler = createAgentProposalHandler({ env: environment, fetch: fetcher, timeoutMs: 10 })
    const first = request(handler)
    expect((await request(handler).done).status).toBe(429)
    expect((await first.done).status).toBe(504)
    expect(fetcher.mock.calls[0][1].signal.aborted).toBe(true)
    expect((await request(handler).done).status).toBe(200)
  })

  it('cannot configure a deadline longer than 15 seconds', async () => {
    vi.useFakeTimers()
    try {
      const fetcher = vi.fn().mockImplementation(() => new Promise(() => {}))
      const handler = createAgentProposalHandler({ env: environment, fetch: fetcher, timeoutMs: 60_000 })
      const pending = request(handler)
      await vi.advanceTimersByTimeAsync(15_000)
      expect((await pending.done).status).toBe(504)
    } finally { vi.useRealTimers() }
  })

  it('aborts upstream on browser disconnection', async () => {
    let signal: AbortSignal | undefined
    const fetcher = vi.fn().mockImplementation((_url, options) => {
      signal = options.signal
      return new Promise((_resolve, reject) => signal!.addEventListener('abort', () => reject(new Error('aborted')), { once: true }))
    })
    const pending = request(createAgentProposalHandler({ env: environment, fetch: fetcher }))
    await vi.waitFor(() => expect(fetcher).toHaveBeenCalled())
    pending.res.emit('close')
    await pending.done
    expect(signal?.aborted).toBe(true)
  })
})
