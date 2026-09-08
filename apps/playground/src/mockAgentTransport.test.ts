import { describe, expect, it, vi } from 'vitest'
import { createMockAgentProposal } from './mockAgentProposal'
import { requestMockAgentProposal } from './mockAgentTransport'

const body = {
  request: 'Mark Mobile as On track', capabilities: [{ name: 'xlsx.cell.set_value' }],
  context: { constraints: { maxOperations: 1, allowedValues: ['Ready', 'On track'], allowedTargets: [
    { sheetId: 'sample', row: 2, column: 2, ref: 'C3', workstream: 'Mobile' },
  ] } },
}

describe('bundled mock transport', () => {
  it('uses the identical pure proposal in static builds without any fetch', async () => {
    const fetcher = vi.fn()
    expect(await requestMockAgentProposal(body, new AbortController().signal, { http: false, fetcher })).toEqual(createMockAgentProposal(body))
    expect(fetcher).not.toHaveBeenCalled()
  })

  it('uses only the same-origin mock endpoint in development', async () => {
    const fetcher = vi.fn().mockResolvedValue(Response.json(createMockAgentProposal(body)))
    const signal = new AbortController().signal
    expect(await requestMockAgentProposal(body, signal, { http: true, fetcher })).toEqual(createMockAgentProposal(body))
    expect(fetcher).toHaveBeenCalledExactlyOnceWith('/api/agent/mock-propose', {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body), signal,
    })
  })

  it('refuses unsupported static requests with the useful mock message', async () => {
    await expect(requestMockAgentProposal({ ...body, request: 'Write a poem' }, new AbortController().signal, { http: false })).rejects.toThrow('mock supports one status edit')
  })

  it('preserves HTTP refusal and never falls back to a different proposer', async () => {
    const fetcher = vi.fn().mockResolvedValue(Response.json({ error: 'No disclosed editable workstream matches that request.' }, { status: 422 }))
    await expect(requestMockAgentProposal(body, new AbortController().signal, { http: true, fetcher })).rejects.toThrow('No disclosed editable workstream')
    expect(fetcher).toHaveBeenCalledTimes(1)
  })

  it('reports unavailable HTTP endpoints without treating a static HTML page as a proposal', async () => {
    const fetcher = vi.fn().mockResolvedValue(new Response('<html>demo</html>', { headers: { 'Content-Type': 'text/html' } }))
    await expect(requestMockAgentProposal(body, new AbortController().signal, { http: true, fetcher })).rejects.toThrow('mock endpoint is unavailable')
  })

  it.each([true, false])('honors cancellation before starting http=%s', async (http) => {
    const controller = new AbortController()
    const fetcher = vi.fn()
    controller.abort()
    await expect(requestMockAgentProposal(body, controller.signal, { http, fetcher })).rejects.toMatchObject({ name: 'AbortError' })
    expect(fetcher).not.toHaveBeenCalled()
  })

  it('drops an HTTP response arriving after cancellation', async () => {
    const controller = new AbortController()
    const fetcher = vi.fn().mockImplementation(async () => {
      controller.abort()
      return Response.json(createMockAgentProposal(body))
    })
    await expect(requestMockAgentProposal(body, controller.signal, { http: true, fetcher })).rejects.toMatchObject({ name: 'AbortError' })
  })
})
