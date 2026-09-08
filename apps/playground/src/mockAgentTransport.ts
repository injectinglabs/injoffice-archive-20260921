import { createMockAgentProposal, MockAgentProposalError } from './mockAgentProposal'

/** Same mock contract in both deployments; never falls back to a live provider. */
export async function requestMockAgentProposal(body: Record<string, unknown>, signal: AbortSignal, options: {
  http?: boolean; fetcher?: typeof fetch
} = {}): Promise<unknown> {
  signal.throwIfAborted()
  let response: Response
  if (options.http ?? import.meta.env.DEV) {
    response = await (options.fetcher ?? fetch)('/api/agent/mock-propose', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body), signal,
    })
  } else {
    // Static hosting has no API server. Simulate that response locally, using
    // exactly the same proposal generator as the development HTTP endpoint.
    await Promise.resolve()
    signal.throwIfAborted()
    try {
      response = Response.json(createMockAgentProposal(body))
    } catch (error) {
      if (!(error instanceof MockAgentProposalError)) throw error
      response = Response.json({ error: error.message }, { status: error.status })
    }
  }
  signal.throwIfAborted()
  if (!response.headers.get('content-type')?.includes('application/json')) throw new Error('The mock endpoint is unavailable. Restart the demo server and try again.')
  const result: unknown = await response.json()
  signal.throwIfAborted()
  if (!response.ok) {
    const message = result && typeof result === 'object' && 'error' in result && typeof result.error === 'string'
      ? result.error : 'The mock could not propose this request. Try one supported workstream status edit.'
    throw new Error(message)
  }
  return result
}
