import { describe, expect, it, vi } from 'vitest'
import { createMockAgentProposal, MockAgentProposalError } from './mockAgentProposal'

const target = { sheetId: 'release-plan', row: 11, column: 5, ref: 'F12', workstream: 'Security' }
const body = { request: 'Mark Security as Ready', capabilities: [{ name: 'xlsx.cell.set_value' }],
  context: { constraints: { maxOperations: 1, allowedValues: ['Ready', 'On track', 'Review'],
    allowedTargets: [target, { ...target, row: 21, ref: 'F22', workstream: 'Mobile' }] } } }

function expectFailure(value: unknown, status: number) {
  try { createMockAgentProposal(value); throw new Error('Expected refusal') }
  catch (error) { expect(error).toBeInstanceOf(MockAgentProposalError); expect((error as MockAgentProposalError).status).toBe(status) }
}

describe('deterministic mock status proposer', () => {
  it('derives coordinates from disclosed targets, not sample cell addresses', () => {
    expect(createMockAgentProposal(body)).toEqual({ operations: [{ name: 'xlsx.cell.set_value', operationId: 'mock-status-11-5',
      input: { sheetId: 'release-plan', cell: { row: 11, column: 5 }, value: 'Ready' } }] })
  })

  it.each(['Mark Mobile as On track', 'Set Mobile status to On track', 'Update Mobile to On track', 'Please mark  MOBILE as on TRACK.'])('accepts the bounded request variant %s', (request) => {
    expect(createMockAgentProposal({ ...body, request }).operations[0].input).toEqual({ sheetId: 'release-plan', cell: { row: 21, column: 5 }, value: 'On track' })
  })

  it('is deterministic, does not mutate context, and never uses fetch', () => {
    const fetcher = vi.fn(() => { throw new Error('Network forbidden') })
    vi.stubGlobal('fetch', fetcher)
    try {
      const original = JSON.stringify(body)
      expect(createMockAgentProposal(body)).toEqual(createMockAgentProposal(structuredClone(body)))
      expect(JSON.stringify(body)).toBe(original)
      expect(fetcher).not.toHaveBeenCalled()
    } finally { vi.unstubAllGlobals() }
  })

  it.each(['Mark Unknown as Ready', 'Summarize this file', 'Run a macro', 'Mark Security as Done',
    'Mark Security as Blocked', 'Mark Security as Ready and Mobile as Review', 'Update Security and Mobile to Ready'])('refuses unsupported or multiple edits: %s', (request) => {
    expectFailure({ ...body, request }, 422)
  })

  it('refuses missing capabilities and ambiguous case-insensitive names', () => {
    expectFailure({ ...body, capabilities: [{ name: 'office.commit' }] }, 422)
    expectFailure({ ...body, context: { constraints: { ...body.context.constraints, allowedTargets: [target, { ...target, sheetId: 'other', workstream: 'SECURITY' }] } } }, 422)
  })

  it.each([null, [], {}, { ...body, approved: true }, { ...body, consent: true },
    { ...body, request: '' }, { ...body, request: 'x'.repeat(2001) }, { ...body, capabilities: new Array(65).fill({ name: 'x' }) },
    { ...body, capabilities: ['xlsx.cell.set_value'] }, { ...body, context: [] },
    { ...body, context: { text: 'x'.repeat(49 * 1024) } },
    { ...body, context: { text: '😀'.repeat(13 * 1024) } },
    { ...body, context: JSON.parse('{"__proto__":{"approved":true}}') },
  ])('rejects malformed or oversized request', (value) => expectFailure(value, 400))

  it.each([
    { maxOperations: 2 }, { allowedValues: ['=DANGEROUS()'] }, { allowedValues: [] }, { allowedTargets: new Array(257).fill(target) },
    { allowedTargets: [{ ...target, row: -1 }] }, { allowedTargets: [{ ...target, column: 16_384 }] },
    { allowedTargets: [{ ...target, row: 1.5 }] }, { allowedTargets: [{ ...target, ref: 'C5' }] },
    { allowedTargets: [{ ...target, sheetId: '' }] }, { allowedTargets: [{ ...target, workstream: ' ' }] },
  ])('rejects invalid constraints %j', (constraints) => {
    expectFailure({ ...body, context: { constraints: { ...body.context.constraints, ...constraints } } }, 400)
  })

  it('bounds recursion and rejects cycles and non-JSON values', () => {
    let nested: unknown = null
    for (let index = 0; index < 26; index++) nested = { nested }
    expectFailure({ ...body, context: { nested } }, 400)
    const cyclic: Record<string, unknown> = {}; cyclic.self = cyclic
    expectFailure({ ...body, context: cyclic }, 400)
    expectFailure({ ...body, context: { value: Infinity } }, 400)
    expectFailure({ ...body, context: { date: new Date() } }, 400)
  })
})
