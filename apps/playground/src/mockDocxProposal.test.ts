import { describe, expect, it } from 'vitest'
import { createMockDocxProposal } from './mockDocxProposal'

function body(request = 'Replace "Northstar Launch Brief" with "Northstar Beta Launch Brief"') {
  return { request, capabilities: [{ name: 'docx.text.replace' }], context: { format: 'docx', constraints: { maxOperations: 1, maxTextLength: 1000,
    allowedTargets: [{ targetKind: 'run', targetId: 'run:1', expectedText: 'Northstar Launch Brief', blockId: 'paragraph:1' }] } } }
}
describe('bounded DOCX mock proposer', () => {
  it('selects a uniquely disclosed exact run and retains its expected text', () => {
    expect(createMockDocxProposal(body())).toEqual({ operations: [{ name: 'docx.text.replace', operationId: 'mock-docx-text', input: {
      targetKind: 'run', targetId: 'run:1', expectedText: 'Northstar Launch Brief', text: 'Northstar Beta Launch Brief',
    } }] })
    expect(createMockDocxProposal(body('Please replace “Northstar Launch Brief” with “A reviewed launch brief”.')).operations[0].input.text).toBe('A reviewed launch brief')
  })
  it.each(['Summarize this document', 'Replace "Missing" with "New"', 'Replace "Northstar Launch Brief" with "Northstar Launch Brief"',
    'Replace "Northstar Launch Brief" with "Bad\u0000XML"', 'Replace "Northstar Launch Brief" with ""',
    'Replace "Northstar Launch Brief" with "New" and replace "Another" with "Other"'])('refuses unsupported or unsafe request %s', (request) => {
    expect(() => createMockDocxProposal(body(request))).toThrow()
  })
  it('never guesses an ambiguous target or falls back after missing capabilities', () => {
    const value = body(); value.context.constraints.allowedTargets.push({ ...value.context.constraints.allowedTargets[0], targetId: 'run:2' })
    expect(() => createMockDocxProposal(value)).toThrow('ambiguous')
    expect(() => createMockDocxProposal({ ...body(), capabilities: [] })).toThrow('capabilities')
  })
  it.each([
    { format: 'xlsx' },
    { format: 'docx', constraints: { maxOperations: 1, maxTextLength: 1000, allowedTargets: [{ targetKind: 'paragraph', targetId: 'p', expectedText: 'text' }] } },
    { format: 'docx', constraints: { maxOperations: 2, maxTextLength: 1000, allowedTargets: [] } },
    { format: 'docx', constraints: { maxOperations: 1, maxTextLength: 9999, allowedTargets: [] } },
  ])('rejects malformed format constraints', (context) => {
    expect(() => createMockDocxProposal({ ...body(), context })).toThrow('Invalid DOCX')
  })
  it('bounds replacement length and rejects unmatched UTF-16 surrogates', () => {
    expect(() => createMockDocxProposal(body(`Replace "Northstar Launch Brief" with "${'a'.repeat(1001)}"`))).toThrow('1,000')
    expect(() => createMockDocxProposal(body('Replace "Northstar Launch Brief" with "\ud800"'))).toThrow('XML-safe')
  })
})
