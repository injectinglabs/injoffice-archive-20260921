import { MockAgentProposalError, type MockAgentProposal } from './mockAgentContract.ts'

const object = (value: unknown): value is Record<string, unknown> => !!value && typeof value === 'object' && !Array.isArray(value)
export const DOCX_MOCK_TEXT_LIMIT = 1000
export function isMockDocxText(value: unknown): value is string {
  if (typeof value !== 'string' || !value.trim() || value.length > DOCX_MOCK_TEXT_LIMIT) return false
  for (const character of value) {
    const code = character.codePointAt(0)!
    if (!(code === 9 || code === 10 || code === 13 || code >= 0x20 && code <= 0xd7ff || code >= 0xe000 && code <= 0xfffd || code >= 0x10000 && code <= 0x10ffff)) return false
  }
  return true
}
export function createMockDocxProposal(body: Record<string, unknown>): MockAgentProposal {
  function fail(message: string): never { throw new MockAgentProposalError(400, message) }
  function refuse(message: string): never { throw new MockAgentProposalError(422, message) }
  if (typeof body.request !== 'string' || body.request.length > 2000 || !object(body.context) || body.context.format !== 'docx' || !Array.isArray(body.capabilities)) fail('Invalid DOCX mock context')
  if (!body.capabilities.some((item) => object(item) && item.name === 'docx.text.replace')) refuse('The disclosed capabilities do not support DOCX text replacement.')
  const constraints = body.context.constraints
  if (!object(constraints) || constraints.maxOperations !== 1 || constraints.maxTextLength !== DOCX_MOCK_TEXT_LIMIT || !Array.isArray(constraints.allowedTargets) || constraints.allowedTargets.length > 128) fail('Invalid DOCX mock constraints')
  const targets = constraints.allowedTargets.map((item) => {
    if (!object(item) || item.targetKind !== 'run' || typeof item.targetId !== 'string' || !item.targetId || item.targetId.length > 256 || !isMockDocxText(item.expectedText) || typeof item.blockId !== 'string' || !item.blockId) fail('Invalid DOCX mock target')
    return { targetKind: 'run', targetId: item.targetId as string, expectedText: item.expectedText as string }
  })
  const match = /^\s*(?:please\s+)?replace\s+["“]([^"”]+)["”]\s+with\s+["“]([^"”]+)["”]\s*\.?\s*$/i.exec(body.request)
  if (!match) refuse('The mock supports one exact text replacement: Replace "Northstar Launch Brief" with "Northstar Beta Launch Brief". It is not a language model.')
  if (!isMockDocxText(match[1]) || !isMockDocxText(match[2])) refuse('Use nonempty XML-safe text of at most 1,000 characters.')
  if (match[1] === match[2]) refuse('The replacement is unchanged; no edit was proposed.')
  const matches = targets.filter((item) => item.expectedText === match[1])
  if (matches.length !== 1) refuse(matches.length ? 'The disclosed text is ambiguous; no edit was proposed.' : 'No disclosed editable run exactly matches the requested text.')
  return { operations: [{ name: 'docx.text.replace', operationId: 'mock-docx-text', input: { ...matches[0], text: match[2] } }] }
}
