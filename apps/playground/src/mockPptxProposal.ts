import { MockAgentProposalError, type MockAgentProposal } from './mockAgentContract.ts'

export type MockPptxTarget = { elementId: string; currentText: string; slideId: string }
const object = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value)
function fail(message: string): never { throw new MockAgentProposalError(400, message) }
function refuse(message: string): never { throw new MockAgentProposalError(422, message) }

/** A deliberately small simulated proposer. It never approves or writes a presentation. */
export function createMockPptxProposal(body: Record<string, unknown>): MockAgentProposal {
  if (!object(body.context) || body.context.format !== 'pptx' || typeof body.request !== 'string' || body.request.length > 2000 ||
      !Array.isArray(body.capabilities)) fail('Invalid presentation proposal context')
  if (!body.capabilities.some((item) => object(item) && item.name === 'pptx.native.text.replace')) refuse('The disclosed capabilities do not allow native presentation text edits.')
  const constraints = body.context.constraints
  if (!object(constraints) || constraints.maxOperations !== 1 || !Array.isArray(constraints.allowedTargets) || constraints.allowedTargets.length > 256) fail('Invalid presentation proposal constraints')
  const ids = new Set<string>()
  const targets: MockPptxTarget[] = constraints.allowedTargets.map((target) => {
    if (!object(target) || typeof target.elementId !== 'string' || !target.elementId || target.elementId.length > 300 ||
        typeof target.slideId !== 'string' || !target.slideId || target.slideId.length > 300 ||
        typeof target.currentText !== 'string' || !target.currentText || target.currentText.length > 1000 || /[\t\r\n]/.test(target.currentText) || ids.has(target.elementId)) fail('Invalid or duplicate presentation target')
    ids.add(target.elementId)
    return { elementId: target.elementId, slideId: target.slideId, currentText: target.currentText }
  })
  const match = /^\s*(?:please\s+)?replace\s+"([^"\r\n]+)"\s+with\s+"([^"\r\n]+)"\s*\.?\s*$/i.exec(body.request)
  if (!match) refuse('The mock supports one exact text edit: Replace "existing text" with "new text". It is not a language model.')
  if (match[2].length > 1000 || /[\t\r\n]/.test(match[2]) || match[1] === match[2]) refuse('Choose a changed, single-line replacement of at most 1,000 characters.')
  const matches = targets.filter((target) => target.currentText === match[1])
  if (matches.length !== 1) refuse('The requested text must match exactly one disclosed editable presentation element.')
  return { operations: [{ operationId: 'mock-pptx-text', name: 'pptx.native.text.replace', input: { elementId: matches[0].elementId, text: match[2] } }] }
}
