import { MockAgentProposalError, type MockAgentProposal } from './mockAgentContract.ts'

const object = (value: unknown): value is Record<string, unknown> => typeof value === 'object' && value !== null && !Array.isArray(value)
function fail(message: string): never { throw new MockAgentProposalError(400, message) }
function refuse(message: string): never { throw new MockAgentProposalError(422, message) }
const DEGREES = [90, 180, 270, -90, -180, -270]

/** Bounded, deterministic page-rotation proposer. No model, network, approval, or write. */
export function createMockPdfProposal(body: Record<string, unknown>): MockAgentProposal {
  if (!object(body.context) || body.context.format !== 'pdf' || typeof body.request !== 'string' || !Array.isArray(body.capabilities)) fail('Invalid PDF proposal context')
  if (!body.capabilities.some((item) => object(item) && item.name === 'pdf.page.rotate')) refuse('The disclosed PDF capabilities do not support page rotation.')
  const constraints = body.context.constraints
  if (!object(constraints) || constraints.maxOperations !== 1 || !Array.isArray(constraints.allowedPages) || constraints.allowedPages.length < 1 || constraints.allowedPages.length > 100 ||
      constraints.allowedPages.some((page) => !Number.isSafeInteger(page) || Number(page) < 1) || new Set(constraints.allowedPages).size !== constraints.allowedPages.length ||
      !Array.isArray(constraints.allowedDegrees) || !constraints.allowedDegrees.length || constraints.allowedDegrees.some((value) => !DEGREES.includes(Number(value)) || typeof value !== 'number')) fail('Invalid PDF proposal constraints')
  const match = /^(?:please\s+)?rotate\s+page\s+(\d+)\s+by\s+(-?\d+)\s+degrees\.?$/i.exec(body.request.trim())
  if (!match) refuse('The PDF mock supports one rotation, for example “Rotate page 2 by 90 degrees”. It is not a language model.')
  const page = Number(match[1]); const degrees = Number(match[2])
  if (!constraints.allowedPages.includes(page)) refuse('The requested page is outside the disclosed PDF pages.')
  if (!constraints.allowedDegrees.includes(degrees)) refuse('Choose a disclosed rotation in 90-degree increments. No edit was proposed.')
  return { operations: [{ name: 'pdf.page.rotate', operationId: `mock-rotate-${page}`, input: { pages: [page], degrees } }] }
}
