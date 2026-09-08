/** Deterministic demo proposer shared by the local HTTP route and static preview.
 * It does not run an LLM, access the network, approve changes, or edit a workbook.
 */
import { MockAgentProposalError, type MockAgentProposal } from './mockAgentContract.ts'
import { createMockDocxProposal } from './mockDocxProposal.ts'
import { createMockPptxProposal } from './mockPptxProposal.ts'
import { createMockPdfProposal } from './mockPdfProposal.ts'
export { MockAgentProposalError, type MockAgentProposal } from './mockAgentContract.ts'

const PAYLOAD_LIMIT = 48 * 1024
const STATUS_VALUES = ['Ready', 'On track', 'At risk', 'Review', 'Blocked']
type JsonObject = Record<string, unknown>

function object(value: unknown): value is JsonObject {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

function validJson(value: unknown, budget: { nodes: number; bytes: number }, depth = 0): boolean {
  if (depth > 24 || --budget.nodes < 0) return false
  if (value === null || typeof value === 'boolean') return true
  if (typeof value === 'string') {
    if (value.length > PAYLOAD_LIMIT) return false
    budget.bytes -= new TextEncoder().encode(value).byteLength
    return budget.bytes >= 0
  }
  if (typeof value === 'number') return Number.isFinite(value)
  if (Array.isArray(value)) return value.length <= 5000 && value.every((entry) => validJson(entry, budget, depth + 1))
  if (!object(value) || ![Object.prototype, null].includes(Object.getPrototypeOf(value))) return false
  return Object.entries(value).every(([key, entry]) => !['__proto__', 'constructor', 'prototype'].includes(key)
    && validJson(key, budget, depth + 1) && validJson(entry, budget, depth + 1))
}

function fail(message = 'Invalid mock proposal request'): never {
  throw new MockAgentProposalError(400, message)
}

function refuse(message: string): never { throw new MockAgentProposalError(422, message) }
function normalized(value: string): string { return value.trim().replace(/\s+/g, ' ').toLowerCase() }
function cellRef(row: number, column: number): string {
  let letters = ''
  for (let index = column + 1; index > 0; index = Math.floor((index - 1) / 26)) letters = String.fromCharCode(65 + (index - 1) % 26) + letters
  return `${letters}${row + 1}`
}

export function createMockAgentProposal(body: unknown): MockAgentProposal {
  if (!object(body) || !validJson(body, { nodes: 20_000, bytes: PAYLOAD_LIMIT }) ||
      new TextEncoder().encode(JSON.stringify(body)).byteLength > PAYLOAD_LIMIT ||
      Object.keys(body).some((key) => !['request', 'context', 'capabilities'].includes(key))) fail()
  if (typeof body.request !== 'string' || !body.request.trim() || body.request.length > 2000 ||
      !object(body.context) || !Array.isArray(body.capabilities) || body.capabilities.length > 64 ||
      !body.capabilities.every((capability) => object(capability) && typeof capability.name === 'string')) fail()
  if (body.context.format === 'docx') return createMockDocxProposal(body)
  if (body.context.format === 'pptx') return createMockPptxProposal(body)
  if (body.context.format === 'pdf') return createMockPdfProposal(body)
  if (body.context.format !== undefined && body.context.format !== 'xlsx') refuse('The mock supports XLSX, DOCX, PPTX, and PDF samples only.')
  if (!body.capabilities.some((capability) => capability.name === 'xlsx.cell.set_value')) {
    refuse('The disclosed capabilities do not support status edits.')
  }
  const constraints = body.context.constraints
  if (!object(constraints) || constraints.maxOperations !== 1 || !Array.isArray(constraints.allowedValues) ||
      constraints.allowedValues.length < 1 || constraints.allowedValues.length > 16 ||
      !constraints.allowedValues.every((value) => typeof value === 'string' && STATUS_VALUES.includes(value)) ||
      !Array.isArray(constraints.allowedTargets) || constraints.allowedTargets.length > 256) fail('Invalid mock proposal constraints')
  const targets = constraints.allowedTargets.map((target) => {
    if (!object(target) || typeof target.sheetId !== 'string' || !target.sheetId || target.sheetId.length > 200 ||
        typeof target.workstream !== 'string' || !target.workstream.trim() || target.workstream.length > 200 ||
        typeof target.row !== 'number' || !Number.isInteger(target.row) || target.row < 0 || target.row >= 1_048_576 ||
        typeof target.column !== 'number' || !Number.isInteger(target.column) || target.column < 0 || target.column >= 16_384 ||
        target.ref !== cellRef(target.row, target.column)) fail('Invalid mock proposal target')
    return { sheetId: target.sheetId, row: target.row, column: target.column, workstream: target.workstream }
  })
  const request = body.request.trim().replace(/^please\s+/i, '').replace(/\.$/, '').trim()
  const match = /^(?:mark\s+(.+?)\s+(?:as|to)|set\s+(.+?)\s+status\s+to|update\s+(.+?)\s+to)\s+(.+)$/i.exec(request)
  if (!match) refuse('The mock supports one status edit, for example “Mark Security as Ready”. It is not a language model.')
  const requestedName = normalized(match[1] ?? match[2] ?? match[3])
  const value = constraints.allowedValues.find((allowed) => normalized(allowed as string) === normalized(match[4])) as string | undefined
  if (!value) refuse('Choose one disclosed status: Ready, On track, At risk, Review, or Blocked. Multiple edits are not supported.')
  const matches = targets.filter((target) => normalized(target.workstream) === requestedName)
  if (matches.length === 0) refuse('No disclosed editable workstream matches that request. Use a workstream shown in the sample.')
  if (matches.length !== 1) refuse('That workstream is ambiguous in the disclosed context; no edit was proposed.')
  const target = matches[0]
  return { operations: [{ name: 'xlsx.cell.set_value', operationId: `mock-status-${target.row}-${target.column}`,
    input: { sheetId: target.sheetId, cell: { row: target.row, column: target.column }, value } }] }
}
