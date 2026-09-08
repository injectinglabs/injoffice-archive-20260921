import type { JsonObject } from '@injoffice/agent-tools'
import { createMockAgentProposal } from './mockAgentProposal'
import { isMockDocxText } from './mockDocxProposal'
import { agentFormatFromTool, type AgentTool } from './route'

export const GUIDED_AGENT_TASKS = [
  { tool: 'sheets', title: 'Update a workstream status', description: 'Choose a workstream and set its status in the launch tracker.', result: 'A verified XLSX workbook' },
  { tool: 'docs', title: 'Revise document text', description: 'Replace an exact piece of text in the launch brief.', result: 'A verified DOCX document' },
  { tool: 'slides', title: 'Edit presentation text', description: 'Update an exact text element in the launch review.', result: 'A verified PPTX presentation' },
  { tool: 'pdf', title: 'Rotate a PDF page', description: 'Choose a page and turn it in 90-degree increments.', result: 'A verified PDF file' },
] as const satisfies ReadonlyArray<{ tool: AgentTool; title: string; description: string; result: string }>

export type GuidedTaskValues = { target: string; value: string; page: string; degrees: string }
export type GuidedTaskTarget = { id: string; label: string; currentValue: string; unavailable?: string }
export type GuidedTaskOptions = { targets: GuidedTaskTarget[]; values: string[]; pages: number[]; degrees: number[] }
export type GuidedTaskRequest = { request: string; error?: string }
const object = (value: unknown): value is Record<string, unknown> => !!value && typeof value === 'object' && !Array.isArray(value)
const normalized = (value: string) => value.trim().replace(/\s+/g, ' ').toLowerCase()
const quotedText = (value: string) => isMockDocxText(value) && !/["“”\t\r\n]/.test(value)
const sheetKey = (target: Record<string, unknown>) => JSON.stringify([target.sheetId, target.row, target.column])

/** Options come only from the current inspected context; the form never creates targets. */
export function getGuidedTaskOptions(tool: AgentTool, context: JsonObject): GuidedTaskOptions {
  const empty: GuidedTaskOptions = { targets: [], values: [], pages: [], degrees: [] }
  if (context.format !== agentFormatFromTool(tool) && !(tool === 'sheets' && context.format === undefined)) return empty
  const constraints = context.constraints
  if (!object(constraints) || constraints.maxOperations !== 1) return empty
  if (tool === 'pdf') return { ...empty,
    pages: Array.isArray(constraints.allowedPages) ? constraints.allowedPages.filter((page): page is number => typeof page === 'number' && Number.isSafeInteger(page) && page > 0) : [],
    degrees: Array.isArray(constraints.allowedDegrees) ? constraints.allowedDegrees.filter((degrees): degrees is number => typeof degrees === 'number' && [90, 180, 270, -90, -180, -270].includes(degrees)) : [],
  }
  const targets: GuidedTaskTarget[] = []
  for (const target of Array.isArray(constraints.allowedTargets) ? constraints.allowedTargets : []) {
    if (!object(target)) continue
    if (tool === 'sheets' && typeof target.sheetId === 'string' && typeof target.row === 'number' && typeof target.column === 'number' && typeof target.workstream === 'string' && typeof target.currentValue === 'string') {
      targets.push({ id: sheetKey(target), label: target.workstream, currentValue: target.currentValue,
        ...(!target.workstream.trim() || /[\r\n\t]/.test(target.workstream) ? { unavailable: 'This workstream cannot be represented by the sample task.' } : {}) })
    } else if (tool === 'docs' && typeof target.targetId === 'string' && typeof target.expectedText === 'string') {
      targets.push({ id: target.targetId, label: target.expectedText, currentValue: target.expectedText,
        ...(!quotedText(target.expectedText) ? { unavailable: 'This text contains quotes, line breaks, or unsupported characters.' } : {}) })
    } else if (tool === 'slides' && typeof target.elementId === 'string' && typeof target.currentText === 'string') {
      targets.push({ id: target.elementId, label: target.currentText, currentValue: target.currentText,
        ...(!quotedText(target.currentText) ? { unavailable: 'This text contains quotes, line breaks, or unsupported characters.' } : {}) })
    }
  }
  for (const target of targets) {
    const sameText = (other: GuidedTaskTarget) => tool === 'sheets' ? normalized(other.label) === normalized(target.label) : other.label === target.label
    if (targets.filter(sameText).length !== 1 || targets.filter((other) => other.id === target.id).length !== 1) target.unavailable = 'This text matches more than one inspected target.'
  }
  return { ...empty, targets, values: tool === 'sheets' && Array.isArray(constraints.allowedValues)
    ? constraints.allowedValues.filter((value): value is string => typeof value === 'string' && ['Ready', 'On track', 'At risk', 'Review', 'Blocked'].includes(value)) : [] }
}

function proposalFor(context: JsonObject, request: string) {
  const { capabilities, ...documentContext } = context
  return createMockAgentProposal({ request, context: documentContext, capabilities }).operations[0]
}

/** Validate the exact same bounded grammar as the mock before enabling a preview. */
export function buildGuidedTaskRequest(tool: AgentTool, context: JsonObject, values: GuidedTaskValues): GuidedTaskRequest {
  const fail = (error: string): GuidedTaskRequest => ({ request: '', error })
  const options = getGuidedTaskOptions(tool, context)
  let request: string
  let target: GuidedTaskTarget | undefined
  if (tool === 'pdf') {
    if (!options.pages.some((page) => String(page) === values.page)) return fail('Choose a page from the inspected PDF.')
    if (!options.degrees.some((degrees) => String(degrees) === values.degrees)) return fail('Choose one of the available rotations.')
    request = `Rotate page ${values.page} by ${values.degrees} degrees`
  } else {
    const matches = options.targets.filter((item) => item.id === values.target)
    if (matches.length !== 1) return fail('Choose one editable target from the inspected file.')
    target = matches[0]
    if (target.unavailable) return fail(target.unavailable)
    if (tool === 'sheets') {
      if (!options.values.includes(values.value)) return fail('Choose an available status.')
      if (normalized(values.value) === normalized(target.currentValue)) return fail('Choose a different status to propose a change.')
      request = `Mark ${target.label} as ${values.value}`
    } else {
      if (!values.value.trim()) return fail('Enter replacement text to propose a change.')
      if (!quotedText(values.value)) return fail('Use a single line of at most 1,000 characters without double quotes or unsupported characters.')
      if (target.currentValue === values.value) return fail('Change the text before requesting a preview.')
      request = `Replace "${target.currentValue}" with "${values.value}"`
    }
  }
  if (request.length > 2000) return fail('Shorten the replacement text: the combined task must fit within 2,000 characters.')
  try {
    const operation = proposalFor(context, request)
    // A name containing grammar keywords must not silently resolve to another target.
    if (tool === 'sheets' && object(operation.input.cell) && sheetKey({ sheetId: operation.input.sheetId, ...operation.input.cell }) !== target?.id ||
        tool === 'docs' && operation.input.targetId !== target?.id || tool === 'slides' && operation.input.elementId !== target?.id) {
      return fail('This target cannot be represented unambiguously by the sample task.')
    }
    return { request }
  } catch {
    return fail('This task is not available in the inspected sample. Choose another target or reload the sample.')
  }
}

/** Defaults are read back through the mock contract, not guessed from fixture content. */
export function initialGuidedTaskValues(tool: AgentTool, context: JsonObject, defaultRequest: string): GuidedTaskValues {
  const options = getGuidedTaskOptions(tool, context)
  const first = options.targets.find((target) => !target.unavailable)
  const values: GuidedTaskValues = { target: first?.id ?? '', value: tool === 'sheets' ? options.values.find((value) => normalized(value) !== normalized(first?.currentValue ?? '')) ?? '' : '', page: String(options.pages[0] ?? ''), degrees: String(options.degrees[0] ?? '') }
  try {
    const { input } = proposalFor(context, defaultRequest)
    if (tool === 'pdf') {
      if (Array.isArray(input.pages)) values.page = String(input.pages[0])
      values.degrees = String(input.degrees)
    } else {
      const id = tool === 'sheets' && object(input.cell) ? sheetKey({ sheetId: input.sheetId, ...input.cell }) : tool === 'docs' ? input.targetId : input.elementId
      if (typeof id === 'string' && options.targets.some((target) => target.id === id && !target.unavailable)) {
        values.target = id
        values.value = String(tool === 'sheets' ? input.value : input.text)
      }
    }
  } catch { /* Invalid defaults leave editable fields empty or on a disclosed option. */ }
  return values
}
