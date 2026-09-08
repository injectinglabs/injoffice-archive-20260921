import { describe, expect, it } from 'vitest'
import type { JsonObject } from '@injoffice/agent-tools'
import { buildGuidedTaskRequest, getGuidedTaskOptions, GUIDED_AGENT_TASKS, initialGuidedTaskValues, type GuidedTaskValues } from './guidedAgentTasks'
import { createMockAgentProposal } from './mockAgentProposal'
import type { AgentTool } from './route'

const contexts: Record<AgentTool, JsonObject> = {
  sheets: { capabilities: [{ name: 'xlsx.cell.set_value' }], constraints: { maxOperations: 1, allowedValues: ['Ready', 'On track', 'At risk'], allowedTargets: [
    { sheetId: 'sheet-1', row: 1, column: 2, ref: 'C2', workstream: 'Security', currentValue: 'At risk' },
    { sheetId: 'sheet-1', row: 2, column: 2, ref: 'C3', workstream: 'Mobile', currentValue: 'Ready' },
  ] } },
  docs: { format: 'docx', capabilities: [{ name: 'docx.text.replace' }], constraints: { maxOperations: 1, maxTextLength: 1000, allowedTargets: [
    { targetKind: 'run', targetId: 'run-1', expectedText: 'Northstar Launch Brief', blockId: 'block-1' },
    { targetKind: 'run', targetId: 'run-2', expectedText: 'Release milestones', blockId: 'block-2' },
  ] } },
  slides: { format: 'pptx', capabilities: [{ name: 'pptx.native.text.replace' }], constraints: { maxOperations: 1, allowedTargets: [
    { elementId: 'element-1', currentText: 'Northstar launch review', slideId: 'slide-1' },
    { elementId: 'element-2', currentText: 'Next quarter', slideId: 'slide-2' },
  ] } },
  pdf: { format: 'pdf', capabilities: [{ name: 'pdf.page.rotate' }], constraints: { maxOperations: 1, allowedPages: [1, 2, 3, 4], allowedDegrees: [90, 180, 270, -90, -180, -270] } },
}
const requests: Record<AgentTool, string> = {
  sheets: 'Mark Security as Ready', docs: 'Replace "Northstar Launch Brief" with "Northstar Beta Launch Brief"',
  slides: 'Replace "Northstar launch review" with "Northstar: ready for launch"', pdf: 'Rotate page 2 by 90 degrees',
}
function values(tool: AgentTool, patch: Partial<GuidedTaskValues> = {}, context = contexts[tool]) {
  return { ...initialGuidedTaskValues(tool, context, requests[tool]), ...patch }
}
function modifiedTargets(tool: 'sheets' | 'docs' | 'slides', mutate: (targets: JsonObject[]) => void): JsonObject {
  const context = structuredClone(contexts[tool])
  mutate((context.constraints as JsonObject).allowedTargets as JsonObject[])
  return context
}

describe('guided sample tasks', () => {
  it('offers exactly one bounded task for each real file format', () => {
    expect(GUIDED_AGENT_TASKS.map((task) => task.tool)).toEqual(['sheets', 'docs', 'slides', 'pdf'])
  })
  it.each(['sheets', 'docs', 'slides', 'pdf'] as const)('round trips the inspected %s default through the real mock contract', (tool) => {
    const result = buildGuidedTaskRequest(tool, contexts[tool], values(tool))
    expect(result).toEqual({ request: requests[tool] })
    const { capabilities, ...context } = contexts[tool]
    expect(createMockAgentProposal({ request: result.request, context, capabilities }).operations).toHaveLength(1)
  })
  it('uses actual workstream targets and disclosed status values', () => {
    const options = getGuidedTaskOptions('sheets', contexts.sheets)
    expect(options.targets.map((target) => target.label)).toEqual(['Security', 'Mobile'])
    expect(buildGuidedTaskRequest('sheets', contexts.sheets, values('sheets', { target: options.targets[1].id, value: 'On track' })))
      .toEqual({ request: 'Mark Mobile as On track' })
  })
  it.each(['docs', 'slides'] as const)('switches to a different inspected %s text target', (tool) => {
    const target = getGuidedTaskOptions(tool, contexts[tool]).targets[1]
    expect(buildGuidedTaskRequest(tool, contexts[tool], values(tool, { target: target.id, value: 'Updated milestones' })))
      .toEqual({ request: `Replace "${target.currentValue}" with "Updated milestones"` })
  })
  it.each(['sheets', 'docs', 'slides'] as const)('rejects a fabricated %s target', (tool) => {
    expect(buildGuidedTaskRequest(tool, contexts[tool], values(tool, { target: 'not-inspected' }))).toMatchObject({ request: '', error: expect.any(String) })
  })
  it.each(['sheets', 'docs', 'slides'] as const)('rejects unchanged %s values', (tool) => {
    const target = getGuidedTaskOptions(tool, contexts[tool]).targets[0]
    expect(buildGuidedTaskRequest(tool, contexts[tool], values(tool, { value: target.currentValue }))).toMatchObject({ request: '', error: expect.any(String) })
  })
  it('rejects normalized duplicate workstreams even when their cell coordinates differ', () => {
    const context = modifiedTargets('sheets', (targets) => { targets[1].workstream = ' security ' })
    expect(getGuidedTaskOptions('sheets', context).targets.every((target) => !!target.unavailable)).toBe(true)
    expect(buildGuidedTaskRequest('sheets', context, values('sheets', { target: JSON.stringify(['sheet-1', 1, 2]) }, context)).request).toBe('')
  })
  it.each(['docs', 'slides'] as const)('disables ambiguous %s text rather than guessing a target', (tool) => {
    const context = modifiedTargets(tool, (targets) => {
      const textKey = tool === 'docs' ? 'expectedText' : 'currentText'
      targets[1][textKey] = targets[0][textKey]
    })
    expect(getGuidedTaskOptions(tool, context).targets.every((target) => !!target.unavailable)).toBe(true)
    expect(buildGuidedTaskRequest(tool, context, values(tool, {}, context)).request).toBe('')
  })
  it.each(['docs', 'slides'] as const)('disables duplicate %s identifiers even for distinct text', (tool) => {
    const context = modifiedTargets(tool, (targets) => {
      const key = tool === 'docs' ? 'targetId' : 'elementId'
      targets[1][key] = targets[0][key]
    })
    expect(getGuidedTaskOptions(tool, context).targets.every((target) => !!target.unavailable)).toBe(true)
  })
  it.each(['docs', 'slides'] as const)('rejects empty, quoted, multiline, oversized, and XML-unsafe %s replacement text', (tool) => {
    for (const value of ['', '   ', 'A "quoted" title', 'A “quoted” title', 'First\nSecond', 'First\rSecond', 'First\tSecond', 'x'.repeat(1001), 'bad\u0000text', '\ud800']) {
      expect(buildGuidedTaskRequest(tool, contexts[tool], values(tool, { value })), JSON.stringify(value)).toMatchObject({ request: '', error: expect.any(String) })
    }
  })
  it.each(['docs', 'slides'] as const)('accepts apostrophes and Unicode in %s replacement text', (tool) => {
    expect(buildGuidedTaskRequest(tool, contexts[tool], values(tool, { value: "Northstar’s launch — ready 🚀" })).error).toBeUndefined()
  })
  it.each(['docs', 'slides'] as const)('does not offer %s text that the request grammar cannot quote', (tool) => {
    const context = modifiedTargets(tool, (targets) => { targets[0][tool === 'docs' ? 'expectedText' : 'currentText'] = 'A "quoted" title' })
    expect(getGuidedTaskOptions(tool, context).targets[0].unavailable).toBeTruthy()
  })
  it('enforces the combined 2,000-character request limit', () => {
    const context = modifiedTargets('docs', (targets) => { targets[0].expectedText = 'a'.repeat(1000) })
    const result = buildGuidedTaskRequest('docs', context, values('docs', { target: 'run-1', value: 'b'.repeat(1000) }, context))
    expect(result).toMatchObject({ request: '', error: expect.stringContaining('2,000') })
  })
  it('rejects unsupported statuses, pages and degrees', () => {
    expect(buildGuidedTaskRequest('sheets', contexts.sheets, values('sheets', { value: 'Done' })).request).toBe('')
    for (const patch of [{ page: '0' }, { page: '5' }, { page: '2.0' }, { page: '2; delete' }, { degrees: '0' }, { degrees: '360' }, { degrees: '45' }]) {
      expect(buildGuidedTaskRequest('pdf', contexts.pdf, values('pdf', patch)).request).toBe('')
    }
    expect(buildGuidedTaskRequest('pdf', contexts.pdf, values('pdf', { page: '4', degrees: '-90' }))).toEqual({ request: 'Rotate page 4 by -90 degrees' })
  })
  it.each(['sheets', 'docs', 'slides', 'pdf'] as const)('fails closed for empty %s discovery or missing capability', (tool) => {
    expect(buildGuidedTaskRequest(tool, {}, values(tool)).request).toBe('')
    expect(buildGuidedTaskRequest(tool, { ...contexts[tool], capabilities: [] }, values(tool)).request).toBe('')
    expect(getGuidedTaskOptions(tool, { ...contexts[tool], format: 'unsupported' }).targets).toEqual([])
  })
  it('rejects malformed disclosed target coordinates through the shared mock validator', () => {
    const context = modifiedTargets('sheets', (targets) => { targets[0].ref = 'Z99' })
    expect(buildGuidedTaskRequest('sheets', context, values('sheets', {}, context)).request).toBe('')
  })
  it('initializes replacement fields empty when no valid default is available', () => {
    expect(initialGuidedTaskValues('docs', contexts.docs, 'Invent a document')).toMatchObject({ target: 'run-1', value: '' })
  })
  it('does not mutate inspected contexts or user values', () => {
    const before = structuredClone(contexts)
    const input = values('docs')
    const beforeInput = { ...input }
    getGuidedTaskOptions('docs', contexts.docs)
    buildGuidedTaskRequest('docs', contexts.docs, input)
    expect(contexts).toEqual(before)
    expect(input).toEqual(beforeInput)
  })
})
