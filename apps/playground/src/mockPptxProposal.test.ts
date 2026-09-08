import { describe, expect, it } from 'vitest'
import { createMockPptxProposal } from './mockPptxProposal'
const body = () => ({ request: 'Replace "Northstar review" with "Northstar: ready for launch"', capabilities: [{ name: 'pptx.native.text.replace' }], context: { format: 'pptx', constraints: { maxOperations: 1, allowedTargets: [{ elementId: 'actual-id', slideId: 'actual-slide', currentText: 'Northstar review' }] } } })
describe('bounded presentation mock proposal', () => {
  it('uses exact disclosed ids and produces only a proposal', () => {
    expect(createMockPptxProposal(body())).toEqual({ operations: [{ name: 'pptx.native.text.replace', operationId: 'mock-pptx-text', input: { elementId: 'actual-id', text: 'Northstar: ready for launch' } }] })
  })
  it.each(['Rewrite the deck', 'Replace "missing" with "changed"', 'Replace "Northstar review" with "Northstar review"', 'Replace "Northstar review" with "bad\ntext"'])('refuses unsupported or unsafe request %s', (request) => { expect(() => createMockPptxProposal({ ...body(), request })).toThrow() })
  it('requires the exact format, capability and one-operation constraint', () => {
    const request = body(); request.context.format = 'xlsx'; expect(() => createMockPptxProposal(request)).toThrow()
    expect(() => createMockPptxProposal({ ...body(), capabilities: [] })).toThrow()
    const multiple = body(); multiple.context.constraints.maxOperations = 2; expect(() => createMockPptxProposal(multiple)).toThrow()
  })
  it('refuses duplicate ids and ambiguous text instead of picking the first target', () => {
    const request = body(), target = request.context.constraints.allowedTargets[0]
    request.context.constraints.allowedTargets.push({ ...target }); expect(() => createMockPptxProposal(request)).toThrow('duplicate')
    request.context.constraints.allowedTargets[1].elementId = 'different-id'; expect(() => createMockPptxProposal(request)).toThrow('exactly one')
  })
  it('refuses overly long and empty targets', () => {
    const request = body(); request.context.constraints.allowedTargets[0].currentText = ''; expect(() => createMockPptxProposal(request)).toThrow()
    const tooLong = body(); tooLong.context.constraints.allowedTargets[0].elementId = 'x'.repeat(301); expect(() => createMockPptxProposal(tooLong)).toThrow()
  })
})
