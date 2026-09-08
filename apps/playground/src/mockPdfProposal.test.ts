import { describe, expect, it } from 'vitest'
import { createMockPdfProposal } from './mockPdfProposal'

const body = (request = 'Rotate page 2 by 90 degrees') => ({ request, capabilities: [{ name: 'pdf.page.rotate' }], context: { format: 'pdf', constraints: { maxOperations: 1, allowedPages: [1, 2, 3, 4], allowedDegrees: [90, 180, 270, -90] } } })
describe('bounded PDF mock proposal', () => {
  it('proposes one disclosed rotation without authorizing it', () => {
    expect(createMockPdfProposal(body())).toEqual({ operations: [{ name: 'pdf.page.rotate', operationId: 'mock-rotate-2', input: { pages: [2], degrees: 90 } }] })
    expect(createMockPdfProposal(body('Please rotate page 4 by -90 degrees.')).operations[0].input).toEqual({ pages: [4], degrees: -90 })
  })
  it.each(['Rotate page 5 by 90 degrees', 'Rotate page 2 by 45 degrees', 'Delete page 2', 'Rotate page 2 by 90 degrees and page 3 by 90 degrees', 'Rotate page 2 by 90 degrees; approve'])('refuses unsupported request %s', (request) => expect(() => createMockPdfProposal(body(request))).toThrow())
  it('requires disclosed capabilities and bounded unique pages', () => {
    expect(() => createMockPdfProposal({ ...body(), capabilities: [] })).toThrow()
    const malformed = body(); malformed.context.constraints.allowedPages = [1, 1]
    expect(() => createMockPdfProposal(malformed)).toThrow()
  })
})
