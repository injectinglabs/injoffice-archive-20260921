import { createAgentSession } from '@injoffice/agent-tools'
import { PDFDocument } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { createPdfAgentAdapter, type PdfAgentArtifact } from './pdf.js'
async function artifact(): Promise<PdfAgentArtifact> { const document = await PDFDocument.create(); document.addPage([100, 200]); document.addPage([200, 300]); return { kind: 'injoffice.pdf', artifactId: 'pdf:1', bytes: await document.save() } }

describe('PDF agent session', () => {
  it('runs bounded discovery through a confirmed, freshly verified byte transform', async () => {
    const input = await artifact(); const session = await createAgentSession({ artifact: input, actor: { id: 'agent-1', kind: 'agent' }, adapter: createPdfAgentAdapter(), confirmDestructive: () => true })
    const capabilities = await session.capabilities(); expect(capabilities.capabilities.some(({ name }) => name === 'pdf.page.delete')).toBe(true); expect(capabilities.capabilities.find(({ name }) => name === 'pdf.page.crop')?.inputSchema).toMatchObject({ properties: { box: { required: ['x', 'y', 'width', 'height'], additionalProperties: false } } }); expect((await session.inspect({ maxItems: 5, maxBytes: 2_000 })).itemCount).toBe(2); expect((await session.read({ query: { page: 1 }, maxItems: 1, maxBytes: 2_000 })).itemCount).toBe(1)
    const change = await session.plan([{ name: 'pdf.page.delete', input: { pages: [2] } }]); expect((await change.validate()).valid).toBe(true); expect((await change.preview()).data).toMatchObject({ pageCount: 1 }); expect((await change.diff()).data).toHaveProperty('before')
    const committed = await change.commit({ idempotencyKey: 'pdf-save-1', confirmation: true }); expect(committed.verification.verified).toBe(true); expect(session.artifact.bytes).not.toEqual(input.bytes)
  })
  it.each([
    [{ pages: [1], degrees: 45 }], [{ pages: [0], degrees: 90 }], [{ pages: [1], degrees: 90, extra: true }],
  ])('refuses malformed runtime input %#', async (input) => { const value = await artifact(); const session = await createAgentSession({ artifact: value, actor: { id: 'agent-1', kind: 'agent' }, adapter: createPdfAgentAdapter() }); const change = await session.plan([{ name: 'pdf.page.rotate', input }]); expect((await change.validate()).valid).toBe(false) })
  it('fails bounded reads instead of returning a non-advancing cursor', async () => { const value = await artifact(); const session = await createAgentSession({ artifact: value, actor: { id: 'agent-1', kind: 'agent' }, adapter: createPdfAgentAdapter() }); await expect(session.read({ maxItems: 1, maxBytes: 5 })).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' }) })
})
