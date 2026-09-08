import { afterEach, describe, expect, it, vi } from 'vitest'
import * as pdfAdapter from '@injoffice/agent-office/pdf'
import { readInfo } from '@injoffice/pdf/browser'
import { createNativePdfAgentSessionInput } from './agentPdfDemo'
import { createPdfDemoFixture } from './pdfDemoFixture'

async function prepare(input: Awaited<ReturnType<typeof createNativePdfAgentSessionInput>>) {
  const inspection = await input.session.inspect({ selection: 'document', maxItems: 100, maxBytes: 32_000 })
  const plan = await input.session.plan(input.operations, { expectedRevision: inspection.revision })
  return { inspection, plan, commit: () => plan.commit({ expectedRevision: inspection.revision, idempotencyKey: `test-${plan.id}`, confirmation: 'approved' }) }
}

describe('real PDF agent demo host', () => {
  afterEach(() => vi.restoreAllMocks())
  it.each(['error', 'missing-pages', 'rotation', 'page-count', 'geometry'])('rejects failed or mismatched PDF preview proof (%s)', async (failure) => {
    const adapter = pdfAdapter.createPdfAgentAdapter()
    const originalPreview = adapter.preview.bind(adapter)
    vi.spyOn(pdfAdapter, 'createPdfAgentAdapter').mockReturnValue({ ...adapter, async preview(request) {
      const result = await originalPreview(request)
      if (failure === 'error') return { ...result, issues: [{ severity: 'error', code: 'PREVIEW_FAILED', message: 'Injected failed proof' }] }
      const data = result.data as unknown as { pageCount: number; pages?: Array<{ rotation: number; width: number }> }
      if (failure === 'missing-pages') delete data.pages
      if (failure === 'rotation') data.pages![1].rotation = 0
      if (failure === 'page-count') data.pageCount++
      if (failure === 'geometry') data.pages![1].width++
      return result
    } })
    const input = await createNativePdfAgentSessionInput('safe'); const { plan } = await prepare(input)
    expect((await plan.validate()).ok).toBe(true)
    await expect(plan.preview()).rejects.toThrow(/PDF preview failed/)
    expect(input.stats().nativeWrites).toBe(0); expect(input.download()).toBeNull(); input.dispose()
  })
  it('revokes previously verified download bytes when the cached receipt no longer matches current identity', async () => {
    const adapter = pdfAdapter.createPdfAgentAdapter()
    let mismatch = false
    vi.spyOn(pdfAdapter, 'createPdfAgentAdapter').mockReturnValue({ ...adapter, async identity(request) {
      const identity = await adapter.identity(request)
      return mismatch ? { ...identity, fingerprint: 'sha256:' + '0'.repeat(64) } : identity
    } })
    const input = await createNativePdfAgentSessionInput('safe'); const { plan, commit } = await prepare(input)
    await input.approve(plan.id); await commit(); expect(input.download()).not.toBeNull()
    mismatch = true
    await commit()
    expect(input.download()).toBeNull(); expect(input.stats().nativeWrites).toBe(1); input.dispose()
  })
  it('previews without source writes, approves, downloads exact verified PDF bytes, and caches retries', async () => {
    const original = await createPdfDemoFixture(); const snapshot = original.slice()
    const input = await createNativePdfAgentSessionInput('safe', async () => original)
    const { inspection, plan, commit } = await prepare(input)
    expect((await plan.validate()).ok).toBe(true)
    expect((await plan.preview()).artifact.content.rotation).toBe(90)
    expect((await plan.diff()).changes).toEqual([{ target: 'page-2', before: '0°', after: '90°' }])
    expect(input.stats().nativeWrites).toBe(0)
    expect(input.download()).toBeNull()
    expect((await input.session.inspect({ selection: 'document', maxItems: 100, maxBytes: 32_000 })).revision).toBe(inspection.revision)
    expect(original).toEqual(snapshot)
    await expect(commit()).rejects.toThrow(/CONFIRMATION/)
    await input.approve(plan.id)
    const receipt = await commit()
    expect(receipt.verification.verified).toBe(true)
    expect(input.stats().nativeWrites).toBe(1)
    const bytes = new Uint8Array(await input.download()!.arrayBuffer())
    expect(new TextDecoder().decode(bytes.subarray(0, 5))).toBe('%PDF-')
    const output = await readInfo(bytes)
    expect(output.pageCount).toBe(4)
    expect(output.pages.map((page) => page.rotation)).toEqual([0, 90, 0, 0])
    const fingerprint = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), (byte) => byte.toString(16).padStart(2, '0')).join('')
    expect(receipt.fingerprint).toContain(fingerprint)
    expect(await commit()).toEqual(receipt)
    expect(input.stats().nativeWrites).toBe(1)
    expect(input.trace().map((item) => item.request.method)).toEqual(expect.arrayContaining(['office.capabilities', 'office.inspect', 'office.read', 'office.plan', 'office.preview', 'office.diff', 'office.validate', 'office.commit']))
    input.dispose()
    expect(input.download()).toBeNull()
    await expect(input.propose('Rotate page 2 by 90 degrees')).rejects.toThrow(/closed/)
  })
  it('rejects a stale approved plan after a real concurrent edit', async () => {
    const input = await createNativePdfAgentSessionInput('safe'); const { plan, commit } = await prepare(input)
    await input.approve(plan.id); await input.simulateConcurrentEdit()
    await expect(commit()).rejects.toThrow(/STALE|REVISION/)
    expect(input.stats().nativeWrites).toBe(1); expect(input.download()).toBeNull(); input.dispose()
  })
  it('pins proposals to disclosed source revisions and rejects injected authorization or targets', async () => {
    const input = await createNativePdfAgentSessionInput('safe')
    const proposal = { operations: [{ name: 'pdf.page.rotate', input: { pages: [2], degrees: 90 } }] }
    expect(() => input.acceptProposal({ ...proposal, confirmation: 'approved' })).toThrow()
    expect(() => input.acceptProposal({ operations: [{ name: 'pdf.page.rotate', input: { pages: [200], degrees: 90 } }] })).toThrow()
    const accepted = input.acceptProposal(proposal)
    await input.simulateConcurrentEdit()
    const inspection = await input.session.inspect({ selection: 'document', maxItems: 100, maxBytes: 32_000 })
    await expect(input.session.plan(accepted, { expectedRevision: inspection.revision })).rejects.toThrow(/revision/)
    input.dispose()
  })
  it('shows committed-but-unverified fault without a verified download', async () => {
    const input = await createNativePdfAgentSessionInput('safe'); const { plan, commit } = await prepare(input)
    input.setVerificationFailure(true); await input.approve(plan.id)
    const receipt = await commit()
    expect(receipt.verification.verified).toBe(false); expect(receipt.content?.rotation).toBe(90)
    expect(input.stats().nativeWrites).toBe(1); expect(input.download()).toBeNull()
    expect((await commit()).fingerprint).toBe(receipt.fingerprint); expect(input.stats().nativeWrites).toBe(1)
    input.dispose()
  })
  it('keeps the real-file refusal path inspectable without transforming or writing', async () => {
    const input = await createNativePdfAgentSessionInput('refusal'); const { plan, commit } = await prepare(input)
    expect((await plan.validate()).ok).toBe(false)
    expect((await plan.preview()).artifact.content.rotation).toBe(0)
    expect((await plan.diff()).changes).toEqual([])
    await input.approve(plan.id)
    await expect(commit()).rejects.toThrow()
    expect(input.stats().nativeWrites).toBe(0); expect(input.download()).toBeNull()
    await expect(input.propose('Rewrite all text')).rejects.toThrow(/mock/)
    input.dispose()
  })
})
