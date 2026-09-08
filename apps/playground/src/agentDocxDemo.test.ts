import { readFileSync } from 'node:fs'
import { describe, expect, it, vi } from 'vitest'
import type { NativeDocxDocumentV1 } from '../../../packages/docs/src/nativeContract'
import { createNativeDocxAgentSessionInput, AGENT_DOCX_PROMPT } from './agentDocxDemo'
import type { DocxRoundTripRuntime } from './docxRoundTripRuntime'

const record = JSON.parse(readFileSync(new URL('../../../go/officecompat/corpus/generated/expected/docx-strict-relocated.json', import.meta.url), 'utf8')) as { native: NativeDocxDocumentV1 }
function setup(options: { wrongText?: boolean } = {}) {
  const source = structuredClone(record.native)
  source.body.blocks[0].paragraph!.runs[0].text = 'Northstar Launch Brief'
  const original = new Uint8Array([80, 75, 1, 2]), records = new Map([[1, source]])
  let generation = 1
  const runtime: DocxRoundTripRuntime = {
    extract: vi.fn(async (bytes) => ({ document: structuredClone(records.get(bytes[2])!), artifactId: '' })),
    apply: vi.fn(async ({ original: input, envelope }) => {
      const output = structuredClone(records.get(input[2])!)
      generation++
      output.source.package_sha256 = `sha256:${generation.toString(16).padStart(64, '0')}`
      output.revision = `rev:${generation.toString(16).padStart(32, '0')}`
      for (const mutation of envelope.payload.mutations) {
        const run = output.body.blocks.flatMap((block) => block.paragraph?.runs ?? []).find((item) => item.id === mutation.target_id)!
        run.text = options.wrongText ? 'Wrong native result' : mutation.text
      }
      records.set(generation, output)
      return { bytes: new Uint8Array([80, 75, generation, 4]) }
    }),
    terminate: vi.fn(),
  }
  const fetcher = vi.fn(async () => new Response(Buffer.from(original).toString('base64')))
  return { source, records, runtime, fetcher }
}
type Input = Awaited<ReturnType<typeof createNativeDocxAgentSessionInput>>
async function planned(input: Input, operations = input.operations) {
  const inspected = await input.session.inspect({ selection: 'document', maxItems: 20, maxBytes: 16_000 })
  return input.session.plan(operations, { expectedRevision: inspected.revision })
}
describe('native DOCX agent demo through the public dispatcher', () => {
  it('previews an isolated copy, requires exact host approval, verifies the real replacement, and deduplicates retry', async () => {
    const fixture = setup(), input = await createNativeDocxAgentSessionInput('safe', fixture.runtime, fixture.fetcher)
    expect(input.scenario.prompt).toBe(AGENT_DOCX_PROMPT)
    const change = await planned(input)
    expect((await change.preview()).artifact.content.title).toBe('Northstar Beta Launch Brief')
    expect((await change.diff()).changes[0]).toMatchObject({ before: 'Northstar Launch Brief', after: 'Northstar Beta Launch Brief' })
    expect(input.stats().nativeWrites).toBe(0); expect(input.download()).toBeNull()
    expect((await input.session.inspect({ selection: 'document', maxItems: 20, maxBytes: 16_000 })).content.title).toBe('Northstar Launch Brief')
    const commit = { expectedRevision: change.expectedRevision, idempotencyKey: 'approved-docx', confirmation: 'approved' as const }
    await expect(change.commit(commit)).rejects.toThrow('CONFIRMATION_DENIED')
    await expect(input.approve('unknown')).rejects.toThrow('exact change set')
    await input.approve(change.id)
    const receipt = await change.commit(commit)
    expect((await change.verify(receipt)).ok).toBe(true)
    expect(receipt.content?.title).toBe('Northstar Beta Launch Brief')
    expect(input.stats().nativeWrites).toBe(1)
    const bytes = new Uint8Array(await input.download()!.arrayBuffer())
    expect(fixture.records.get(bytes[2])!.source.package_sha256).toBe(receipt.fingerprint)
    expect(fixture.source.body.blocks[0].paragraph!.runs[0].text).toBe('Northstar Launch Brief')
    expect(await change.commit(commit)).toEqual(receipt)
    expect(input.stats().nativeWrites).toBe(1)
    expect(input.trace().some(({ request }) => request.method === 'office.read')).toBe(true)
    expect(input.trace().some(({ request }) => request.method === 'office.verify')).toBe(false)
    input.dispose(); input.dispose()
    expect(fixture.runtime.terminate).toHaveBeenCalledOnce(); expect(input.download()).toBeNull()
    await expect(input.propose(AGENT_DOCX_PROMPT)).rejects.toThrow('closed')
  })
  it('does not reuse an approval for a different exact text replacement', async () => {
    const fixture = setup(), input = await createNativeDocxAgentSessionInput('safe', fixture.runtime, fixture.fetcher)
    const first = await planned(input), second = await planned(input, await input.propose('Replace "Northstar Launch Brief" with "Another title"'))
    await input.approve(first.id)
    await expect(second.commit({ expectedRevision: second.expectedRevision, idempotencyKey: 'other', confirmation: 'approved' })).rejects.toThrow('CONFIRMATION_DENIED')
    expect(input.stats().nativeWrites).toBe(0); input.dispose()
  })
  it('refuses macros before native preview or commit writes', async () => {
    const fixture = setup(), input = await createNativeDocxAgentSessionInput('refusal', fixture.runtime, fixture.fetcher)
    const change = await planned(input)
    expect((await change.validate()).ok).toBe(false)
    expect((await change.preview()).artifact.format).toBe('docx')
    expect((await change.diff()).changes).toEqual([])
    await input.approve(change.id)
    await expect(change.commit({ expectedRevision: change.expectedRevision, idempotencyKey: 'no', confirmation: 'approved' })).rejects.toThrow('VALIDATION_FAILED')
    expect(fixture.runtime.apply).not.toHaveBeenCalled(); input.dispose()
  })
  it('rejects stale approval and pins accepted proposal discovery to its source revision', async () => {
    const fixture = setup(), input = await createNativeDocxAgentSessionInput('safe', fixture.runtime, fixture.fetcher)
    const change = await planned(input)
    await input.approve(change.id); await input.simulateConcurrentEdit()
    await expect(input.approve(change.id)).rejects.toThrow('changed revision')
    await expect(change.commit({ expectedRevision: change.expectedRevision, idempotencyKey: 'stale', confirmation: 'approved' })).rejects.toThrow('STALE_REVISION')
    await expect(planned(input)).rejects.toThrow('after proposal discovery')
    expect(input.stats().nativeWrites).toBe(1); expect(input.download()).toBeNull(); input.dispose()
  })
  it('fails closed on native preview mismatch without a source write', async () => {
    const fixture = setup({ wrongText: true }), input = await createNativeDocxAgentSessionInput('safe', fixture.runtime, fixture.fetcher)
    const change = await planned(input)
    expect((await change.validate()).ok).toBe(true)
    await expect(change.preview()).rejects.toThrow('PREVIEW_MISMATCH')
    expect(input.stats().nativeWrites).toBe(0); expect(input.download()).toBeNull(); input.dispose()
  })
  it.each(['fault', 'wrong-result'])('retains successful write but blocks download on failed verification (%s)', async (failure) => {
    const fixture = setup({ wrongText: failure === 'wrong-result' }), input = await createNativeDocxAgentSessionInput('safe', fixture.runtime, fixture.fetcher)
    const change = await planned(input)
    await input.approve(change.id); input.setVerificationFailure(failure === 'fault')
    const commit = { expectedRevision: change.expectedRevision, idempotencyKey: 'failed-proof', confirmation: 'approved' as const }
    const receipt = await change.commit(commit)
    expect(receipt.verification.verified).toBe(false); expect(input.stats().nativeWrites).toBe(1); expect(input.download()).toBeNull()
    expect(await change.commit(commit)).toEqual(receipt); expect(input.stats().nativeWrites).toBe(1); input.dispose()
  })
  it('rejects untrusted extra fields, wrong expected text, undisclosed runs, and malformed text', async () => {
    const fixture = setup(), input = await createNativeDocxAgentSessionInput('safe', fixture.runtime, fixture.fetcher)
    const operation = { name: input.operations[0].op, input: input.operations[0].input }
    expect(() => input.acceptProposal({ operations: [operation], confirmation: 'approved' })).toThrow('top-level')
    for (const patch of [{ expectedText: 'Wrong' }, { targetId: 'missing' }, { targetKind: 'paragraph' }, { text: '\u0000' }, { text: 'a'.repeat(1001) }, { authorization: 'approved' }]) {
      expect(() => input.acceptProposal({ operations: [{ ...operation, input: { ...operation.input, ...patch } }] })).toThrow()
    }
    expect(() => input.acceptProposal({ operations: [operation, operation] })).toThrow('exactly one')
    expect(input.stats().nativeWrites).toBe(0); input.dispose()
  })
  it('terminates the worker if the real fixture cannot be loaded', async () => {
    const fixture = setup(), fetcher = vi.fn(async () => new Response('missing', { status: 404 }))
    await expect(createNativeDocxAgentSessionInput('safe', fixture.runtime, fetcher)).rejects.toThrow('HTTP 404')
    expect(fixture.runtime.terminate).toHaveBeenCalledOnce(); expect(fixture.runtime.extract).not.toHaveBeenCalled()
  })
})
