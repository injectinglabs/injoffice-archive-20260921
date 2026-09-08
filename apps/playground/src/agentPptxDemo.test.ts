import { readFileSync } from 'node:fs'
import { describe, expect, it, vi } from 'vitest'
import type { NativeElement, NativePptxDeck } from '@injoffice/pptx-native'
import { createNativePptxAgentSessionInput } from './agentPptxDemo'
import type { PptxRoundTripRuntime } from './pptxRoundTripRuntime'

const model = JSON.parse(readFileSync(new URL('../../../go/officecompat/corpus/generated/expected/pptx-transitional-common.json', import.meta.url), 'utf8')).native as NativePptxDeck
function flatten(deck: NativePptxDeck) { const all: NativeElement[] = []; const visit = (element: NativeElement) => { all.push(element); if (element.kind === 'group') element.children.forEach(visit) }; deck.slides.forEach((slide) => slide.elements.forEach(visit)); return all }
function setup(options: { wrongText?: boolean; failApply?: boolean } = {}) {
  const source = structuredClone(model), original = new Uint8Array([80, 75, 1, 2]), records = new Map([[1, source]])
  let generation = 1
  const runtime: PptxRoundTripRuntime = {
    extract: vi.fn(async (bytes) => ({ deck: structuredClone(records.get(bytes[2])!), artifactId: '' })),
    apply: vi.fn(async ({ original: input, mutation }) => {
      if (options.failApply) throw new Error('Native write failed')
      const output = structuredClone(records.get(input[2])!)
      generation++
      output.sourceRevision = `rev-${generation.toString(16).padStart(64, '0')}`
      for (const operation of mutation.operations) {
        const target = flatten(output).find((element) => element.id === operation.elementId)!
        if (operation.kind !== 'text.replace' || target.kind !== 'text' && target.kind !== 'shape') throw new Error('Unexpected test mutation')
        target.paragraphs = structuredClone(operation.paragraphs) as typeof target.paragraphs
        if (options.wrongText) target.paragraphs[0].runs[0].text = 'Wrong'
        target.source!.fingerprintSha256 = generation.toString(16).padStart(64, '0')
      }
      records.set(generation, output)
      return { bytes: new Uint8Array([80, 75, generation, 4]) }
    }), terminate: vi.fn(),
  }
  return { runtime, source, records, fetcher: vi.fn(async () => new Response(original)) }
}
async function planned(input: Awaited<ReturnType<typeof createNativePptxAgentSessionInput>>, operations = input.operations) {
  const inspected = await input.session.inspect({ selection: 'document', maxItems: 20, maxBytes: 16_000 })
  return input.session.plan(operations, { expectedRevision: inspected.revision })
}
const commitArgs = (change: { expectedRevision: string }) => ({ expectedRevision: change.expectedRevision, idempotencyKey: 'test-commit', confirmation: 'approved' as const })

describe('public native PPTX demo session', () => {
  it('uses public discovery, native isolated preview, trusted approval, fresh readback, and idempotent retry', async () => {
    const fixture = setup(), input = await createNativePptxAgentSessionInput('safe', fixture.runtime, fixture.fetcher)
    const change = await planned(input)
    expect((await change.validate()).ok).toBe(true)
    expect((await change.preview()).artifact.content.blocks).toContainEqual(expect.objectContaining({ text: expect.stringContaining('Northstar: ready for launch') }))
    expect(input.stats().nativeWrites).toBe(0)
    expect(input.download()).toBeNull()
    const original = structuredClone(fixture.source)
    await expect(change.commit(commitArgs(change))).rejects.toThrow('CONFIRMATION_DENIED')
    await input.approve(change.id)
    const receipt = await change.commit(commitArgs(change))
    expect(receipt.verification.verified).toBe(true)
    expect((await change.verify(receipt)).ok).toBe(true)
    expect(receipt.content?.blocks).toContainEqual(expect.objectContaining({ text: expect.stringContaining('Northstar: ready for launch') }))
    expect(input.stats().nativeWrites).toBe(1)
    const downloaded = new Uint8Array(await input.download()!.arrayBuffer())
    expect(fixture.records.get(downloaded[2])!.sourceRevision).toBe(receipt.fingerprint)
    const beforeTarget = flatten(original).find((element) => element.id === input.operations[0].input.elementId)!
    const afterTarget = flatten(fixture.records.get(downloaded[2])!).find((element) => element.id === beforeTarget.id)!
    if (beforeTarget.kind !== 'text' || afterTarget.kind !== 'text') throw new Error('Expected the fixture text target')
    const expectedParagraphs = structuredClone(beforeTarget.paragraphs)
    expectedParagraphs[0].runs[0].text = 'Northstar: ready for launch'
    expect(afterTarget.paragraphs).toEqual(expectedParagraphs)
    expect(fixture.source).toEqual(original)
    expect(await change.commit(commitArgs(change))).toEqual(receipt)
    expect(input.stats().nativeWrites).toBe(1)
    expect(input.trace().some(({ request }) => request.method === 'office.read')).toBe(true)
    expect(input.trace().some(({ request }) => request.method === 'office.verify')).toBe(false)
    input.dispose(); input.dispose()
    expect(fixture.runtime.terminate).toHaveBeenCalledOnce()
    expect(input.download()).toBeNull()
    await expect(input.propose('anything')).rejects.toThrow('closed')
  })
  it('binds trusted approval to exactly one retained plan', async () => {
    const fixture = setup(), input = await createNativePptxAgentSessionInput('safe', fixture.runtime, fixture.fetcher)
    const first = await planned(input)
    const target = (await input.proposalContext()).constraints as { allowedTargets: Array<{ elementId: string }> }
    const second = await planned(input, input.acceptProposal({ operations: [{ name: 'pptx.native.text.replace', input: { elementId: target.allowedTargets[0].elementId, text: 'Different proposal' } }] }))
    await input.approve(first.id)
    await expect(second.commit(commitArgs(second))).rejects.toThrow('CONFIRMATION_DENIED')
    await expect(input.approve('unknown')).rejects.toThrow('exact change set')
    expect(input.stats().nativeWrites).toBe(0)
    input.dispose()
  })
  it('does not expose wrong native preview output as a successful preview', async () => {
    const fixture = setup({ wrongText: true }), input = await createNativePptxAgentSessionInput('safe', fixture.runtime, fixture.fetcher)
    const change = await planned(input)
    await expect(change.preview()).rejects.toThrow('PREVIEW_MISMATCH')
    expect(input.stats().nativeWrites).toBe(0)
    expect(input.download()).toBeNull()
    input.dispose()
  })
  it('refuses unsupported actions before any native write', async () => {
    const fixture = setup(), input = await createNativePptxAgentSessionInput('refusal', fixture.runtime, fixture.fetcher)
    const change = await planned(input)
    expect((await change.validate()).ok).toBe(false)
    await change.preview()
    await input.approve(change.id)
    await expect(change.commit(commitArgs(change))).rejects.toThrow('VALIDATION_FAILED')
    expect(fixture.runtime.apply).not.toHaveBeenCalled()
    input.dispose()
  })
  it('rejects stale approved plans and proposals after a concurrent real host edit', async () => {
    const fixture = setup(), input = await createNativePptxAgentSessionInput('safe', fixture.runtime, fixture.fetcher)
    const change = await planned(input)
    await input.approve(change.id)
    await input.simulateConcurrentEdit()
    await expect(input.approve(change.id)).rejects.toThrow('revision changed')
    await expect(change.commit(commitArgs(change))).rejects.toThrow('STALE_REVISION')
    await expect(planned(input)).rejects.toThrow('after proposal discovery')
    expect(input.stats().nativeWrites).toBe(1)
    expect(input.download()).toBeNull()
    input.dispose()
  })
  it.each(['fault', 'wrong-text'])('records a successful write but withholds download after failed verification (%s)', async (failure) => {
    const fixture = setup({ wrongText: failure === 'wrong-text' }), input = await createNativePptxAgentSessionInput('safe', fixture.runtime, fixture.fetcher)
    const change = await planned(input)
    await input.approve(change.id)
    input.setVerificationFailure(failure === 'fault')
    const receipt = await change.commit(commitArgs(change))
    expect(receipt.verification.verified).toBe(false)
    expect(input.stats().nativeWrites).toBe(1)
    expect(input.download()).toBeNull()
    expect(await change.commit(commitArgs(change))).toEqual(receipt)
    expect(input.stats().nativeWrites).toBe(1)
    input.dispose()
  })
  it('keeps the original bytes authoritative after a native write failure', async () => {
    const fixture = setup({ failApply: true }), input = await createNativePptxAgentSessionInput('safe', fixture.runtime, fixture.fetcher)
    const change = await planned(input)
    await input.approve(change.id)
    await expect(change.commit(commitArgs(change))).rejects.toThrow('Native write failed')
    expect(input.stats().nativeWrites).toBe(0)
    expect(input.download()).toBeNull()
    input.dispose()
  })
  it('rejects an invalid reopened replacement before adopting its bytes', async () => {
    const fixture = setup(), originalExtract = fixture.runtime.extract
    fixture.runtime.extract = vi.fn(async (bytes) => {
      const result = await originalExtract(bytes)
      if (bytes[2] !== 1) result.deck.contractVersion = 'invalid' as NativePptxDeck['contractVersion']
      return result
    })
    const input = await createNativePptxAgentSessionInput('safe', fixture.runtime, fixture.fetcher), change = await planned(input)
    await input.approve(change.id)
    await expect(change.commit(commitArgs(change))).rejects.toThrow()
    expect(input.stats().nativeWrites).toBe(0)
    expect((await input.session.inspect({ selection: 'document', maxItems: 20, maxBytes: 16_000 })).revision).toBe(fixture.source.sourceRevision)
    expect(input.download()).toBeNull()
    input.dispose()
  })
  it('clamps untrusted proposals to one disclosed target and rejects approval metadata', async () => {
    const fixture = setup(), input = await createNativePptxAgentSessionInput('safe', fixture.runtime, fixture.fetcher)
    const valid = { name: 'pptx.native.text.replace', input: input.operations[0].input }
    expect(() => input.acceptProposal({ operations: [valid], approved: true })).toThrow()
    expect(() => input.acceptProposal({ operations: [valid, valid] })).toThrow()
    expect(() => input.acceptProposal({ operations: [{ ...valid, input: { elementId: 'unknown', text: 'Change' } }] })).toThrow('disclosed')
    expect(() => input.acceptProposal({ operations: [{ ...valid, input: { ...valid.input, text: 'bad\ntext' } }] })).toThrow('single-line')
    expect(() => input.acceptProposal({ operations: [{ ...valid, confirmation: 'approved' }] })).toThrow()
    expect(input.stats().nativeWrites).toBe(0)
    input.dispose()
  })
  it('terminates the worker if sample loading fails', async () => {
    const fixture = setup()
    await expect(createNativePptxAgentSessionInput('safe', fixture.runtime, vi.fn(async () => new Response('', { status: 404 })))).rejects.toThrow('HTTP 404')
    expect(fixture.runtime.terminate).toHaveBeenCalledOnce()
  })
})
