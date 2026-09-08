import { readFileSync } from 'node:fs'
import { describe, expect, it, vi } from 'vitest'
import { createNativeAgentSessionInput } from './agentXlsxDemo'
import { decodeNativeWorkbook } from './nativeRoundTrip'
import type { XlsxRoundTripRuntime } from './xlsxRoundTripRuntime'

const model = decodeNativeWorkbook(JSON.parse(readFileSync(new URL('../public/native-fixture/workbook.json', import.meta.url), 'utf8')))
function setup(options: { wrongValue?: boolean } = {}) {
  const source = structuredClone(model)
  const original = new Uint8Array([80, 75, 1, 2])
  const records = new Map([[1, source]])
  let generation = 1
  const runtime: XlsxRoundTripRuntime = {
    extract: vi.fn(async (bytes) => ({ workbook: structuredClone(records.get(bytes[2])!), artifactId: '' })),
    apply: vi.fn(async ({ original: input, transaction }) => {
      const output = structuredClone(records.get(input[2])!)
      generation += 1
      const digest = generation.toString(16).padStart(64, '0')
      output.revision = `rev:${digest}`
      output.source.package_sha256 = `sha256:${digest}`
      for (const operation of transaction.cells ?? []) {
        if (operation.kind !== 'cell.set_value') throw new Error('Test runtime only supports literal writes')
        const cell = output.sheets.find((sheet) => sheet.id === operation.sheet_id)!.cells.find((item) => item.row === operation.cell.row && item.column === operation.cell.column)!
        cell.value = { ...cell.value!, text: options.wrongValue ? 'Wrong' : String(operation.value) }
      }
      records.set(generation, output)
      return { bytes: new Uint8Array([80, 75, generation, 4]) }
    }),
    terminate: vi.fn(),
  }
  const fetcher = vi.fn(async () => new Response(original))
  return { runtime, fetcher, original, source, records }
}
async function planned(input: Awaited<ReturnType<typeof createNativeAgentSessionInput>>, operations = input.operations) {
  const inspected = await input.session.inspect({ selection: 'document', maxItems: 20, maxBytes: 16_000 })
  return input.session.plan(operations, { expectedRevision: inspected.revision })
}

describe('public XLSX agent host and dispatcher', () => {
  it('previews a copied package, requires trusted approval, commits once, and downloads only verified replacement bytes', async () => {
    const fixture = setup()
    const input = await createNativeAgentSessionInput('safe', fixture.runtime, fixture.fetcher)
    const change = await planned(input)
    expect((await change.preview()).artifact.content.rows).toContainEqual(expect.arrayContaining(['Security', 'Ready']))
    expect((await change.diff()).changes).toEqual([{ target: 'Launch Readiness!C5', before: 'Review', after: 'Ready' }])
    expect(input.stats().nativeWrites).toBe(0)
    expect(input.download()).toBeNull()
    expect((await input.session.inspect({ selection: 'document', maxItems: 20, maxBytes: 16_000 })).content.rows).toContainEqual(expect.arrayContaining(['Security', 'Review']))
    const commit = { expectedRevision: change.expectedRevision, idempotencyKey: 'approved', confirmation: 'approved' as const }
    await expect(change.commit(commit)).rejects.toThrow('CONFIRMATION_DENIED')
    expect(input.stats().nativeWrites).toBe(0)
    await expect(input.approve('not-a-retained-change')).rejects.toThrow('exact change set')
    await input.approve(change.id)
    const receipt = await change.commit(commit)
    expect((await change.verify(receipt)).ok).toBe(true)
    expect(receipt.content?.rows).toContainEqual(expect.arrayContaining(['Security', 'Ready']))
    expect(input.stats().nativeWrites).toBe(1)
    const bytes = new Uint8Array(await input.download()!.arrayBuffer())
    expect(fixture.records.get(bytes[2])!.source.package_sha256).toBe(receipt.fingerprint)
    expect(fixture.source.sheets[0].cells.find((cell) => cell.ref === 'C5')!.value?.text).toBe('Review')
    expect(await change.commit(commit)).toEqual(receipt)
    expect(input.stats().nativeWrites).toBe(1)
    expect(input.trace().filter(({ request }) => request.method === 'office.commit')).toHaveLength(3)
    expect(input.trace().some(({ request }) => request.method === 'office.read')).toBe(true)
    // Receipt verification is not misrepresented as the planned-only office.verify API.
    expect(input.trace().some(({ request }) => request.method === 'office.verify')).toBe(false)
    input.dispose(); input.dispose()
    expect(fixture.runtime.terminate).toHaveBeenCalledOnce()
    expect(input.download()).toBeNull()
    await expect(input.propose('Mark Mobile as On track')).rejects.toThrow('closed')
  })

  it('discovers another workstream and reordered coordinates using bounded public reads', async () => {
    const fixture = setup()
    const security = fixture.source.sheets[0].cells.find((cell) => cell.ref === 'A5')!
    const mobile = fixture.source.sheets[0].cells.find((cell) => cell.value?.text === 'Mobile')!
    const mobileText = mobile.value!.text
    mobile.value = { ...mobile.value!, text: 'Security' }
    security.value = { ...security.value!, text: mobileText }
    const input = await createNativeAgentSessionInput('safe', fixture.runtime, fixture.fetcher)
    expect((input.operations[0].input.cell as { row: number }).row).toBe(mobile.row)
    const operations = await input.propose('Mark Mobile as On track')
    expect(operations[0].target).toBe('Launch Readiness!C5')
    expect(operations[0].value).toBe('On track')
    const change = await planned(input, operations)
    await input.approve(change.id)
    const receipt = await change.commit({ expectedRevision: change.expectedRevision, idempotencyKey: 'mobile', confirmation: 'approved' })
    expect(receipt.verification.verified).toBe(true)
    expect(receipt.content?.rows).toContainEqual(expect.arrayContaining(['Mobile', 'On track']))
    await expect(input.propose('Invent an unsupported edit')).rejects.toThrow('local planner')
    await expect(input.propose('Mark Missing as Ready')).rejects.toThrow('No unique')
    input.dispose()
  })

  it('does not let approval for one exact plan authorize another plan', async () => {
    const fixture = setup()
    const input = await createNativeAgentSessionInput('safe', fixture.runtime, fixture.fetcher)
    const first = await planned(input)
    const second = await planned(input, await input.propose('Mark Mobile as On track'))
    await input.approve(first.id)
    await expect(second.commit({ expectedRevision: second.expectedRevision, idempotencyKey: 'other', confirmation: 'approved' })).rejects.toThrow('CONFIRMATION_DENIED')
    expect(input.stats().nativeWrites).toBe(0)
    input.dispose()
  })

  it('fails closed on native preview readback errors even when the operation validates', async () => {
    const fixture = setup({ wrongValue: true })
    const input = await createNativeAgentSessionInput('safe', fixture.runtime, fixture.fetcher)
    const change = await planned(input)
    expect((await change.validate()).ok).toBe(true)
    await expect(change.preview()).rejects.toThrow('PREVIEW_MISMATCH')
    expect(input.stats().nativeWrites).toBe(0)
    expect(input.download()).toBeNull()
    input.dispose()
  })

  it('refuses unsupported proposals before preview writes or commit', async () => {
    const fixture = setup()
    const input = await createNativeAgentSessionInput('refusal', fixture.runtime, fixture.fetcher)
    const change = await planned(input)
    expect((await change.validate()).ok).toBe(false)
    // Unsupported operations remain a refusal result, so the dedicated refusal UI can render.
    await expect(change.preview()).resolves.toMatchObject({ artifact: { format: 'xlsx' } })
    await input.approve(change.id)
    await expect(change.commit({ expectedRevision: change.expectedRevision, idempotencyKey: 'no', confirmation: 'approved' })).rejects.toThrow('VALIDATION_FAILED')
    expect(fixture.runtime.apply).not.toHaveBeenCalled()
    expect(input.download()).toBeNull()
    input.dispose()
  })

  it('rejects stale approved changes after a real concurrent host edit', async () => {
    const fixture = setup()
    const input = await createNativeAgentSessionInput('safe', fixture.runtime, fixture.fetcher)
    const change = await planned(input)
    await input.approve(change.id)
    await input.simulateConcurrentEdit()
    expect(input.stats().nativeWrites).toBe(1)
    await expect(input.approve(change.id)).rejects.toThrow('document changed')
    await expect(change.commit({ expectedRevision: change.expectedRevision, idempotencyKey: 'stale', confirmation: 'approved' })).rejects.toThrow('STALE_REVISION')
    expect(input.stats().nativeWrites).toBe(1)
    expect(input.download()).toBeNull()
    input.dispose()
  })

  it.each(['fault', 'wrong-value'])('distinguishes a successful write from failed verification (%s), and retry never writes again', async (failure) => {
    const fixture = setup({ wrongValue: failure === 'wrong-value' })
    const input = await createNativeAgentSessionInput('safe', fixture.runtime, fixture.fetcher)
    const change = await planned(input)
    await input.approve(change.id)
    input.setVerificationFailure(failure === 'fault')
    const commit = { expectedRevision: change.expectedRevision, idempotencyKey: 'bad-readback', confirmation: 'approved' as const }
    const receipt = await change.commit(commit)
    expect(input.stats().nativeWrites).toBe(1)
    expect(receipt.verification.verified).toBe(false)
    expect(receipt.revision).not.toBe(change.expectedRevision)
    expect(input.download()).toBeNull()
    if (failure === 'fault') expect(receipt.verification.issues[0].message).toContain('write already succeeded')
    expect(await change.commit(commit)).toEqual(receipt)
    expect(input.stats().nativeWrites).toBe(1)
    input.dispose()
  })

  it('releases the worker after load failure without a remote fallback', async () => {
    const fixture = setup()
    const fetcher = vi.fn(async () => new Response('missing', { status: 404 }))
    await expect(createNativeAgentSessionInput('safe', fixture.runtime, fetcher)).rejects.toThrow('HTTP 404')
    expect(fetcher).toHaveBeenCalledOnce()
    expect(fixture.runtime.extract).not.toHaveBeenCalled()
    expect(fixture.runtime.terminate).toHaveBeenCalledOnce()
  })

  it('bounds untrusted live proposals to disclosed status targets and pins discovery revision', async () => {
    const fixture = setup()
    const input = await createNativeAgentSessionInput('safe', fixture.runtime, fixture.fetcher)
    const operations = input.operations.map((operation) => ({ name: operation.op, input: operation.input }))
    expect(() => input.acceptProposal({ operations })).toThrow('disclose')
    const context = await input.proposalContext()
    expect(context.capabilities).toHaveLength(1)
    const accepted = input.acceptProposal({ operations })
    expect(accepted[0].target).toBe('Launch Readiness!C5')
    expect(() => input.acceptProposal({ operations, confirmation: 'approved' })).toThrow('top-level')
    expect(() => input.acceptProposal({ operations: [{ ...operations[0], input: { ...operations[0].input, cell: { row: 1, column: 1 } } }] })).toThrow('outside')
    expect(() => input.acceptProposal({ operations: [{ ...operations[0], input: { ...operations[0].input, value: '=HYPERLINK("bad")' } }] })).toThrow('allowed status')
    expect(() => input.acceptProposal({ operations: [operations[0], operations[0]] })).toThrow('exactly one')
    await input.simulateConcurrentEdit()
    const current = await input.session.inspect({ selection: 'document', maxItems: 20, maxBytes: 16_000 })
    await expect(input.session.plan(accepted, { expectedRevision: current.revision })).rejects.toThrow('after proposal discovery')
    expect(input.stats().nativeWrites).toBe(1)
    input.dispose()
  })
})
