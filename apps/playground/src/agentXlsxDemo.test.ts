import { readFileSync } from 'node:fs'
import { describe, expect, it, vi } from 'vitest'
import { createNativeAgentSessionInput } from './agentXlsxDemo'
import { decodeNativeWorkbook } from './nativeRoundTrip'
import type { XlsxRoundTripRuntime } from './xlsxRoundTripRuntime'

const model = decodeNativeWorkbook(JSON.parse(readFileSync(new URL('../public/native-fixture/workbook.json', import.meta.url), 'utf8')))
function setup(options: { wrongValue?: boolean; verifyThrows?: boolean } = {}) {
  const source = structuredClone(model)
  const output = structuredClone(model)
  output.revision = `rev:${'b'.repeat(64)}`
  output.source.package_sha256 = `sha256:${'b'.repeat(64)}`
  output.sheets[0].cells.find((cell) => cell.ref === 'C5')!.value = { kind: 'string', text: options.wrongValue ? 'Wrong' : 'Ready' }
  const original = new Uint8Array([80, 75, 1, 2])
  const replacement = new Uint8Array([80, 75, 3, 4])
  let extracts = 0
  const runtime: XlsxRoundTripRuntime = {
    extract: vi.fn(async (bytes) => {
      extracts += 1
      if (options.verifyThrows && extracts === 3) throw new Error('Readback unavailable')
      return { workbook: structuredClone(bytes[2] === 1 ? source : output), artifactId: '' }
    }),
    apply: vi.fn(async () => ({ bytes: replacement })),
    terminate: vi.fn(),
  }
  const fetcher = vi.fn(async () => new Response(original))
  return { runtime, fetcher, original, replacement, source }
}

describe('real-file agent XLSX adapter', () => {
  it('keeps preview isolated, requires approval, writes once, reopens exact bytes and releases download only after verification', async () => {
    const fixture = setup()
    const input = await createNativeAgentSessionInput('safe', fixture.runtime, fixture.fetcher)
    expect(input.scenario.artifact.name).toBe('launch-readiness-plan.xlsx')
    expect(input.download()).toBeNull()
    const inspected = await input.session.inspect({ selection: 'A1:F8', maxItems: 20, maxBytes: 16_000 })
    expect(inspected.fingerprint).toBe(model.source.package_sha256)
    const planned = await input.session.plan(input.operations, { expectedRevision: inspected.revision })
    expect((await planned.preview()).artifact.content.rows).toContainEqual(expect.arrayContaining(['Security', 'Ready']))
    expect((await planned.diff()).changes).toEqual([{ target: 'Launch Readiness!C5', before: 'Review', after: 'Ready' }])
    expect(fixture.runtime.apply).not.toHaveBeenCalled()
    await expect(planned.commit({ expectedRevision: inspected.revision, idempotencyKey: 'unapproved', confirmation: 'no' as 'approved' })).rejects.toThrow()
    expect(fixture.runtime.apply).not.toHaveBeenCalled()
    const receipt = await planned.commit({ expectedRevision: inspected.revision, idempotencyKey: 'approved', confirmation: 'approved' })
    expect((await planned.verify(receipt)).ok).toBe(true)
    expect(receipt.content?.rows).toContainEqual(expect.arrayContaining(['Security', 'Ready']))
    expect(fixture.runtime.apply).toHaveBeenCalledOnce()
    expect(fixture.runtime.extract).toHaveBeenCalledTimes(3)
    expect(fixture.runtime.extract).toHaveBeenLastCalledWith(fixture.replacement)
    expect(new Uint8Array(await input.download()!.arrayBuffer())).toEqual(fixture.replacement)
    expect(fixture.source.sheets[0].cells.find((cell) => cell.ref === 'C5')!.value?.text).toBe('Review')
    await planned.commit({ expectedRevision: inspected.revision, idempotencyKey: 'approved', confirmation: 'approved' })
    expect(fixture.runtime.apply).toHaveBeenCalledOnce()
    input.dispose()
    expect(fixture.runtime.terminate).toHaveBeenCalledOnce()
    expect(input.download()).toBeNull()
  })

  it('refuses unsupported proposals and stale revisions before native write', async () => {
    const fixture = setup()
    const input = await createNativeAgentSessionInput('refusal', fixture.runtime, fixture.fetcher)
    const inspected = await input.session.inspect({ selection: 'document', maxItems: 20, maxBytes: 16_000 })
    await expect(input.session.plan(input.operations, { expectedRevision: 'stale' })).rejects.toThrow()
    const planned = await input.session.plan(input.operations, { expectedRevision: inspected.revision })
    expect((await planned.validate()).ok).toBe(false)
    await expect(planned.commit({ expectedRevision: inspected.revision, idempotencyKey: 'no', confirmation: 'approved' })).rejects.toThrow()
    expect(fixture.runtime.apply).not.toHaveBeenCalled()
    expect(input.download()).toBeNull()
    input.dispose()
  })

  it.each([{ wrongValue: true }, { verifyThrows: true }])('does not offer unverified output: %j', async (options) => {
    const fixture = setup(options)
    const input = await createNativeAgentSessionInput('safe', fixture.runtime, fixture.fetcher)
    const inspected = await input.session.inspect({ selection: 'document', maxItems: 20, maxBytes: 16_000 })
    const planned = await input.session.plan(input.operations, { expectedRevision: inspected.revision })
    const receipt = await planned.commit({ expectedRevision: inspected.revision, idempotencyKey: 'bad-readback', confirmation: 'approved' })
    expect((await planned.verify(receipt)).ok).toBe(false)
    expect(input.download()).toBeNull()
    input.dispose()
  })

  it('cleans up the worker when sample loading fails, without remote fallback', async () => {
    const fixture = setup()
    const fetcher = vi.fn(async () => new Response('missing', { status: 404 }))
    await expect(createNativeAgentSessionInput('safe', fixture.runtime, fetcher)).rejects.toThrow('HTTP 404')
    expect(fetcher).toHaveBeenCalledOnce()
    expect(fixture.runtime.extract).not.toHaveBeenCalled()
    expect(fixture.runtime.terminate).toHaveBeenCalledOnce()
  })

  it('refuses a reordered fixture rather than silently editing another workstream', async () => {
    const fixture = setup()
    fixture.source.sheets[0].cells.find((cell) => cell.ref === 'A5')!.value = { kind: 'string', text: 'Not Security' }
    await expect(createNativeAgentSessionInput('safe', fixture.runtime, fixture.fetcher)).rejects.toThrow('expected editable Security')
    expect(fixture.runtime.apply).not.toHaveBeenCalled()
    expect(fixture.runtime.terminate).toHaveBeenCalledOnce()
  })
})
