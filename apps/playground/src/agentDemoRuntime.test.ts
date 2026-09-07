import { describe, expect, it } from 'vitest'
import { createDemoSessionInput } from './agentDemoRuntime'
import type { AgentDemoFormat } from './agentDemoScenario'

describe('agent change set playground proof', () => {
  it.each<AgentDemoFormat>(['xlsx', 'docx', 'pptx', 'pdf'])('runs the guarded %s workflow end to end', async (format) => {
    const { session, operations } = createDemoSessionInput(format, 'safe')
    const capabilities = await session.capabilities()
    expect(capabilities.operations.some((capability) => capability.operation === operations[0].op)).toBe(true)

    const inspection = await session.inspect({ selection: 'document', maxItems: 20, maxBytes: 16_000 })
    const changeSet = await session.plan(operations, { expectedRevision: inspection.revision })
    const [preview, diff, validation] = await Promise.all([changeSet.preview(), changeSet.diff(), changeSet.validate()])

    expect(preview.artifact.revision).toBe(inspection.revision)
    expect(diff.changes).toHaveLength(1)
    expect(validation).toEqual({ ok: true, issues: [] })

    const receipt = await changeSet.commit({ expectedRevision: inspection.revision, idempotencyKey: `test-${format}`, confirmation: 'approved' })
    const verification = await changeSet.verify(receipt)
    expect(receipt.previousRevision).toBe(inspection.revision)
    expect(receipt.revision).not.toBe(inspection.revision)
    expect(verification.ok).toBe(true)
    expect(verification.fingerprint).toBe(receipt.fingerprint)
  })

  it.each<AgentDemoFormat>(['xlsx', 'docx', 'pptx', 'pdf'])('refuses unsupported %s work before commit', async (format) => {
    const { session, operations } = createDemoSessionInput(format, 'refusal')
    const inspection = await session.inspect({ selection: 'document', maxItems: 20, maxBytes: 16_000 })
    const changeSet = await session.plan(operations, { expectedRevision: inspection.revision })
    const validation = await changeSet.validate()

    expect(validation.ok).toBe(false)
    expect(validation.issues[0]).toMatchObject({ code: 'unsupported_operation' })
    expect(validation.issues[0].path).toContain('/operations/')
    await expect(changeSet.commit({ expectedRevision: inspection.revision, idempotencyKey: `refusal-${format}`, confirmation: 'approved' })).rejects.toThrow(/validation failed/i)
  })

  it('enforces bounded reads, expected revisions, and idempotent commits', async () => {
    const { session, operations } = createDemoSessionInput('xlsx', 'safe')
    await expect(session.inspect({ selection: 'workbook', maxItems: 20, maxBytes: 16_001 })).rejects.toThrow(/maxBytes/)

    await expect(session.plan(operations, { expectedRevision: 'rev-stale' })).rejects.toThrow(/expected revision and fingerprint/i)

    const inspection = await session.inspect({ selection: 'sheet-forecast!A1:D5', maxItems: 20, maxBytes: 16_000 })
    const changeSet = await session.plan(operations, { expectedRevision: inspection.revision })
    const input = { expectedRevision: inspection.revision, idempotencyKey: 'retry-safe', confirmation: 'approved' as const }
    const first = await changeSet.commit(input)
    const retry = await changeSet.commit(input)
    expect(retry).toEqual(first)
  })
})
