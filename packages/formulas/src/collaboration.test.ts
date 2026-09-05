import { describe, expect, it, vi } from 'vitest'
import { CALCULATION_PROTOCOL_VERSION, type CalculationResult } from './calculation'
import {
  FORMULA_COLLABORATION_PROTOCOL,
  FormulaCollaborationSession,
  planCollaborativeFormulaApply,
  type FormulaCellAddress,
  type FormulaCellOccupancy,
  type FormulaCollaborationTarget,
  type FormulaCollaborationTransport,
  type FormulaResultEnvelope,
} from './collaboration'

const identity = { workbookId: 'book-1', revision: 'rev-7', fingerprint: 'sha256:seven' }

function result(overrides: Partial<CalculationResult> = {}): CalculationResult {
  return {
    protocolVersion: CALCULATION_PROTOCOL_VERSION,
    jobId: 'job-1', workbookId: 'book-1', sourceRevision: 'rev-7', sourceFingerprint: 'sha256:seven',
    requestFingerprint: 'fnv1a32:request', engineFingerprint: 'fnv1a32:engine',
    cells: [{ sheetId: 'sheet-1', row: 0, column: 0, result: { kind: 'spill', values: [[1, 2], [3, 4]] } }],
    diagnostics: [],
    ...overrides,
  }
}

function occupancy(addresses: readonly FormulaCellAddress[], overrides: Record<string, Partial<FormulaCellOccupancy>> = {}): FormulaCellOccupancy[] {
  return addresses.map((address) => {
    const key = `${address.row}:${address.column}`
    return { ...address, kind: address.row === 0 && address.column === 0 ? 'formula-anchor' : 'empty', ...overrides[key] } as FormulaCellOccupancy
  })
}

function envelope(overrides: Partial<FormulaResultEnvelope> = {}): FormulaResultEnvelope {
  return {
    protocol: FORMULA_COLLABORATION_PROTOCOL, room: 'room-1', sequence: 1,
    authorityId: 'calc-primary', authorityEpoch: 'epoch-1', result: result(), ...overrides,
  }
}

function harness(role: 'authority' | 'follower' = 'follower', targetOverrides: Partial<FormulaCollaborationTarget> = {}) {
  let listener: ((value: unknown) => void) | undefined
  const published: FormulaResultEnvelope[] = []
  const transport: FormulaCollaborationTransport = {
    async publish(value) { published.push(structuredClone(value)) },
    subscribe(handler) { listener = handler; return () => { listener = undefined } },
  }
  const applied: unknown[] = []
  const target: FormulaCollaborationTarget = {
    currentIdentity: () => identity,
    inspect: async (addresses) => occupancy(addresses),
    applyDerived: async (writes, value) => { applied.push({ writes: structuredClone(writes), envelope: structuredClone(value) }) },
    ...targetOverrides,
  }
  const events: unknown[] = []
  const recalculate = vi.fn()
  const resync = vi.fn()
  const session = new FormulaCollaborationSession({
    room: 'room-1', authorityId: 'calc-primary', authorityEpoch: 'epoch-1', role,
    transport, target, onEvent: (event) => events.push(event),
    onRecalculationRequired: recalculate, onResyncRequired: resync,
  })
  return { applied, events, listener: () => listener, published, recalculate, resync, session, target, transport }
}

describe('planCollaborativeFormulaApply', () => {
  it('flattens a spill when every destination is empty or owned by its anchor', () => {
    const addresses = [
      { sheetId: 'sheet-1', row: 0, column: 0 }, { sheetId: 'sheet-1', row: 0, column: 1 },
      { sheetId: 'sheet-1', row: 1, column: 0 }, { sheetId: 'sheet-1', row: 1, column: 1 },
    ]
    const plan = planCollaborativeFormulaApply(result(), identity, occupancy(addresses, {
      '1:1': { kind: 'spill', owner: addresses[0] },
    }))
    expect(plan).toMatchObject({ status: 'ready', writes: [
      { row: 0, column: 0, value: 1 }, { row: 0, column: 1, value: 2 },
      { row: 1, column: 0, value: 3 }, { row: 1, column: 1, value: 4 },
    ] })
  })

  it('refuses dirty, occupied, missing-anchor, and overlapping spill cells', () => {
    const addresses = [
      { sheetId: 'sheet-1', row: 0, column: 0 }, { sheetId: 'sheet-1', row: 0, column: 1 },
      { sheetId: 'sheet-1', row: 1, column: 0 }, { sheetId: 'sheet-1', row: 1, column: 1 },
    ]
    const unsafe = occupancy(addresses, {
      '0:0': { kind: 'empty' }, '0:1': { kind: 'value' }, '1:0': { dirty: true },
    })
    const overlap = result({ cells: [
      ...result().cells,
      { sheetId: 'sheet-1', row: 1, column: 1, result: { kind: 'value', value: 9 } },
    ] })
    const plan = planCollaborativeFormulaApply(overlap, identity, unsafe)
    expect(plan.status).toBe('conflict')
    if (plan.status === 'conflict') expect(new Set(plan.conflicts.map(({ code }) => code))).toEqual(new Set(['missing-anchor', 'occupied', 'dirty', 'overlap']))
  })
})

describe('FormulaCollaborationSession', () => {
  it('applies an ordered authority result atomically and ignores duplicates', async () => {
    const subject = harness()
    expect(await subject.session.receive(envelope())).toBe('applied')
    expect(subject.applied).toHaveLength(1)
    expect(subject.session.state.sequence).toBe(1)
    expect(await subject.session.receive(envelope())).toBe('duplicate')
  })

  it('advances stale/conflicting derived messages and requests recalculation', async () => {
    const stale = harness('follower', { currentIdentity: () => ({ ...identity, revision: 'rev-8' }) })
    expect(await stale.session.receive(envelope())).toBe('stale')
    expect(stale.session.state.sequence).toBe(1)
    expect(stale.recalculate).toHaveBeenCalledWith('stale', expect.anything())

    const conflict = harness('follower', { inspect: async (addresses) => occupancy(addresses, { '0:1': { kind: 'value' } }) })
    expect(await conflict.session.receive(envelope())).toBe('conflict')
    expect(conflict.recalculate).toHaveBeenCalledWith('spill-conflict', expect.anything())
    expect(conflict.applied).toHaveLength(0)
  })

  it('blocks on sequence gaps until an exact-epoch resync', async () => {
    const subject = harness()
    expect(await subject.session.receive(envelope({ sequence: 2 }))).toBe('gap')
    expect(subject.session.state.blocked).toBe(true)
    expect(subject.resync).toHaveBeenCalledWith(1, 2)
    expect(subject.session.resync(1, 'wrong-epoch')).toBe(false)
    expect(subject.session.resync(1)).toBe(true)
    expect(await subject.session.receive(envelope({ sequence: 2 }))).toBe('applied')
  })

  it('rejects forged authority and malformed result envelopes', async () => {
    const subject = harness()
    expect(await subject.session.receive(envelope({ authorityId: 'forged' }))).toBe('rejected')
    expect(await subject.session.receive({ ...envelope(), result: { ...result(), protocolVersion: 'future' } })).toBe('rejected')
    expect(subject.session.state.sequence).toBe(0)
  })

  it('permits only the authority to publish current validated results', async () => {
    const authority = harness('authority')
    const sent = await authority.session.publish(result())
    expect(sent.sequence).toBe(1)
    expect(authority.published).toEqual([sent])
    await expect(harness('follower').session.publish(result())).rejects.toThrow('Only the negotiated')
    await expect(authority.session.publish(result({ sourceRevision: 'stale' }))).rejects.toMatchObject({ code: 'STALE_RESULT' })
  })

  it('serializes concurrent authority publications into unique sequences', async () => {
    const authority = harness('authority')
    const [first, second] = await Promise.all([
      authority.session.publish(result({ jobId: 'job-1' })),
      authority.session.publish(result({ jobId: 'job-2' })),
    ])
    expect([first.sequence, second.sequence]).toEqual([1, 2])
    expect(authority.published.map(({ sequence }) => sequence)).toEqual([1, 2])
  })

  it('serializes subscribed envelopes and disposes the transport listener', async () => {
    const applied: number[] = []
    const subject = harness('follower', { applyDerived: async (_writes, value) => { applied.push(value.sequence) } })
    subject.session.start()
    subject.listener()?.(envelope({ sequence: 1 }))
    subject.listener()?.(envelope({ sequence: 2 }))
    await vi.waitFor(() => expect(applied).toEqual([1, 2]))
    subject.session.dispose()
    expect(subject.listener()).toBeUndefined()
  })
})
