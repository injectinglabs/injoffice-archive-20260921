import type { FUniver } from '@univerjs/core/lib/facade'
import { describe, expect, it, vi } from 'vitest'
import { PivotCommandController, type PivotSnapshotV1 } from './commands'
import {
  PIVOT_COLLABORATION_PROTOCOL,
  PivotCollaborationSession,
  applyPivotCollaborationOperation,
  fingerprintPivot,
  type PivotCollaborationOperation,
  type PivotCollaborationTransport,
} from './collaboration'
import { PivotManager } from './manager'
import type { PivotSpec } from './types'

const GRID = [
  ['Region', 'Product', 'Quarter', 'Sales'],
  ['West', 'Widget', 'Q1', 20],
  ['East', 'Gadget', 'Q2', 35],
]

function harness(): PivotCommandController {
  const sheet = {
    getSheetId: () => 'sheet-1',
    getRange: () => ({ getValues: () => GRID, setValues: () => undefined }),
  }
  const api = { getActiveWorkbook: () => ({
    getActiveSheet: () => sheet,
    getSheetBySheetId: (id: string) => id === 'sheet-1' ? sheet : null,
  }) } as unknown as FUniver
  return new PivotCommandController(new PivotManager(api))
}

function spec(id: string, overrides: Partial<PivotSpec> = {}): PivotSpec {
  return {
    id,
    source: { sheetId: 'sheet-1', startRow: 0, startColumn: 0, endRow: 2, endColumn: 3 },
    rows: ['Region'],
    columns: [],
    values: [{ field: 'Sales', agg: 'sum' }],
    target: { sheetId: 'sheet-1', startRow: 0, startColumn: 6 },
    ...overrides,
  }
}

const empty = (): PivotSnapshotV1 => ({ version: 1, pivots: [] })
const base = { protocol: PIVOT_COLLABORATION_PROTOCOL, clientId: 'client-a' } as const

describe('pivot collaboration reducer', () => {
  it('applies fingerprinted lifecycle operations without mutating input', () => {
    const original = empty()
    const created = applyPivotCollaborationOperation(original, { ...base, opId: '1', kind: 'create', value: spec('a') })
    expect(original.pivots).toEqual([])
    const current = created.pivots[0]
    const updated = applyPivotCollaborationOperation(created, {
      ...base, opId: '2', kind: 'update', id: 'a', expectedFingerprint: fingerprintPivot(current),
      value: { ...current, values: [{ field: 'Sales', agg: 'avg' }] },
    })
    expect(updated.pivots[0].values[0].agg).toBe('avg')
    expect(() => applyPivotCollaborationOperation(updated, {
      ...base, opId: '3', kind: 'remove', id: 'a', expectedFingerprint: fingerprintPivot(current),
    })).toThrowError(expect.objectContaining({ code: 'conflict' }))
    expect(applyPivotCollaborationOperation(updated, {
      ...base, opId: '4', kind: 'remove', id: 'a', expectedFingerprint: fingerprintPivot(updated.pivots[0]),
    }).pivots).toEqual([])
  })

  it('preserves stable native identity and validates exact nested fields', () => {
    const current = spec('a', { nativeIdentity: { part: 'xl/pivotTables/pivotTable1.xml' } })
    const snapshot = { version: 1 as const, pivots: [current] }
    expect(() => applyPivotCollaborationOperation(snapshot, {
      ...base, opId: 'bad', kind: 'update', id: 'a', expectedFingerprint: fingerprintPivot(current),
      value: { ...current, nativeIdentity: { part: 'xl/pivotTables/pivotTable2.xml' } },
    })).toThrowError(expect.objectContaining({ code: 'conflict' }))
    expect(() => applyPivotCollaborationOperation(empty(), {
      ...base, opId: 'extra', kind: 'create', value: { ...spec('a'), target: { ...spec('a').target, unsafe: true } },
    } as PivotCollaborationOperation)).toThrowError(expect.objectContaining({ code: 'invalid' }))
  })

  it('canonicalizes fingerprints and rejects widened envelopes', () => {
    expect(fingerprintPivot(spec('a', { filters: { Region: ['West'], Quarter: ['Q1'] } })))
      .toBe(fingerprintPivot(spec('a', { filters: { Quarter: ['Q1'], Region: ['West'] } })))
    expect(() => applyPivotCollaborationOperation(empty(), {
      ...base, opId: 'bad', kind: 'create', value: spec('a'), privileged: true,
    } as PivotCollaborationOperation)).toThrowError(expect.objectContaining({ code: 'invalid' }))
  })
})

class Hub {
  snapshot = empty()
  sequence = 0
  private handlers = new Map<string, (entry: unknown) => void>()

  transport(clientId: string): PivotCollaborationTransport {
    return {
      submit: async (room, operation, baseSequence) => {
        if (baseSequence !== this.sequence) throw new Error('STALE_BASE')
        this.snapshot = applyPivotCollaborationOperation(this.snapshot, operation)
        const entry = { room, operation, sequence: ++this.sequence }
        queueMicrotask(() => { for (const [id, handler] of this.handlers) if (id !== clientId) handler(structuredClone(entry)) })
        return structuredClone(entry)
      },
      subscribe: (_room, handler) => { this.handlers.set(clientId, handler); return () => { this.handlers.delete(clientId) } },
    }
  }
}

function session(controller: PivotCommandController, transport: PivotCollaborationTransport, clientId: string, initialSequence = 0, options: Record<string, unknown> = {}) {
  let id = 0
  return new PivotCollaborationSession(controller, transport, {
    room: 'book-1', clientId, initialSequence, idFactory: () => `${clientId}-${++id}`, ...options,
  })
}

describe('PivotCollaborationSession', () => {
  it('converges two clients and marks remote derived-grid application', async () => {
    const hub = new Hub()
    const firstController = harness(); const secondController = harness()
    const contexts: Array<{ remote: boolean }> = []
    const first = session(firstController, hub.transport('first'), 'first')
    const second = session(secondController, hub.transport('second'), 'second', 0, {
      applySnapshot: (snapshot: PivotSnapshotV1, context: { remote: boolean }) => {
        contexts.push(context)
        return secondController.restore(snapshot)
      },
    })
    first.start(); second.start()
    await first.create(spec('a'))
    await new Promise((resolve) => setTimeout(resolve, 0))
    await second.update('a', { pageFields: [{ field: 'Product', selectedItem: 'Widget' }] })
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(firstController.snapshot()).toEqual(hub.snapshot)
    expect(secondController.snapshot()).toEqual(hub.snapshot)
    expect(contexts.map(({ remote }) => remote)).toEqual([true, false])
    first.dispose(); second.dispose()
  })

  it('does not change specs or baked grids before acknowledgement', async () => {
    let acknowledge!: () => void
    const transport: PivotCollaborationTransport = {
      submit: (room, operation) => new Promise((resolve) => { acknowledge = () => resolve({ room, sequence: 1, operation }) }),
      subscribe: () => () => undefined,
    }
    const controller = harness()
    const client = session(controller, transport, 'first')
    const pending = client.create(spec('a'))
    await Promise.resolve()
    expect(controller.snapshot().pivots).toEqual([])
    acknowledge()
    await pending
    expect(controller.snapshot().pivots[0].id).toBe('a')
  })

  it('preflights invalid specs and enforces host authorization', async () => {
    const submit = vi.fn()
    const denied = session(harness(), { submit, subscribe: () => () => undefined }, 'first', 0, { authorize: () => false })
    await expect(denied.create(spec('a'))).rejects.toMatchObject({ code: 'permission' })
    await expect(denied.create(spec('bad', { values: [] }))).rejects.toMatchObject({ code: 'invalid' })
    expect(submit).not.toHaveBeenCalled()
  })

  it('blocks stale bases and resumes only after successful host resync', async () => {
    const hub = new Hub()
    hub.snapshot = applyPivotCollaborationOperation(empty(), { ...base, opId: 'seed', kind: 'create', value: spec('seed') })
    hub.sequence = 1
    const controller = harness()
    const required = vi.fn()
    const client = session(controller, hub.transport('late'), 'late', 0, { onResyncRequired: required })
    await expect(client.create(spec('a'))).rejects.toMatchObject({ code: 'stale-base' })
    expect(client.state.blocked).toBe(true)
    expect(required).toHaveBeenCalledWith(1, -1, expect.anything())
    await expect(client.resync(hub.snapshot, 1)).resolves.toBe(true)
    await client.create(spec('a'))
    expect(controller.snapshot()).toEqual(hub.snapshot)
  })

  it('blocks altered acknowledgements, gaps, and failed atomic application', async () => {
    const altered: PivotCollaborationTransport = {
      submit: async (room, operation) => ({ room, sequence: 1, operation: { ...operation, value: spec('forged') } as PivotCollaborationOperation }),
      subscribe: () => () => undefined,
    }
    const forgedController = harness()
    const forged = session(forgedController, altered, 'first')
    await expect(forged.create(spec('a'))).rejects.toMatchObject({ code: 'conflict' })
    expect(forgedController.snapshot().pivots).toEqual([])

    const operation: PivotCollaborationOperation = { ...base, clientId: 'remote', opId: 'remote-1', kind: 'create', value: spec('a') }
    const gap = session(harness(), { submit: vi.fn(), subscribe: () => () => undefined }, 'gap')
    await expect(gap.receive({ room: 'book-1', sequence: 2, operation })).resolves.toBe('blocked')

    const failedController = harness()
    const failed = session(failedController, { submit: vi.fn(), subscribe: () => () => undefined }, 'failed', 0, { applySnapshot: () => false })
    await expect(failed.receive({ room: 'book-1', sequence: 1, operation })).resolves.toBe('blocked')
    expect(failedController.snapshot().pivots).toEqual([])
  })

  it('blocks fingerprint conflicts without partial manager mutation', async () => {
    const controller = harness(); controller.add(spec('a'))
    const before = controller.snapshot()
    const client = session(controller, { submit: vi.fn(), subscribe: () => () => undefined }, 'first')
    const operation: PivotCollaborationOperation = {
      ...base, clientId: 'remote', opId: 'remote-1', kind: 'remove', id: 'a', expectedFingerprint: 'fnv1a32:00000000',
    }
    await expect(client.receive({ room: 'book-1', sequence: 1, operation })).resolves.toBe('blocked')
    expect(controller.snapshot()).toEqual(before)
  })
})
