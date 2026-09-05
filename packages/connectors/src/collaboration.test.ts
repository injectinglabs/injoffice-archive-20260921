import type { ICellData } from '@univerjs/core'
import type { FUniver } from '@univerjs/core/lib/facade'
import { describe, expect, it, vi } from 'vitest'
import { ConnectorCommandController, type ConnectorCommandSnapshotV1 } from './commands'
import {
  CONNECTOR_COLLABORATION_PROTOCOL,
  ConnectorCollaborationSession,
  ConnectorCollaborationTransportError,
  applyConnectorCollaborationOperation,
  fingerprintConnector,
  fingerprintConnectorRange,
  verifyConnectorCollaborationCellPatch,
  type ConnectorCollaborationCellPatch,
  type ConnectorCollaborationEntry,
  type ConnectorCollaborationOperation,
  type ConnectorCollaborationTransport,
  type ConnectorCollaborativeRefreshRequest,
  type ConnectorLifecycleCollaborationOperation,
} from './collaboration'
import { ConnectorManager, type ConnectorManagerOptions, type ConnectorManagerSnapshotV1, type ConnectorMountedSnapshot, type ConnectorRangeSnapshot } from './manager'
import { sha256Hex } from './sha256'
import type { ConnectorSpec } from './types'

function spec(id = 'weather', patch: Partial<ConnectorSpec> = {}): ConnectorSpec {
  return {
    id,
    name: 'Weather',
    source: { kind: 'http', url: `/gateway/${id}`, format: 'json' },
    target: { sheetId: 'sheet-1', startRow: 0, startColumn: 0 },
    refresh: 'manual',
    ...patch,
  }
}

function mounted(value: ConnectorSpec, rows = 0, columns = 0): ConnectorMountedSnapshot {
  return {
    spec: value,
    status: rows || columns ? { lastRefreshTs: 1, rowCount: rows, columnCount: columns, cacheHit: false } : {},
    lastRows: rows,
    lastColumns: columns,
  }
}

const empty = (): ConnectorManagerSnapshotV1 => ({ version: 1, connectors: [] })
const base = { protocol: CONNECTOR_COLLABORATION_PROTOCOL, clientId: 'client-a' } as const
const nullRange = (rows: number, columns: number): ConnectorCollaborationCellPatch['after'] => ({
  sheetId: 'sheet-1', startRow: 0, startColumn: 0, rowCount: rows, columnCount: columns,
  values: Array.from({ length: rows }, () => Array.from({ length: columns }, () => null)),
})

describe('connector collaboration reducer', () => {
  it('orders credential-free lifecycle changes and rejects target overlap', () => {
    const created = applyConnectorCollaborationOperation(empty(), { ...base, opId: '1', kind: 'create', value: spec('a') })
    const current = created.connectors[0]
    const updatedSpec = spec('a', { name: 'Renamed', target: { sheetId: 'sheet-1', startRow: 3, startColumn: 3 } })
    const moved = applyConnectorCollaborationOperation(created, {
      ...base, opId: '2', kind: 'update', id: 'a', expectedFingerprint: fingerprintConnector(current), value: updatedSpec,
    })
    expect(moved.connectors[0]).toMatchObject({ spec: { name: 'Renamed' }, lastRows: 0, lastColumns: 0 })
    expect(() => applyConnectorCollaborationOperation(moved, {
      ...base, opId: '3', kind: 'create', value: spec('b', { target: updatedSpec.target }),
    })).toThrowError(expect.objectContaining({ code: 'conflict' }))
    expect(() => applyConnectorCollaborationOperation(empty(), {
      ...base, opId: 'secret', kind: 'create', value: spec('bad', { source: { kind: 'http', url: '/gateway?token=secret', format: 'json' } }),
    })).toThrowError(expect.objectContaining({ code: 'invalid' }))
  })

  it('requires exact clear patches for moves and removes with existing data', () => {
    const current = mounted(spec('a'), 2, 2)
    const snapshot = { version: 1 as const, connectors: [current] }
    expect(() => applyConnectorCollaborationOperation(snapshot, {
      ...base, opId: 'remove', kind: 'remove', id: 'a', expectedFingerprint: fingerprintConnector(current),
    })).toThrowError(expect.objectContaining({ code: 'invalid' }))
    const patch = { expectedFingerprint: `sha256:${'0'.repeat(64)}`, after: nullRange(2, 2) }
    expect(applyConnectorCollaborationOperation(snapshot, {
      ...base, opId: 'remove-2', kind: 'remove', id: 'a', expectedFingerprint: fingerprintConnector(current), cellPatch: patch,
    }).connectors).toEqual([])
    expect(() => applyConnectorCollaborationOperation(snapshot, {
      ...base, opId: 'move', kind: 'update', id: 'a', expectedFingerprint: fingerprintConnector(current),
      value: spec('a', { target: { sheetId: 'sheet-1', startRow: 4, startColumn: 4 } }),
      cellPatch: { ...patch, after: { ...patch.after, rowCount: 1, values: [[null, null]] } },
    })).toThrowError(expect.objectContaining({ code: 'invalid' }))
  })

  it('validates authoritative refresh extent, shrink clearing, and cell preconditions', () => {
    const current = mounted(spec('a'), 2, 2)
    const snapshot = { version: 1 as const, connectors: [current] }
    const before: ConnectorRangeSnapshot = { ...nullRange(2, 2), values: [[{ v: 1 }, { f: '=1+1' }], [{ v: 3 }, { v: 4 }]] }
    const cellPatch: ConnectorCollaborationCellPatch = {
      expectedFingerprint: fingerprintConnectorRange(before),
      after: { ...nullRange(2, 2), values: [[9, null], [null, null]] },
    }
    expect(verifyConnectorCollaborationCellPatch(before, cellPatch)).toBe(true)
    const operation: ConnectorCollaborationOperation = {
      ...base, opId: 'refresh', kind: 'refresh', id: 'a', expectedFingerprint: fingerprintConnector(current),
      sourceRevision: 'source-1', preprocessFingerprint: 'pipeline-1',
      value: mounted(current.spec, 1, 1), cellPatch,
    }
    expect(applyConnectorCollaborationOperation(snapshot, operation).connectors[0].lastRows).toBe(1)
    expect(() => applyConnectorCollaborationOperation(snapshot, {
      ...operation, cellPatch: { ...cellPatch, after: { ...cellPatch.after, values: [[9, 8], [null, null]] } },
    })).toThrowError(expect.objectContaining({ code: 'invalid' }))
    expect(() => applyConnectorCollaborationOperation(snapshot, {
      ...operation, cellPatch: { ...cellPatch, after: { ...cellPatch.after, values: [[{ f: '=evil()' } as never, null], [null, null]] } },
    })).toThrowError(expect.objectContaining({ code: 'invalid' }))
  })

  it('validates authoritative refresh cells against the connector schema', () => {
    const constrained = spec('a', { schema: { columns: [{ index: 0, type: 'number' }] } })
    const current = mounted(constrained)
    const before: ConnectorRangeSnapshot = { ...nullRange(1, 1), values: [[{ v: null }]] }
    const operation: ConnectorCollaborationOperation = {
      ...base, opId: 'schema', kind: 'refresh', id: 'a', expectedFingerprint: fingerprintConnector(current), sourceRevision: 'source-1',
      value: mounted(constrained, 1, 1),
      cellPatch: { expectedFingerprint: fingerprintConnectorRange(before), after: { ...nullRange(1, 1), values: [['wrong']] } },
    }
    expect(() => applyConnectorCollaborationOperation({ version: 1, connectors: [current] }, operation))
      .toThrowError(expect.objectContaining({ code: 'invalid' }))
  })

  it('bounds authoritative result dimensions before application', () => {
    const current = mounted(spec('a'))
    const operation: ConnectorCollaborationOperation = {
      ...base, opId: 'huge', kind: 'refresh', id: 'a', expectedFingerprint: fingerprintConnector(current), sourceRevision: 'source-1',
      value: mounted(current.spec, 1, 1),
      cellPatch: {
        expectedFingerprint: `sha256:${'0'.repeat(64)}`,
        after: { ...nullRange(1, 1_025), values: [Array.from({ length: 1_025 }, () => null)] },
      },
    }
    expect(() => applyConnectorCollaborationOperation({ version: 1, connectors: [current] }, operation))
      .toThrowError(expect.objectContaining({ code: 'invalid' }))
  })
})

function harness(options: Omit<ConnectorManagerOptions, 'executionAuthority'> = {}) {
  const cells: ICellData[][] = Array.from({ length: 8 }, () => Array.from({ length: 8 }, () => ({ v: null })))
  const sheet = {
    getSheetId: () => 'sheet-1',
    getRange: (startRow: number, startColumn: number, rowCount: number, columnCount: number) => ({
      getValues: () => Array.from({ length: rowCount }, (_, row) => Array.from({ length: columnCount }, (_, column) => cells[startRow + row][startColumn + column].v ?? null)),
      getFormulas: () => Array.from({ length: rowCount }, (_, row) => Array.from({ length: columnCount }, (_, column) => cells[startRow + row][startColumn + column].f ?? null)),
      setValues: (values: ICellData[][]) => { for (let row = 0; row < rowCount; row++) for (let column = 0; column < columnCount; column++) cells[startRow + row][startColumn + column] = structuredClone(values[row][column]) },
    }),
  }
  const api = { getActiveWorkbook: () => ({ getActiveSheet: () => sheet, getSheetBySheetId: (id: string) => id === 'sheet-1' ? sheet : null }) } as unknown as FUniver
  const fetcher = vi.fn(async () => { throw new Error('client fetch must not run') })
  const manager = new ConnectorManager(api, fetcher, { ...options, executionAuthority: 'server' })
  return { cells, manager, controller: new ConnectorCommandController(manager), fetcher }
}

function seed(manager: ConnectorManager, value: ConnectorMountedSnapshot): void {
  expect(manager.restoreState({ version: 1, connectors: [value] }, { source: 'collaboration' })).toBe(true)
}

class Hub {
  state = empty()
  sequence = 0
  cells: ICellData[][] = Array.from({ length: 8 }, () => Array.from({ length: 8 }, () => ({ v: null })))
  private handlers = new Map<string, (entry: unknown) => void>()

  transport(clientId: string): ConnectorCollaborationTransport {
    return {
      submit: async (room, operation, baseSequence) => this.commit(room, clientId, operation, baseSequence),
      requestRefresh: async (room, request, baseSequence) => {
        if (baseSequence !== this.sequence) throw new ConnectorCollaborationTransportError('STALE_BASE')
        const current = this.state.connectors.find(({ spec }) => spec.id === request.connectorId)
        if (!current || fingerprintConnector(current) !== request.expectedFingerprint) throw new Error('CONFLICT')
        const range = { sheetId: current.spec.target.sheetId, startRow: current.spec.target.startRow, startColumn: current.spec.target.startColumn, rowCount: 2, columnCount: 2 }
        const before = this.capture(range)
        const operation: ConnectorCollaborationOperation = {
          protocol: CONNECTOR_COLLABORATION_PROTOCOL,
          opId: request.requestId,
          clientId: request.clientId,
          kind: 'refresh',
          id: request.connectorId,
          expectedFingerprint: request.expectedFingerprint,
          sourceRevision: 'source-1',
          ...(request.expectedPreprocessFingerprint ? { preprocessFingerprint: request.expectedPreprocessFingerprint } : {}),
          value: mounted(current.spec, 2, 2),
          cellPatch: { expectedFingerprint: fingerprintConnectorRange(before), after: { ...range, values: [[9, 8], [7, 6]] } },
        }
        return this.commit(room, clientId, operation, baseSequence)
      },
      subscribe: (_room, handler) => { this.handlers.set(clientId, handler); return () => { this.handlers.delete(clientId) } },
    }
  }

  private commit(room: string, sender: string, operation: ConnectorCollaborationOperation, baseSequence: number): ConnectorCollaborationEntry {
    if (baseSequence !== this.sequence) throw new ConnectorCollaborationTransportError('STALE_BASE')
    const next = applyConnectorCollaborationOperation(this.state, operation)
    const patch = operation.kind === 'create' ? undefined : operation.cellPatch
    if (patch) {
      const before = this.capture(patch.after)
      if (!verifyConnectorCollaborationCellPatch(before, patch)) throw new Error('CELL_CONFLICT')
      this.write(patch.after)
    }
    this.state = next
    const entry = { room, sequence: ++this.sequence, operation }
    queueMicrotask(() => { for (const [id, handler] of this.handlers) if (id !== sender) handler(structuredClone(entry)) })
    return structuredClone(entry)
  }

  private capture(range: { sheetId: string; startRow: number; startColumn: number; rowCount: number; columnCount: number }): ConnectorRangeSnapshot {
    return { ...range, values: Array.from({ length: range.rowCount }, (_, row) => Array.from({ length: range.columnCount }, (_, column) => structuredClone(this.cells[range.startRow + row][range.startColumn + column]))) }
  }

  private write(range: ConnectorCollaborationCellPatch['after']): void {
    for (let row = 0; row < range.rowCount; row++) for (let column = 0; column < range.columnCount; column++) this.cells[range.startRow + row][range.startColumn + column] = { v: range.values[row][column] }
  }
}

function session(controller: ConnectorCommandController, transport: ConnectorCollaborationTransport, clientId: string, initialSequence = 0, options: Record<string, unknown> = {}) {
  let id = 0
  return new ConnectorCollaborationSession(controller, transport, {
    room: 'book-1', clientId, initialSequence, idFactory: () => `${clientId}-${++id}`,
    applySnapshot: (snapshot) => controller.restore(snapshot, { source: 'collaboration' }),
    ...options,
  })
}

describe('ConnectorCollaborationSession', () => {
  it('disables every local execution path under server authority', async () => {
    const setInterval = vi.fn(() => 1)
    const host = harness({ scheduler: { setInterval, clearInterval: vi.fn() } })
    const interval = mounted(spec('weather', { refresh: 'interval', schedule: { intervalMs: 1_000 } }))
    seed(host.manager, interval)
    expect(setInterval).not.toHaveBeenCalled()
    await expect(host.manager.refresh('weather')).resolves.toBe('skipped')
    expect(host.manager.add(spec('other'))).toBe(false)
    expect(host.manager.update('weather', { name: 'local' })).toBe(false)
    expect(host.manager.remove('weather')).toBe(false)
    expect(host.controller.restore({ version: 1, manager: empty(), ranges: [] })).toBe(false)
    expect(host.manager.restoreState(empty())).toBe(false)
    expect(host.manager.applyRangeSnapshot({ sheetId: 'sheet-1', startRow: 0, startColumn: 0, rowCount: 1, columnCount: 1, values: [[{ v: 9 }]] })).toBe(false)
    expect(host.manager.snapshotState()).toEqual({ version: 1, connectors: [interval] })
    expect(host.fetcher).not.toHaveBeenCalled()
  })

  it('converges lifecycle and server-owned refresh results without client fetching', async () => {
    const hub = new Hub()
    const firstHost = harness(); const secondHost = harness()
    const contexts: Array<{ kind: string; remote: boolean }> = []
    const first = session(firstHost.controller, hub.transport('first'), 'first')
    const second = session(secondHost.controller, hub.transport('second'), 'second', 0, {
      applySnapshot: (snapshot: ConnectorCommandSnapshotV1, context: { kind: string; remote: boolean }) => {
        contexts.push(context)
        return secondHost.controller.restore(snapshot, { source: 'collaboration' })
      },
    })
    first.start(); second.start()
    await first.create(spec())
    await new Promise((resolve) => setTimeout(resolve, 0))
    await second.refresh('weather', 'pipeline-1')
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(firstHost.manager.snapshotState()).toEqual(hub.state)
    expect(secondHost.manager.snapshotState()).toEqual(hub.state)
    expect(firstHost.cells[1][1].v).toBe(6)
    expect(secondHost.cells[1][1].v).toBe(6)
    expect(firstHost.fetcher).not.toHaveBeenCalled()
    expect(secondHost.fetcher).not.toHaveBeenCalled()
    expect(contexts.map(({ kind, remote }) => `${kind}:${remote}`)).toEqual(['entry:true', 'entry:false'])
    first.dispose(); second.dispose()
  })

  it('does not mutate local state before a lifecycle acknowledgement', async () => {
    let acknowledge!: () => void
    const transport: ConnectorCollaborationTransport = {
      submit: (room, operation) => new Promise((resolve) => { acknowledge = () => resolve({ room, sequence: 1, operation }) }),
      requestRefresh: vi.fn(), subscribe: () => () => undefined,
    }
    const host = harness()
    const client = session(host.controller, transport, 'first')
    const pending = client.create(spec())
    await Promise.resolve()
    expect(host.manager.list()).toEqual([])
    acknowledge()
    await pending
    expect(host.manager.get('weather')).toBeDefined()
  })

  it('authorizes lifecycle and refresh requests before transport', async () => {
    const submit = vi.fn(); const requestRefresh = vi.fn()
    const transport: ConnectorCollaborationTransport = { submit, requestRefresh, subscribe: () => () => undefined }
    const host = harness(); seed(host.manager, mounted(spec()))
    const denied = session(host.controller, transport, 'first', 0, { authorize: () => false, authorizeRefresh: () => false })
    await expect(denied.create(spec('other', { target: { sheetId: 'sheet-1', startRow: 4, startColumn: 4 } }))).rejects.toMatchObject({ code: 'permission' })
    await expect(denied.refresh('weather')).rejects.toMatchObject({ code: 'permission' })
    expect(submit).not.toHaveBeenCalled(); expect(requestRefresh).not.toHaveBeenCalled()
  })

  it('blocks stale bases and recovers through an atomic resync hook', async () => {
    const hub = new Hub()
    hub.state = applyConnectorCollaborationOperation(empty(), { ...base, opId: 'seed', kind: 'create', value: spec() })
    hub.sequence = 1
    const host = harness(); const apply = vi.fn((snapshot: ConnectorCommandSnapshotV1) => host.controller.restore(snapshot, { source: 'collaboration' }))
    const required = vi.fn()
    const client = session(host.controller, hub.transport('late'), 'late', 0, { applySnapshot: apply, onResyncRequired: required })
    await expect(client.create(spec('other', { target: { sheetId: 'sheet-1', startRow: 4, startColumn: 4 } }))).rejects.toMatchObject({ code: 'stale-base' })
    expect(client.state.blocked).toBe(true)
    const snapshot = { protocol: CONNECTOR_COLLABORATION_PROTOCOL, room: 'book-1', sequence: 1, revision: 'snapshot-1', manager: hub.state, ranges: [] }
    await expect(client.resync(snapshot)).resolves.toBe(true)
    expect(apply).toHaveBeenCalledWith({ version: 1, manager: hub.state, ranges: [] }, { kind: 'resync', remote: true })
    expect(required).toHaveBeenCalled()
  })

  it('clears stale connector-owned cells while resynchronizing a removal', async () => {
    const host = harness()
    seed(host.manager, mounted(spec(), 2, 2))
    host.cells[0][0] = { v: 1 }; host.cells[0][1] = { v: 2 }
    host.cells[1][0] = { v: 3 }; host.cells[1][1] = { v: 4 }
    const client = session(host.controller, { submit: vi.fn(), requestRefresh: vi.fn(), subscribe: () => () => undefined }, 'resync')
    await expect(client.resync({
      protocol: CONNECTOR_COLLABORATION_PROTOCOL,
      room: 'book-1', sequence: 1, revision: 'removed-1', manager: empty(), ranges: [],
    })).resolves.toBe(true)
    expect(host.manager.list()).toEqual([])
    expect(host.cells.slice(0, 2).map((row) => row.slice(0, 2).map(({ v }) => v))).toEqual([[null, null], [null, null]])
  })

  it('blocks after an ambiguous transport failure until authoritative resync', async () => {
    let committed: ConnectorCollaborationEntry | undefined
    const transport: ConnectorCollaborationTransport = {
      submit: async (room, operation) => {
        committed = { room, sequence: 1, operation }
        throw new ConnectorCollaborationTransportError('UNAVAILABLE', 'acknowledgement lost')
      },
      requestRefresh: vi.fn(), subscribe: () => () => undefined,
    }
    const required = vi.fn()
    const host = harness(); const client = session(host.controller, transport, 'first', 0, { onResyncRequired: required })
    await expect(client.create(spec())).rejects.toMatchObject({ code: 'transport' })
    expect(committed).toBeDefined()
    expect(client.state.blocked).toBe(true)
    await expect(client.create(spec('retry'))).rejects.toMatchObject({ code: 'blocked' })
    await expect(client.receive(committed)).resolves.toBe('blocked')
    expect(required).toHaveBeenCalledOnce()
    expect(host.manager.list()).toEqual([])
  })

  it('blocks altered acknowledgements, sequence gaps, and cell conflicts', async () => {
    const altered: ConnectorCollaborationTransport = {
      submit: async (room, operation) => ({ room, sequence: 1, operation: { ...operation, value: spec('forged') } as ConnectorCollaborationOperation }),
      requestRefresh: vi.fn(), subscribe: () => () => undefined,
    }
    const forgedHost = harness(); const forged = session(forgedHost.controller, altered, 'first')
    await expect(forged.create(spec())).rejects.toMatchObject({ code: 'conflict' })
    expect(forgedHost.manager.list()).toEqual([])

    const operation: ConnectorLifecycleCollaborationOperation = { ...base, clientId: 'remote', opId: 'remote-1', kind: 'create', value: spec() }
    const gap = session(harness().controller, { submit: vi.fn(), requestRefresh: vi.fn(), subscribe: () => () => undefined }, 'gap')
    await expect(gap.receive({ room: 'book-1', sequence: 2, operation })).resolves.toBe('blocked')

    const conflictHost = harness(); seed(conflictHost.manager, mounted(spec()))
    const current = conflictHost.manager.snapshotState().connectors[0]
    const refresh: ConnectorCollaborationOperation = {
      ...base, clientId: 'remote', opId: 'refresh', kind: 'refresh', id: 'weather', expectedFingerprint: fingerprintConnector(current),
      sourceRevision: 'source-1', value: mounted(current.spec, 1, 1),
      cellPatch: { expectedFingerprint: `sha256:${'0'.repeat(64)}`, after: { ...nullRange(1, 1), values: [[5]] } },
    }
    await expect(session(conflictHost.controller, { submit: vi.fn(), requestRefresh: vi.fn(), subscribe: () => () => undefined }, 'conflict')
      .receive({ room: 'book-1', sequence: 1, operation: refresh })).resolves.toBe('blocked')
    expect(conflictHost.cells[0][0].v).toBeNull()
  })

  it('blocks malformed acknowledgements and isolates a throwing resync observer', async () => {
    const malformed: ConnectorCollaborationTransport = {
      submit: async (room, operation) => ({ room, sequence: 1, operation, extra: true } as never),
      requestRefresh: vi.fn(), subscribe: () => () => undefined,
    }
    const host = harness()
    const client = session(host.controller, malformed, 'first', 0, { onResyncRequired: () => { throw new Error('observer failed') } })
    await expect(client.create(spec())).rejects.toMatchObject({ code: 'conflict' })
    expect(client.state.blocked).toBe(true)
    expect(host.manager.list()).toEqual([])
  })

  it('serializes resync behind an in-flight acknowledgement', async () => {
    let acknowledge!: () => void
    const transport: ConnectorCollaborationTransport = {
      submit: (room, operation) => new Promise((resolve) => { acknowledge = () => resolve({ room, sequence: 1, operation }) }),
      requestRefresh: vi.fn(), subscribe: () => () => undefined,
    }
    const host = harness()
    const client = session(host.controller, transport, 'first')
    const pending = client.create(spec())
    await Promise.resolve()
    const staleResync = client.resync({ protocol: CONNECTOR_COLLABORATION_PROTOCOL, room: 'book-1', sequence: 0, revision: 'old', manager: empty(), ranges: [] })
    acknowledge()
    await pending
    await expect(staleResync).resolves.toBe(false)
    expect(host.manager.get('weather')).toBeDefined()
    expect(client.state.sequence).toBe(1)
  })

  it('checks apply postconditions and refuses unbound resync ranges', async () => {
    const hub = new Hub()
    const host = harness()
    const lying = session(host.controller, hub.transport('liar'), 'liar', 0, { applySnapshot: () => true })
    await expect(lying.create(spec())).rejects.toMatchObject({ code: 'conflict' })
    expect(lying.state.blocked).toBe(true)

    const resyncHost = harness()
    const resyncClient = session(resyncHost.controller, { submit: vi.fn(), requestRefresh: vi.fn(), subscribe: () => () => undefined }, 'resync')
    const state = { version: 1 as const, connectors: [mounted(spec(), 1, 1)] }
    await expect(resyncClient.resync({
      protocol: CONNECTOR_COLLABORATION_PROTOCOL,
      room: 'book-1', sequence: 1, revision: 'snapshot-1', manager: state,
      ranges: [{ sheetId: 'sheet-1', startRow: 5, startColumn: 5, rowCount: 1, columnCount: 1, values: [[1]] }],
    })).resolves.toBe(false)
    expect(resyncHost.manager.list()).toEqual([])

    const submit = vi.fn()
    const corruptHost = harness()
    const corrupt = session(corruptHost.controller, { submit, requestRefresh: vi.fn(), subscribe: () => () => undefined }, 'corrupt', 0, {
      applySnapshot: (snapshot: ConnectorCommandSnapshotV1) => corruptHost.manager.restoreState(snapshot.manager, { source: 'collaboration' }),
    })
    await expect(corrupt.resync({
      protocol: CONNECTOR_COLLABORATION_PROTOCOL,
      room: 'book-1', sequence: 1, revision: 'snapshot-2', manager: state,
      ranges: [{ ...nullRange(1, 1), values: [[7]] }],
    })).resolves.toBe(false)
    expect(corrupt.state.blocked).toBe(true)
    await expect(corrupt.create(spec('other'))).rejects.toMatchObject({ code: 'blocked' })
    expect(submit).not.toHaveBeenCalled()
  })
})

describe('connector SHA-256', () => {
  it('matches standard UTF-8 vectors', () => {
    expect(sha256Hex('')).toBe('e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855')
    expect(sha256Hex('abc')).toBe('ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad')
    expect(sha256Hex('✓')).toBe('1dabba21cdad44541f6b15796f8d22978fc7ea10c46aeceeeeb66c23b3ac7604')
  })
})
