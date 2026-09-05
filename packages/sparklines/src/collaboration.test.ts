import { describe, expect, it, vi } from 'vitest'
import {
  SPARKLINE_COLLABORATION_PROTOCOL,
  SparklineCollaborationError,
  SparklineCollaborationSession,
  applySparklineCollaborationOperation,
  fingerprintSparkline,
  type SparklineCollaborationEntry,
  type SparklineCollaborationOperation,
  type SparklineCollaborationTransport,
} from './collaboration'
import { SparklineManager } from './manager'
import type { SparklineSnapshotV1, SparklineSpec } from './types'

const spec = (id: string, row: number): SparklineSpec => ({
  id,
  type: 'line',
  source: { sheetId: 'sheet-1', startRow: row, endRow: row, startColumn: 0, endColumn: 2 },
  target: { sheetId: 'sheet-1', row, column: 3 },
})
const empty = (): SparklineSnapshotV1 => ({ version: 1, sparklines: [], groups: [] })
const base = { protocol: SPARKLINE_COLLABORATION_PROTOCOL, clientId: 'client-a' } as const

describe('sparkline collaboration reducer', () => {
  it('applies create, fingerprinted update, and remove operations', () => {
    const created = applySparklineCollaborationOperation(empty(), { ...base, opId: '1', kind: 'create', value: spec('a', 0) })
    const current = created.sparklines[0]
    const updated = applySparklineCollaborationOperation(created, {
      ...base,
      opId: '2',
      kind: 'update',
      id: 'a',
      expectedFingerprint: fingerprintSparkline(current),
      value: { ...current, options: { showMarkers: true } },
    })
    expect(updated.sparklines[0].options).toEqual({ showMarkers: true })
    expect(() => applySparklineCollaborationOperation(updated, {
      ...base, opId: '3', kind: 'remove', id: 'a', expectedFingerprint: fingerprintSparkline(current),
    })).toThrowError(expect.objectContaining({ code: 'conflict' }))
    const removed = applySparklineCollaborationOperation(updated, {
      ...base, opId: '4', kind: 'remove', id: 'a', expectedFingerprint: fingerprintSparkline(updated.sparklines[0]),
    })
    expect(removed.sparklines).toEqual([])
  })

  it('requires fingerprints for the complete group mutation closure', () => {
    const manager = new SparklineManager()
    manager.add(spec('a', 0)); manager.add(spec('b', 1)); manager.add(spec('c', 2))
    manager.group(['a', 'b'], 'old-group')
    const snapshot = manager.serialize()
    const fingerprints = Object.fromEntries(snapshot.sparklines.map((value) => [value.id, fingerprintSparkline(value)]))
    const operation: SparklineCollaborationOperation = { ...base, opId: 'group', kind: 'group', groupId: 'new-group', memberIds: ['a', 'c'], expectedFingerprints: fingerprints }
    expect(() => applySparklineCollaborationOperation(snapshot, { ...operation, expectedFingerprints: { a: fingerprints.a, c: fingerprints.c } })).toThrowError(expect.objectContaining({ code: 'conflict' }))
    const next = applySparklineCollaborationOperation(snapshot, operation)
    expect(next.groups).toEqual([{ id: 'new-group', memberIds: ['a', 'c'] }])
    expect(next.sparklines.find((value) => value.id === 'b')?.groupId).toBeUndefined()
  })

  it('uses canonical fingerprints and refuses widened transport envelopes', () => {
    expect(fingerprintSparkline({ ...spec('a', 0), options: { showLow: true, showHigh: true } }))
      .toBe(fingerprintSparkline({ ...spec('a', 0), options: { showHigh: true, showLow: true } }))
    expect(() => applySparklineCollaborationOperation(empty(), {
      ...base, opId: 'bad', kind: 'create', value: spec('a', 0), admin: true,
    } as SparklineCollaborationOperation)).toThrowError(expect.objectContaining({ code: 'invalid' }))
  })
})

class Hub {
  snapshot = empty()
  sequence = 0
  private handlers = new Map<string, (entry: unknown) => void>()

  transport(clientId: string): SparklineCollaborationTransport {
    return {
      submit: async (room, operation, baseSequence) => {
        if (baseSequence !== this.sequence) throw new Error('STALE_BASE')
        this.snapshot = applySparklineCollaborationOperation(this.snapshot, operation)
        const entry = { room, operation, sequence: ++this.sequence }
        queueMicrotask(() => { for (const [id, handler] of this.handlers) if (id !== clientId) handler(structuredClone(entry)) })
        return structuredClone(entry)
      },
      subscribe: (_room, handler) => { this.handlers.set(clientId, handler); return () => { this.handlers.delete(clientId) } },
    }
  }
}

const session = (manager: SparklineManager, transport: SparklineCollaborationTransport, clientId: string, initialSequence = 0, options: Record<string, unknown> = {}) => {
  let id = 0
  return new SparklineCollaborationSession(manager, transport, {
    room: 'book-1', clientId, initialSequence, idFactory: () => `${clientId}-${++id}`, ...options,
  })
}

describe('SparklineCollaborationSession', () => {
  it('converges two clients from server-ordered operations', async () => {
    const hub = new Hub()
    const firstManager = new SparklineManager(); const secondManager = new SparklineManager()
    const first = session(firstManager, hub.transport('first'), 'first')
    const second = session(secondManager, hub.transport('second'), 'second')
    first.start(); second.start()
    await first.create(spec('a', 0))
    await new Promise((resolve) => setTimeout(resolve, 0))
    await second.update('a', { options: { showMarkers: true } })
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(firstManager.serialize()).toEqual(hub.snapshot)
    expect(secondManager.serialize()).toEqual(hub.snapshot)
    expect(first.state.sequence).toBe(2)
    first.dispose(); second.dispose()
  })

  it('does not mutate local state before the authoritative acknowledgement', async () => {
    let acknowledge!: (entry: SparklineCollaborationEntry) => void
    const transport: SparklineCollaborationTransport = {
      submit: (_room, operation) => new Promise((resolve) => { acknowledge = (entry) => resolve(entry) }),
      subscribe: () => () => undefined,
    }
    const manager = new SparklineManager()
    const client = session(manager, transport, 'first')
    const pending = client.create(spec('a', 0))
    await Promise.resolve()
    expect(manager.list()).toEqual([])
    acknowledge({ room: 'book-1', sequence: 1, operation: { ...base, clientId: 'first', opId: 'first-1', kind: 'create', value: spec('a', 0) } })
    await pending
    expect(manager.get('a')).toBeDefined()
  })

  it('preflights invalid operations and enforces host authorization', async () => {
    const submit = vi.fn()
    const transport: SparklineCollaborationTransport = { submit, subscribe: () => () => undefined }
    const manager = new SparklineManager(); manager.add(spec('a', 0)); manager.add(spec('b', 1))
    const denied = session(manager, transport, 'first', 0, { authorize: () => false })
    await expect(denied.remove('a')).rejects.toMatchObject({ code: 'permission' })
    await expect(denied.group(['a'], 'group')).rejects.toMatchObject({ code: 'invalid' })
    expect(submit).not.toHaveBeenCalled()
    expect(manager.list()).toHaveLength(2)
  })

  it('blocks on a stale base and recovers only through an explicit resync', async () => {
    const hub = new Hub()
    hub.snapshot = applySparklineCollaborationOperation(empty(), { ...base, opId: 'seed', kind: 'create', value: spec('seed', 0) })
    hub.sequence = 1
    const manager = new SparklineManager()
    const resync = vi.fn()
    const client = session(manager, hub.transport('late'), 'late', 0, { onResyncRequired: resync })
    await expect(client.create(spec('a', 1))).rejects.toMatchObject({ code: 'stale-base' })
    expect(client.state.blocked).toBe(true)
    expect(resync).toHaveBeenCalledWith(1, -1, expect.anything())
    expect(client.resync(hub.snapshot, hub.sequence)).toBe(true)
    await client.create(spec('a', 1))
    expect(manager.serialize()).toEqual(hub.snapshot)
  })

  it('blocks forged acknowledgements and remote sequence gaps', async () => {
    const changedAck: SparklineCollaborationTransport = {
      submit: async (room, operation) => ({ room, sequence: 1, operation: { ...operation, value: spec('forged', 1) } as SparklineCollaborationOperation }),
      subscribe: () => () => undefined,
    }
    const forgedManager = new SparklineManager()
    const forged = session(forgedManager, changedAck, 'first')
    await expect(forged.create(spec('a', 0))).rejects.toMatchObject({ code: 'conflict' })
    expect(forged.state.blocked).toBe(true)
    expect(forgedManager.list()).toEqual([])

    const gap = session(new SparklineManager(), { submit: vi.fn(), subscribe: () => () => undefined }, 'gap')
    const operation: SparklineCollaborationOperation = { ...base, clientId: 'remote', opId: 'remote-1', kind: 'create', value: spec('a', 0) }
    await expect(gap.receive({ room: 'book-1', sequence: 2, operation })).resolves.toBe('blocked')
    expect(gap.state).toMatchObject({ sequence: 0, blocked: true })
  })

  it('rejects conflicting remote fingerprints without partial mutation', async () => {
    const manager = new SparklineManager(); manager.add(spec('a', 0))
    const client = session(manager, { submit: vi.fn(), subscribe: () => () => undefined }, 'first')
    const before = manager.serialize()
    const operation: SparklineCollaborationOperation = {
      ...base, clientId: 'remote', opId: 'remote-1', kind: 'remove', id: 'a', expectedFingerprint: 'fnv1a32:00000000',
    }
    await expect(client.receive({ room: 'book-1', sequence: 1, operation })).resolves.toBe('blocked')
    expect(manager.serialize()).toEqual(before)
  })
})
