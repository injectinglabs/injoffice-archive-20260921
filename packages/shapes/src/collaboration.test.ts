import type { FUniver } from '@univerjs/core/lib/facade'
import { describe, expect, it, vi } from 'vitest'
import { ShapeCommandController, type ShapeSnapshotV1 } from './commands'
import {
  SHAPE_COLLABORATION_PROTOCOL,
  ShapeCollaborationSession,
  applyShapeCollaborationOperation,
  fingerprintShape,
  type ShapeCollaborationEntry,
  type ShapeCollaborationOperation,
  type ShapeCollaborationTransport,
} from './collaboration'
import { ShapeManager } from './manager'
import type { ShapeSpec } from './types'

function harness(accept = true): ShapeCommandController {
  const sheet = {
    getSheetId: () => 'sheet-1',
    addFloatDomToPosition: (_config: unknown, id: string) => accept ? { id, dispose: () => undefined } : null,
  }
  const api = { getActiveWorkbook: () => ({
    getActiveSheet: () => sheet,
    getSheetBySheetId: (id: string) => id === 'sheet-1' ? sheet : null,
  }) } as unknown as FUniver
  return new ShapeCommandController(new ShapeManager(api))
}

const spec = (id: string, overrides: Partial<ShapeSpec> = {}): ShapeSpec => ({ id, kind: 'rect', text: 'Original', fill: '#ffffff', ...overrides })
const value = (id: string, overrides: Partial<ShapeSpec> = {}) => ({ spec: spec(id, overrides), sheetId: 'sheet-1' })
const empty = (): ShapeSnapshotV1 => ({ version: 1, shapes: [] })
const base = { protocol: SHAPE_COLLABORATION_PROTOCOL, clientId: 'client-a' } as const

describe('shape collaboration reducer', () => {
  it('applies fingerprinted lifecycle operations without mutating its input', () => {
    const original = empty()
    const created = applyShapeCollaborationOperation(original, { ...base, opId: '1', kind: 'create', value: value('a') })
    expect(original.shapes).toEqual([])
    const current = created.shapes[0]
    const updated = applyShapeCollaborationOperation(created, {
      ...base, opId: '2', kind: 'update', id: 'a', expectedFingerprint: fingerprintShape(current),
      value: value('a', { text: 'Updated' }),
    })
    expect(updated.shapes[0].spec.text).toBe('Updated')
    expect(() => applyShapeCollaborationOperation(updated, {
      ...base, opId: '3', kind: 'remove', id: 'a', expectedFingerprint: fingerprintShape(current),
    })).toThrowError(expect.objectContaining({ code: 'conflict' }))
    expect(applyShapeCollaborationOperation(updated, {
      ...base, opId: '4', kind: 'remove', id: 'a', expectedFingerprint: fingerprintShape(updated.shapes[0]),
    }).shapes).toEqual([])
  })

  it('preserves stable kind and native identity while allowing placement updates', () => {
    const identity = { drawingPart: 'xl/drawings/drawing1.xml', objectId: 7 }
    const original = { version: 1 as const, shapes: [{ ...value('a', { nativeIdentity: identity }), cellAnchor: { FromCol: 0, FromRow: 0, ToCol: 2, ToRow: 3 } }] }
    const replacement = { ...original.shapes[0], spec: { ...original.shapes[0].spec, nativeIdentity: { ...identity, objectId: 8 } }, sheetId: 'sheet-2' }
    expect(() => applyShapeCollaborationOperation(original, {
      ...base, opId: 'bad', kind: 'update', id: 'a', expectedFingerprint: fingerprintShape(original.shapes[0]), value: replacement,
    })).toThrowError(expect.objectContaining({ code: 'conflict' }))
    const moved = applyShapeCollaborationOperation(original, {
      ...base, opId: 'move', kind: 'update', id: 'a', expectedFingerprint: fingerprintShape(original.shapes[0]),
      value: { ...original.shapes[0], cellAnchor: { FromCol: 2, FromRow: 2, ToCol: 5, ToRow: 8 } },
    })
    expect(moved.shapes[0].cellAnchor?.FromCol).toBe(2)
  })

  it('canonicalizes property order and rejects widened envelopes', () => {
    expect(fingerprintShape(value('a', { text: 'x', fill: '#fff' })))
      .toBe(fingerprintShape({ sheetId: 'sheet-1', spec: { fill: '#fff', text: 'x', kind: 'rect', id: 'a' } }))
    expect(() => applyShapeCollaborationOperation(empty(), {
      ...base, opId: 'bad', kind: 'create', value: value('a'), privileged: true,
    } as ShapeCollaborationOperation)).toThrowError(expect.objectContaining({ code: 'invalid' }))
  })
})

class Hub {
  snapshot = empty()
  sequence = 0
  private handlers = new Map<string, (entry: unknown) => void>()

  transport(clientId: string): ShapeCollaborationTransport {
    return {
      submit: async (room, operation, baseSequence) => {
        if (baseSequence !== this.sequence) throw new Error('STALE_BASE')
        this.snapshot = applyShapeCollaborationOperation(this.snapshot, operation)
        const entry = { room, operation, sequence: ++this.sequence }
        queueMicrotask(() => { for (const [id, handler] of this.handlers) if (id !== clientId) handler(structuredClone(entry)) })
        return structuredClone(entry)
      },
      subscribe: (_room, handler) => { this.handlers.set(clientId, handler); return () => { this.handlers.delete(clientId) } },
    }
  }
}

function session(controller: ShapeCommandController, transport: ShapeCollaborationTransport, clientId: string, initialSequence = 0, options: Record<string, unknown> = {}) {
  let id = 0
  return new ShapeCollaborationSession(controller, transport, {
    room: 'book-1', clientId, initialSequence, idFactory: () => `${clientId}-${++id}`, ...options,
  })
}

describe('ShapeCollaborationSession', () => {
  it('converges two mounted clients through server-ordered operations', async () => {
    const hub = new Hub()
    const firstController = harness(); const secondController = harness()
    const first = session(firstController, hub.transport('first'), 'first')
    const second = session(secondController, hub.transport('second'), 'second')
    first.start(); second.start()
    await first.create(value('a'))
    await new Promise((resolve) => setTimeout(resolve, 0))
    await second.update('a', { spec: { text: 'Shared' }, cellAnchor: { FromCol: 1, FromRow: 2, ToCol: 4, ToRow: 7 } })
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(firstController.snapshot()).toEqual(hub.snapshot)
    expect(secondController.snapshot()).toEqual(hub.snapshot)
    expect(first.state.sequence).toBe(2)
    first.dispose(); second.dispose()
  })

  it('does not mount local state before authoritative acknowledgement', async () => {
    let acknowledge!: () => void
    let submitted!: ShapeCollaborationOperation
    const transport: ShapeCollaborationTransport = {
      submit: (room, operation) => new Promise((resolve) => {
        submitted = operation
        acknowledge = () => resolve({ room, sequence: 1, operation })
      }),
      subscribe: () => () => undefined,
    }
    const controller = harness()
    const client = session(controller, transport, 'first')
    const pending = client.create(value('a'))
    await Promise.resolve()
    expect(submitted.kind).toBe('create')
    expect(controller.snapshot().shapes).toEqual([])
    acknowledge()
    await pending
    expect(controller.snapshot().shapes[0].spec.id).toBe('a')
  })

  it('preflights invalid values and enforces host authorization', async () => {
    const submit = vi.fn()
    const transport: ShapeCollaborationTransport = { submit, subscribe: () => () => undefined }
    const denied = session(harness(), transport, 'first', 0, { authorize: () => false })
    await expect(denied.create(value('a'))).rejects.toMatchObject({ code: 'permission' })
    await expect(denied.create({ spec: spec('bad'), sheetId: '' })).rejects.toMatchObject({ code: 'invalid' })
    expect(submit).not.toHaveBeenCalled()
  })

  it('blocks stale bases and resumes only from a validated resync', async () => {
    const hub = new Hub()
    hub.snapshot = applyShapeCollaborationOperation(empty(), { ...base, opId: 'seed', kind: 'create', value: value('seed') })
    hub.sequence = 1
    const controller = harness()
    const resync = vi.fn()
    const client = session(controller, hub.transport('late'), 'late', 0, { onResyncRequired: resync })
    await expect(client.create(value('a'))).rejects.toMatchObject({ code: 'stale-base' })
    expect(client.state.blocked).toBe(true)
    expect(resync).toHaveBeenCalledWith(1, -1, expect.anything())
    expect(client.resync(hub.snapshot, 1)).toBe(true)
    await client.create(value('a'))
    expect(controller.snapshot()).toEqual(hub.snapshot)
  })

  it('blocks altered acknowledgements, sequence gaps, and mount failures atomically', async () => {
    const altered: ShapeCollaborationTransport = {
      submit: async (room, operation) => ({ room, sequence: 1, operation: { ...operation, value: value('forged') } as ShapeCollaborationOperation }),
      subscribe: () => () => undefined,
    }
    const forgedController = harness()
    const forged = session(forgedController, altered, 'first')
    await expect(forged.create(value('a'))).rejects.toMatchObject({ code: 'conflict' })
    expect(forgedController.snapshot().shapes).toEqual([])

    const operation: ShapeCollaborationOperation = { ...base, clientId: 'remote', opId: 'remote-1', kind: 'create', value: value('a') }
    const gap = session(harness(), { submit: vi.fn(), subscribe: () => () => undefined }, 'gap')
    await expect(gap.receive({ room: 'book-1', sequence: 2, operation })).resolves.toBe('blocked')

    const rejecting = harness(false)
    const failedMount = session(rejecting, { submit: vi.fn(), subscribe: () => () => undefined }, 'reject')
    await expect(failedMount.receive({ room: 'book-1', sequence: 1, operation })).resolves.toBe('blocked')
    expect(rejecting.snapshot().shapes).toEqual([])
  })

  it('blocks conflicting remote fingerprints without partial state', async () => {
    const controller = harness(); controller.add(spec('a'))
    const before = controller.snapshot()
    const client = session(controller, { submit: vi.fn(), subscribe: () => () => undefined }, 'first')
    const operation: ShapeCollaborationOperation = {
      ...base, clientId: 'remote', opId: 'remote-1', kind: 'remove', id: 'a', expectedFingerprint: 'fnv1a32:00000000',
    }
    await expect(client.receive({ room: 'book-1', sequence: 1, operation })).resolves.toBe('blocked')
    expect(controller.snapshot()).toEqual(before)
  })
})
