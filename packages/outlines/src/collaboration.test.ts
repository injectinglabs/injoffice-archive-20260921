import { describe, expect, it, vi } from 'vitest'
import { OutlineCommandController, type OutlineUndoRecord } from './commands'
import {
  OUTLINE_COLLABORATION_PROTOCOL,
  OutlineCollaborationSession,
  applyOutlineCollaborationOperation,
  fingerprintOutlineSnapshot,
  type OutlineCollaborationOperation,
  type OutlineCollaborationTransport,
} from './collaboration'
import { OutlineManager } from './manager'
import { sha256Hex } from './sha256'
import type { OutlineGroup } from './types'

const rows = (id = 'rows', start = 2, end = 5, collapsed = false): OutlineGroup => ({ id, sheetId: 'sheet-1', axis: 'row', start, end, collapsed })
const visibility = () => ({ hide: vi.fn(), show: vi.fn() })

function subject(authority: 'local' | 'collaboration' = 'collaboration') {
  const undo: OutlineUndoRecord[] = []
  const controller = new OutlineCommandController(new OutlineManager(visibility(), { mutationAuthority: authority }), { push: (record) => undo.push(record) })
  return { controller, undo }
}

function operation(before: OutlineGroup[], value: OutlineGroup[], action: OutlineCollaborationOperation['action'], overrides: Partial<OutlineCollaborationOperation> = {}): OutlineCollaborationOperation {
  return {
    protocol: OUTLINE_COLLABORATION_PROTOCOL,
    opId: 'op-1',
    clientId: 'client-a',
    action,
    expectedFingerprint: fingerprintOutlineSnapshot(before),
    value,
    ...overrides,
  }
}

class Hub {
  snapshot: OutlineGroup[]
  sequence = 0
  private handlers = new Map<string, (entry: unknown) => void>()

  constructor(snapshot: OutlineGroup[] = []) { this.snapshot = structuredClone(snapshot) }

  transport(clientId: string): OutlineCollaborationTransport {
    return {
      submit: async (room, value, baseSequence) => {
        if (baseSequence !== this.sequence) throw Object.assign(new Error('stale'), { code: 'STALE_BASE' })
        this.snapshot = applyOutlineCollaborationOperation(this.snapshot, value)
        const entry = { room, sequence: ++this.sequence, operation: value }
        queueMicrotask(() => { for (const [id, handler] of this.handlers) if (id !== clientId) handler(structuredClone(entry)) })
        return structuredClone(entry)
      },
      subscribe: (_room, handler) => { this.handlers.set(clientId, handler); return () => { this.handlers.delete(clientId) } },
    }
  }
}

function session(controller: OutlineCommandController, transport: OutlineCollaborationTransport, clientId: string, initialSequence = 0, options: Record<string, unknown> = {}) {
  let id = 0
  return new OutlineCollaborationSession(controller, transport, {
    room: 'book-1', clientId, initialSequence, idFactory: () => `${clientId}-${++id}`, ...options,
  })
}

describe('outline collaboration reducer', () => {
  it('enforces declared stable-ID create, update, collapse, remove, and clear differences', () => {
    const first = applyOutlineCollaborationOperation([], operation([], [rows()], 'create'))
    const updated = [rows('rows', 1, 6)]
    expect(applyOutlineCollaborationOperation(first, operation(first, updated, 'update'))).toEqual(updated)
    const collapsed = [rows('rows', 1, 6, true)]
    expect(applyOutlineCollaborationOperation(updated, operation(updated, collapsed, 'set-collapsed'))).toEqual(collapsed)
    expect(applyOutlineCollaborationOperation(collapsed, operation(collapsed, [], 'remove'))).toEqual([])

    const two = [rows('a', 1, 2), rows('b', 4, 5)]
    expect(applyOutlineCollaborationOperation(two, operation(two, [], 'clear'))).toEqual([])
    expect(() => applyOutlineCollaborationOperation(first, operation(first, [], 'update'))).toThrowError(expect.objectContaining({ code: 'invalid' }))
    expect(() => applyOutlineCollaborationOperation(updated, operation(first, [], 'remove'))).toThrowError(expect.objectContaining({ code: 'conflict' }))
  })

  it('canonicalizes snapshots and rejects unknown nested or operation fields', () => {
    const first = rows('a', 1, 3)
    const second = rows('b', 5, 7)
    expect(fingerprintOutlineSnapshot([first, second])).toBe(fingerprintOutlineSnapshot([second, first]))
    expect(() => applyOutlineCollaborationOperation([], { ...operation([], [first], 'create'), admin: true } as OutlineCollaborationOperation))
      .toThrowError(expect.objectContaining({ code: 'invalid' }))
    expect(() => applyOutlineCollaborationOperation([], operation([], [{ ...first, hidden: true } as OutlineGroup], 'create')))
      .toThrowError(expect.objectContaining({ code: 'invalid' }))
  })

  it('matches published SHA-256 vectors', () => {
    expect(sha256Hex('')).toBe('e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855')
    expect(sha256Hex('abc')).toBe('ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad')
    expect(sha256Hex('✓')).toBe('1dabba21cdad44541f6b15796f8d22978fc7ea10c46aeceeeeb66c23b3ac7604')
  })
})

describe('OutlineCollaborationSession', () => {
  it('converges two authority-gated clients through ordered lifecycle and collapse operations', async () => {
    const leftState = subject()
    const rightState = subject()
    const hub = new Hub()
    const left = session(leftState.controller, hub.transport('left'), 'left')
    const right = session(rightState.controller, hub.transport('right'), 'right')
    left.start(); right.start()

    expect(leftState.controller.add(rows()).ok).toBe(false)
    await expect(left.add(rows())).resolves.toBe(true)
    await vi.waitFor(() => expect(right.snapshot()).toEqual([rows()]))
    await expect(right.update('rows', { start: 1, end: 6 })).resolves.toBe(true)
    await vi.waitFor(() => expect(left.snapshot()[0]).toMatchObject({ start: 1, end: 6 }))
    await expect(left.setCollapsed('rows', true)).resolves.toBe(true)
    await vi.waitFor(() => expect(right.snapshot()[0]?.collapsed).toBe(true))
    await expect(right.remove('rows')).resolves.toBe(true)
    await vi.waitFor(() => expect(left.snapshot()).toEqual([]))

    expect(left.snapshot()).toEqual(hub.snapshot)
    expect(right.snapshot()).toEqual(hub.snapshot)
    expect(left.state.sequence).toBe(4)
    expect(right.state.sequence).toBe(4)
    expect(leftState.undo.map(({ label }) => label)).toEqual(['Add outline', 'Set outline collapsed state'])
    expect(rightState.undo.map(({ label }) => label)).toEqual(['Update outline', 'Remove outline'])
  })

  it('does not expose a local group before authoritative acknowledgement', async () => {
    const { controller, undo } = subject()
    let release!: () => void
    const gate = new Promise<void>((resolve) => { release = resolve })
    const transport: OutlineCollaborationTransport = {
      submit: async (room, value) => { await gate; return { room, sequence: 1, operation: value } },
      subscribe: () => () => undefined,
    }
    const target = session(controller, transport, 'local')
    const pending = target.add(rows())
    await Promise.resolve()
    expect(controller.snapshot()).toEqual([])
    expect(undo).toEqual([])
    release()
    await expect(pending).resolves.toBe(true)
    expect(controller.snapshot()).toEqual([rows()])
    expect(undo).toHaveLength(1)
  })

  it('serializes concurrent local candidates from the latest acknowledgement', async () => {
    const { controller } = subject()
    const hub = new Hub()
    const target = session(controller, hub.transport('local'), 'local')
    await expect(Promise.all([target.add(rows('a', 1, 2)), target.add(rows('b', 4, 5))])).resolves.toEqual([true, true])
    expect(target.snapshot()).toEqual(hub.snapshot)
    expect(target.state.sequence).toBe(2)
  })

  it('enforces submit and receive permissions without partial application', async () => {
    const { controller } = subject()
    const submit = vi.fn()
    const denied = session(controller, { submit, subscribe: () => () => undefined }, 'denied', 0, { authorize: () => false })
    await expect(denied.add(rows())).rejects.toMatchObject({ code: 'permission' })
    expect(submit).not.toHaveBeenCalled()
    expect(controller.snapshot()).toEqual([])

    const remote = session(controller, { submit, subscribe: () => () => undefined }, 'remote', 0, {
      authorize: (_operation: OutlineCollaborationOperation, context: { direction: string }) => context.direction === 'submit',
    })
    await expect(remote.receive({ room: 'book-1', sequence: 1, operation: operation([], [rows()], 'create', { clientId: 'other' }) })).resolves.toBe('blocked')
    expect(remote.state.blocked).toBe(true)
    expect(controller.snapshot()).toEqual([])
  })

  it('handles duplicates, blocks gaps and typed stale bases, and accepts only room-bound resync', async () => {
    const { controller, undo } = subject()
    const hub = new Hub([rows('server')])
    hub.sequence = 2
    const target = session(controller, hub.transport('late'), 'late')
    await expect(target.add(rows('local'))).rejects.toMatchObject({ code: 'stale-base' })
    expect(target.state.blocked).toBe(true)
    await expect(target.resync({ protocol: OUTLINE_COLLABORATION_PROTOCOL, room: 'wrong', sequence: 2, value: hub.snapshot })).resolves.toBe(false)
    await expect(target.resync({ protocol: OUTLINE_COLLABORATION_PROTOCOL, room: 'book-1', sequence: 2, value: hub.snapshot })).resolves.toBe(true)
    expect(target.snapshot()).toEqual([rows('server')])
    expect(undo).toEqual([])

    await expect(target.receive({ room: 'book-1', sequence: 2, operation: operation([], [rows()], 'create') })).resolves.toBe('duplicate')
    const next = [rows('server', 1, 6)]
    await expect(target.receive({ room: 'book-1', sequence: 4, operation: operation(target.snapshot(), next, 'update', { clientId: 'other' }) })).resolves.toBe('blocked')
    expect(target.state.sequence).toBe(2)
  })

  it('queues room resync behind an in-flight acknowledgement', async () => {
    const { controller, undo } = subject()
    let release!: () => void
    const gate = new Promise<void>((resolve) => { release = resolve })
    const transport: OutlineCollaborationTransport = {
      submit: async (room, value) => { await gate; return { room, sequence: 1, operation: value } },
      subscribe: () => () => undefined,
    }
    const target = session(controller, transport, 'local')
    const local = target.add(rows('local'))
    const resync = target.resync({ protocol: OUTLINE_COLLABORATION_PROTOCOL, room: 'book-1', sequence: 2, value: [rows('server', 8, 10, true)] })
    await Promise.resolve()
    expect(target.snapshot()).toEqual([])
    release()
    await expect(local).resolves.toBe(true)
    await expect(resync).resolves.toBe(true)
    expect(target.snapshot()).toEqual([rows('server', 8, 10, true)])
    expect(target.state.sequence).toBe(2)
    expect(undo).toHaveLength(1)
  })

  it('routes undo restore through the server without recursive history', async () => {
    const { controller, undo } = subject()
    const hub = new Hub()
    const target = session(controller, hub.transport('local'), 'local')
    await target.add(rows())
    const before = undo[0]!.before
    undo.length = 0
    await expect(target.restore(before)).resolves.toBe(true)
    expect(target.snapshot()).toEqual([])
    expect(hub.snapshot).toEqual([])
    expect(undo).toEqual([])
  })

  it('blocks changed acknowledgements and incomplete resync application', async () => {
    const first = subject()
    const altered = session(first.controller, {
      submit: async (room, value) => ({ room, sequence: 1, operation: { ...value, action: 'restore' } }),
      subscribe: () => () => undefined,
    }, 'local')
    await expect(altered.add(rows())).rejects.toMatchObject({ code: 'conflict' })
    expect(altered.state.blocked).toBe(true)
    expect(first.controller.snapshot()).toEqual([])

    const second = subject()
    const submit = vi.fn()
    const incomplete = session(second.controller, { submit, subscribe: () => () => undefined }, 'local', 0, {
      applySnapshot: () => true,
    })
    await expect(incomplete.resync({ protocol: OUTLINE_COLLABORATION_PROTOCOL, room: 'book-1', sequence: 3, value: [rows()] })).resolves.toBe(false)
    expect(incomplete.state).toEqual({ sequence: 0, blocked: true, disposed: false })
    await expect(incomplete.add(rows())).rejects.toMatchObject({ code: 'blocked' })
    expect(submit).not.toHaveBeenCalled()
  })

  it('isolates observers and disposes its subscription', async () => {
    const { controller } = subject()
    const dispose = vi.fn()
    const hub = new Hub()
    const target = session(controller, { ...hub.transport('local'), subscribe: () => dispose }, 'local', 0, {
      onEvent: () => { throw new Error('observer') },
    })
    target.start(); target.start()
    await expect(target.add(rows())).resolves.toBe(true)
    target.dispose(); target.dispose()
    expect(dispose).toHaveBeenCalledOnce()
    await expect(target.add(rows('other'))).rejects.toMatchObject({ code: 'blocked' })
  })
})
