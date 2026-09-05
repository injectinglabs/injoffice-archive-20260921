import { describe, expect, it, vi } from 'vitest'
import { PdfPresenceManager, type PdfSelection } from './pdfManager'
import type { CollabEvent, CollabTransport, JoinResult } from './types'

function fakeTransport(peers: JoinResult<PdfSelection>['peers'] = []) {
  let handler: ((ev: CollabEvent) => void) | null = null
  const t: CollabTransport<PdfSelection> = {
    join: vi.fn(async () => ({
      room: 'r1',
      self: { client_id: 'me', user_id: 'u', name: 'Me', color: '#0f9f88', joined_at: 1 },
      peers,
      file: { mtime: 1, size: 2 },
    })),
    leave: vi.fn(async () => undefined),
    presence: vi.fn(async () => undefined),
    onEvent: (h) => {
      handler = h
      return () => {
        handler = null
      }
    },
    onReconnect: () => () => {},
  }
  return { t, emit: (ev: CollabEvent) => handler?.(ev) }
}

describe('PdfPresenceManager', () => {
  it('joins, tracks peers, and surfaces remote saves (never its own)', async () => {
    const { t, emit } = fakeTransport([{ client_id: 'a', user_id: 'ua', name: 'Ann', color: '#333', joined_at: 0 }])
    const m = new PdfPresenceManager(t, { path: '/v1/files/x.pdf', name: 'Me' })
    const changes: number[] = []
    const files: string[] = []
    m.onChange(() => changes.push(m.peers().length))
    m.onFileChanged((c) => files.push(c.action))
    await m.start()
    expect(m.clientId).toBe('me')
    expect(m.peers().map((p) => p.name)).toEqual(['Ann'])

    emit({ event: 'collab.peer.joined', payload: { room: 'r1', peer: { client_id: 'b', user_id: 'ub', name: 'Bob', color: '#444', joined_at: 2 } } })
    expect(m.peers()).toHaveLength(2)
    emit({ event: 'collab.file.changed', payload: { room: 'r1', path: '/x.pdf', author: 'user', action: 'saved', origin: 'a', mtime: 3, size: 4 } })
    emit({ event: 'collab.file.changed', payload: { room: 'r1', path: '/x.pdf', author: 'user', action: 'saved', origin: 'me', mtime: 5, size: 6 } })
    expect(files).toEqual(['saved']) // own save filtered

    m.stop()
    expect(t.leave).toHaveBeenCalledWith('/v1/files/x.pdf')
  })

  it('survives a transport that rejects join (old server)', async () => {
    const t: CollabTransport<PdfSelection> = {
      join: async () => {
        throw new Error('unknown method')
      },
      leave: async () => undefined,
      presence: async () => undefined,
      onEvent: () => () => {},
      onReconnect: () => () => {},
    }
    const m = new PdfPresenceManager(t, { path: '/x.pdf', name: 'Me' })
    await m.start()
    expect(m.clientId).toBeNull()
    expect(m.peers()).toEqual([])
    m.stop()
  })

  it('publish calls transport.presence', async () => {
    const { t } = fakeTransport()
    const m = new PdfPresenceManager(t, { path: '/v1/files/x.pdf', name: 'Me' })
    m.publish({ page: 1 })
    expect(t.presence).not.toHaveBeenCalled()
    await m.start()
    m.publish({ page: 2, annotLocalId: 'ann-1' })
    expect(t.presence).toHaveBeenCalledWith('/v1/files/x.pdf', { page: 2, annotLocalId: 'ann-1' })
    m.stop()
  })

  it('remote collab.presence updates selection and drops bad input', async () => {
    const { t, emit } = fakeTransport([{ client_id: 'a', user_id: 'ua', name: 'Ann', color: '#333', joined_at: 0 }])
    const m = new PdfPresenceManager(t, { path: '/v1/files/x.pdf', name: 'Me' })
    await m.start()

    emit({ event: 'collab.presence', payload: { room: 'r1', client_id: 'a', selection: { page: 3, annotLocalId: 'a1' } } })
    expect(m.peers()[0].selection).toEqual({ page: 3, annotLocalId: 'a1' })

    emit({ event: 'collab.presence', payload: { room: 'r1', client_id: 'a', selection: { page: 0 } } })
    expect(m.peers()[0].selection).toBeNull()
    emit({ event: 'collab.presence', payload: { room: 'r1', client_id: 'a', selection: { page: 1.5 } } })
    expect(m.peers()[0].selection).toBeNull()
    emit({ event: 'collab.presence', payload: { room: 'r1', client_id: 'a', selection: { page: 4 } } })
    expect(m.peers()[0].selection).toEqual({ page: 4, annotLocalId: null })

    m.stop()
  })
})

describe('PdfPresenceManager sync', () => {
  it('startSync bootstraps and submitSoon flushes pending', async () => {
    const opSubmit = vi.fn(async () => 3)
    const opSince = vi.fn(async () => ({
      ops: [{ room: 'r1', seq: 2, client_id: 'a', ops: [{ remote: true } as never] }],
      head: 2,
      reset: false,
    }))
    const { t } = fakeTransport()
    t.join = async () => ({
      room: 'r1',
      self: { client_id: 'me', user_id: 'u', name: 'Me', color: '#0f9f88', joined_at: 1 },
      peers: [],
      file: { mtime: 1, size: 2 },
      log: { seq: 1, saved_seq: 1 },
    })
    t.opSince = opSince
    t.opSubmit = opSubmit
    const m = new PdfPresenceManager(t, { path: '/v1/files/x.pdf', name: 'Me' })
    await m.start()
    let pending: { ops: unknown[] } | null = { ops: [{ step: 'typed-during-bootstrap' }] }
    const applyRemote = vi.fn()
    await m.startSync({
      applyRemote,
      getPending: () => pending,
      onAcked: () => {
        pending = null
      },
      onResync: vi.fn(),
    })
    expect(opSince).toHaveBeenCalledWith('/v1/files/x.pdf', 1)
    expect(applyRemote).toHaveBeenCalledWith([{ remote: true }], 'a')
    await vi.waitFor(() => expect(opSubmit).toHaveBeenCalledTimes(1))
    expect(opSubmit).toHaveBeenCalledWith(
      '/v1/files/x.pdf',
      [{ step: 'typed-during-bootstrap' }],
      2,
    )
    m.stop()
  })

  it('remote collab.op calls applyRemote', async () => {
    const { t, emit } = fakeTransport()
    t.join = async () => ({
      room: 'r1',
      self: { client_id: 'me', user_id: 'u', name: 'Me', color: '#0f9f88', joined_at: 1 },
      peers: [],
      file: { mtime: 1, size: 2 },
      log: { seq: 0, saved_seq: 0 },
    })
    t.opSince = async () => ({ ops: [], head: 0, reset: false })
    t.opSubmit = async () => 1
    const m = new PdfPresenceManager(t, { path: '/v1/files/x.pdf', name: 'Me' })
    await m.start()
    const applyRemote = vi.fn()
    await m.startSync({
      applyRemote,
      getPending: () => null,
      onAcked: vi.fn(),
      onResync: vi.fn(),
    })
    emit({ event: 'collab.op', payload: { room: 'r1', seq: 1, client_id: 'a', ops: [{ x: 1 } as never] } })
    await vi.waitFor(() => expect(applyRemote).toHaveBeenCalledWith([{ x: 1 }], 'a'))
    await vi.waitFor(() => expect(m.appliedSeq).toBe(1))
    m.stop()
  })

  it('ignores peer saves already covered by engine.applied unless reset', async () => {
    const { t, emit } = fakeTransport()
    t.join = async () => ({
      room: 'r1',
      self: { client_id: 'me', user_id: 'u', name: 'Me', color: '#0f9f88', joined_at: 1 },
      peers: [],
      file: { mtime: 1, size: 2 },
      log: { seq: 4, saved_seq: 4 },
    })
    t.opSince = async () => ({ ops: [], head: 4, reset: false })
    t.opSubmit = async () => 5
    const m = new PdfPresenceManager(t, { path: '/v1/files/x.pdf', name: 'Me' })
    const files: Array<{ action: string; reset?: boolean }> = []
    m.onFileChanged((c) => files.push({ action: c.action, reset: c.reset }))
    await m.start()
    await m.startSync({
      applyRemote: vi.fn(),
      getPending: () => null,
      onAcked: vi.fn(),
      onResync: vi.fn(),
    })
    expect(m.appliedSeq).toBe(4)
    emit({ event: 'collab.file.changed', payload: { room: 'r1', path: '/x.pdf', author: 'user', action: 'saved', origin: 'a', mtime: 3, size: 4, saved_seq: 4 } })
    emit({ event: 'collab.file.changed', payload: { room: 'r1', path: '/x.pdf', author: 'agent', action: 'delivered', origin: 'a', mtime: 5, size: 6, saved_seq: 4, reset: true } })
    expect(files).toEqual([{ action: 'delivered', reset: true }])
    m.stop()
  })
})
