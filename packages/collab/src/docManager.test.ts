import { describe, expect, it, vi } from 'vitest'
import { DocPresenceManager, type DocSelection } from './docManager'
import type { CollabEvent, CollabTransport, JoinResult } from './types'

function fakeTransport(peers: JoinResult<DocSelection>['peers'] = []) {
  let handler: ((ev: CollabEvent) => void) | null = null
  const t: CollabTransport<DocSelection> = {
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

describe('DocPresenceManager', () => {
  it('joins, tracks peers, and surfaces remote saves (never its own)', async () => {
    const { t, emit } = fakeTransport([{ client_id: 'a', user_id: 'ua', name: 'Ann', color: '#333', joined_at: 0 }])
    const m = new DocPresenceManager(t, { path: '/v1/files/x.md', name: 'Me' })
    const changes: number[] = []
    const files: string[] = []
    m.onChange(() => changes.push(m.peers().length))
    m.onFileChanged((c) => files.push(c.action))
    await m.start()
    expect(m.clientId).toBe('me')
    expect(m.peers().map((p) => p.name)).toEqual(['Ann'])

    emit({ event: 'collab.peer.joined', payload: { room: 'r1', peer: { client_id: 'b', user_id: 'ub', name: 'Bob', color: '#444', joined_at: 2 } } })
    expect(m.peers()).toHaveLength(2)
    emit({ event: 'collab.file.changed', payload: { room: 'r1', path: '/x.md', author: 'user', action: 'saved', origin: 'a', mtime: 3, size: 4 } })
    emit({ event: 'collab.file.changed', payload: { room: 'r1', path: '/x.md', author: 'user', action: 'saved', origin: 'me', mtime: 5, size: 6 } })
    expect(files).toEqual(['saved']) // own save filtered

    m.stop()
    expect(t.leave).toHaveBeenCalledWith('/v1/files/x.md')
  })

  it('publishes cursor presence over the transport when joined', async () => {
    const { t } = fakeTransport()
    const m = new DocPresenceManager(t, { path: '/v1/files/x.md', name: 'Me' })
    m.publish({ from: 1, to: 2 })
    expect(t.presence).not.toHaveBeenCalled()
    await m.start()
    m.publish({ from: 3, to: 8 })
    expect(t.presence).toHaveBeenCalledWith('/v1/files/x.md', { from: 3, to: 8 })
    m.stop()
  })

  it('applies remote cursor presence and drops invalid payloads', async () => {
    const { t, emit } = fakeTransport([{ client_id: 'a', user_id: 'ua', name: 'Ann', color: '#333', joined_at: 0 }])
    const m = new DocPresenceManager(t, { path: '/v1/files/x.md', name: 'Me' })
    await m.start()
    emit({ event: 'collab.presence', payload: { room: 'r1', client_id: 'a', selection: { from: 4, to: 9 } } })
    expect(m.peers()[0].selection).toEqual({ from: 4, to: 9 })
    emit({ event: 'collab.presence', payload: { room: 'r1', client_id: 'a', selection: { from: -1, to: 2 } } })
    expect(m.peers()[0].selection).toBeNull()
    m.stop()
  })

  it('survives a transport that rejects join (old server)', async () => {
    const t: CollabTransport<DocSelection> = {
      join: async () => {
        throw new Error('unknown method')
      },
      leave: async () => undefined,
      presence: async () => undefined,
      onEvent: () => () => {},
      onReconnect: () => () => {},
    }
    const m = new DocPresenceManager(t, { path: '/x.md', name: 'Me' })
    await m.start()
    expect(m.clientId).toBeNull()
    expect(m.peers()).toEqual([])
    m.stop()
  })
})

describe('restartSync', () => {
  it('submits steps authored while the sync engine was attaching', async () => {
    const opSubmit = vi.fn(async () => 1)
    const { t } = fakeTransport()
    t.join = async () => ({
      room: 'r1',
      self: { client_id: 'me', user_id: 'u', name: 'Me', color: '#0f9f88', joined_at: 1 },
      peers: [],
      file: { mtime: 1, size: 2 },
      log: { seq: 0, saved_seq: 0 },
    })
    t.opSince = async () => ({ ops: [], head: 0, reset: false })
    t.opSubmit = opSubmit
    const m = new DocPresenceManager(t, { path: '/v1/files/x.md', name: 'Me' })
    await m.start()
    let pending: { ops: unknown[] } | null = { ops: [{ step: 'typed-during-bootstrap' }] }
    await m.startSync({
      applyRemote: vi.fn(),
      getPending: () => pending,
      onAcked: () => {
        pending = null
      },
      onResync: vi.fn(),
    })
    await vi.waitFor(() => expect(opSubmit).toHaveBeenCalledTimes(1))
    expect(opSubmit).toHaveBeenCalledWith(
      '/v1/files/x.md',
      [{ step: 'typed-during-bootstrap' }],
      0,
    )
    m.stop()
  })

  it('drops the dead engine and bootstraps a fresh one from the re-joined log', async () => {
    let joinCount = 0
    let handler: ((ev: CollabEvent) => void) | null = null
    const t: CollabTransport<DocSelection> = {
      join: async (): Promise<JoinResult<DocSelection>> => {
        joinCount++
        return {
          room: 'r1',
          self: { client_id: 'me', user_id: 'u', name: 'Me', color: '#0f9f88', joined_at: 1 },
          peers: [],
          file: { mtime: 1, size: 2 },
          // Second join (post-reset) reports the fresh log base.
          log: { seq: joinCount === 1 ? 4 : 9, saved_seq: joinCount === 1 ? 4 : 9 },
        }
      },
      leave: async () => undefined,
      presence: async () => undefined,
      onEvent: (h) => {
        handler = h
        return () => {
          handler = null
        }
      },
      onReconnect: () => () => {},
      opSubmit: async () => 1,
      opSince: async () => ({ ops: [], head: 0, reset: false }),
    }
    const m = new DocPresenceManager(t, { path: '/v1/files/x.md', name: 'Me' })
    await m.start()
    const hooks = {
      applyRemote: vi.fn(),
      getPending: () => null,
      onAcked: vi.fn(),
      onResync: vi.fn(),
    }
    const first = await m.startSync(hooks)
    expect(first).not.toBeNull()
    expect(m.appliedSeq).toBe(4)

    const second = await m.restartSync(hooks)
    expect(second).not.toBeNull()
    expect(second).not.toBe(first)
    expect(joinCount).toBe(2)
    expect(m.appliedSeq).toBe(9)
    // The dead engine ignores stale entries; the new one applies fresh ones.
    if (handler) (handler as (ev: CollabEvent) => void)({ event: 'collab.op', payload: { room: 'r1', seq: 10, client_id: 'a', ops: [{ x: 1 } as never] } })
    await vi.waitFor(() => expect(hooks.applyRemote).toHaveBeenCalledWith([{ x: 1 }], 'a'))
    await vi.waitFor(() => expect(m.appliedSeq).toBe(10))
    m.stop()
  })
})
