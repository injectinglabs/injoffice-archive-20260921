import { describe, expect, it, vi } from 'vitest'
import { DocSyncEngine, type DocSyncHooks } from './docSync'
import type { CollabTransport } from './types'

function harness(server: { head: number; entries: Array<{ seq: number; ops: unknown[]; client_id: string }> }) {
  const applied: Array<{ seq?: number; ops: unknown[]; clientID: string }> = []
  const acked: unknown[][] = []
  let pending: { ops: unknown[] } | null = null
  let resyncs = 0
  const hooks: DocSyncHooks = {
    applyRemote: (ops, clientID) => {
      applied.push({ ops, clientID })
    },
    getPending: () => pending,
    onAcked: (ops) => {
      acked.push(ops as unknown[])
      pending = null
    },
    onResync: () => resyncs++,
  }
  const transport: CollabTransport = {
    join: async () => ({ room: 'r', self: { client_id: 'me', user_id: 'u', name: 'Me', color: '#000', joined_at: 0 }, peers: [], file: { mtime: 0, size: 0 } }),
    leave: async () => undefined,
    presence: async () => undefined,
    onEvent: () => () => {},
    onReconnect: () => () => {},
    opSubmit: vi.fn(async (_p, ops, base) => {
      if (base !== server.head) throw new Error('STALE_BASE: catch up and resubmit')
      server.head++
      server.entries.push({ seq: server.head, ops: ops as unknown[], client_id: 'me' })
      return server.head
    }),
    opSince: vi.fn(async (_p, since) => ({
      ops: server.entries.filter((e) => e.seq > since) as never,
      head: server.head,
      reset: false,
    })),
  }
  const engine = new DocSyncEngine(transport, '/doc.md', hooks)
  return { engine, hooks, transport, appliedLog: applied, acked, setPending: (p: { ops: unknown[] } | null) => (pending = p), resyncCount: () => resyncs }
}

describe('DocSyncEngine', () => {
  it('bootstraps from saved_seq and applies live entries in order', async () => {
    const srv = { head: 3, entries: [{ seq: 2, ops: ['b'], client_id: 'p' }, { seq: 3, ops: ['c'], client_id: 'p' }] }
    const h = harness(srv)
    await h.engine.bootstrap(1)
    expect(h.appliedLog.map((a) => a.ops[0])).toEqual(['b', 'c'])
    expect(h.engine.applied).toBe(3)
    // Out-of-order live entries buffer then drain.
    void h.engine.receiveEntry(5, ['e'], 'p')
    srv.entries.push({ seq: 4, ops: ['d'], client_id: 'p' }, { seq: 5, ops: ['e'], client_id: 'p' })
    srv.head = 5
    await h.engine.receiveEntry(4, ['d'], 'p')
    expect(h.appliedLog.map((a) => a.ops[0])).toEqual(['b', 'c', 'd', 'e'])
    expect(h.engine.applied).toBe(5)
  })

  it('submits pending against the applied base and confirms on ack', async () => {
    const srv = { head: 0, entries: [] as never[] }
    const h = harness(srv)
    await h.engine.bootstrap(0)
    h.setPending({ ops: ['s1'] })
    h.engine.submitSoon()
    await new Promise((r) => setTimeout(r, 5))
    expect(h.acked).toEqual([['s1']])
    expect(h.engine.applied).toBe(1)
  })

  it('on STALE_BASE it catches up (editor rebases) and resubmits', async () => {
    const srv = { head: 1, entries: [{ seq: 1, ops: ['remote'], client_id: 'p' }] }
    const h = harness(srv)
    await h.engine.bootstrap(1)
    // Server moves ahead while we type.
    srv.head = 2
    srv.entries.push({ seq: 2, ops: ['r2'], client_id: 'p' })
    h.setPending({ ops: ['mine'] })
    h.engine.submitSoon()
    await new Promise((r) => setTimeout(r, 10))
    // Caught up (applied r2), then the resubmit landed as seq 3.
    expect(h.appliedLog.map((a) => a.ops[0])).toEqual(['r2'])
    expect(h.acked).toEqual([['mine']])
    expect(h.engine.applied).toBe(3)
  })

  it('a failing remote apply triggers resync exactly once', async () => {
    const srv = { head: 1, entries: [{ seq: 1, ops: ['boom'], client_id: 'p' }] }
    const h = harness(srv)
    h.hooks.applyRemote = () => {
      throw new Error('bad step')
    }
    await h.engine.bootstrap(0)
    expect(h.resyncCount()).toBe(1)
    h.engine.receiveEntry(2, ['x'], 'p')
    expect(h.resyncCount()).toBe(1)
  })

  it('reset from the log triggers resync', async () => {
    const h = harness({ head: 0, entries: [] })
    ;(h.transport.opSince as ReturnType<typeof vi.fn>).mockResolvedValue({ ops: [], head: 9, reset: true })
    await h.engine.bootstrap(0)
    expect(h.resyncCount()).toBe(1)
  })

  it('awaits async applyRemote before advancing applied', async () => {
    let release!: () => void
    const gate = new Promise<void>((resolve) => {
      release = resolve
    })
    const srv = { head: 1, entries: [{ seq: 1, ops: ['slow'], client_id: 'p' }] }
    const h = harness(srv)
    h.hooks.applyRemote = async (ops, clientID) => {
      await gate
      h.appliedLog.push({ ops, clientID })
    }
    const boot = h.engine.bootstrap(0)
    await new Promise((r) => setTimeout(r, 5))
    expect(h.engine.applied).toBe(0)
    release()
    await boot
    expect(h.engine.applied).toBe(1)
    expect(h.appliedLog.map((a) => a.ops[0])).toEqual(['slow'])
  })

  it('skips catch-up apply for our own client id and never rewinds applied', async () => {
    const srv = {
      head: 2,
      entries: [
        { seq: 1, ops: ['mine'], client_id: 'me' },
        { seq: 2, ops: ['theirs'], client_id: 'p' },
      ],
    }
    const applied: unknown[] = []
    const hooks: DocSyncHooks = {
      applyRemote: (ops) => {
        applied.push(ops[0])
      },
      getPending: () => null,
      onAcked: () => undefined,
      onResync: () => undefined,
    }
    const transport: CollabTransport = {
      join: async () => ({ room: 'r', self: { client_id: 'me', user_id: 'u', name: 'Me', color: '#000', joined_at: 0 }, peers: [], file: { mtime: 0, size: 0 } }),
      leave: async () => undefined,
      presence: async () => undefined,
      onEvent: () => () => {},
      onReconnect: () => () => {},
      opSubmit: async () => 1,
      opSince: async () => ({ ops: srv.entries as never, head: srv.head, reset: false }),
    }
    const engine = new DocSyncEngine(transport, '/doc.md', hooks, 'me')
    await engine.bootstrap(0)
    expect(applied).toEqual(['theirs'])
    expect(engine.applied).toBe(2)
  })
})
