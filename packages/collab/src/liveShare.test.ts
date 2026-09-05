import { describe, expect, it, vi } from 'vitest'
import { LiveShareSession, type LiveShareEvent, type LiveShareIntent, type LiveShareTransport } from './liveShare'
import { CollabPermissionManager, PermissionDeniedError, type CollabPermissionSnapshot } from './permissions'

class Hub {
  revision = 0
  listeners = new Set<(event: unknown) => void>()
  transport(): LiveShareTransport {
    return {
      send: (intent) => {
        const event: LiveShareEvent = { ...intent, revision: ++this.revision }
        for (const listener of this.listeners) listener(event)
      },
      subscribe: (listener) => { this.listeners.add(listener); return () => this.listeners.delete(listener) },
    }
  }
  emit(event: unknown): void { for (const listener of this.listeners) listener(event) }
}

const permissionSnapshot = (revision = 1): CollabPermissionSnapshot => ({
  room: 'book', revision, roles: { alice: 'editor', bob: 'viewer', eve: 'viewer' },
})

function client(hub: Hub, userId: string, clientId: string, permission = permissionSnapshot()) {
  const applied: unknown[] = []
  const permissions = new CollabPermissionManager(permission)
  const session = new LiveShareSession(hub.transport(), permissions, { apply: (viewport) => applied.push(viewport) }, { sessionIdFactory: () => `${clientId}-session` })
  expect(session.join('book', clientId, userId, hub.revision)).toBe(true)
  return { applied, permissions, session }
}

const viewport = { sheetId: 'sheet', startRow: 2, startColumn: 3, endRow: 8, endColumn: 9, zoom: 1.25 }

describe('LiveShareSession multi-client lifecycle', () => {
  it('presents, follows, applies ordered viewports, and leaves cleanly', async () => {
    const hub = new Hub()
    const alice = client(hub, 'alice', 'alice-tab')
    const bob = client(hub, 'bob', 'bob-tab')
    await alice.session.startPresenting()
    expect(alice.session.state.mode).toBe('presenting')
    expect(bob.session.state.presenterId).toBe('alice-tab')
    expect(bob.session.follow('alice-tab')).toBe(true)
    await alice.session.publishViewport(viewport)
    expect(bob.applied).toEqual([viewport])
    await alice.session.stopPresenting()
    expect(bob.session.state).toMatchObject({ mode: 'idle', presenterId: null, followingId: null })
    await alice.session.leave()
    expect(alice.session.state.joined).toBe(false)
  })

  it('enforces presenter permissions before transport publication', async () => {
    const hub = new Hub()
    const bob = client(hub, 'bob', 'bob-tab')
    await expect(bob.session.startPresenting()).rejects.toBeInstanceOf(PermissionDeniedError)
    expect(hub.revision).toBe(0)
  })

  it('freezes on a sequence gap and resumes only from a trusted resync', async () => {
    const hub = new Hub()
    const resync = vi.fn()
    const rejected = vi.fn()
    const permissions = new CollabPermissionManager(permissionSnapshot())
    const session = new LiveShareSession(hub.transport(), permissions, { apply: vi.fn() }, { onResyncRequired: resync, onRejected: rejected })
    session.join('book', 'bob-tab', 'bob', 0)
    const event: LiveShareEvent = {
      type: 'presenter.started', room: 'book', senderClientId: 'alice-tab', senderUserId: 'alice', sessionId: 's',
      permissionRevision: 1, revision: 2,
    }
    hub.emit(event)
    expect(session.state.desynchronized).toBe(true)
    expect(resync).toHaveBeenCalledWith(1, 2)
    hub.emit({ ...event, revision: 3 })
    expect(session.state.revision).toBe(0)
    expect(session.resync(3, { clientId: 'alice-tab', userId: 'alice', sessionId: 's' })).toBe(true)
    expect(session.state).toMatchObject({ desynchronized: false, presenterId: 'alice-tab', revision: 3 })
    hub.emit({ ...event, revision: 3 })
    expect(rejected).toHaveBeenCalledWith('stale', expect.anything())
  })

  it('advances but rejects unauthorized, stale-permission, and forged presenter events', async () => {
    const hub = new Hub()
    const rejected = vi.fn()
    const permissions = new CollabPermissionManager(permissionSnapshot())
    const session = new LiveShareSession(hub.transport(), permissions, { apply: vi.fn() }, { onRejected: rejected })
    session.join('book', 'bob-tab', 'bob', 0)
    const base = { room: 'book', sessionId: 's', permissionRevision: 1 }
    hub.emit({ ...base, type: 'presenter.started', senderClientId: 'eve-tab', senderUserId: 'eve', revision: 1 })
    expect(rejected).toHaveBeenCalledWith('unauthorized', expect.anything())
    hub.emit({ ...base, type: 'presenter.started', senderClientId: 'alice-tab', senderUserId: 'alice', permissionRevision: 0, revision: 2 })
    expect(rejected).toHaveBeenCalledWith('permission-revision', expect.anything())
    hub.emit({ ...base, type: 'viewport.changed', senderClientId: 'alice-tab', senderUserId: 'alice', revision: 3, viewport })
    expect(rejected).toHaveBeenCalledWith('not-presenter', expect.anything())
    expect(session.state.revision).toBe(3)
  })

  it('ends presentation and following when a role is revoked', async () => {
    const hub = new Hub()
    const alice = client(hub, 'alice', 'alice-tab')
    const bob = client(hub, 'bob', 'bob-tab')
    await alice.session.startPresenting()
    bob.session.follow('alice-tab')
    expect(alice.permissions.applyRoleChange({ room: 'book', revision: 2, userId: 'alice', role: 'viewer' }).ok).toBe(true)
    expect(bob.permissions.applyRoleChange({ room: 'book', revision: 2, userId: 'alice', role: 'viewer' }).ok).toBe(true)
    expect(alice.session.state.mode).toBe('idle')
    expect(bob.session.state).toMatchObject({ mode: 'idle', presenterId: null })
  })

  it('validates viewports and disposes subscriptions', async () => {
    const hub = new Hub()
    const alice = client(hub, 'alice', 'alice-tab')
    await alice.session.startPresenting()
    await expect(alice.session.publishViewport({ ...viewport, zoom: 99 })).rejects.toBeInstanceOf(TypeError)
    alice.session.dispose()
    expect(hub.listeners.size).toBe(0)
    await expect(alice.session.publishViewport(viewport)).rejects.toThrow('not ready')
  })

  it('unsubscribes and clears local state when a presenter leave send fails', async () => {
    let listener: ((event: unknown) => void) | undefined
    const unsubscribe = vi.fn()
    const transport: LiveShareTransport = {
      send: () => Promise.reject(new Error('network down')),
      subscribe: (next) => { listener = next; return unsubscribe },
    }
    const permissions = new CollabPermissionManager(permissionSnapshot())
    const session = new LiveShareSession(transport, permissions, { apply: vi.fn() })
    session.join('book', 'alice-tab', 'alice', 0)
    listener!({
      type: 'presenter.started', room: 'book', senderClientId: 'alice-tab', senderUserId: 'alice',
      sessionId: 's', permissionRevision: 1, revision: 1,
    })
    await expect(session.leave()).rejects.toThrow('network down')
    expect(unsubscribe).toHaveBeenCalledOnce()
    expect(session.state).toMatchObject({ joined: false, mode: 'idle', presenterId: null })
  })
})
