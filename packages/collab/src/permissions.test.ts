import { describe, expect, it, vi } from 'vitest'
import {
  CollabPermissionManager,
  PermissionDeniedError,
  validatePermissionSnapshot,
  type CollabPermissionSnapshot,
} from './permissions'

const snapshot = (revision = 1): CollabPermissionSnapshot => ({
  room: 'book',
  revision,
  roles: { owner: 'owner', editor: 'editor', commenter: 'commenter', viewer: 'viewer' },
  rules: [{ userId: 'editor', deny: ['present'], allow: ['manage-members'] }],
})

describe('CollabPermissionManager', () => {
  it('uses deny-first capabilities and refuses unknown principals', () => {
    const subject = new CollabPermissionManager(snapshot())
    expect(subject.decide('owner', 'manage-members').allowed).toBe(true)
    expect(subject.decide('editor', 'present')).toMatchObject({ allowed: false, reason: 'explicit-deny' })
    expect(subject.decide('editor', 'manage-members').allowed).toBe(true)
    expect(subject.decide('viewer', 'edit')).toMatchObject({ allowed: false, reason: 'role-denied' })
    expect(subject.decide('missing', 'view')).toMatchObject({ allowed: false, reason: 'unknown-user' })
    expect(subject.decide('toString', 'view')).toMatchObject({ allowed: false, reason: 'unknown-user' })
  })

  it('keeps snapshots atomic across invalid, stale, conflict, and gap updates', () => {
    const subject = new CollabPermissionManager(snapshot())
    const before = subject.snapshot
    expect(subject.replace({ ...snapshot(2), roles: { x: 'administrator' } })).toMatchObject({ ok: false, code: 'invalid' })
    expect(subject.replace(snapshot(0))).toMatchObject({ ok: false, code: 'stale' })
    expect(subject.replace({ ...snapshot(), roles: { owner: 'viewer' } })).toMatchObject({ ok: false, code: 'conflict' })
    expect(subject.applyRoleChange({ room: 'book', revision: 3, userId: 'viewer', role: 'editor' })).toMatchObject({ ok: false, code: 'gap' })
    expect(subject.snapshot).toEqual(before)
    expect(subject.applyRoleChange({ room: 'book', revision: 2, userId: 'viewer', role: 'editor' })).toMatchObject({ ok: true, revision: 2 })
    expect(subject.decide('viewer', 'edit').allowed).toBe(true)
  })

  it('provides enforcement hooks and clears listeners on disposal', () => {
    const subject = new CollabPermissionManager(snapshot())
    const listener = vi.fn()
    subject.onChange(listener)
    expect(subject.enforce('owner', 'edit', () => 42)).toBe(42)
    expect(() => subject.enforce('viewer', 'edit', () => 0)).toThrow(PermissionDeniedError)
    subject.dispose()
    expect(subject.applyRoleChange({ room: 'book', revision: 2, userId: 'viewer', role: 'editor' })).toMatchObject({ ok: false, code: 'disposed' })
    expect(subject.decide('owner', 'view')).toMatchObject({ allowed: false, reason: 'disposed' })
    expect(listener).not.toHaveBeenCalled()
  })

  it('rejects duplicate rules and malformed capability lists', () => {
    expect(validatePermissionSnapshot({ ...snapshot(), rules: [{ userId: 'viewer' }, { userId: 'viewer' }] })).toBe(false)
    expect(validatePermissionSnapshot({ ...snapshot(), rules: [{ userId: 'viewer', allow: ['fly'] }] })).toBe(false)
  })

  it('returns detached snapshots', () => {
    const subject = new CollabPermissionManager(snapshot())
    const detached = subject.snapshot
    detached.roles.owner = 'viewer'
    detached.rules![0].deny!.length = 0
    expect(subject.decide('owner', 'manage-members').allowed).toBe(true)
    expect(subject.decide('editor', 'present').allowed).toBe(false)
  })
})
