export const COLLAB_CAPABILITIES = [
  'view', 'edit', 'comment', 'present', 'follow', 'manage-members',
] as const

export type CollabCapability = typeof COLLAB_CAPABILITIES[number]
export type CollabRole = 'owner' | 'editor' | 'commenter' | 'viewer'

export interface CollabPermissionRule {
  userId: string
  allow?: CollabCapability[]
  deny?: CollabCapability[]
}

export interface CollabPermissionSnapshot {
  room: string
  revision: number
  roles: Record<string, CollabRole>
  rules?: CollabPermissionRule[]
}

export interface CollabRoleChange {
  room: string
  revision: number
  userId: string
  role: CollabRole | null
}

export type PermissionRejectCode = 'disposed' | 'invalid' | 'room' | 'stale' | 'conflict' | 'gap'
export type PermissionUpdateResult =
  | { ok: true; revision: number; changed: boolean }
  | { ok: false; code: PermissionRejectCode; message: string }

export interface PermissionDecision {
  allowed: boolean
  userId: string
  capability: CollabCapability
  revision: number
  role?: CollabRole
  reason?: 'unknown-user' | 'role-denied' | 'explicit-deny' | 'disposed'
}

const ROLE_CAPABILITIES: Readonly<Record<CollabRole, ReadonlySet<CollabCapability>>> = {
  owner: new Set(COLLAB_CAPABILITIES),
  editor: new Set(['view', 'edit', 'comment', 'present', 'follow']),
  commenter: new Set(['view', 'comment', 'follow']),
  viewer: new Set(['view', 'follow']),
}
const CAPABILITY_SET = new Set<string>(COLLAB_CAPABILITIES)
const ROLE_SET = new Set<string>(Object.keys(ROLE_CAPABILITIES))

function clone(snapshot: CollabPermissionSnapshot): CollabPermissionSnapshot {
  return {
    room: snapshot.room,
    revision: snapshot.revision,
    roles: { ...snapshot.roles },
    rules: snapshot.rules?.map((rule) => ({
      userId: rule.userId,
      allow: rule.allow ? [...rule.allow] : undefined,
      deny: rule.deny ? [...rule.deny] : undefined,
    })),
  }
}

export function validatePermissionSnapshot(value: unknown): value is CollabPermissionSnapshot {
  if (!value || typeof value !== 'object') return false
  const input = value as Partial<CollabPermissionSnapshot>
  if (typeof input.room !== 'string' || !input.room.trim() || !Number.isSafeInteger(input.revision) || input.revision! < 0) return false
  if (!input.roles || typeof input.roles !== 'object' || Array.isArray(input.roles)) return false
  for (const [userId, role] of Object.entries(input.roles)) {
    if (!userId.trim() || !ROLE_SET.has(role)) return false
  }
  if (input.rules !== undefined && !Array.isArray(input.rules)) return false
  const seen = new Set<string>()
  for (const rule of input.rules ?? []) {
    if (!rule || typeof rule.userId !== 'string' || !rule.userId.trim() || seen.has(rule.userId)) return false
    seen.add(rule.userId)
    for (const list of [rule.allow, rule.deny]) {
      if (list !== undefined && (!Array.isArray(list) || new Set(list).size !== list.length || list.some((item) => !CAPABILITY_SET.has(item)))) return false
    }
  }
  return true
}

export class PermissionDeniedError extends Error {
  readonly decision: PermissionDecision
  constructor(decision: PermissionDecision) {
    super(`Permission denied: ${decision.userId} cannot ${decision.capability}`)
    this.name = 'PermissionDeniedError'
    this.decision = decision
  }
}

/**
 * Client-side capability projection. Snapshots must come from a trusted host;
 * this class is an enforcement hook, not a replacement for server authz.
 */
export class CollabPermissionManager {
  private current: CollabPermissionSnapshot
  private disposed = false
  private readonly listeners = new Set<(snapshot: CollabPermissionSnapshot) => void>()

  constructor(initial: CollabPermissionSnapshot) {
    if (!validatePermissionSnapshot(initial)) throw new TypeError('Invalid collaboration permission snapshot')
    this.current = clone(initial)
  }

  get room(): string { return this.current.room }
  get revision(): number { return this.current.revision }
  get snapshot(): CollabPermissionSnapshot { return clone(this.current) }

  onChange(listener: (snapshot: CollabPermissionSnapshot) => void): () => void {
    if (this.disposed) return () => undefined
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }

  replace(next: unknown): PermissionUpdateResult {
    if (this.disposed) return { ok: false, code: 'disposed', message: 'permission manager is disposed' }
    if (!validatePermissionSnapshot(next)) return { ok: false, code: 'invalid', message: 'invalid permission snapshot' }
    if (next.room !== this.current.room) return { ok: false, code: 'room', message: 'permission snapshot belongs to another room' }
    if (next.revision < this.current.revision) return { ok: false, code: 'stale', message: 'permission snapshot revision is stale' }
    if (next.revision === this.current.revision) {
      const identical = JSON.stringify(clone(next)) === JSON.stringify(this.current)
      return identical
        ? { ok: true, revision: next.revision, changed: false }
        : { ok: false, code: 'conflict', message: 'permission snapshot conflicts at the current revision' }
    }
    this.current = clone(next)
    this.emit()
    return { ok: true, revision: next.revision, changed: true }
  }

  applyRoleChange(change: CollabRoleChange): PermissionUpdateResult {
    if (this.disposed) return { ok: false, code: 'disposed', message: 'permission manager is disposed' }
    if (!change || change.room !== this.current.room) return { ok: false, code: 'room', message: 'role change belongs to another room' }
    if (!Number.isSafeInteger(change.revision) || change.revision < 0 || !change.userId?.trim() || (change.role !== null && !ROLE_SET.has(change.role))) {
      return { ok: false, code: 'invalid', message: 'invalid role change' }
    }
    if (change.revision <= this.current.revision) return { ok: false, code: 'stale', message: 'role change revision is stale' }
    if (change.revision !== this.current.revision + 1) return { ok: false, code: 'gap', message: 'role change revision has a gap' }
    const roles = { ...this.current.roles }
    if (change.role === null) delete roles[change.userId]
    else roles[change.userId] = change.role
    this.current = { ...this.current, revision: change.revision, roles }
    this.emit()
    return { ok: true, revision: change.revision, changed: true }
  }

  decide(userId: string, capability: CollabCapability): PermissionDecision {
    if (this.disposed) return { allowed: false, userId, capability, revision: this.current.revision, reason: 'disposed' }
    const role = Object.hasOwn(this.current.roles, userId) ? this.current.roles[userId] : undefined
    if (!role) return { allowed: false, userId, capability, revision: this.current.revision, reason: 'unknown-user' }
    const rule = this.current.rules?.find((candidate) => candidate.userId === userId)
    if (rule?.deny?.includes(capability)) return { allowed: false, userId, capability, revision: this.current.revision, role, reason: 'explicit-deny' }
    if (rule?.allow?.includes(capability) || ROLE_CAPABILITIES[role].has(capability)) {
      return { allowed: true, userId, capability, revision: this.current.revision, role }
    }
    return { allowed: false, userId, capability, revision: this.current.revision, role, reason: 'role-denied' }
  }

  enforce<T>(userId: string, capability: CollabCapability, operation: () => T): T {
    const decision = this.decide(userId, capability)
    if (!decision.allowed) throw new PermissionDeniedError(decision)
    return operation()
  }

  dispose(): void {
    this.disposed = true
    this.listeners.clear()
  }

  private emit(): void {
    const snapshot = this.snapshot
    for (const listener of this.listeners) listener(snapshot)
  }
}
