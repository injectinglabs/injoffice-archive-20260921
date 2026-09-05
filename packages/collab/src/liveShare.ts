import { CollabPermissionManager, PermissionDeniedError } from './permissions'

export interface LiveShareViewport {
  sheetId: string
  startRow: number
  startColumn: number
  endRow: number
  endColumn: number
  zoom?: number
  active?: [row: number, column: number]
}

export type LiveShareIntent =
  | { type: 'presenter.started'; room: string; senderClientId: string; senderUserId: string; sessionId: string; permissionRevision: number }
  | { type: 'presenter.stopped'; room: string; senderClientId: string; senderUserId: string; sessionId: string; permissionRevision: number }
  | { type: 'viewport.changed'; room: string; senderClientId: string; senderUserId: string; sessionId: string; permissionRevision: number; viewport: LiveShareViewport }

export type LiveShareEvent = LiveShareIntent & { revision: number }

export interface LiveShareTransport {
  send(intent: LiveShareIntent): void | Promise<void>
  subscribe(listener: (event: unknown) => void): () => void
}

export interface LiveShareViewportAdapter {
  apply(viewport: LiveShareViewport): void
}

export type LiveShareMode = 'idle' | 'presenting' | 'following'
export type LiveShareRejectReason = 'malformed' | 'stale' | 'gap' | 'permission-revision' | 'unauthorized' | 'not-presenter'

export interface LiveShareState {
  joined: boolean
  room: string | null
  clientId: string | null
  userId: string | null
  revision: number
  desynchronized: boolean
  mode: LiveShareMode
  presenterId: string | null
  presenterUserId: string | null
  presenterSessionId: string | null
  followingId: string | null
}

export interface LiveShareOptions {
  onRejected?: (reason: LiveShareRejectReason, event: unknown) => void
  onResyncRequired?: (expectedRevision: number, receivedRevision: number) => void
  sessionIdFactory?: () => string
}

let sessionSequence = 0
const defaultSessionId = (): string => `live-${Date.now().toString(36)}-${(++sessionSequence).toString(36)}`
const copyViewport = (viewport: LiveShareViewport): LiveShareViewport => ({ ...viewport, active: viewport.active ? [...viewport.active] as [number, number] : undefined })

export function validateLiveShareViewport(value: unknown): value is LiveShareViewport {
  if (!value || typeof value !== 'object') return false
  const viewport = value as Partial<LiveShareViewport>
  const indexes = [viewport.startRow, viewport.startColumn, viewport.endRow, viewport.endColumn]
  if (typeof viewport.sheetId !== 'string' || !viewport.sheetId.trim() || indexes.some((index) => !Number.isSafeInteger(index) || index! < 0)) return false
  if (viewport.endRow! < viewport.startRow! || viewport.endColumn! < viewport.startColumn!) return false
  if (viewport.zoom !== undefined && (!Number.isFinite(viewport.zoom) || viewport.zoom < 0.1 || viewport.zoom > 8)) return false
  return viewport.active === undefined || (Array.isArray(viewport.active) && viewport.active.length === 2 && viewport.active.every((index) => Number.isSafeInteger(index) && index >= 0))
}

export function validateLiveShareEvent(value: unknown): value is LiveShareEvent {
  if (!value || typeof value !== 'object') return false
  const event = value as Partial<LiveShareEvent>
  if (!['presenter.started', 'presenter.stopped', 'viewport.changed'].includes(event.type ?? '')) return false
  if (![event.room, event.senderClientId, event.senderUserId, event.sessionId].every((item) => typeof item === 'string' && item.trim())) return false
  if (!Number.isSafeInteger(event.revision) || event.revision! < 1 || !Number.isSafeInteger(event.permissionRevision) || event.permissionRevision! < 0) return false
  return event.type !== 'viewport.changed' || validateLiveShareViewport(event.viewport)
}

/** Ordered presenter/follower state; the host owns transport and authority. */
export class LiveShareSession {
  private stateValue: LiveShareState = {
    joined: false, room: null, clientId: null, userId: null, revision: 0, desynchronized: false,
    mode: 'idle', presenterId: null, presenterUserId: null, presenterSessionId: null, followingId: null,
  }
  private unsubscribeTransport?: () => void
  private unsubscribePermissions?: () => void
  private disposed = false
  private readonly listeners = new Set<(state: LiveShareState) => void>()
  private readonly sessionIdFactory: () => string

  constructor(
    private readonly transport: LiveShareTransport,
    private readonly permissions: CollabPermissionManager,
    private readonly viewport: LiveShareViewportAdapter,
    private readonly options: LiveShareOptions = {},
  ) {
    this.sessionIdFactory = options.sessionIdFactory ?? defaultSessionId
  }

  get state(): LiveShareState { return { ...this.stateValue } }

  onChange(listener: (state: LiveShareState) => void): () => void {
    if (this.disposed) return () => undefined
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }

  join(room: string, clientId: string, userId: string, revision = 0): boolean {
    if (this.disposed || this.stateValue.joined || room !== this.permissions.room || !clientId.trim() || !userId.trim() || !Number.isSafeInteger(revision) || revision < 0) return false
    this.stateValue = { ...this.stateValue, joined: true, room, clientId, userId, revision }
    this.unsubscribeTransport = this.transport.subscribe((event) => this.receive(event))
    this.unsubscribePermissions = this.permissions.onChange(() => this.reconcilePermissions())
    this.emit()
    return true
  }

  async leave(): Promise<void> {
    if (!this.stateValue.joined) return
    try {
      if (this.stateValue.mode === 'presenting') await this.stopPresenting()
    } finally {
      this.unsubscribeTransport?.()
      this.unsubscribePermissions?.()
      this.unsubscribeTransport = undefined
      this.unsubscribePermissions = undefined
      this.stateValue = {
        joined: false, room: null, clientId: null, userId: null, revision: 0, desynchronized: false,
        mode: 'idle', presenterId: null, presenterUserId: null, presenterSessionId: null, followingId: null,
      }
      this.emit()
    }
  }

  async startPresenting(): Promise<string> {
    this.assertReady('present')
    const sessionId = this.sessionIdFactory()
    if (typeof sessionId !== 'string' || !sessionId.trim()) throw new TypeError('Invalid live-share session id')
    await this.transport.send(this.intent('presenter.started', sessionId))
    return sessionId
  }

  async stopPresenting(): Promise<void> {
    this.assertReady('present')
    if (this.stateValue.mode !== 'presenting' || !this.stateValue.presenterSessionId) throw new Error('Local client is not presenting')
    await this.transport.send(this.intent('presenter.stopped', this.stateValue.presenterSessionId))
  }

  follow(presenterId: string): boolean {
    if (!this.ready() || !this.stateValue.clientId || !this.stateValue.userId) return false
    if (!this.permissions.decide(this.stateValue.userId, 'follow').allowed || this.stateValue.presenterId !== presenterId || presenterId === this.stateValue.clientId) return false
    this.stateValue = { ...this.stateValue, mode: 'following', followingId: presenterId }
    this.emit()
    return true
  }

  stopFollowing(): boolean {
    if (this.stateValue.mode !== 'following') return false
    this.stateValue = { ...this.stateValue, mode: 'idle', followingId: null }
    this.emit()
    return true
  }

  async publishViewport(viewport: LiveShareViewport): Promise<void> {
    this.assertReady('present')
    if (this.stateValue.mode !== 'presenting' || !this.stateValue.presenterSessionId) throw new Error('Local client is not presenting')
    if (!validateLiveShareViewport(viewport)) throw new TypeError('Invalid live-share viewport')
    await this.transport.send({
      type: 'viewport.changed',
      room: this.stateValue.room!,
      senderClientId: this.stateValue.clientId!,
      senderUserId: this.stateValue.userId!,
      sessionId: this.stateValue.presenterSessionId,
      permissionRevision: this.permissions.revision,
      viewport: copyViewport(viewport),
    })
  }

  /** Apply a trusted room snapshot after a sequence gap. */
  resync(revision: number, presenter?: { clientId: string; userId: string; sessionId: string } | null): boolean {
    if (!this.stateValue.joined || !Number.isSafeInteger(revision) || revision < this.stateValue.revision) return false
    if (presenter && (!presenter.clientId.trim() || !presenter.userId.trim() || !presenter.sessionId.trim())) return false
    const presenterAllowed = !presenter || this.permissions.decide(presenter.userId, 'present').allowed
    if (!presenterAllowed) return false
    const localId = this.stateValue.clientId
    const mode: LiveShareMode = presenter?.clientId === localId ? 'presenting' : 'idle'
    this.stateValue = {
      ...this.stateValue,
      revision,
      desynchronized: false,
      presenterId: presenter?.clientId ?? null,
      presenterUserId: presenter?.userId ?? null,
      presenterSessionId: presenter?.sessionId ?? null,
      followingId: null,
      mode,
    }
    this.emit()
    return true
  }

  dispose(): void {
    if (this.disposed) return
    this.unsubscribeTransport?.()
    this.unsubscribePermissions?.()
    this.listeners.clear()
    this.disposed = true
    this.stateValue = {
      joined: false, room: null, clientId: null, userId: null, revision: 0, desynchronized: false,
      mode: 'idle', presenterId: null, presenterUserId: null, presenterSessionId: null, followingId: null,
    }
  }

  private receive(value: unknown): void {
    if (!this.stateValue.joined || this.stateValue.desynchronized) return
    if (!validateLiveShareEvent(value) || value.room !== this.stateValue.room) return this.reject('malformed', value)
    if (value.revision <= this.stateValue.revision) return this.reject('stale', value)
    if (value.revision !== this.stateValue.revision + 1) {
      this.stateValue = { ...this.stateValue, desynchronized: true }
      this.options.onResyncRequired?.(this.stateValue.revision + 1, value.revision)
      this.reject('gap', value)
      this.emit()
      return
    }
    this.stateValue = { ...this.stateValue, revision: value.revision }
    if (value.permissionRevision !== this.permissions.revision) {
      this.reject('permission-revision', value)
      return this.emit()
    }
    if (!this.permissions.decide(value.senderUserId, 'present').allowed) {
      this.reject('unauthorized', value)
      return this.emit()
    }

    if (value.type === 'presenter.started') {
      const local = value.senderClientId === this.stateValue.clientId
      this.stateValue = {
        ...this.stateValue,
        presenterId: value.senderClientId,
        presenterUserId: value.senderUserId,
        presenterSessionId: value.sessionId,
        followingId: null,
        mode: local ? 'presenting' : 'idle',
      }
      return this.emit()
    }
    if (value.senderClientId !== this.stateValue.presenterId || value.sessionId !== this.stateValue.presenterSessionId) {
      this.reject('not-presenter', value)
      return this.emit()
    }
    if (value.type === 'presenter.stopped') {
      this.stateValue = { ...this.stateValue, presenterId: null, presenterUserId: null, presenterSessionId: null, followingId: null, mode: 'idle' }
      return this.emit()
    }
    if (this.stateValue.mode === 'following' && this.stateValue.followingId === value.senderClientId) this.viewport.apply(copyViewport(value.viewport))
    this.emit()
  }

  private reconcilePermissions(): void {
    const presenterUser = this.stateValue.presenterUserId
    if (presenterUser && !this.permissions.decide(presenterUser, 'present').allowed) {
      this.stateValue = { ...this.stateValue, presenterId: null, presenterUserId: null, presenterSessionId: null, followingId: null, mode: 'idle' }
      return this.emit()
    }
    const localUser = this.stateValue.userId
    if (localUser && this.stateValue.mode === 'following' && !this.permissions.decide(localUser, 'follow').allowed) {
      this.stateValue = { ...this.stateValue, followingId: null, mode: 'idle' }
      this.emit()
    }
  }

  private ready(): boolean { return !this.disposed && this.stateValue.joined && !this.stateValue.desynchronized }

  private assertReady(capability: 'present'): void {
    if (!this.ready() || !this.stateValue.clientId || !this.stateValue.userId) throw new Error('Live-share session is not ready')
    const decision = this.permissions.decide(this.stateValue.userId, capability)
    if (!decision.allowed) throw new PermissionDeniedError(decision)
  }

  private intent(type: 'presenter.started' | 'presenter.stopped', sessionId: string): LiveShareIntent {
    return {
      type,
      room: this.stateValue.room!,
      senderClientId: this.stateValue.clientId!,
      senderUserId: this.stateValue.userId!,
      sessionId,
      permissionRevision: this.permissions.revision,
    } as LiveShareIntent
  }

  private reject(reason: LiveShareRejectReason, event: unknown): void { this.options.onRejected?.(reason, event) }
  private emit(): void { const state = this.state; for (const listener of this.listeners) listener(state) }
}
