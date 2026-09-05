import { SparklineManager } from './manager'
import type { SparklineSnapshotV1, SparklineSpec } from './types'

export const SPARKLINE_COLLABORATION_PROTOCOL = 'injoffice.sparkline-ops.v1' as const

interface OperationBase { protocol: typeof SPARKLINE_COLLABORATION_PROTOCOL; opId: string; clientId: string }
export type SparklineCollaborationOperation =
  | OperationBase & { kind: 'create'; value: SparklineSpec }
  | OperationBase & { kind: 'update'; id: string; expectedFingerprint: string; value: SparklineSpec }
  | OperationBase & { kind: 'remove'; id: string; expectedFingerprint: string }
  | OperationBase & { kind: 'group'; groupId: string; memberIds: string[]; expectedFingerprints: Record<string, string> }
  | OperationBase & { kind: 'ungroup'; memberIds: string[]; expectedFingerprints: Record<string, string> }

export interface SparklineCollaborationEntry {
  room: string
  sequence: number
  operation: SparklineCollaborationOperation
}

export interface SparklineCollaborationTransport {
  /** The server must validate baseSequence and apply the operation's content
   * preconditions before appending and returning the next ordered entry. */
  submit(room: string, operation: SparklineCollaborationOperation, baseSequence: number, signal: AbortSignal): Promise<SparklineCollaborationEntry>
  subscribe(room: string, handler: (entry: unknown) => void): () => void
}

export type SparklineCollaborationErrorCode = 'blocked' | 'conflict' | 'invalid' | 'permission' | 'stale-base' | 'transport'

export class SparklineCollaborationError extends Error {
  constructor(readonly code: SparklineCollaborationErrorCode, message: string, options?: ErrorOptions) {
    super(message, options)
    this.name = 'SparklineCollaborationError'
  }
}

export type SparklineCollaborationEvent =
  | { type: 'submitted' | 'applied'; entry: SparklineCollaborationEntry }
  | { type: 'duplicate'; sequence: number }
  | { type: 'blocked'; expectedSequence: number; receivedSequence: number; cause?: unknown }

export interface SparklineCollaborationOptions {
  room: string
  clientId: string
  initialSequence?: number
  idFactory: () => string
  authorize?: (operation: SparklineCollaborationOperation) => boolean | Promise<boolean>
  onEvent?: (event: SparklineCollaborationEvent) => void
  onResyncRequired?: (expectedSequence: number, receivedSequence: number, cause?: unknown) => void
}

/** Pure reducer suitable for both the authoritative server and clients. It
 * validates stable object fingerprints before every destructive operation. */
export function applySparklineCollaborationOperation(
  snapshotInput: SparklineSnapshotV1,
  operationInput: SparklineCollaborationOperation,
): SparklineSnapshotV1 {
  const snapshot = clone(snapshotInput)
  const operation = validateOperation(operationInput)
  const manager = new SparklineManager()
  manager.hydrate(snapshot)
  if (operation.kind === 'create') {
    manager.add(operation.value)
  } else if (operation.kind === 'update') {
    const current = requireFingerprint(manager, operation.id, operation.expectedFingerprint)
    if (operation.value.id !== operation.id || operation.value.groupId !== current.groupId) conflict('Update cannot change stable id or group membership')
    const { id: _id, groupId: _groupId, ...replacement } = operation.value
    manager.update(operation.id, replacement)
  } else if (operation.kind === 'remove') {
    requireFingerprint(manager, operation.id, operation.expectedFingerprint)
    if (!manager.remove(operation.id)) conflict(`Sparkline ${operation.id} disappeared`)
  } else {
    const affected = affectedIds(manager, operation.memberIds)
    if (canonical([...affected].sort()) !== canonical(Object.keys(operation.expectedFingerprints).sort())) conflict('Group preconditions do not cover every affected member')
    for (const id of affected) requireFingerprint(manager, id, operation.expectedFingerprints[id] ?? '')
    if (operation.kind === 'group') manager.group(operation.memberIds, operation.groupId)
    else {
      if (!operation.memberIds.some((id) => manager.get(id)?.groupId)) conflict('Selected sparklines are not grouped')
      manager.ungroup(operation.memberIds)
    }
  }
  return manager.serialize()
}

export function fingerprintSparkline(spec: SparklineSpec): string {
  return hash(canonical(spec))
}

/** Server-first session: local state changes only after an authoritative ack;
 * remote entries use the same pure reducer and block on gaps or conflicts. */
export class SparklineCollaborationSession {
  private sequence: number
  private blocked = false
  private disposed = false
  private unsubscribe?: () => void
  private tail: Promise<void> = Promise.resolve()

  constructor(
    readonly manager: SparklineManager,
    private readonly transport: SparklineCollaborationTransport,
    private readonly options: SparklineCollaborationOptions,
  ) {
    if (!validId(options.room) || !validId(options.clientId) || typeof options.idFactory !== 'function') throw new TypeError('room, clientId, and idFactory are required')
    this.sequence = options.initialSequence ?? 0
    if (!Number.isSafeInteger(this.sequence) || this.sequence < 0) throw new TypeError('initialSequence must be a non-negative safe integer')
  }

  get state(): { sequence: number; blocked: boolean; disposed: boolean } { return { sequence: this.sequence, blocked: this.blocked, disposed: this.disposed } }

  start(): void {
    this.assertUsable()
    if (this.unsubscribe) return
    const unsubscribe = this.transport.subscribe(this.options.room, (entry) => { void this.receive(entry) })
    if (typeof unsubscribe !== 'function') throw new TypeError('transport.subscribe must return an unsubscribe function')
    this.unsubscribe = unsubscribe
  }

  create(value: SparklineSpec, signal?: AbortSignal): Promise<SparklineCollaborationEntry> {
    return this.submit(this.base({ kind: 'create', value: clone(value) }), signal)
  }

  update(id: string, patch: Partial<Omit<SparklineSpec, 'id' | 'groupId'>>, signal?: AbortSignal): Promise<SparklineCollaborationEntry> {
    const current = this.require(id)
    return this.submit(this.base({ kind: 'update', id, expectedFingerprint: fingerprintSparkline(current), value: { ...current, ...clone(patch), id, groupId: current.groupId } }), signal)
  }

  remove(id: string, signal?: AbortSignal): Promise<SparklineCollaborationEntry> {
    const current = this.require(id)
    return this.submit(this.base({ kind: 'remove', id, expectedFingerprint: fingerprintSparkline(current) }), signal)
  }

  group(memberIds: readonly string[], groupId: string, signal?: AbortSignal): Promise<SparklineCollaborationEntry> {
    const members = unique(memberIds)
    return this.submit(this.base({ kind: 'group', groupId, memberIds: members, expectedFingerprints: this.fingerprints(affectedIds(this.manager, members)) }), signal)
  }

  ungroup(memberIds: readonly string[], signal?: AbortSignal): Promise<SparklineCollaborationEntry> {
    const members = unique(memberIds)
    return this.submit(this.base({ kind: 'ungroup', memberIds: members, expectedFingerprints: this.fingerprints(affectedIds(this.manager, members)) }), signal)
  }

  receive(entryInput: unknown): Promise<'applied' | 'duplicate' | 'blocked'> {
    const result = this.tail.then(() => this.applyEntry(entryInput, false))
    this.tail = result.then(() => undefined, () => undefined)
    return result
  }

  resync(snapshot: SparklineSnapshotV1, sequence: number): boolean {
    if (this.disposed || !Number.isSafeInteger(sequence) || sequence < this.sequence) return false
    try { this.manager.hydrate(snapshot) } catch { return false }
    this.sequence = sequence
    this.blocked = false
    return true
  }

  dispose(): void { if (!this.disposed) { this.disposed = true; this.unsubscribe?.(); this.unsubscribe = undefined } }

  private submit(operation: SparklineCollaborationOperation, signal = new AbortController().signal): Promise<SparklineCollaborationEntry> {
    const result = this.tail.then(async () => {
      this.assertUsable()
      // Fail locally-invalid operations before asking an authoritative service
      // to append them. The service must still repeat this reducer check because
      // another accepted operation may race this preflight.
      applySparklineCollaborationOperation(this.manager.serialize(), operation)
      if (this.options.authorize && !await this.options.authorize(clone(operation))) throw new SparklineCollaborationError('permission', 'Sparkline collaboration operation is not authorized')
      let entry: SparklineCollaborationEntry
      try { entry = await this.transport.submit(this.options.room, clone(operation), this.sequence, signal) } catch (cause) {
        const stale = String(cause).includes('STALE_BASE')
        if (stale) this.block(this.sequence + 1, -1, cause)
        throw new SparklineCollaborationError(stale ? 'stale-base' : 'transport', 'Sparkline operation submission failed', { cause })
      }
      if (canonical(validateEntry(entry, this.options.room).operation) !== canonical(operation)) {
        this.block(this.sequence + 1, entry.sequence, new Error('Acknowledgement operation changed'))
        throw new SparklineCollaborationError('conflict', 'Authoritative acknowledgement changed the submitted operation')
      }
      const status = await this.applyEntry(entry, true)
      if (status !== 'applied') throw new SparklineCollaborationError('conflict', 'Authoritative sparkline acknowledgement could not be applied')
      this.emit({ type: 'submitted', entry: clone(entry) })
      return clone(entry)
    })
    this.tail = result.then(() => undefined, () => undefined)
    return result
  }

  private async applyEntry(entryInput: unknown, own: boolean): Promise<'applied' | 'duplicate' | 'blocked'> {
    if (this.disposed || this.blocked) return 'blocked'
    let entry: SparklineCollaborationEntry
    try { entry = validateEntry(entryInput, this.options.room) } catch (cause) { return this.block(this.sequence + 1, -1, cause) }
    if (entry.sequence <= this.sequence) { this.emit({ type: 'duplicate', sequence: entry.sequence }); return 'duplicate' }
    if (entry.sequence !== this.sequence + 1) return this.block(this.sequence + 1, entry.sequence)
    if (own && entry.operation.clientId !== this.options.clientId) return this.block(this.sequence + 1, entry.sequence, new Error('Acknowledgement client identity changed'))
    try {
      const next = applySparklineCollaborationOperation(this.manager.serialize(), entry.operation)
      this.manager.hydrate(next)
    } catch (cause) { return this.block(this.sequence + 1, entry.sequence, cause) }
    this.sequence = entry.sequence
    this.emit({ type: 'applied', entry: clone(entry) })
    return 'applied'
  }

  private block(expectedSequence: number, receivedSequence: number, cause?: unknown): 'blocked' {
    this.blocked = true
    this.emit({ type: 'blocked', expectedSequence, receivedSequence, cause })
    this.options.onResyncRequired?.(expectedSequence, receivedSequence, cause)
    return 'blocked'
  }

  private base<T extends Omit<SparklineCollaborationOperation, keyof OperationBase>>(value: T): T & OperationBase {
    const opId = this.options.idFactory()
    if (!validId(opId)) throw new TypeError('idFactory returned an invalid operation id')
    return { protocol: SPARKLINE_COLLABORATION_PROTOCOL, opId, clientId: this.options.clientId, ...value }
  }

  private require(id: string): SparklineSpec {
    const spec = this.manager.get(id)
    if (!spec) throw new SparklineCollaborationError('conflict', `Unknown sparkline ${id}`)
    return spec
  }

  private fingerprints(ids: readonly string[]): Record<string, string> {
    const result: Record<string, string> = {}
    for (const id of unique(ids)) result[id] = fingerprintSparkline(this.require(id))
    return result
  }

  private assertUsable(): void {
    if (this.disposed) throw new SparklineCollaborationError('blocked', 'Sparkline collaboration session is disposed')
    if (this.blocked) throw new SparklineCollaborationError('blocked', 'Sparkline collaboration session requires resynchronization')
  }

  private emit(event: SparklineCollaborationEvent): void { try { this.options.onEvent?.(event) } catch { /* observer isolation */ } }
}

function validateOperation(input: SparklineCollaborationOperation): SparklineCollaborationOperation {
  const operation = clone(input)
  if (!operation || typeof operation !== 'object' || operation.protocol !== SPARKLINE_COLLABORATION_PROTOCOL || !validId(operation.opId) || !validId(operation.clientId)) invalid('Invalid sparkline operation identity')
  if (!['create', 'update', 'remove', 'group', 'ungroup'].includes(operation.kind)) invalid('Unsupported sparkline operation')
  const baseKeys = ['clientId', 'kind', 'opId', 'protocol']
  if (operation.kind === 'create') {
    exactKeys(operation, [...baseKeys, 'value'])
    if (!operation.value || operation.value.groupId !== undefined) invalid('Created sparkline must be ungrouped')
  }
  if (operation.kind === 'update' || operation.kind === 'remove') {
    exactKeys(operation, [...baseKeys, 'expectedFingerprint', 'id', ...(operation.kind === 'update' ? ['value'] : [])])
    if (!validId(operation.id) || !validFingerprint(operation.expectedFingerprint)) invalid('Invalid object precondition')
  }
  if (operation.kind === 'update' && !operation.value) invalid('Update replacement is required')
  if (operation.kind === 'group' || operation.kind === 'ungroup') {
    exactKeys(operation, [...baseKeys, 'expectedFingerprints', 'memberIds', ...(operation.kind === 'group' ? ['groupId'] : [])])
    if (!Array.isArray(operation.memberIds) || operation.memberIds.length < (operation.kind === 'group' ? 2 : 1) || operation.memberIds.some((id) => !validId(id)) || unique(operation.memberIds).length !== operation.memberIds.length) invalid('Invalid group members')
    if (operation.kind === 'group' && !validId(operation.groupId)) invalid('Invalid group id')
    if (!plainRecord(operation.expectedFingerprints) || Object.entries(operation.expectedFingerprints).some(([id, fingerprint]) => !validId(id) || !validFingerprint(fingerprint)) || operation.memberIds.some((id) => !validFingerprint(operation.expectedFingerprints[id]))) invalid('Group fingerprints must cover selected members')
  }
  return operation
}

function validateEntry(input: unknown, room: string): SparklineCollaborationEntry {
  if (!input || typeof input !== 'object') invalid('Sparkline entry must be an object')
  const entry = input as SparklineCollaborationEntry
  exactKeys(entry, ['operation', 'room', 'sequence'])
  if (entry.room !== room || !Number.isSafeInteger(entry.sequence) || entry.sequence < 1) invalid('Invalid sparkline entry room or sequence')
  return { room, sequence: entry.sequence, operation: validateOperation(entry.operation) }
}

function requireFingerprint(manager: SparklineManager, id: string, expected: string): SparklineSpec {
  const spec = manager.get(id)
  if (!spec || fingerprintSparkline(spec) !== expected) conflict(`Sparkline ${id} changed concurrently`)
  return spec
}

function affectedIds(manager: SparklineManager, memberIds: readonly string[]): string[] {
  const ids = new Set(memberIds)
  const groupIds = new Set(memberIds.map((id) => manager.get(id)?.groupId).filter((id): id is string => !!id))
  for (const group of manager.listGroups()) if (groupIds.has(group.id)) for (const id of group.memberIds) ids.add(id)
  return [...ids]
}

function unique(ids: readonly string[]): string[] { return [...new Set(ids)] }
function validId(value: unknown): value is string { return typeof value === 'string' && /^[A-Za-z0-9][A-Za-z0-9._:/-]{0,255}$/.test(value) }
function validFingerprint(value: unknown): value is string { return typeof value === 'string' && /^fnv1a32:[0-9a-f]{8}$/.test(value) }
function plainRecord(value: unknown): value is Record<string, unknown> { return !!value && typeof value === 'object' && !Array.isArray(value) && Object.getPrototypeOf(value) === Object.prototype }
function exactKeys(value: object, expected: readonly string[]): void { if (canonical(Object.keys(value).sort()) !== canonical([...expected].sort())) invalid('Sparkline collaboration payload contains missing or unsupported fields') }
function clone<T>(value: T): T { try { return structuredClone(value) } catch (cause) { throw new SparklineCollaborationError('invalid', 'Sparkline collaboration payload must be structured-cloneable', { cause }) } }
function canonical(value: unknown): string { if (Array.isArray(value)) return `[${value.map(canonical).join(',')}]`; if (value && typeof value === 'object') return `{${Object.keys(value).sort().map((key) => `${JSON.stringify(key)}:${canonical((value as Record<string, unknown>)[key])}`).join(',')}}`; return JSON.stringify(value) }
function hash(value: string): string { let hash = 0x811c9dc5; for (let index = 0; index < value.length; index++) { hash ^= value.charCodeAt(index); hash = Math.imul(hash, 0x01000193) } return `fnv1a32:${(hash >>> 0).toString(16).padStart(8, '0')}` }
function invalid(message: string): never { throw new SparklineCollaborationError('invalid', message) }
function conflict(message: string): never { throw new SparklineCollaborationError('conflict', message) }
