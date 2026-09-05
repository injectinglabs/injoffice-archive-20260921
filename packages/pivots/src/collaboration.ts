import { isPivotSnapshot, type PivotCommandController, type PivotSnapshotV1 } from './commands'
import type { PivotSpec } from './types'

export const PIVOT_COLLABORATION_PROTOCOL = 'injoffice.pivot-ops.v1' as const

interface OperationBase { protocol: typeof PIVOT_COLLABORATION_PROTOCOL; opId: string; clientId: string }
export type PivotCollaborationOperation =
  | OperationBase & { kind: 'create'; value: PivotSpec }
  | OperationBase & { kind: 'update'; id: string; expectedFingerprint: string; value: PivotSpec }
  | OperationBase & { kind: 'remove'; id: string; expectedFingerprint: string }

export interface PivotCollaborationEntry {
  room: string
  sequence: number
  operation: PivotCollaborationOperation
}

export interface PivotCollaborationTransport {
  /** The service must authenticate/authorize, require the exact base sequence,
   * deduplicate opId per client, run the exported reducer, durably append, and
   * only then return and broadcast the accepted entry. */
  submit(room: string, operation: PivotCollaborationOperation, baseSequence: number, signal: AbortSignal): Promise<PivotCollaborationEntry>
  subscribe(room: string, handler: (entry: unknown) => void): () => void
}

export type PivotCollaborationErrorCode = 'blocked' | 'conflict' | 'invalid' | 'permission' | 'stale-base' | 'transport'

export class PivotCollaborationError extends Error {
  constructor(readonly code: PivotCollaborationErrorCode, message: string, options?: ErrorOptions) {
    super(message, options)
    this.name = 'PivotCollaborationError'
  }
}

export type PivotCollaborationEvent =
  | { type: 'submitted' | 'applied'; entry: PivotCollaborationEntry }
  | { type: 'duplicate'; sequence: number }
  | { type: 'blocked'; expectedSequence: number; receivedSequence: number; cause?: unknown }

export interface PivotCollaborationApplyContext {
  kind: 'entry' | 'resync'
  remote: boolean
  entry?: PivotCollaborationEntry
}

export interface PivotCollaborationOptions {
  room: string
  clientId: string
  initialSequence?: number
  idFactory: () => string
  authorize?: (operation: PivotCollaborationOperation) => boolean | Promise<boolean>
  /** Override snapshot application when the host must transactionally mark
   * derived pivot-grid writes as local-only/from-collaboration for its generic
   * cell mutation transport. Return false after rolling back any partial work. */
  applySnapshot?: (snapshot: PivotSnapshotV1, context: PivotCollaborationApplyContext) => boolean | Promise<boolean>
  onEvent?: (event: PivotCollaborationEvent) => void
  onResyncRequired?: (expectedSequence: number, receivedSequence: number, cause?: unknown) => void
}

/** Pure object reducer shared by authoritative services and clients. A full
 * before-state fingerprint makes same-pivot races explicit while allowing
 * independently ordered changes to different pivots. */
export function applyPivotCollaborationOperation(
  snapshotInput: PivotSnapshotV1,
  operationInput: PivotCollaborationOperation,
): PivotSnapshotV1 {
  const snapshot = clone(snapshotInput)
  if (!isPivotSnapshot(snapshot)) invalid('Invalid pivot snapshot')
  const operation = validateOperation(operationInput)
  const index = operation.kind === 'create' ? -1 : snapshot.pivots.findIndex(({ id }) => id === operation.id)
  if (operation.kind === 'create') {
    if (snapshot.pivots.some(({ id }) => id === operation.value.id)) conflict(`Pivot ${operation.value.id} already exists`)
    snapshot.pivots.push(clone(operation.value))
  } else {
    if (index < 0) conflict(`Pivot ${operation.id} disappeared`)
    const current = snapshot.pivots[index]
    if (fingerprintPivot(current) !== operation.expectedFingerprint) conflict(`Pivot ${operation.id} changed concurrently`)
    if (operation.kind === 'remove') snapshot.pivots.splice(index, 1)
    else {
      if (operation.value.id !== operation.id) conflict('Update cannot change stable pivot id')
      if (canonical(operation.value.nativeIdentity) !== canonical(current.nativeIdentity)) conflict('Update cannot change native pivot identity')
      snapshot.pivots[index] = clone(operation.value)
    }
  }
  if (!isPivotSnapshot(snapshot)) invalid('Pivot operation produced an invalid snapshot')
  return snapshot
}

export function fingerprintPivot(value: PivotSpec): string { return hash(canonical(clone(value))) }

/** Server-first session. Managed specs and their deterministic baked grids
 * change only after an unchanged authoritative acknowledgement. */
export class PivotCollaborationSession {
  private sequence: number
  private blocked = false
  private disposed = false
  private unsubscribe?: () => void
  private tail: Promise<void> = Promise.resolve()

  constructor(
    readonly controller: PivotCommandController,
    private readonly transport: PivotCollaborationTransport,
    private readonly options: PivotCollaborationOptions,
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

  create(value: PivotSpec, signal?: AbortSignal): Promise<PivotCollaborationEntry> {
    return this.submit(this.base({ kind: 'create', value: clone(value) }), signal)
  }

  update(id: string, patch: Partial<Omit<PivotSpec, 'id' | 'nativeIdentity'>>, signal?: AbortSignal): Promise<PivotCollaborationEntry> {
    const current = this.require(id)
    if (!patch || typeof patch !== 'object' || Array.isArray(patch)) return Promise.reject(new PivotCollaborationError('invalid', 'Pivot update patch must be an object'))
    const untrusted = patch as Partial<PivotSpec>
    const { id: _id, nativeIdentity: _identity, ...safePatch } = untrusted
    const value = { ...current, ...clone(safePatch), id, ...(current.nativeIdentity ? { nativeIdentity: clone(current.nativeIdentity) } : {}) }
    return this.submit(this.base({ kind: 'update', id, expectedFingerprint: fingerprintPivot(current), value }), signal)
  }

  remove(id: string, signal?: AbortSignal): Promise<PivotCollaborationEntry> {
    const current = this.require(id)
    return this.submit(this.base({ kind: 'remove', id, expectedFingerprint: fingerprintPivot(current) }), signal)
  }

  receive(entryInput: unknown): Promise<'applied' | 'duplicate' | 'blocked'> {
    const result = this.tail.then(() => this.applyEntry(entryInput, false))
    this.tail = result.then(() => undefined, () => undefined)
    return result
  }

  async resync(snapshot: PivotSnapshotV1, sequence: number): Promise<boolean> {
    if (this.disposed || !Number.isSafeInteger(sequence) || sequence < this.sequence || !isPivotSnapshot(snapshot)) return false
    try {
      if (!await this.restore(clone(snapshot), { kind: 'resync', remote: true })) return false
    } catch { return false }
    this.sequence = sequence
    this.blocked = false
    return true
  }

  dispose(): void { if (!this.disposed) { this.disposed = true; this.unsubscribe?.(); this.unsubscribe = undefined } }

  private submit(operation: PivotCollaborationOperation, signal = new AbortController().signal): Promise<PivotCollaborationEntry> {
    const result = this.tail.then(async () => {
      this.assertUsable()
      applyPivotCollaborationOperation(this.controller.snapshot(), operation)
      if (this.options.authorize && !await this.options.authorize(clone(operation))) throw new PivotCollaborationError('permission', 'Pivot collaboration operation is not authorized')
      let entry: PivotCollaborationEntry
      try { entry = await this.transport.submit(this.options.room, clone(operation), this.sequence, signal) } catch (cause) {
        const stale = String(cause).includes('STALE_BASE')
        if (stale) this.block(this.sequence + 1, -1, cause)
        throw new PivotCollaborationError(stale ? 'stale-base' : 'transport', 'Pivot operation submission failed', { cause })
      }
      if (canonical(validateEntry(entry, this.options.room).operation) !== canonical(operation)) {
        this.block(this.sequence + 1, entry.sequence, new Error('Acknowledgement operation changed'))
        throw new PivotCollaborationError('conflict', 'Authoritative acknowledgement changed the submitted operation')
      }
      const status = await this.applyEntry(entry, true)
      if (status !== 'applied') throw new PivotCollaborationError('conflict', 'Authoritative pivot acknowledgement could not be applied')
      this.emit({ type: 'submitted', entry: clone(entry) })
      return clone(entry)
    })
    this.tail = result.then(() => undefined, () => undefined)
    return result
  }

  private async applyEntry(entryInput: unknown, own: boolean): Promise<'applied' | 'duplicate' | 'blocked'> {
    if (this.disposed || this.blocked) return 'blocked'
    let entry: PivotCollaborationEntry
    try { entry = validateEntry(entryInput, this.options.room) } catch (cause) { return this.block(this.sequence + 1, -1, cause) }
    if (entry.sequence <= this.sequence) { this.emit({ type: 'duplicate', sequence: entry.sequence }); return 'duplicate' }
    if (entry.sequence !== this.sequence + 1) return this.block(this.sequence + 1, entry.sequence)
    if (own && entry.operation.clientId !== this.options.clientId) return this.block(this.sequence + 1, entry.sequence, new Error('Acknowledgement client identity changed'))
    try {
      const next = applyPivotCollaborationOperation(this.controller.snapshot(), entry.operation)
      if (!await this.restore(next, { kind: 'entry', remote: !own, entry: clone(entry) })) conflict('Pivot snapshot and baked grid could not be applied atomically')
    } catch (cause) { return this.block(this.sequence + 1, entry.sequence, cause) }
    this.sequence = entry.sequence
    this.emit({ type: 'applied', entry: clone(entry) })
    return 'applied'
  }

  private restore(snapshot: PivotSnapshotV1, context: PivotCollaborationApplyContext): boolean | Promise<boolean> {
    return this.options.applySnapshot ? this.options.applySnapshot(snapshot, context) : this.controller.restore(snapshot)
  }

  private base<T extends Omit<PivotCollaborationOperation, keyof OperationBase>>(value: T): T & OperationBase {
    const opId = this.options.idFactory()
    if (!validId(opId)) throw new TypeError('idFactory returned an invalid operation id')
    return { protocol: PIVOT_COLLABORATION_PROTOCOL, opId, clientId: this.options.clientId, ...value }
  }

  private require(id: string): PivotSpec {
    const value = this.controller.snapshot().pivots.find((pivot) => pivot.id === id)
    if (!value) throw new PivotCollaborationError('conflict', `Unknown pivot ${id}`)
    return value
  }

  private assertUsable(): void {
    if (this.disposed) throw new PivotCollaborationError('blocked', 'Pivot collaboration session is disposed')
    if (this.blocked) throw new PivotCollaborationError('blocked', 'Pivot collaboration session requires resynchronization')
  }

  private block(expectedSequence: number, receivedSequence: number, cause?: unknown): 'blocked' {
    this.blocked = true
    this.emit({ type: 'blocked', expectedSequence, receivedSequence, cause })
    this.options.onResyncRequired?.(expectedSequence, receivedSequence, cause)
    return 'blocked'
  }

  private emit(event: PivotCollaborationEvent): void { try { this.options.onEvent?.(event) } catch { /* observer isolation */ } }
}

function validateOperation(input: PivotCollaborationOperation): PivotCollaborationOperation {
  const operation = clone(input)
  if (!operation || typeof operation !== 'object' || operation.protocol !== PIVOT_COLLABORATION_PROTOCOL || !validId(operation.opId) || !validId(operation.clientId)) invalid('Invalid pivot operation identity')
  if (!['create', 'update', 'remove'].includes(operation.kind)) invalid('Unsupported pivot operation')
  const baseKeys = ['clientId', 'kind', 'opId', 'protocol']
  if (operation.kind === 'create') exactKeys(operation, [...baseKeys, 'value'])
  else {
    exactKeys(operation, [...baseKeys, 'expectedFingerprint', 'id', ...(operation.kind === 'update' ? ['value'] : [])])
    if (!validId(operation.id) || !validFingerprint(operation.expectedFingerprint)) invalid('Invalid pivot object precondition')
  }
  if ((operation.kind === 'create' || operation.kind === 'update') && !isPivotSnapshot({ version: 1, pivots: [operation.value] })) invalid('Invalid pivot operation value')
  return operation
}

function validateEntry(input: unknown, room: string): PivotCollaborationEntry {
  if (!input || typeof input !== 'object') invalid('Pivot entry must be an object')
  const entry = input as PivotCollaborationEntry
  exactKeys(entry, ['operation', 'room', 'sequence'])
  if (entry.room !== room || !Number.isSafeInteger(entry.sequence) || entry.sequence < 1) invalid('Invalid pivot entry room or sequence')
  return { room, sequence: entry.sequence, operation: validateOperation(entry.operation) }
}

function validId(value: unknown): value is string { return typeof value === 'string' && /^[A-Za-z0-9][A-Za-z0-9._:/-]{0,255}$/.test(value) }
function validFingerprint(value: unknown): value is string { return typeof value === 'string' && /^fnv1a32:[0-9a-f]{8}$/.test(value) }
function exactKeys(value: object, expected: readonly string[]): void { if (canonical(Object.keys(value).sort()) !== canonical([...expected].sort())) invalid('Pivot collaboration payload contains missing or unsupported fields') }
function clone<T>(value: T): T { try { return JSON.parse(JSON.stringify(value)) as T } catch (cause) { throw new PivotCollaborationError('invalid', 'Pivot collaboration payload must be plain JSON', { cause }) } }
function canonical(value: unknown): string { if (Array.isArray(value)) return `[${value.map(canonical).join(',')}]`; if (value && typeof value === 'object') return `{${Object.keys(value).sort().map((key) => `${JSON.stringify(key)}:${canonical((value as Record<string, unknown>)[key])}`).join(',')}}`; return JSON.stringify(value) }
function hash(value: string): string { let result = 0x811c9dc5; for (let index = 0; index < value.length; index++) { result ^= value.charCodeAt(index); result = Math.imul(result, 0x01000193) } return `fnv1a32:${(result >>> 0).toString(16).padStart(8, '0')}` }
function invalid(message: string): never { throw new PivotCollaborationError('invalid', message) }
function conflict(message: string): never { throw new PivotCollaborationError('conflict', message) }
