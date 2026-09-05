import {
  isShapeSnapshot,
  type ShapeCommandController,
  type ShapeSnapshotEntry,
  type ShapeSnapshotV1,
} from './commands'
import type { FileShapeAnchor } from './fromFile'
import type { ShapeSpec } from './types'

export const SHAPE_COLLABORATION_PROTOCOL = 'injoffice.shape-ops.v1' as const

interface OperationBase { protocol: typeof SHAPE_COLLABORATION_PROTOCOL; opId: string; clientId: string }
export type ShapeCollaborationOperation =
  | OperationBase & { kind: 'create'; value: ShapeSnapshotEntry }
  | OperationBase & { kind: 'update'; id: string; expectedFingerprint: string; value: ShapeSnapshotEntry }
  | OperationBase & { kind: 'remove'; id: string; expectedFingerprint: string }

export interface ShapeCollaborationEntry {
  room: string
  sequence: number
  operation: ShapeCollaborationOperation
}

export interface ShapeCollaborationTransport {
  /** The authoritative service must authenticate the caller, validate the
   * exact base sequence, deduplicate opId per client, run the exported reducer,
   * durably append, and only then return/broadcast the accepted entry. */
  submit(room: string, operation: ShapeCollaborationOperation, baseSequence: number, signal: AbortSignal): Promise<ShapeCollaborationEntry>
  subscribe(room: string, handler: (entry: unknown) => void): () => void
}

export type ShapeCollaborationErrorCode = 'blocked' | 'conflict' | 'invalid' | 'permission' | 'stale-base' | 'transport'

export class ShapeCollaborationError extends Error {
  constructor(readonly code: ShapeCollaborationErrorCode, message: string, options?: ErrorOptions) {
    super(message, options)
    this.name = 'ShapeCollaborationError'
  }
}

export type ShapeCollaborationEvent =
  | { type: 'submitted' | 'applied'; entry: ShapeCollaborationEntry }
  | { type: 'duplicate'; sequence: number }
  | { type: 'blocked'; expectedSequence: number; receivedSequence: number; cause?: unknown }

export interface ShapeCollaborationOptions {
  room: string
  clientId: string
  initialSequence?: number
  idFactory: () => string
  authorize?: (operation: ShapeCollaborationOperation) => boolean | Promise<boolean>
  onEvent?: (event: ShapeCollaborationEvent) => void
  onResyncRequired?: (expectedSequence: number, receivedSequence: number, cause?: unknown) => void
}

export interface ShapeCollaborationPatch {
  spec?: Partial<Omit<ShapeSpec, 'id' | 'kind' | 'nativeIdentity'>>
  /** `null` clears the persisted cell anchor. Pixel-level float-DOM transform
   * capture remains a host integration concern. */
  cellAnchor?: FileShapeAnchor | null
  sheetId?: string
}

/** Pure reducer shared by authoritative services and clients. Operations use
 * stable shape IDs and exact before-state fingerprints instead of merging
 * whole snapshots with last-writer-wins behavior. */
export function applyShapeCollaborationOperation(
  snapshotInput: ShapeSnapshotV1,
  operationInput: ShapeCollaborationOperation,
): ShapeSnapshotV1 {
  const snapshot = clone(snapshotInput)
  if (!isShapeSnapshot(snapshot)) invalid('Invalid shape snapshot')
  const operation = validateOperation(operationInput)
  const index = operation.kind === 'create' ? -1 : snapshot.shapes.findIndex((entry) => entry.spec.id === operation.id)
  if (operation.kind === 'create') {
    if (snapshot.shapes.some((entry) => entry.spec.id === operation.value.spec.id)) conflict(`Shape ${operation.value.spec.id} already exists`)
    snapshot.shapes.push(clone(operation.value))
  } else {
    if (index < 0) conflict(`Shape ${operation.id} disappeared`)
    const current = snapshot.shapes[index]
    if (fingerprintShape(current) !== operation.expectedFingerprint) conflict(`Shape ${operation.id} changed concurrently`)
    if (operation.kind === 'remove') snapshot.shapes.splice(index, 1)
    else {
      if (operation.value.spec.id !== operation.id) conflict('Update cannot change stable shape id')
      if (operation.value.spec.kind !== current.spec.kind) conflict('Update cannot change shape kind')
      if (canonical(operation.value.spec.nativeIdentity) !== canonical(current.spec.nativeIdentity)) conflict('Update cannot change native shape identity')
      snapshot.shapes[index] = clone(operation.value)
    }
  }
  if (!isShapeSnapshot(snapshot)) invalid('Shape operation produced an invalid snapshot')
  return snapshot
}

export function fingerprintShape(value: ShapeSnapshotEntry): string { return hash(canonical(clone(value))) }

/** Server-first client: local mounted state changes only after an unchanged,
 * ordered authoritative acknowledgement. Any gap or conflict requires an
 * explicit validated snapshot resync before editing can continue. */
export class ShapeCollaborationSession {
  private sequence: number
  private blocked = false
  private disposed = false
  private unsubscribe?: () => void
  private tail: Promise<void> = Promise.resolve()

  constructor(
    readonly controller: ShapeCommandController,
    private readonly transport: ShapeCollaborationTransport,
    private readonly options: ShapeCollaborationOptions,
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

  create(value: ShapeSnapshotEntry, signal?: AbortSignal): Promise<ShapeCollaborationEntry> {
    return this.submit(this.base({ kind: 'create', value: clone(value) }), signal)
  }

  update(id: string, patch: ShapeCollaborationPatch, signal?: AbortSignal): Promise<ShapeCollaborationEntry> {
    const current = this.require(id)
    const untrusted = (patch.spec ?? {}) as Partial<ShapeSpec>
    const { id: _id, kind: _kind, nativeIdentity: _identity, ...safeSpec } = untrusted
    const value: ShapeSnapshotEntry = {
      spec: { ...current.spec, ...clone(safeSpec), id, kind: current.spec.kind },
      sheetId: patch.sheetId ?? current.sheetId,
    }
    const nextAnchor = 'cellAnchor' in patch ? patch.cellAnchor : current.cellAnchor
    if (nextAnchor) value.cellAnchor = clone(nextAnchor)
    return this.submit(this.base({ kind: 'update', id, expectedFingerprint: fingerprintShape(current), value }), signal)
  }

  remove(id: string, signal?: AbortSignal): Promise<ShapeCollaborationEntry> {
    const current = this.require(id)
    return this.submit(this.base({ kind: 'remove', id, expectedFingerprint: fingerprintShape(current) }), signal)
  }

  receive(entryInput: unknown): Promise<'applied' | 'duplicate' | 'blocked'> {
    const result = this.tail.then(() => this.applyEntry(entryInput, false))
    this.tail = result.then(() => undefined, () => undefined)
    return result
  }

  resync(snapshot: ShapeSnapshotV1, sequence: number): boolean {
    if (this.disposed || !Number.isSafeInteger(sequence) || sequence < this.sequence || !isShapeSnapshot(snapshot)) return false
    if (!this.controller.restore(clone(snapshot))) return false
    this.sequence = sequence
    this.blocked = false
    return true
  }

  dispose(): void { if (!this.disposed) { this.disposed = true; this.unsubscribe?.(); this.unsubscribe = undefined } }

  private submit(operation: ShapeCollaborationOperation, signal = new AbortController().signal): Promise<ShapeCollaborationEntry> {
    const result = this.tail.then(async () => {
      this.assertUsable()
      applyShapeCollaborationOperation(this.controller.snapshot(), operation)
      if (this.options.authorize && !await this.options.authorize(clone(operation))) throw new ShapeCollaborationError('permission', 'Shape collaboration operation is not authorized')
      let entry: ShapeCollaborationEntry
      try { entry = await this.transport.submit(this.options.room, clone(operation), this.sequence, signal) } catch (cause) {
        const stale = String(cause).includes('STALE_BASE')
        if (stale) this.block(this.sequence + 1, -1, cause)
        throw new ShapeCollaborationError(stale ? 'stale-base' : 'transport', 'Shape operation submission failed', { cause })
      }
      if (canonical(validateEntry(entry, this.options.room).operation) !== canonical(operation)) {
        this.block(this.sequence + 1, entry.sequence, new Error('Acknowledgement operation changed'))
        throw new ShapeCollaborationError('conflict', 'Authoritative acknowledgement changed the submitted operation')
      }
      const status = this.applyEntry(entry, true)
      if (status !== 'applied') throw new ShapeCollaborationError('conflict', 'Authoritative shape acknowledgement could not be applied')
      this.emit({ type: 'submitted', entry: clone(entry) })
      return clone(entry)
    })
    this.tail = result.then(() => undefined, () => undefined)
    return result
  }

  private applyEntry(entryInput: unknown, own: boolean): 'applied' | 'duplicate' | 'blocked' {
    if (this.disposed || this.blocked) return 'blocked'
    let entry: ShapeCollaborationEntry
    try { entry = validateEntry(entryInput, this.options.room) } catch (cause) { return this.block(this.sequence + 1, -1, cause) }
    if (entry.sequence <= this.sequence) { this.emit({ type: 'duplicate', sequence: entry.sequence }); return 'duplicate' }
    if (entry.sequence !== this.sequence + 1) return this.block(this.sequence + 1, entry.sequence)
    if (own && entry.operation.clientId !== this.options.clientId) return this.block(this.sequence + 1, entry.sequence, new Error('Acknowledgement client identity changed'))
    try {
      const next = applyShapeCollaborationOperation(this.controller.snapshot(), entry.operation)
      if (!this.controller.restore(next)) conflict('Shape snapshot could not be mounted atomically')
    } catch (cause) { return this.block(this.sequence + 1, entry.sequence, cause) }
    this.sequence = entry.sequence
    this.emit({ type: 'applied', entry: clone(entry) })
    return 'applied'
  }

  private base<T extends Omit<ShapeCollaborationOperation, keyof OperationBase>>(value: T): T & OperationBase {
    const opId = this.options.idFactory()
    if (!validId(opId)) throw new TypeError('idFactory returned an invalid operation id')
    return { protocol: SHAPE_COLLABORATION_PROTOCOL, opId, clientId: this.options.clientId, ...value }
  }

  private require(id: string): ShapeSnapshotEntry {
    const entry = this.controller.snapshot().shapes.find((value) => value.spec.id === id)
    if (!entry) throw new ShapeCollaborationError('conflict', `Unknown shape ${id}`)
    return entry
  }

  private assertUsable(): void {
    if (this.disposed) throw new ShapeCollaborationError('blocked', 'Shape collaboration session is disposed')
    if (this.blocked) throw new ShapeCollaborationError('blocked', 'Shape collaboration session requires resynchronization')
  }

  private block(expectedSequence: number, receivedSequence: number, cause?: unknown): 'blocked' {
    this.blocked = true
    this.emit({ type: 'blocked', expectedSequence, receivedSequence, cause })
    this.options.onResyncRequired?.(expectedSequence, receivedSequence, cause)
    return 'blocked'
  }

  private emit(event: ShapeCollaborationEvent): void { try { this.options.onEvent?.(event) } catch { /* observer isolation */ } }
}

function validateOperation(input: ShapeCollaborationOperation): ShapeCollaborationOperation {
  const operation = clone(input)
  if (!operation || typeof operation !== 'object' || operation.protocol !== SHAPE_COLLABORATION_PROTOCOL || !validId(operation.opId) || !validId(operation.clientId)) invalid('Invalid shape operation identity')
  if (!['create', 'update', 'remove'].includes(operation.kind)) invalid('Unsupported shape operation')
  const baseKeys = ['clientId', 'kind', 'opId', 'protocol']
  if (operation.kind === 'create') exactKeys(operation, [...baseKeys, 'value'])
  else {
    exactKeys(operation, [...baseKeys, 'expectedFingerprint', 'id', ...(operation.kind === 'update' ? ['value'] : [])])
    if (!validId(operation.id) || !validFingerprint(operation.expectedFingerprint)) invalid('Invalid shape object precondition')
  }
  if ((operation.kind === 'create' || operation.kind === 'update') && !isShapeSnapshot({ version: 1, shapes: [operation.value] })) invalid('Invalid shape operation value')
  return operation
}

function validateEntry(input: unknown, room: string): ShapeCollaborationEntry {
  if (!input || typeof input !== 'object') invalid('Shape entry must be an object')
  const entry = input as ShapeCollaborationEntry
  exactKeys(entry, ['operation', 'room', 'sequence'])
  if (entry.room !== room || !Number.isSafeInteger(entry.sequence) || entry.sequence < 1) invalid('Invalid shape entry room or sequence')
  return { room, sequence: entry.sequence, operation: validateOperation(entry.operation) }
}

function validId(value: unknown): value is string { return typeof value === 'string' && /^[A-Za-z0-9][A-Za-z0-9._:/-]{0,255}$/.test(value) }
function validFingerprint(value: unknown): value is string { return typeof value === 'string' && /^fnv1a32:[0-9a-f]{8}$/.test(value) }
function exactKeys(value: object, expected: readonly string[]): void { if (canonical(Object.keys(value).sort()) !== canonical([...expected].sort())) invalid('Shape collaboration payload contains missing or unsupported fields') }
function clone<T>(value: T): T { try { return JSON.parse(JSON.stringify(value)) as T } catch (cause) { throw new ShapeCollaborationError('invalid', 'Shape collaboration payload must be plain JSON', { cause }) } }
function canonical(value: unknown): string { if (Array.isArray(value)) return `[${value.map(canonical).join(',')}]`; if (value && typeof value === 'object') return `{${Object.keys(value).sort().map((key) => `${JSON.stringify(key)}:${canonical((value as Record<string, unknown>)[key])}`).join(',')}}`; return JSON.stringify(value) }
function hash(value: string): string { let result = 0x811c9dc5; for (let index = 0; index < value.length; index++) { result ^= value.charCodeAt(index); result = Math.imul(result, 0x01000193) } return `fnv1a32:${(result >>> 0).toString(16).padStart(8, '0')}` }
function invalid(message: string): never { throw new ShapeCollaborationError('invalid', message) }
function conflict(message: string): never { throw new ShapeCollaborationError('conflict', message) }
