import { OUTLINE_COLLABORATION_APPLY, type OutlineCommandController } from './commands'
import { sha256Hex } from './sha256'
import type { OutlineAxis, OutlineGroup } from './types'
import { validateOutlineSnapshot } from './validation'

export const OUTLINE_COLLABORATION_PROTOCOL = 'injoffice.outline-ops.v1' as const

export type OutlineCollaborationAction = 'create' | 'update' | 'set-collapsed' | 'remove' | 'clear' | 'restore'

export interface OutlineCollaborationOperation {
  protocol: typeof OUTLINE_COLLABORATION_PROTOCOL
  opId: string
  clientId: string
  action: OutlineCollaborationAction
  expectedFingerprint: string
  value: OutlineGroup[]
}

export interface OutlineCollaborationEntry {
  room: string
  sequence: number
  operation: OutlineCollaborationOperation
}

export interface OutlineCollaborationResync {
  protocol: typeof OUTLINE_COLLABORATION_PROTOCOL
  room: string
  sequence: number
  value: OutlineGroup[]
}

export interface OutlineCollaborationTransport {
  /** The service must authenticate and authorize the actor, require the exact
   * base sequence, deduplicate opId per client, run the exported reducer,
   * durably append, and only then acknowledge and broadcast the entry. */
  submit(room: string, operation: OutlineCollaborationOperation, baseSequence: number, signal: AbortSignal): Promise<OutlineCollaborationEntry>
  subscribe(room: string, handler: (entry: unknown) => void): () => void
}

export interface OutlineCollaborationCommandTarget {
  add(group: OutlineGroup): Promise<boolean> | boolean
  update(id: string, patch: Partial<Omit<OutlineGroup, 'id'>>): Promise<boolean> | boolean
  setCollapsed(id: string, collapsed: boolean): Promise<boolean> | boolean
  remove(id: string): Promise<boolean> | boolean
  clear(sheetId: string, axis?: OutlineAxis, start?: number, end?: number): Promise<boolean> | boolean
  restore(snapshot: readonly OutlineGroup[]): Promise<boolean> | boolean
  dispose?(): void
}

export type OutlineCollaborationErrorCode = 'blocked' | 'conflict' | 'invalid' | 'permission' | 'stale-base' | 'transport'

export class OutlineCollaborationError extends Error {
  constructor(readonly code: OutlineCollaborationErrorCode, message: string, options?: ErrorOptions) {
    super(message, options)
    this.name = 'OutlineCollaborationError'
  }
}

export type OutlineCollaborationEvent =
  | { type: 'submitted' | 'applied'; entry: OutlineCollaborationEntry }
  | { type: 'duplicate'; sequence: number }
  | { type: 'blocked'; expectedSequence: number; receivedSequence: number; cause?: unknown }

export interface OutlineCollaborationApplyContext {
  kind: 'entry' | 'resync'
  remote: boolean
  recordUndo: boolean
  entry?: OutlineCollaborationEntry
}

export interface OutlineCollaborationAuthorizationContext {
  direction: 'submit' | 'receive'
  room: string
}

export interface OutlineCollaborationOptions {
  room: string
  clientId: string
  initialSequence?: number
  idFactory: () => string
  /** Immediate client policy only. The authoritative service must enforce its
   * own authenticated authorization decision. */
  authorize?: (operation: OutlineCollaborationOperation, context: OutlineCollaborationAuthorizationContext) => boolean | Promise<boolean>
  /** Override when outline model and visibility must join a larger host
   * transaction. It must apply completely or roll back before returning false. */
  applySnapshot?: (snapshot: OutlineGroup[], context: OutlineCollaborationApplyContext) => boolean | Promise<boolean>
  onEvent?: (event: OutlineCollaborationEvent) => void
  onResyncRequired?: (expectedSequence: number, receivedSequence: number, cause?: unknown) => void
}

/** Pure reducer shared by clients and the authoritative service. The complete
 * outline collection is one conflict domain, and the declared action must
 * exactly describe the stable-ID snapshot difference. */
export function applyOutlineCollaborationOperation(
  snapshotInput: readonly OutlineGroup[],
  operationInput: OutlineCollaborationOperation,
): OutlineGroup[] {
  const snapshot = validateSnapshot(snapshotInput)
  const operation = validateOperation(operationInput)
  if (fingerprintOutlineSnapshot(snapshot) !== operation.expectedFingerprint) conflict('Outline collection changed concurrently')
  assertAction(snapshot, operation.value, operation.action)
  return clone(operation.value)
}

export function fingerprintOutlineSnapshot(snapshot: readonly OutlineGroup[]): string {
  return `sha256:${sha256Hex(canonical(validateSnapshot(snapshot)))}`
}

/** Server-first ordered command target. Local changes and undo records appear
 * only after an unchanged consecutive acknowledgement. */
export class OutlineCollaborationSession implements OutlineCollaborationCommandTarget {
  private sequence: number
  private blocked = false
  private disposed = false
  private unsubscribe?: () => void
  private tail: Promise<void> = Promise.resolve()

  constructor(
    readonly controller: OutlineCommandController,
    private readonly transport: OutlineCollaborationTransport,
    private readonly options: OutlineCollaborationOptions,
  ) {
    if (!validId(options.room) || !validId(options.clientId) || typeof options.idFactory !== 'function') throw new TypeError('room, clientId, and idFactory are required')
    this.sequence = options.initialSequence ?? 0
    if (!Number.isSafeInteger(this.sequence) || this.sequence < 0) throw new TypeError('initialSequence must be a non-negative safe integer')
  }

  get manager() { return this.controller.manager }
  get state(): { sequence: number; blocked: boolean; disposed: boolean } { return { sequence: this.sequence, blocked: this.blocked, disposed: this.disposed } }
  snapshot(): OutlineGroup[] { return this.controller.snapshot() }

  start(): void {
    this.assertUsable()
    if (this.unsubscribe) return
    const unsubscribe = this.transport.subscribe(this.options.room, (entry) => { void this.receive(entry) })
    if (typeof unsubscribe !== 'function') throw new TypeError('transport.subscribe must return an unsubscribe function')
    this.unsubscribe = unsubscribe
  }

  add(group: OutlineGroup, signal?: AbortSignal): Promise<boolean> {
    let value: OutlineGroup
    try { value = validateSnapshot([group])[0]! } catch { return Promise.resolve(false) }
    return this.submit('create', (before) => [...before, value], true, signal)
  }

  update(id: string, patch: Partial<Omit<OutlineGroup, 'id'>>, signal?: AbortSignal): Promise<boolean> {
    if (!validGroupId(id) || !plainObject(patch)) return Promise.resolve(false)
    let copied: Partial<Omit<OutlineGroup, 'id'>>
    try { copied = clone(patch) } catch { return Promise.resolve(false) }
    return this.submit('update', (before) => {
      const index = before.findIndex((group) => group.id === id)
      if (index < 0) conflict(`Outline ${id} does not exist`)
      const next = clone(before)
      next[index] = { ...next[index]!, ...copied, id }
      return next
    }, true, signal)
  }

  setCollapsed(id: string, collapsed: boolean, signal?: AbortSignal): Promise<boolean> {
    if (!validGroupId(id) || typeof collapsed !== 'boolean') return Promise.resolve(false)
    return this.submit('set-collapsed', (before) => {
      const index = before.findIndex((group) => group.id === id)
      if (index < 0) conflict(`Outline ${id} does not exist`)
      const next = clone(before)
      next[index] = { ...next[index]!, collapsed }
      return next
    }, true, signal)
  }

  remove(id: string, signal?: AbortSignal): Promise<boolean> {
    if (!validGroupId(id)) return Promise.resolve(false)
    return this.submit('remove', (before) => {
      if (!before.some((group) => group.id === id)) conflict(`Outline ${id} does not exist`)
      return before.filter((group) => group.id !== id)
    }, true, signal)
  }

  clear(sheetId: string, axis?: OutlineAxis, start?: number, end?: number, signal?: AbortSignal): Promise<boolean> {
    const maximum = axis === 'column' ? 16_383 : 1_048_575
    if (typeof sheetId !== 'string' || !sheetId || sheetId.length > 128 || (axis !== undefined && axis !== 'row' && axis !== 'column')
      || (start !== undefined && (!Number.isSafeInteger(start) || start < 0 || start > maximum))
      || (end !== undefined && (!Number.isSafeInteger(end) || end < 0 || end > maximum))
      || (start !== undefined && end !== undefined && end < start)) return Promise.resolve(false)
    return this.submit('clear', (before) => before.filter((group) => !(
      group.sheetId === sheetId
      && (axis === undefined || group.axis === axis)
      && (start === undefined || group.start >= start)
      && (end === undefined || group.end <= end)
    )), true, signal)
  }

  /** Undo/redo is an authoritative replacement and does not recursively add
   * another local undo record after acknowledgement. */
  restore(snapshot: readonly OutlineGroup[], signal?: AbortSignal): Promise<boolean> {
    let checked: OutlineGroup[]
    try { checked = validateSnapshot(snapshot) } catch { return Promise.resolve(false) }
    return this.submit('restore', () => checked, false, signal)
  }

  receive(entryInput: unknown): Promise<'applied' | 'duplicate' | 'blocked'> {
    const result = this.tail.then(() => this.applyEntry(entryInput, false))
    this.tail = result.then(() => undefined, () => undefined)
    return result
  }

  resync(envelopeInput: unknown): Promise<boolean> {
    let envelope: OutlineCollaborationResync
    try { envelope = validateResync(envelopeInput, this.options.room) } catch { return Promise.resolve(false) }
    const result = this.tail.then(async () => {
      if (this.disposed || envelope.sequence < this.sequence) return false
      try {
        if (!await this.apply(envelope.value, { kind: 'resync', remote: true, recordUndo: false })) {
          this.block(this.sequence + 1, envelope.sequence, new Error('Outline resync was not applied'))
          return false
        }
        if (fingerprintOutlineSnapshot(this.snapshot()) !== fingerprintOutlineSnapshot(envelope.value)) {
          this.block(this.sequence + 1, envelope.sequence, new Error('Outline resync did not reach the authoritative state'))
          return false
        }
      } catch (cause) {
        this.block(this.sequence + 1, envelope.sequence, cause)
        return false
      }
      this.sequence = envelope.sequence
      this.blocked = false
      return true
    })
    this.tail = result.then(() => undefined, () => undefined)
    return result
  }

  dispose(): void {
    if (this.disposed) return
    this.disposed = true
    this.unsubscribe?.()
    this.unsubscribe = undefined
  }

  private submit(
    action: OutlineCollaborationAction,
    candidate: (before: OutlineGroup[]) => OutlineGroup[],
    recordUndo: boolean,
    signal = new AbortController().signal,
  ): Promise<boolean> {
    const result = this.tail.then(async () => {
      this.assertUsable()
      const before = this.snapshot()
      const operation = this.operation(action, candidate(before), before)
      applyOutlineCollaborationOperation(before, operation)
      if (!await this.authorized(operation, 'submit')) throw new OutlineCollaborationError('permission', 'Outline operation is not authorized')
      let entry: OutlineCollaborationEntry
      try { entry = await this.transport.submit(this.options.room, clone(operation), this.sequence, signal) } catch (cause) {
        const stale = isStaleBase(cause)
        if (stale) this.block(this.sequence + 1, -1, cause)
        throw new OutlineCollaborationError(stale ? 'stale-base' : 'transport', 'Outline operation submission failed', { cause })
      }
      let validated: OutlineCollaborationEntry
      try { validated = validateEntry(entry, this.options.room) } catch (cause) {
        this.block(this.sequence + 1, -1, cause)
        throw new OutlineCollaborationError('conflict', 'Authoritative acknowledgement is invalid', { cause })
      }
      if (canonical(validated.operation) !== canonical(operation)) {
        this.block(this.sequence + 1, validated.sequence, new Error('Acknowledgement operation changed'))
        throw new OutlineCollaborationError('conflict', 'Authoritative acknowledgement changed the submitted operation')
      }
      const status = await this.applyEntry(validated, true, recordUndo)
      if (status !== 'applied') throw new OutlineCollaborationError('conflict', 'Authoritative outline acknowledgement could not be applied')
      this.emit({ type: 'submitted', entry: clone(validated) })
      return true
    })
    this.tail = result.then(() => undefined, () => undefined)
    return result.catch((cause) => {
      if (cause instanceof OutlineCollaborationError && cause.code === 'invalid') return false
      throw cause
    })
  }

  private async applyEntry(entryInput: unknown, own: boolean, recordUndo = false): Promise<'applied' | 'duplicate' | 'blocked'> {
    if (this.disposed || this.blocked) return 'blocked'
    let entry: OutlineCollaborationEntry
    try { entry = validateEntry(entryInput, this.options.room) } catch (cause) { return this.block(this.sequence + 1, -1, cause) }
    if (entry.sequence <= this.sequence) { this.emit({ type: 'duplicate', sequence: entry.sequence }); return 'duplicate' }
    if (entry.sequence !== this.sequence + 1) return this.block(this.sequence + 1, entry.sequence)
    if (own && entry.operation.clientId !== this.options.clientId) return this.block(this.sequence + 1, entry.sequence, new Error('Acknowledgement client identity changed'))
    if (!own) {
      try { if (!await this.authorized(entry.operation, 'receive')) return this.block(this.sequence + 1, entry.sequence, new OutlineCollaborationError('permission', 'Received outline operation violates local policy')) }
      catch (cause) { return this.block(this.sequence + 1, entry.sequence, cause) }
    }
    try {
      const next = applyOutlineCollaborationOperation(this.snapshot(), entry.operation)
      const context = { kind: 'entry' as const, remote: !own, recordUndo: own && recordUndo, entry: clone(entry) }
      if (!await this.apply(next, context)) conflict('Outline model and visibility could not be applied atomically')
      if (fingerprintOutlineSnapshot(this.snapshot()) !== fingerprintOutlineSnapshot(next)) conflict('Outline apply did not reach the acknowledged state')
    } catch (cause) { return this.block(this.sequence + 1, entry.sequence, cause) }
    this.sequence = entry.sequence
    this.emit({ type: 'applied', entry: clone(entry) })
    return 'applied'
  }

  private operation(action: OutlineCollaborationAction, value: OutlineGroup[], before: OutlineGroup[]): OutlineCollaborationOperation {
    this.assertUsable()
    const opId = this.options.idFactory()
    if (!validId(opId)) throw new TypeError('idFactory returned an invalid operation id')
    return validateOperation({
      protocol: OUTLINE_COLLABORATION_PROTOCOL,
      opId,
      clientId: this.options.clientId,
      action,
      expectedFingerprint: fingerprintOutlineSnapshot(before),
      value,
    })
  }

  private apply(snapshot: OutlineGroup[], context: OutlineCollaborationApplyContext): boolean | Promise<boolean> {
    if (this.options.applySnapshot) return this.options.applySnapshot(clone(snapshot), context)
    return this.controller[OUTLINE_COLLABORATION_APPLY](snapshot, { label: labelFor(context.entry?.operation.action), recordUndo: context.recordUndo })
  }

  private async authorized(operation: OutlineCollaborationOperation, direction: 'submit' | 'receive'): Promise<boolean> {
    if (!this.options.authorize) return true
    try { return await this.options.authorize(clone(operation), { direction, room: this.options.room }) === true }
    catch (cause) { throw new OutlineCollaborationError('permission', 'Outline collaboration authorization failed', { cause }) }
  }

  private assertUsable(): void {
    if (this.disposed) throw new OutlineCollaborationError('blocked', 'Outline collaboration session is disposed')
    if (this.blocked) throw new OutlineCollaborationError('blocked', 'Outline collaboration session requires resynchronization')
  }

  private block(expectedSequence: number, receivedSequence: number, cause?: unknown): 'blocked' {
    this.blocked = true
    this.emit({ type: 'blocked', expectedSequence, receivedSequence, cause })
    try { this.options.onResyncRequired?.(expectedSequence, receivedSequence, cause) } catch { /* observer isolation */ }
    return 'blocked'
  }

  private emit(event: OutlineCollaborationEvent): void { try { this.options.onEvent?.(event) } catch { /* observer isolation */ } }
}

function labelFor(action: OutlineCollaborationAction | undefined): string {
  if (action === 'create') return 'Add outline'
  if (action === 'update') return 'Update outline'
  if (action === 'set-collapsed') return 'Set outline collapsed state'
  if (action === 'remove') return 'Remove outline'
  if (action === 'clear') return 'Ungroup outline'
  return 'Restore outline snapshot'
}

function assertAction(before: OutlineGroup[], after: OutlineGroup[], action: OutlineCollaborationAction): void {
  if (action === 'restore') return
  const left = new Map(before.map((group) => [group.id, group]))
  const right = new Map(after.map((group) => [group.id, group]))
  const added = [...right.keys()].filter((id) => !left.has(id))
  const removed = [...left.keys()].filter((id) => !right.has(id))
  const changed = [...left.keys()].filter((id) => right.has(id) && canonical(left.get(id)) !== canonical(right.get(id)))
  if (action === 'create' && added.length === 1 && removed.length === 0 && changed.length === 0) return
  if (action === 'remove' && removed.length === 1 && added.length === 0 && changed.length === 0) return
  if (action === 'clear' && removed.length > 0 && added.length === 0 && changed.length === 0) return
  if (action === 'update' && added.length === 0 && removed.length === 0 && changed.length === 1) return
  if (action === 'set-collapsed' && added.length === 0 && removed.length === 0 && changed.length === 1) {
    const previous = left.get(changed[0]!)!
    const next = right.get(changed[0]!)!
    if (previous.collapsed !== next.collapsed && canonical({ ...previous, collapsed: false }) === canonical({ ...next, collapsed: false })) return
  }
  invalid(`Outline ${action} payload does not match its declared action`)
}

function validateSnapshot(input: unknown): OutlineGroup[] {
  if (!Array.isArray(input) || input.length > 10_000) invalid('Outline snapshot must contain at most 10000 groups')
  const groups = clone(input) as OutlineGroup[]
  for (const group of groups) {
    if (!plainObject(group)) invalid('Outline group must be a plain object')
    exactKeys(group, ['axis', 'collapsed', 'end', 'id', 'sheetId', 'start'])
    if (!validGroupId(group.id) || typeof group.sheetId !== 'string' || typeof group.collapsed !== 'boolean'
      || (group.axis !== 'row' && group.axis !== 'column') || !Number.isSafeInteger(group.start) || !Number.isSafeInteger(group.end)) invalid('Invalid outline group fields')
  }
  const checked = validateOutlineSnapshot(groups)
  if (!checked.ok) invalid('Invalid outline snapshot')
  const ordered = checked.value.map((group) => ({ ...group })).sort(compareGroups)
  if (new TextEncoder().encode(canonical(ordered)).byteLength > 2_000_000) invalid('Outline snapshot exceeds the 2000000-byte collaboration limit')
  return ordered
}

function validateOperation(input: unknown): OutlineCollaborationOperation {
  if (!plainObject(input)) invalid('Outline operation must be a plain object')
  exactKeys(input, ['action', 'clientId', 'expectedFingerprint', 'opId', 'protocol', 'value'])
  const operation = input as unknown as OutlineCollaborationOperation
  if (operation.protocol !== OUTLINE_COLLABORATION_PROTOCOL || !validId(operation.opId) || !validId(operation.clientId)
    || !['create', 'update', 'set-collapsed', 'remove', 'clear', 'restore'].includes(operation.action)
    || !validFingerprint(operation.expectedFingerprint)) invalid('Invalid outline operation identity or precondition')
  return { ...clone(operation), value: validateSnapshot(operation.value) }
}

function validateEntry(input: unknown, room: string): OutlineCollaborationEntry {
  if (!plainObject(input)) invalid('Outline entry must be a plain object')
  exactKeys(input, ['operation', 'room', 'sequence'])
  const entry = input as unknown as OutlineCollaborationEntry
  if (entry.room !== room || !Number.isSafeInteger(entry.sequence) || entry.sequence < 1) invalid('Invalid outline entry room or sequence')
  return { room, sequence: entry.sequence, operation: validateOperation(entry.operation) }
}

function validateResync(input: unknown, room: string): OutlineCollaborationResync {
  if (!plainObject(input)) invalid('Outline resync must be a plain object')
  exactKeys(input, ['protocol', 'room', 'sequence', 'value'])
  const envelope = input as unknown as OutlineCollaborationResync
  if (envelope.protocol !== OUTLINE_COLLABORATION_PROTOCOL || envelope.room !== room
    || !Number.isSafeInteger(envelope.sequence) || envelope.sequence < 0) invalid('Invalid outline resync identity')
  return { protocol: OUTLINE_COLLABORATION_PROTOCOL, room, sequence: envelope.sequence, value: validateSnapshot(envelope.value) }
}

function compareGroups(left: OutlineGroup, right: OutlineGroup): number {
  return left.sheetId.localeCompare(right.sheetId) || left.axis.localeCompare(right.axis) || left.start - right.start || right.end - left.end || left.id.localeCompare(right.id)
}
function validGroupId(value: unknown): value is string { return typeof value === 'string' && /^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/.test(value) }
function validId(value: unknown): value is string { return typeof value === 'string' && /^[A-Za-z0-9][A-Za-z0-9._:/-]{0,255}$/.test(value) }
function validFingerprint(value: unknown): value is string { return typeof value === 'string' && /^sha256:[0-9a-f]{64}$/.test(value) }
function plainObject(value: unknown): value is Record<string, unknown> { return !!value && typeof value === 'object' && !Array.isArray(value) && Object.getPrototypeOf(value) === Object.prototype }
function exactKeys(value: object, expected: readonly string[]): void { if (canonical(Object.keys(value).sort()) !== canonical([...expected].sort())) invalid('Outline collaboration payload contains missing or unsupported fields') }
function clone<T>(value: T): T { try { return JSON.parse(JSON.stringify(value)) as T } catch (cause) { throw new OutlineCollaborationError('invalid', 'Outline collaboration payload must be plain JSON', { cause }) } }
function canonical(value: unknown): string { if (Array.isArray(value)) return `[${value.map(canonical).join(',')}]`; if (value && typeof value === 'object') return `{${Object.keys(value).sort().map((key) => `${JSON.stringify(key)}:${canonical((value as Record<string, unknown>)[key])}`).join(',')}}`; return JSON.stringify(value) }
function isStaleBase(cause: unknown): boolean { return !!cause && typeof cause === 'object' && 'code' in cause && (cause as { code?: unknown }).code === 'STALE_BASE' }
function invalid(message: string): never { throw new OutlineCollaborationError('invalid', message) }
function conflict(message: string): never { throw new OutlineCollaborationError('conflict', message) }
