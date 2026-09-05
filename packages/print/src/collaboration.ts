import type { PrintConfigurationCommandTarget } from './commands'
import { PrintConfigurationCommandController } from './commands'
import { sha256Hex } from './collaborationHash'
import type { PrintConfigurationSnapshotV1, PrintLayoutConfig, PrintRenderConfig } from './types'
import { validatePrintConfigurationSnapshot } from './validation'

export const PRINT_COLLABORATION_PROTOCOL = 'injoffice.print-configuration.v1' as const

export type PrintCollaborationAction = 'update-layout' | 'update-render' | 'replace' | 'restore'

export interface PrintCollaborationOperation {
  protocol: typeof PRINT_COLLABORATION_PROTOCOL
  opId: string
  clientId: string
  kind: 'replace'
  action: PrintCollaborationAction
  expectedFingerprint: string
  value: PrintConfigurationSnapshotV1
}

export interface PrintCollaborationEntry {
  room: string
  sequence: number
  operation: PrintCollaborationOperation
}

export interface PrintCollaborationTransport {
  /** The service must authenticate and authorize the actor, require the exact
   * base sequence, deduplicate opId per client, run the exported reducer,
   * durably append, and only then return and broadcast the accepted entry. */
  submit(room: string, operation: PrintCollaborationOperation, baseSequence: number, signal: AbortSignal): Promise<PrintCollaborationEntry>
  subscribe(room: string, handler: (entry: unknown) => void): () => void
}

export type PrintCollaborationErrorCode = 'blocked' | 'conflict' | 'invalid' | 'permission' | 'stale-base' | 'transport'

export class PrintCollaborationError extends Error {
  constructor(readonly code: PrintCollaborationErrorCode, message: string, options?: ErrorOptions) {
    super(message, options)
    this.name = 'PrintCollaborationError'
  }
}

export type PrintCollaborationEvent =
  | { type: 'submitted' | 'applied'; entry: PrintCollaborationEntry }
  | { type: 'duplicate'; sequence: number }
  | { type: 'blocked'; expectedSequence: number; receivedSequence: number; cause?: unknown }

export interface PrintCollaborationApplyContext {
  kind: 'entry' | 'resync'
  remote: boolean
  recordUndo: boolean
  entry?: PrintCollaborationEntry
}

export interface PrintCollaborationAuthorizationContext {
  direction: 'submit' | 'receive'
  room: string
}

export interface PrintCollaborationOptions {
  room: string
  clientId: string
  initialSequence?: number
  idFactory: () => string
  /** A client-side policy boundary. The authoritative transport must still
   * authenticate and authorize every submitted operation. */
  authorize?: (operation: PrintCollaborationOperation, context: PrintCollaborationAuthorizationContext) => boolean | Promise<boolean>
  /** Override when the host must combine durable print persistence with other
   * state. It must apply or roll back the complete snapshot atomically. */
  applySnapshot?: (snapshot: PrintConfigurationSnapshotV1, context: PrintCollaborationApplyContext) => boolean | Promise<boolean>
  onEvent?: (event: PrintCollaborationEvent) => void
  onResyncRequired?: (expectedSequence: number, receivedSequence: number, cause?: unknown) => void
}

/** Pure reducer shared by clients and authoritative services. Print settings
 * form one conflict domain: every accepted replacement names the exact prior
 * complete-snapshot fingerprint. */
export function applyPrintCollaborationOperation(
  snapshotInput: PrintConfigurationSnapshotV1,
  operationInput: PrintCollaborationOperation,
): PrintConfigurationSnapshotV1 {
  const snapshot = validateSnapshot(snapshotInput)
  const operation = validateOperation(operationInput)
  if (fingerprintPrintConfiguration(snapshot) !== operation.expectedFingerprint) conflict('Print configuration changed concurrently')
  return clone(operation.value)
}

export function fingerprintPrintConfiguration(snapshot: PrintConfigurationSnapshotV1): string { return `sha256:${sha256Hex(canonical(validateSnapshot(snapshot)))}` }

/** Server-first ordered session and command target. Local state changes only
 * after an unchanged acknowledgement. Remote entries and resync never create
 * local undo, while acknowledged local edits do. */
export class PrintCollaborationSession implements PrintConfigurationCommandTarget {
  private sequence: number
  private blocked = false
  private disposed = false
  private unsubscribe?: () => void
  private tail: Promise<void> = Promise.resolve()

  constructor(
    readonly controller: PrintConfigurationCommandController,
    private readonly transport: PrintCollaborationTransport,
    private readonly options: PrintCollaborationOptions,
  ) {
    if (!validId(options.room) || !validId(options.clientId) || typeof options.idFactory !== 'function') throw new TypeError('room, clientId, and idFactory are required')
    this.sequence = options.initialSequence ?? 0
    if (!Number.isSafeInteger(this.sequence) || this.sequence < 0) throw new TypeError('initialSequence must be a non-negative safe integer')
  }

  get manager() { return this.controller.manager }
  get state(): { sequence: number; blocked: boolean; disposed: boolean } { return { sequence: this.sequence, blocked: this.blocked, disposed: this.disposed } }
  snapshot(): PrintConfigurationSnapshotV1 { return this.controller.snapshot() }

  start(): void {
    this.assertUsable()
    if (this.unsubscribe) return
    const unsubscribe = this.transport.subscribe(this.options.room, (entry) => { void this.receive(entry) })
    if (typeof unsubscribe !== 'function') throw new TypeError('transport.subscribe must return an unsubscribe function')
    this.unsubscribe = unsubscribe
  }

  updateLayout(patch: Partial<PrintLayoutConfig>, signal?: AbortSignal): Promise<boolean> {
    if (!patch || typeof patch !== 'object' || Array.isArray(patch)) return Promise.resolve(false)
    try {
      const copy = clone(patch)
      return this.submit('update-layout', (before) => ({ ...before, layout: { ...before.layout, ...copy } }), true, signal)
    } catch { return Promise.resolve(false) }
  }

  updateRender(patch: Partial<PrintRenderConfig>, signal?: AbortSignal): Promise<boolean> {
    if (!patch || typeof patch !== 'object' || Array.isArray(patch)) return Promise.resolve(false)
    try {
      const copy = clone(patch)
      return this.submit('update-render', (before) => ({
        ...before,
        render: { ...before.render, ...copy, headerFooterSetting: { ...before.render.headerFooterSetting, ...copy.headerFooterSetting } },
      }), true, signal)
    } catch { return Promise.resolve(false) }
  }

  replace(snapshot: unknown, signal?: AbortSignal): Promise<boolean> {
    try {
      const checked = validateSnapshot(snapshot)
      return this.submit('replace', () => checked, true, signal)
    } catch { return Promise.resolve(false) }
  }

  /** Undo/redo is another authoritative replacement but does not recursively
   * add an undo record after the server accepts it. */
  restore(snapshot: unknown, signal?: AbortSignal): Promise<boolean> {
    try {
      const checked = validateSnapshot(snapshot)
      return this.submit('restore', () => checked, false, signal)
    } catch { return Promise.resolve(false) }
  }

  receive(entryInput: unknown): Promise<'applied' | 'duplicate' | 'blocked'> {
    const result = this.tail.then(() => this.applyEntry(entryInput, false))
    this.tail = result.then(() => undefined, () => undefined)
    return result
  }

  resync(snapshot: PrintConfigurationSnapshotV1, sequence: number): Promise<boolean> {
    let checked: PrintConfigurationSnapshotV1
    try { checked = validateSnapshot(snapshot) } catch { return Promise.resolve(false) }
    const result = this.tail.then(async () => {
      if (this.disposed || !Number.isSafeInteger(sequence) || sequence < this.sequence) return false
      try {
        if (!await this.apply(checked, { kind: 'resync', remote: true, recordUndo: false })) {
          this.block(this.sequence + 1, sequence, new Error('Resynchronization snapshot was not applied'))
          return false
        }
        if (fingerprintPrintConfiguration(this.snapshot()) !== fingerprintPrintConfiguration(checked)) {
          this.block(this.sequence + 1, sequence, new Error('Resynchronization did not reach the authoritative state'))
          return false
        }
      } catch (cause) {
        this.block(this.sequence + 1, sequence, cause)
        return false
      }
      this.sequence = sequence
      this.blocked = false
      return true
    })
    this.tail = result.then(() => undefined, () => undefined)
    return result.catch((cause) => {
      if (cause instanceof PrintCollaborationError && cause.code === 'invalid') return false
      throw cause
    })
  }

  dispose(): void {
    if (this.disposed) return
    this.disposed = true
    this.unsubscribe?.()
    this.unsubscribe = undefined
  }

  private operation(action: PrintCollaborationAction, value: PrintConfigurationSnapshotV1): PrintCollaborationOperation {
    this.assertUsable()
    const opId = this.options.idFactory()
    if (!validId(opId)) throw new TypeError('idFactory returned an invalid operation id')
    return validateOperation({
      protocol: PRINT_COLLABORATION_PROTOCOL,
      opId,
      clientId: this.options.clientId,
      kind: 'replace',
      action,
      expectedFingerprint: fingerprintPrintConfiguration(this.snapshot()),
      value,
    })
  }

  private submit(action: PrintCollaborationAction, candidate: (before: PrintConfigurationSnapshotV1) => PrintConfigurationSnapshotV1, recordUndo: boolean, signal = new AbortController().signal): Promise<boolean> {
    const result = this.tail.then(async () => {
      this.assertUsable()
      const operation = this.operation(action, candidate(this.snapshot()))
      applyPrintCollaborationOperation(this.snapshot(), operation)
      if (!await this.authorized(operation, 'submit')) throw new PrintCollaborationError('permission', 'Print configuration operation is not authorized')
      let entry: PrintCollaborationEntry
      try { entry = await this.transport.submit(this.options.room, clone(operation), this.sequence, signal) } catch (cause) {
        const stale = isStaleBase(cause)
        if (stale) this.block(this.sequence + 1, -1, cause)
        throw new PrintCollaborationError(stale ? 'stale-base' : 'transport', 'Print configuration submission failed', { cause })
      }
      let validated: PrintCollaborationEntry
      try { validated = validateEntry(entry, this.options.room) } catch (cause) {
        this.block(this.sequence + 1, -1, cause)
        throw new PrintCollaborationError('conflict', 'Authoritative acknowledgement is invalid', { cause })
      }
      if (canonical(validated.operation) !== canonical(operation)) {
        this.block(this.sequence + 1, validated.sequence, new Error('Acknowledgement operation changed'))
        throw new PrintCollaborationError('conflict', 'Authoritative acknowledgement changed the submitted operation')
      }
      const status = await this.applyEntry(validated, true, recordUndo)
      if (status !== 'applied') throw new PrintCollaborationError('conflict', 'Authoritative print acknowledgement could not be applied')
      this.emit({ type: 'submitted', entry: clone(validated) })
      return true
    })
    this.tail = result.then(() => undefined, () => undefined)
    return result.catch((cause) => {
      if (cause instanceof PrintCollaborationError && cause.code === 'invalid') return false
      throw cause
    })
  }

  private async applyEntry(entryInput: unknown, own: boolean, recordUndo = false): Promise<'applied' | 'duplicate' | 'blocked'> {
    if (this.disposed || this.blocked) return 'blocked'
    let entry: PrintCollaborationEntry
    try { entry = validateEntry(entryInput, this.options.room) } catch (cause) { return this.block(this.sequence + 1, -1, cause) }
    if (entry.sequence <= this.sequence) { this.emit({ type: 'duplicate', sequence: entry.sequence }); return 'duplicate' }
    if (entry.sequence !== this.sequence + 1) return this.block(this.sequence + 1, entry.sequence)
    if (own && entry.operation.clientId !== this.options.clientId) return this.block(this.sequence + 1, entry.sequence, new Error('Acknowledgement client identity changed'))
    if (!own) {
      try { if (!await this.authorized(entry.operation, 'receive')) return this.block(this.sequence + 1, entry.sequence, new PrintCollaborationError('permission', 'Authoritative print operation violates local policy')) }
      catch (cause) { return this.block(this.sequence + 1, entry.sequence, cause) }
    }
    try {
      const next = applyPrintCollaborationOperation(this.snapshot(), entry.operation)
      const context = { kind: 'entry' as const, remote: !own, recordUndo: own && recordUndo, entry: clone(entry) }
      if (!await this.apply(next, context)) conflict('Print configuration could not be persisted and applied atomically')
      if (fingerprintPrintConfiguration(this.snapshot()) !== fingerprintPrintConfiguration(next)) conflict('Print configuration apply did not reach the acknowledged state')
    } catch (cause) { return this.block(this.sequence + 1, entry.sequence, cause) }
    this.sequence = entry.sequence
    this.emit({ type: 'applied', entry: clone(entry) })
    return 'applied'
  }

  private apply(snapshot: PrintConfigurationSnapshotV1, context: PrintCollaborationApplyContext): boolean | Promise<boolean> {
    if (this.options.applySnapshot) return this.options.applySnapshot(clone(snapshot), context)
    return this.controller.applyAuthoritative(snapshot, {
      label: labelFor(context.entry?.operation.action),
      recordUndo: context.recordUndo,
    })
  }

  private async authorized(operation: PrintCollaborationOperation, direction: 'submit' | 'receive'): Promise<boolean> {
    if (!this.options.authorize) return true
    try { return await this.options.authorize(clone(operation), { direction, room: this.options.room }) === true }
    catch (cause) { throw new PrintCollaborationError('permission', 'Print collaboration authorization failed', { cause }) }
  }

  private assertUsable(): void {
    if (this.disposed) throw new PrintCollaborationError('blocked', 'Print collaboration session is disposed')
    if (this.blocked) throw new PrintCollaborationError('blocked', 'Print collaboration session requires resynchronization')
  }

  private block(expectedSequence: number, receivedSequence: number, cause?: unknown): 'blocked' {
    this.blocked = true
    this.emit({ type: 'blocked', expectedSequence, receivedSequence, cause })
    try { this.options.onResyncRequired?.(expectedSequence, receivedSequence, cause) } catch { /* observer isolation */ }
    return 'blocked'
  }

  private emit(event: PrintCollaborationEvent): void { try { this.options.onEvent?.(event) } catch { /* observer isolation */ } }
}

function labelFor(action: PrintCollaborationAction | undefined): string {
  if (action === 'update-layout') return 'Update print layout'
  if (action === 'update-render') return 'Update print rendering'
  if (action === 'replace') return 'Replace print configuration'
  return 'Restore print configuration'
}

function validateSnapshot(input: unknown): PrintConfigurationSnapshotV1 {
  if (!input || typeof input !== 'object' || Array.isArray(input)) invalid('Invalid print configuration snapshot')
  noUnknown(input, ['layout', 'render', 'version'])
  const value = input as PrintConfigurationSnapshotV1
  if (value.layout && typeof value.layout === 'object') {
    noUnknown(value.layout, ['area', 'customMargins', 'customScale', 'direction', 'fitToHeightPages', 'fitToWidthPages', 'freeze', 'margin', 'maxColumnsEachPage', 'maxRowsEachPage', 'pageSizeCustom', 'paperSize', 'repeatColumns', 'repeatRows', 'scale', 'subUnitIds'])
    if (value.layout.customMargins) noUnknown(value.layout.customMargins, ['bottom', 'footer', 'header', 'left', 'right', 'top'])
    if (value.layout.pageSizeCustom) noUnknown(value.layout.pageSizeCustom, ['height', 'width'])
    if (value.layout.repeatRows) noUnknown(value.layout.repeatRows, ['endRow', 'startRow'])
    if (value.layout.repeatColumns) noUnknown(value.layout.repeatColumns, ['endColumn', 'startColumn'])
    if (Array.isArray(value.layout.subUnitIds)) for (const target of value.layout.subUnitIds) if (target && typeof target === 'object') {
      noUnknown(target, ['id', 'range'])
      if (target.range) noUnknown(target.range, ['endColumn', 'endRow', 'startColumn', 'startRow'])
    }
  }
  if (value.render && typeof value.render === 'object') {
    noUnknown(value.render, ['gridlines', 'hAlign', 'headerFooter', 'headerFooterSetting', 'headings', 'isCustomHeaderFooter', 'vAlign', 'watermark'])
    if (value.render.headerFooterSetting) noUnknown(value.render.headerFooterSetting, ['bottomCenter', 'bottomLeft', 'bottomRight', 'topCenter', 'topLeft', 'topRight'])
    if (value.render.watermark) noUnknown(value.render.watermark, value.render.watermark.kind === 'text' ? ['color', 'fontSize', 'kind', 'opacity', 'rotation', 'text'] : ['kind', 'opacity', 'scale', 'sourceId'])
  }
  const checked = validatePrintConfigurationSnapshot(value)
  if (!checked.ok) invalid('Invalid print configuration snapshot')
  if (new TextEncoder().encode(canonical(checked.value)).byteLength > 1_000_000) invalid('Print configuration snapshot exceeds the 1000000-byte collaboration limit')
  return clone(checked.value)
}

function validateOperation(input: unknown): PrintCollaborationOperation {
  if (!input || typeof input !== 'object' || Array.isArray(input)) invalid('Print operation must be an object')
  noUnknown(input, ['action', 'clientId', 'expectedFingerprint', 'kind', 'opId', 'protocol', 'value'])
  const operation = input as PrintCollaborationOperation
  if (operation.protocol !== PRINT_COLLABORATION_PROTOCOL || operation.kind !== 'replace' || !validId(operation.opId) || !validId(operation.clientId)) invalid('Invalid print operation identity')
  if (!['update-layout', 'update-render', 'replace', 'restore'].includes(operation.action) || !validFingerprint(operation.expectedFingerprint)) invalid('Invalid print operation precondition or action')
  return { ...clone(operation), value: validateSnapshot(operation.value) }
}

function validateEntry(input: unknown, room: string): PrintCollaborationEntry {
  if (!input || typeof input !== 'object' || Array.isArray(input)) invalid('Print entry must be an object')
  noUnknown(input, ['operation', 'room', 'sequence'])
  const entry = input as PrintCollaborationEntry
  if (entry.room !== room || !Number.isSafeInteger(entry.sequence) || entry.sequence < 1) invalid('Invalid print entry room or sequence')
  return { room, sequence: entry.sequence, operation: validateOperation(entry.operation) }
}

function validId(value: unknown): value is string { return typeof value === 'string' && /^[A-Za-z0-9][A-Za-z0-9._:/-]{0,255}$/.test(value) }
function validFingerprint(value: unknown): value is string { return typeof value === 'string' && /^sha256:[0-9a-f]{64}$/.test(value) }
function noUnknown(value: object, allowed: readonly string[]): void { if (Object.keys(value).some((key) => !allowed.includes(key))) invalid('Print collaboration payload contains unsupported fields') }
function clone<T>(value: T): T { try { return JSON.parse(JSON.stringify(value)) as T } catch (cause) { throw new PrintCollaborationError('invalid', 'Print collaboration payload must be plain JSON', { cause }) } }
function canonical(value: unknown): string { if (Array.isArray(value)) return `[${value.map(canonical).join(',')}]`; if (value && typeof value === 'object') return `{${Object.keys(value).sort().map((key) => `${JSON.stringify(key)}:${canonical((value as Record<string, unknown>)[key])}`).join(',')}}`; return JSON.stringify(value) }
function isStaleBase(cause: unknown): boolean { return !!cause && typeof cause === 'object' && 'code' in cause && (cause as { code?: unknown }).code === 'STALE_BASE' }
function invalid(message: string): never { throw new PrintCollaborationError('invalid', message) }
function conflict(message: string): never { throw new PrintCollaborationError('conflict', message) }
