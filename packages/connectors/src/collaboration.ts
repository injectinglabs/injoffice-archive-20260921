import {
  type ConnectorCommandController,
  type ConnectorCommandSnapshotV1,
  isCredentialFreeConnectorSpec,
} from './commands'
import {
  isConnectorManagerSnapshot,
  isConnectorRangeSnapshot,
  type ConnectorManagerSnapshotV1,
  type ConnectorMountedSnapshot,
  type ConnectorRangeSnapshot,
} from './manager'
import type { RangeValue } from './preprocess'
import { sha256Hex } from './sha256'
import type { ConnectorSpec } from './types'

export const CONNECTOR_COLLABORATION_PROTOCOL = 'injoffice.connector-ops.v1' as const

export interface ConnectorCollaborationRange {
  sheetId: string
  startRow: number
  startColumn: number
  rowCount: number
  columnCount: number
  /** Authoritative connector output only. Formulas and rich ICellData are not
   * accepted across this boundary. Null clears cells outside a shrunk result. */
  values: RangeValue[][]
}

export interface ConnectorCollaborationCellPatch {
  expectedFingerprint: string
  after: ConnectorCollaborationRange
}

interface OperationBase { protocol: typeof CONNECTOR_COLLABORATION_PROTOCOL; opId: string; clientId: string }
export type ConnectorLifecycleCollaborationOperation =
  | OperationBase & { kind: 'create'; value: ConnectorSpec }
  | OperationBase & { kind: 'update'; id: string; expectedFingerprint: string; value: ConnectorSpec; cellPatch?: ConnectorCollaborationCellPatch }
  | OperationBase & { kind: 'remove'; id: string; expectedFingerprint: string; cellPatch?: ConnectorCollaborationCellPatch }

export type ConnectorRefreshCollaborationOperation = OperationBase & {
  kind: 'refresh'
  id: string
  expectedFingerprint: string
  sourceRevision: string
  preprocessFingerprint?: string
  value: ConnectorMountedSnapshot
  cellPatch?: ConnectorCollaborationCellPatch
}

export type ConnectorCollaborationOperation = ConnectorLifecycleCollaborationOperation | ConnectorRefreshCollaborationOperation

export interface ConnectorCollaborationEntry {
  room: string
  sequence: number
  operation: ConnectorCollaborationOperation
}

export interface ConnectorCollaborativeRefreshRequest {
  protocol: typeof CONNECTOR_COLLABORATION_PROTOCOL
  requestId: string
  clientId: string
  connectorId: string
  expectedFingerprint: string
  expectedPreprocessFingerprint?: string
}

export interface ConnectorCollaborationTransport {
  /** Lifecycle submission is data-free. The server validates model and cell
   * preconditions before append. */
  submit(room: string, operation: ConnectorLifecycleCollaborationOperation, baseSequence: number, signal: AbortSignal): Promise<ConnectorCollaborationEntry>
  /** The authenticated server, never the browser client, fetches and
   * preprocesses external data, enforces refresh ownership/single-flight, and
   * returns a committed refresh operation containing the bounded result. */
  requestRefresh(room: string, request: ConnectorCollaborativeRefreshRequest, baseSequence: number, signal: AbortSignal): Promise<ConnectorCollaborationEntry>
  subscribe(room: string, handler: (entry: unknown) => void): () => void
}

export type ConnectorCollaborationTransportErrorCode = 'STALE_BASE' | 'CONFLICT' | 'DENIED' | 'UNAVAILABLE'

export class ConnectorCollaborationTransportError extends Error {
  constructor(readonly code: ConnectorCollaborationTransportErrorCode, message: string = code) {
    super(message)
    this.name = 'ConnectorCollaborationTransportError'
  }
}

export type ConnectorCollaborationErrorCode = 'blocked' | 'conflict' | 'invalid' | 'permission' | 'stale-base' | 'transport'

export class ConnectorCollaborationError extends Error {
  constructor(readonly code: ConnectorCollaborationErrorCode, message: string, options?: ErrorOptions) {
    super(message, options)
    this.name = 'ConnectorCollaborationError'
  }
}

export type ConnectorCollaborationEvent =
  | { type: 'submitted' | 'applied'; entry: ConnectorCollaborationEntry }
  | { type: 'duplicate'; sequence: number }
  | { type: 'blocked'; expectedSequence: number; receivedSequence: number; cause?: unknown }

export interface ConnectorCollaborationApplyContext {
  kind: 'entry' | 'resync'
  remote: boolean
  entry?: ConnectorCollaborationEntry
}

export interface ConnectorCollaborationOptions {
  room: string
  clientId: string
  initialSequence?: number
  idFactory: () => string
  authorize?: (operation: ConnectorLifecycleCollaborationOperation) => boolean | Promise<boolean>
  authorizeRefresh?: (request: ConnectorCollaborativeRefreshRequest) => boolean | Promise<boolean>
  /** Use this to wrap model/range restoration in the host transaction that
   * marks range writes from-collaboration so a generic cell transport does not
   * echo them. It must return false only after rolling back partial writes. */
  applySnapshot: (snapshot: ConnectorCommandSnapshotV1, context: ConnectorCollaborationApplyContext) => boolean | Promise<boolean>
  onEvent?: (event: ConnectorCollaborationEvent) => void
  onResyncRequired?: (expectedSequence: number, receivedSequence: number, cause?: unknown) => void
}

export interface ConnectorCollaborationResync {
  protocol: typeof CONNECTOR_COLLABORATION_PROTOCOL
  room: string
  sequence: number
  revision: string
  manager: ConnectorManagerSnapshotV1
  ranges: ConnectorCollaborationRange[]
}

/** Pure connector-state reducer for authoritative and client use. Cell patches
 * are validated structurally here and against workbook cells with
 * verifyConnectorCollaborationCellPatch before commit/apply. */
export function applyConnectorCollaborationOperation(
  snapshotInput: ConnectorManagerSnapshotV1,
  operationInput: ConnectorCollaborationOperation,
): ConnectorManagerSnapshotV1 {
  assertJsonSize(snapshotInput)
  assertJsonSize(operationInput)
  const snapshot = clone(snapshotInput)
  validateManagerSnapshot(snapshot)
  const operation = validateOperation(operationInput)
  const index = operation.kind === 'create' ? -1 : snapshot.connectors.findIndex(({ spec }) => spec.id === operation.id)
  if (operation.kind === 'create') {
    if (snapshot.connectors.some(({ spec }) => spec.id === operation.value.id)) conflict(`Connector ${operation.value.id} already exists`)
    assertAvailableTarget(snapshot, operation.value.id, operation.value.target, 0, 0)
    snapshot.connectors.push({ spec: clone(operation.value), status: {}, lastRows: 0, lastColumns: 0 })
  } else {
    if (index < 0) conflict(`Connector ${operation.id} disappeared`)
    const current = snapshot.connectors[index]
    if (fingerprintConnector(current) !== operation.expectedFingerprint) conflict(`Connector ${operation.id} changed concurrently`)
    if (operation.kind === 'remove') {
      validateClearPatch(current, operation.cellPatch)
      snapshot.connectors.splice(index, 1)
    } else if (operation.kind === 'update') {
      if (operation.value.id !== operation.id) conflict('Update cannot change stable connector id')
      const moved = !sameTarget(current.spec.target, operation.value.target)
      if (moved) validateClearPatch(current, operation.cellPatch)
      else if (operation.cellPatch) invalid('An unchanged connector target must not include a cell patch')
      assertAvailableTarget(snapshot, operation.id, operation.value.target, moved ? 0 : current.lastRows, moved ? 0 : current.lastColumns)
      snapshot.connectors[index] = moved
        ? { spec: clone(operation.value), status: {}, lastRows: 0, lastColumns: 0 }
        : { ...current, spec: clone(operation.value) }
    } else {
      if (canonical(operation.value.spec) !== canonical(current.spec)) conflict('Refresh cannot change connector definition')
      validateRefreshPatch(current, operation.value, operation.cellPatch)
      assertAvailableTarget(snapshot, operation.id, current.spec.target, operation.value.lastRows, operation.value.lastColumns)
      snapshot.connectors[index] = clone(operation.value)
    }
  }
  validateManagerSnapshot(snapshot)
  return snapshot
}

export function fingerprintConnector(value: ConnectorMountedSnapshot): string { return hash(canonical(clone(value))) }
export function fingerprintConnectorRange(value: ConnectorRangeSnapshot): string { return hash(canonical(clone(value))) }

/** Verify a patch against the current workbook before an authoritative server
 * appends it or a client applies it. */
export function verifyConnectorCollaborationCellPatch(
  before: ConnectorRangeSnapshot | null,
  patch: ConnectorCollaborationCellPatch | undefined,
): boolean {
  if (!patch) return before === null
  return !!before && sameRange(before, patch.after) && fingerprintConnectorRange(before) === patch.expectedFingerprint
}

export class ConnectorCollaborationSession {
  private sequence: number
  private blocked = false
  private disposed = false
  private unsubscribe?: () => void
  private tail: Promise<void> = Promise.resolve()

  constructor(
    readonly controller: ConnectorCommandController,
    private readonly transport: ConnectorCollaborationTransport,
    private readonly options: ConnectorCollaborationOptions,
  ) {
    if (!validId(options.room) || !validId(options.clientId) || typeof options.idFactory !== 'function') throw new TypeError('room, clientId, and idFactory are required')
    if (controller.manager.executionAuthority !== 'server' || typeof options.applySnapshot !== 'function') throw new TypeError('Connector collaboration requires server execution authority and an atomic applySnapshot adapter')
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

  create(value: ConnectorSpec, signal?: AbortSignal): Promise<ConnectorCollaborationEntry> {
    return this.submit(this.base({ kind: 'create', value: clone(value) }), signal)
  }

  update(id: string, patch: Partial<Omit<ConnectorSpec, 'id'>>, signal?: AbortSignal): Promise<ConnectorCollaborationEntry> {
    const current = this.require(id)
    if (!patch || typeof patch !== 'object' || Array.isArray(patch)) return Promise.reject(new ConnectorCollaborationError('invalid', 'Connector update patch must be an object'))
    const untrusted = patch as Partial<ConnectorSpec>
    const { id: _id, ...safePatch } = untrusted
    const value = { ...current.spec, ...clone(safePatch), id }
    const moved = !sameTarget(current.spec.target, value.target)
    const operation = this.base({
      kind: 'update' as const,
      id,
      expectedFingerprint: fingerprintConnector(current),
      value,
      ...(moved ? this.clearPatchField(current) : {}),
    })
    return this.submit(operation, signal)
  }

  remove(id: string, signal?: AbortSignal): Promise<ConnectorCollaborationEntry> {
    const current = this.require(id)
    return this.submit(this.base({
      kind: 'remove' as const,
      id,
      expectedFingerprint: fingerprintConnector(current),
      ...this.clearPatchField(current),
    }), signal)
  }

  refresh(id: string, expectedPreprocessFingerprint?: string, signal = new AbortController().signal): Promise<ConnectorCollaborationEntry> {
    const current = this.require(id)
    const request: ConnectorCollaborativeRefreshRequest = {
      protocol: CONNECTOR_COLLABORATION_PROTOCOL,
      requestId: this.nextId(),
      clientId: this.options.clientId,
      connectorId: id,
      expectedFingerprint: fingerprintConnector(current),
      ...(expectedPreprocessFingerprint === undefined ? {} : { expectedPreprocessFingerprint }),
    }
    const result = this.tail.then(async () => {
      this.assertUsable()
      validateRefreshRequest(request)
      const live = this.require(id)
      if (fingerprintConnector(live) !== request.expectedFingerprint) conflict(`Connector ${id} changed before refresh submission`)
      if (this.options.authorizeRefresh && !await this.options.authorizeRefresh(clone(request))) throw new ConnectorCollaborationError('permission', 'Collaborative connector refresh is not authorized')
      let entry: ConnectorCollaborationEntry
      try { entry = await this.transport.requestRefresh(this.options.room, clone(request), this.sequence, signal) } catch (cause) {
        this.transportFailure(cause, 'Collaborative connector refresh request failed')
      }
      const validated = this.authoritativeEntry(entry)
      if (validated.operation.kind !== 'refresh' || validated.operation.opId !== request.requestId
        || validated.operation.clientId !== request.clientId || validated.operation.id !== request.connectorId
        || validated.operation.expectedFingerprint !== request.expectedFingerprint
        || request.expectedPreprocessFingerprint !== undefined && validated.operation.preprocessFingerprint !== request.expectedPreprocessFingerprint) {
        this.block(this.sequence + 1, validated.sequence, new Error('Refresh acknowledgement does not match request'))
        throw new ConnectorCollaborationError('conflict', 'Authoritative refresh acknowledgement does not match the submitted request')
      }
      const status = await this.applyEntry(validated, true)
      if (status !== 'applied') throw new ConnectorCollaborationError('conflict', 'Authoritative connector refresh could not be applied')
      this.emit({ type: 'submitted', entry: clone(validated) })
      return clone(validated)
    })
    this.tail = result.then(() => undefined, () => undefined)
    return result
  }

  receive(entryInput: unknown): Promise<'applied' | 'duplicate' | 'blocked'> {
    const result = this.tail.then(() => this.applyEntry(entryInput, false))
    this.tail = result.then(() => undefined, () => undefined)
    return result
  }

  resync(input: ConnectorCollaborationResync): Promise<boolean> {
    const result = this.tail.then(async () => {
      let resync: ConnectorCollaborationResync
      try { resync = validateResync(input, this.options.room) } catch { return false }
      if (this.disposed || resync.sequence < this.sequence) return false
      // Reconcile the client-owned old extents as well as the authoritative
      // extents. Without these tombstones, a resync after a remove, move, or
      // shrink would replace the model while leaving orphaned cells behind.
      // Subtract authoritative rectangles so clears are disjoint and the
      // final postcondition can verify every range exactly.
      const staleRanges = staleConnectorRanges(this.controller.manager.snapshotState(), resync.ranges)
      const snapshot: ConnectorCommandSnapshotV1 = {
        version: 1,
        manager: resync.manager,
        ranges: [...staleRanges.map(toRangeSnapshot), ...resync.ranges.map(toRangeSnapshot)],
      }
      try {
        const applied = await this.options.applySnapshot(clone(snapshot), { kind: 'resync', remote: true })
        if (!applied) {
          this.block(this.sequence + 1, resync.sequence, new Error('Connector resync snapshot was not applied'))
          return false
        }
      } catch (cause) {
        this.block(this.sequence + 1, resync.sequence, cause)
        return false
      }
      if (!this.postcondition(snapshot)) {
        this.block(this.sequence + 1, resync.sequence, new Error('Connector resync apply postcondition failed'))
        return false
      }
      this.sequence = resync.sequence
      this.blocked = false
      return true
    })
    this.tail = result.then(() => undefined, () => undefined)
    return result
  }

  dispose(): void { if (!this.disposed) { this.disposed = true; this.unsubscribe?.(); this.unsubscribe = undefined } }

  private submit(operation: ConnectorLifecycleCollaborationOperation, signal = new AbortController().signal): Promise<ConnectorCollaborationEntry> {
    const result = this.tail.then(async () => {
      this.assertUsable()
      applyConnectorCollaborationOperation(this.controller.manager.snapshotState(), operation)
      this.verifyPatch(cellPatchOf(operation))
      if (this.options.authorize && !await this.options.authorize(clone(operation))) throw new ConnectorCollaborationError('permission', 'Connector collaboration operation is not authorized')
      let entry: ConnectorCollaborationEntry
      try { entry = await this.transport.submit(this.options.room, clone(operation), this.sequence, signal) } catch (cause) {
        this.transportFailure(cause, 'Connector operation submission failed')
      }
      const validated = this.authoritativeEntry(entry)
      if (canonical(validated.operation) !== canonical(operation)) {
        this.block(this.sequence + 1, validated.sequence, new Error('Acknowledgement operation changed'))
        throw new ConnectorCollaborationError('conflict', 'Authoritative acknowledgement changed the submitted operation')
      }
      const status = await this.applyEntry(validated, true)
      if (status !== 'applied') throw new ConnectorCollaborationError('conflict', 'Authoritative connector acknowledgement could not be applied')
      this.emit({ type: 'submitted', entry: clone(validated) })
      return clone(validated)
    })
    this.tail = result.then(() => undefined, () => undefined)
    return result
  }

  private async applyEntry(entryInput: unknown, own: boolean): Promise<'applied' | 'duplicate' | 'blocked'> {
    if (this.disposed || this.blocked) return 'blocked'
    let entry: ConnectorCollaborationEntry
    try { entry = validateEntry(entryInput, this.options.room) } catch (cause) { return this.block(this.sequence + 1, -1, cause) }
    if (entry.sequence <= this.sequence) { this.emit({ type: 'duplicate', sequence: entry.sequence }); return 'duplicate' }
    if (entry.sequence !== this.sequence + 1) return this.block(this.sequence + 1, entry.sequence)
    if (own && entry.operation.clientId !== this.options.clientId) return this.block(this.sequence + 1, entry.sequence, new Error('Acknowledgement client identity changed'))
    try {
      const next = applyConnectorCollaborationOperation(this.controller.manager.snapshotState(), entry.operation)
      const cellPatch = cellPatchOf(entry.operation)
      this.verifyPatch(cellPatch)
      const snapshot: ConnectorCommandSnapshotV1 = {
        version: 1,
        manager: next,
        ranges: cellPatch ? [toRangeSnapshot(cellPatch.after)] : [],
      }
      const applied = await this.options.applySnapshot(clone(snapshot), { kind: 'entry', remote: !own, entry: clone(entry) })
      if (!applied) conflict('Connector model and bounded cells could not be applied atomically')
      if (!this.postcondition(snapshot)) conflict('Connector collaboration apply postcondition failed')
    } catch (cause) { return this.block(this.sequence + 1, entry.sequence, cause) }
    this.sequence = entry.sequence
    this.emit({ type: 'applied', entry: clone(entry) })
    return 'applied'
  }

  private verifyPatch(patch: ConnectorCollaborationCellPatch | undefined): void {
    if (!patch) return
    const before = this.controller.manager.captureRangeSnapshot(patch.after, patch.after.rowCount, patch.after.columnCount)
    if (!verifyConnectorCollaborationCellPatch(before, patch)) conflict('Connector target cells changed concurrently')
  }

  private clearPatchField(current: ConnectorMountedSnapshot): { cellPatch?: ConnectorCollaborationCellPatch } {
    if (!current.lastRows || !current.lastColumns) return {}
    const before = this.controller.manager.captureExtent(current.spec.id)
    if (!before) conflict('Connector extent is unavailable')
    return {
      cellPatch: {
        expectedFingerprint: fingerprintConnectorRange(before),
        after: {
          sheetId: current.spec.target.sheetId,
          startRow: current.spec.target.startRow,
          startColumn: current.spec.target.startColumn,
          rowCount: current.lastRows,
          columnCount: current.lastColumns,
          values: Array.from({ length: current.lastRows }, () => Array.from({ length: current.lastColumns }, () => null)),
        },
      },
    }
  }

  private require(id: string): ConnectorMountedSnapshot {
    const value = this.controller.manager.snapshotState().connectors.find(({ spec }) => spec.id === id)
    if (!value) throw new ConnectorCollaborationError('conflict', `Unknown connector ${id}`)
    return value
  }

  private base<T extends Omit<ConnectorLifecycleCollaborationOperation, keyof OperationBase>>(value: T): T & OperationBase {
    return { protocol: CONNECTOR_COLLABORATION_PROTOCOL, opId: this.nextId(), clientId: this.options.clientId, ...value }
  }

  private nextId(): string {
    const value = this.options.idFactory()
    if (!validId(value)) throw new TypeError('idFactory returned an invalid operation id')
    return value
  }

  private authoritativeEntry(input: unknown): ConnectorCollaborationEntry {
    try { return validateEntry(input, this.options.room) } catch (cause) {
      this.block(this.sequence + 1, -1, cause)
      throw new ConnectorCollaborationError('conflict', 'Authoritative connector response is invalid', { cause })
    }
  }

  private transportFailure(cause: unknown, message: string): never {
    if (cause instanceof ConnectorCollaborationTransportError) {
      if (cause.code === 'DENIED') throw new ConnectorCollaborationError('permission', message, { cause })
      // Every other failure has an uncertain commit outcome. Stop accepting
      // work until resync proves whether the server appended the operation.
      this.block(this.sequence + 1, -1, cause)
      if (cause.code === 'STALE_BASE') throw new ConnectorCollaborationError('stale-base', message, { cause })
      if (cause.code === 'CONFLICT') throw new ConnectorCollaborationError('conflict', message, { cause })
      throw new ConnectorCollaborationError('transport', message, { cause })
    }
    this.block(this.sequence + 1, -1, cause)
    throw new ConnectorCollaborationError('transport', message, { cause })
  }

  private postcondition(snapshot: ConnectorCommandSnapshotV1): boolean {
    if (canonical(this.controller.manager.snapshotState()) !== canonical(snapshot.manager)) return false
    return snapshot.ranges.every((range) => {
      const actual = this.controller.manager.captureRangeSnapshot(range, range.rowCount, range.columnCount)
      return !!actual && canonical(actual) === canonical(range)
    })
  }

  private assertUsable(): void {
    if (this.disposed) throw new ConnectorCollaborationError('blocked', 'Connector collaboration session is disposed')
    if (this.blocked) throw new ConnectorCollaborationError('blocked', 'Connector collaboration session requires resynchronization')
  }

  private block(expectedSequence: number, receivedSequence: number, cause?: unknown): 'blocked' {
    this.blocked = true
    this.emit({ type: 'blocked', expectedSequence, receivedSequence, cause })
    try { this.options.onResyncRequired?.(expectedSequence, receivedSequence, cause) } catch { /* observer isolation */ }
    return 'blocked'
  }

  private emit(event: ConnectorCollaborationEvent): void { try { this.options.onEvent?.(event) } catch { /* observer isolation */ } }
}

function validateOperation(input: ConnectorCollaborationOperation): ConnectorCollaborationOperation {
  assertJsonSize(input)
  const operation = clone(input)
  if (!operation || typeof operation !== 'object' || operation.protocol !== CONNECTOR_COLLABORATION_PROTOCOL || !validId(operation.opId) || !validId(operation.clientId)) invalid('Invalid connector operation identity')
  if (!['create', 'update', 'remove', 'refresh'].includes(operation.kind)) invalid('Unsupported connector operation')
  const baseKeys = ['clientId', 'kind', 'opId', 'protocol']
  if (operation.kind === 'create') {
    exactKeys(operation, [...baseKeys, 'value'])
    validateSpec(operation.value)
  } else {
    exactKeys(operation, [...baseKeys, 'expectedFingerprint', 'id', ...(operation.kind === 'remove' ? [] : ['value']), ...(operation.kind === 'refresh' ? ['sourceRevision'] : []), ...(operation.cellPatch ? ['cellPatch'] : []), ...(operation.kind === 'refresh' && operation.preprocessFingerprint !== undefined ? ['preprocessFingerprint'] : [])])
    if (!validId(operation.id) || !validFingerprint(operation.expectedFingerprint)) invalid('Invalid connector object precondition')
    if (operation.cellPatch) validateCellPatch(operation.cellPatch)
    if (operation.kind === 'update') validateSpec(operation.value)
    if (operation.kind === 'refresh') {
      if (!validOpaque(operation.sourceRevision)) invalid('Invalid connector source revision')
      if (operation.preprocessFingerprint !== undefined && !validOpaque(operation.preprocessFingerprint)) invalid('Invalid preprocessing fingerprint')
      validateMounted(operation.value)
    }
  }
  return operation
}

function validateRefreshRequest(input: ConnectorCollaborativeRefreshRequest): ConnectorCollaborativeRefreshRequest {
  assertJsonSize(input)
  const request = clone(input)
  exactKeys(request, ['clientId', 'connectorId', 'expectedFingerprint', 'protocol', 'requestId', ...(request.expectedPreprocessFingerprint === undefined ? [] : ['expectedPreprocessFingerprint'])])
  if (request.protocol !== CONNECTOR_COLLABORATION_PROTOCOL || !validId(request.requestId) || !validId(request.clientId)
    || !validId(request.connectorId) || !validFingerprint(request.expectedFingerprint)
    || request.expectedPreprocessFingerprint !== undefined && !validOpaque(request.expectedPreprocessFingerprint)) invalid('Invalid collaborative refresh request')
  return request
}

function validateEntry(input: unknown, room: string): ConnectorCollaborationEntry {
  assertJsonSize(input)
  if (!input || typeof input !== 'object') invalid('Connector entry must be an object')
  const entry = input as ConnectorCollaborationEntry
  exactKeys(entry, ['operation', 'room', 'sequence'])
  if (entry.room !== room || !Number.isSafeInteger(entry.sequence) || entry.sequence < 1) invalid('Invalid connector entry room or sequence')
  return { room, sequence: entry.sequence, operation: validateOperation(entry.operation) }
}

function validateManagerSnapshot(snapshot: ConnectorManagerSnapshotV1): void {
  if (!isConnectorManagerSnapshot(snapshot) || snapshot.connectors.length > MAX_CONNECTORS || snapshot.connectors.some(({ spec }) => !isCredentialFreeConnectorSpec(spec))) invalid('Invalid or credential-bearing connector state')
  let totalCells = 0
  for (const entry of snapshot.connectors) {
    if (entry.spec.target.startRow >= MAX_SHEET_ROWS || entry.spec.target.startColumn >= MAX_SHEET_COLUMNS) invalid('Connector target exceeds worksheet bounds')
    if ((entry.lastRows === 0) !== (entry.lastColumns === 0)) invalid('Connector extent dimensions must both be zero or positive')
    if (entry.lastRows > MAX_ROWS || entry.lastColumns > MAX_COLUMNS || entry.lastRows * entry.lastColumns > MAX_CELLS
      || entry.spec.target.startRow + entry.lastRows > MAX_SHEET_ROWS || entry.spec.target.startColumn + entry.lastColumns > MAX_SHEET_COLUMNS) invalid('Connector extent exceeds collaboration limits')
    totalCells += entry.lastRows * entry.lastColumns
    if (totalCells > MAX_CELLS) invalid('Connector state exceeds the collaborative cell limit')
    if (entry.status.rowCount !== undefined && entry.status.rowCount !== entry.lastRows) invalid('Connector row status does not match its extent')
    if (entry.status.columnCount !== undefined && entry.status.columnCount !== entry.lastColumns) invalid('Connector column status does not match its extent')
    assertAvailableTarget({ version: 1, connectors: snapshot.connectors.filter(({ spec }) => spec.id !== entry.spec.id) }, entry.spec.id, entry.spec.target, entry.lastRows, entry.lastColumns)
  }
}

function validateSpec(spec: ConnectorSpec): void {
  validateManagerSnapshot({ version: 1, connectors: [{ spec, status: {}, lastRows: 0, lastColumns: 0 }] })
}

function validateMounted(value: ConnectorMountedSnapshot): void {
  validateManagerSnapshot({ version: 1, connectors: [value] })
  if ((value.lastRows === 0) !== (value.lastColumns === 0)) invalid('Connector extent dimensions must both be zero or positive')
  if (!Number.isFinite(value.status.lastRefreshTs) || (value.status.lastRefreshTs ?? -1) < 0) invalid('Connector refresh timestamp is required')
  if (value.status.rowCount !== value.lastRows) invalid('Connector row status does not match result extent')
  if (value.status.columnCount !== value.lastColumns) invalid('Connector column status does not match result extent')
}

function validateCellPatch(patch: ConnectorCollaborationCellPatch): void {
  if (!patch || typeof patch !== 'object') invalid('Invalid connector cell patch')
  exactKeys(patch, ['after', 'expectedFingerprint'])
  if (!validFingerprint(patch.expectedFingerprint)) invalid('Invalid connector cell fingerprint')
  const range = patch.after
  if (!range || typeof range !== 'object') invalid('Invalid connector result range')
  exactKeys(range, ['columnCount', 'rowCount', 'sheetId', 'startColumn', 'startRow', 'values'])
  if (!validId(range.sheetId) || !dimensions(range.startRow, range.startColumn, range.rowCount, range.columnCount)
    || range.rowCount > MAX_ROWS || range.columnCount > MAX_COLUMNS || range.rowCount * range.columnCount > MAX_CELLS
    || range.startRow + range.rowCount > MAX_SHEET_ROWS || range.startColumn + range.columnCount > MAX_SHEET_COLUMNS
    || range.rowCount < 1 || range.columnCount < 1 || !Array.isArray(range.values) || range.values.length !== range.rowCount
    || range.values.some((row) => !Array.isArray(row) || row.length !== range.columnCount || row.some((cell) => !validCell(cell)))) invalid('Invalid connector result range')
}

function validateClearPatch(current: ConnectorMountedSnapshot, patch: ConnectorCollaborationCellPatch | undefined): void {
  const hasExtent = current.lastRows > 0 && current.lastColumns > 0
  if (!hasExtent) { if (patch) invalid('Connector without an extent must not clear cells'); return }
  if (!patch || !matchesTarget(patch.after, current.spec.target) || patch.after.rowCount !== current.lastRows || patch.after.columnCount !== current.lastColumns
    || patch.after.values.some((row) => row.some((value) => value !== null))) invalid('Connector removal or move must clear its exact prior extent')
}

function validateRefreshPatch(current: ConnectorMountedSnapshot, next: ConnectorMountedSnapshot, patch: ConnectorCollaborationCellPatch | undefined): void {
  const rows = Math.max(current.lastRows, next.lastRows)
  const columns = Math.max(current.lastColumns, next.lastColumns)
  if (!rows || !columns) { if (patch) invalid('Empty connector refresh must not contain a cell patch'); return }
  if (!patch || !matchesTarget(patch.after, current.spec.target) || patch.after.rowCount !== rows || patch.after.columnCount !== columns) invalid('Refresh patch must cover the union of old and new extents')
  for (let row = 0; row < rows; row++) for (let column = 0; column < columns; column++) {
    if ((row >= next.lastRows || column >= next.lastColumns) && patch.after.values[row][column] !== null) invalid('Refresh shrink remainder must be cleared')
  }
  validateResultSchema(next.spec, patch.after.values.slice(0, next.lastRows).map((row) => row.slice(0, next.lastColumns)))
}

function assertAvailableTarget(snapshot: ConnectorManagerSnapshotV1, id: string, target: ConnectorSpec['target'], rows: number, columns: number): void {
  const candidate = { sheetId: target.sheetId, startRow: target.startRow, startColumn: target.startColumn, rowCount: Math.max(1, rows), columnCount: Math.max(1, columns) }
  for (const entry of snapshot.connectors) {
    if (entry.spec.id === id) continue
    const other = { sheetId: entry.spec.target.sheetId, startRow: entry.spec.target.startRow, startColumn: entry.spec.target.startColumn, rowCount: Math.max(1, entry.lastRows), columnCount: Math.max(1, entry.lastColumns) }
    if (overlaps(candidate, other)) conflict(`Connector ${id} overlaps connector ${entry.spec.id}`)
  }
}

function cellPatchOf(operation: ConnectorCollaborationOperation): ConnectorCollaborationCellPatch | undefined {
  return operation.kind === 'create' ? undefined : operation.cellPatch
}

function toRangeSnapshot(range: ConnectorCollaborationRange): ConnectorRangeSnapshot {
  return { ...clone(range), values: range.values.map((row) => row.map((value) => ({ v: value }))) }
}

function sameTarget(left: ConnectorSpec['target'], right: ConnectorSpec['target']): boolean { return left.sheetId === right.sheetId && left.startRow === right.startRow && left.startColumn === right.startColumn }
function matchesTarget(range: ConnectorCollaborationRange, target: ConnectorSpec['target']): boolean { return range.sheetId === target.sheetId && range.startRow === target.startRow && range.startColumn === target.startColumn }
function sameRange(left: ConnectorRangeSnapshot, right: ConnectorCollaborationRange): boolean { return left.sheetId === right.sheetId && left.startRow === right.startRow && left.startColumn === right.startColumn && left.rowCount === right.rowCount && left.columnCount === right.columnCount }
function overlaps(left: Omit<ConnectorCollaborationRange, 'values'>, right: Omit<ConnectorCollaborationRange, 'values'>): boolean { return left.sheetId === right.sheetId && left.startRow < right.startRow + right.rowCount && right.startRow < left.startRow + left.rowCount && left.startColumn < right.startColumn + right.columnCount && right.startColumn < left.startColumn + left.columnCount }
function dimensions(...values: unknown[]): boolean { return values.every((value) => Number.isSafeInteger(value) && (value as number) >= 0) }
function validCell(value: unknown): value is RangeValue { return value === null || typeof value === 'string' && value.length <= MAX_STRING || typeof value === 'boolean' || typeof value === 'number' && Number.isFinite(value) }
function validId(value: unknown): value is string { return typeof value === 'string' && /^[A-Za-z0-9][A-Za-z0-9._:/-]{0,255}$/.test(value) }
function validOpaque(value: unknown): value is string { return typeof value === 'string' && value.length > 0 && value.length <= 512 && !/[\u0000-\u001f\u007f]/.test(value) }
function validFingerprint(value: unknown): value is string { return typeof value === 'string' && /^sha256:[0-9a-f]{64}$/.test(value) }
function exactKeys(value: object, expected: readonly string[]): void { if (canonical(Object.keys(value).sort()) !== canonical([...expected].sort())) invalid('Connector collaboration payload contains missing or unsupported fields') }
function clone<T>(value: T): T { try { return JSON.parse(JSON.stringify(value)) as T } catch (cause) { throw new ConnectorCollaborationError('invalid', 'Connector collaboration payload must be plain JSON', { cause }) } }
function canonical(value: unknown): string { if (Array.isArray(value)) return `[${value.map(canonical).join(',')}]`; if (value && typeof value === 'object') return `{${Object.keys(value).sort().map((key) => `${JSON.stringify(key)}:${canonical((value as Record<string, unknown>)[key])}`).join(',')}}`; return JSON.stringify(value) }
function hash(value: string): string { return `sha256:${sha256Hex(value)}` }
function invalid(message: string): never { throw new ConnectorCollaborationError('invalid', message) }
function conflict(message: string): never { throw new ConnectorCollaborationError('conflict', message) }

const MAX_OPERATION_BYTES = 8 * 1024 * 1024
const MAX_CONNECTORS = 10_000
const MAX_ROWS = 10_000
const MAX_COLUMNS = 1_024
const MAX_CELLS = 100_000
const MAX_STRING = 1_000_000
const MAX_SHEET_ROWS = 1_048_576
const MAX_SHEET_COLUMNS = 16_384

function assertJsonSize(value: unknown): void {
  let encoded: string
  try { encoded = JSON.stringify(value) } catch (cause) { throw new ConnectorCollaborationError('invalid', 'Connector collaboration payload must be JSON', { cause }) }
  if (encoded === undefined || new TextEncoder().encode(encoded).byteLength > MAX_OPERATION_BYTES) invalid('Connector collaboration payload exceeds the encoded size limit')
}

function validateResultSchema(spec: ConnectorSpec, grid: RangeValue[][]): void {
  const schema = spec.schema
  if (!schema) return
  const width = grid[0]?.length ?? 0
  const declaredWidth = schema.columns.reduce((maximum, column) => Math.max(maximum, column.index + 1), 0)
  if (!schema.allowAdditionalColumns && width > declaredWidth) invalid('Connector refresh exceeds its declared schema width')
  for (const column of schema.columns) for (let row = 0; row < grid.length; row++) {
    const value = grid[row][column.index] ?? null
    if (value === null) { if (!column.nullable) invalid(`Connector refresh has null at row ${row} column ${column.index}`) }
    else if (column.type !== 'any' && typeof value !== column.type) invalid(`Connector refresh violates ${column.type} schema at row ${row} column ${column.index}`)
  }
}

function validateResync(input: ConnectorCollaborationResync, room: string): ConnectorCollaborationResync {
  assertJsonSize(input)
  const value = clone(input)
  exactKeys(value, ['manager', 'protocol', 'ranges', 'revision', 'room', 'sequence'])
  if (value.protocol !== CONNECTOR_COLLABORATION_PROTOCOL || value.room !== room || !Number.isSafeInteger(value.sequence) || value.sequence < 0 || !validOpaque(value.revision)) invalid('Invalid connector resync envelope')
  validateManagerSnapshot(value.manager)
  if (!Array.isArray(value.ranges)) invalid('Invalid connector resync ranges')
  value.ranges.forEach((range) => validateCellPatch({ expectedFingerprint: `sha256:${'0'.repeat(64)}`, after: range }))
  const expected = value.manager.connectors.filter((entry) => entry.lastRows > 0).map((entry) => `${entry.spec.id}:${entry.spec.target.sheetId}:${entry.spec.target.startRow}:${entry.spec.target.startColumn}:${entry.lastRows}:${entry.lastColumns}`).sort()
  const actual = value.ranges.map((range) => {
    const owner = value.manager.connectors.find((entry) => matchesTarget(range, entry.spec.target) && range.rowCount === entry.lastRows && range.columnCount === entry.lastColumns)
    if (!owner) invalid('Connector resync range is not bound to one connector extent')
    validateResultSchema(owner.spec, range.values)
    return `${owner.spec.id}:${range.sheetId}:${range.startRow}:${range.startColumn}:${range.rowCount}:${range.columnCount}`
  }).sort()
  if (canonical(expected) !== canonical(actual)) invalid('Connector resync must cover every and only connector-owned extent')
  return value
}

function staleConnectorRanges(current: ConnectorManagerSnapshotV1, authoritative: ConnectorCollaborationRange[]): ConnectorCollaborationRange[] {
  validateManagerSnapshot(current)
  const result: ConnectorCollaborationRange[] = []
  for (const entry of current.connectors) {
    if (!entry.lastRows || !entry.lastColumns) continue
    let pieces: RangeRect[] = [{
      sheetId: entry.spec.target.sheetId,
      startRow: entry.spec.target.startRow,
      startColumn: entry.spec.target.startColumn,
      rowCount: entry.lastRows,
      columnCount: entry.lastColumns,
    }]
    for (const next of authoritative) pieces = pieces.flatMap((piece) => subtractRange(piece, next))
    for (const piece of pieces) result.push({
      ...piece,
      values: Array.from({ length: piece.rowCount }, () => Array.from({ length: piece.columnCount }, () => null)),
    })
  }
  return result
}

type RangeRect = Omit<ConnectorCollaborationRange, 'values'>

function subtractRange(source: RangeRect, cut: RangeRect): RangeRect[] {
  if (!overlaps(source, cut)) return [source]
  const top = Math.max(source.startRow, cut.startRow)
  const left = Math.max(source.startColumn, cut.startColumn)
  const bottom = Math.min(source.startRow + source.rowCount, cut.startRow + cut.rowCount)
  const right = Math.min(source.startColumn + source.columnCount, cut.startColumn + cut.columnCount)
  const pieces: RangeRect[] = []
  if (source.startRow < top) pieces.push({ ...source, rowCount: top - source.startRow })
  if (bottom < source.startRow + source.rowCount) pieces.push({ ...source, startRow: bottom, rowCount: source.startRow + source.rowCount - bottom })
  if (source.startColumn < left) pieces.push({ ...source, startRow: top, rowCount: bottom - top, columnCount: left - source.startColumn })
  if (right < source.startColumn + source.columnCount) pieces.push({ ...source, startRow: top, startColumn: right, rowCount: bottom - top, columnCount: source.startColumn + source.columnCount - right })
  return pieces.filter(({ rowCount, columnCount }) => rowCount > 0 && columnCount > 0)
}
