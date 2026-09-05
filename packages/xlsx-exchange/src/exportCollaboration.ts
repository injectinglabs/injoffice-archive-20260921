import { sha256HexBytes } from './exportHash.js'
import type { XlsxCollaborationRoomState } from './collaboration.js'
import type { XlsxExchangeManager } from './manager.js'
import type { ExportResult } from './types.js'

export const XLSX_COLLABORATIVE_EXPORT_PROTOCOL = 'injoffice.xlsx-collaborative-export.v1' as const
const XLSX_MEDIA_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
const MAX_CONTROL_BYTES = 64 * 1024

export interface XlsxCollaborativeExportRequest<DestinationRef> {
  readonly room: string
  readonly expectedArtifactVersion: string
  readonly expectedLogEpoch: string
  readonly expectedHeadSequence: number
  readonly destination?: DestinationRef
  readonly signal?: AbortSignal
}

export interface XlsxCollaborativeExportSnapshot {
  readonly room: string
  readonly sourceUnitId: string
  /** Immutable server-side unit captured specifically for this export. */
  readonly exportUnitId: string
  readonly artifactVersion: string
  readonly logEpoch: string
  readonly headSequence: number
  readonly revision: string
}

export interface XlsxCollaborativeExportLease {
  readonly requestId: string
  readonly transactionId: string
  readonly expiresAt: number
  readonly state: XlsxCollaborationRoomState
  readonly snapshot: XlsxCollaborativeExportSnapshot
}

export type XlsxCollaborativeExportBeginResult =
  | { readonly status: 'granted'; readonly lease: XlsxCollaborativeExportLease }
  | { readonly status: 'denied' | 'busy' | 'stale-head' | 'pending-outbound' | 'unsaved-head' | 'invalid-state' }

export interface XlsxCollaborativeExportFileBinding {
  readonly revision: string
  readonly byteLength: number
  readonly sha256: string
  readonly location?: string
}

export interface XlsxCollaborativeExportReceipt {
  readonly requestId: string
  readonly transactionId: string
  readonly snapshot: XlsxCollaborativeExportSnapshot
  readonly file: XlsxCollaborativeExportFileBinding
  /** The room is active again, still at the exact exported head. */
  readonly state: XlsxCollaborationRoomState
}

export type XlsxCollaborativeExportResolution =
  | { readonly status: 'committed'; readonly receipt: XlsxCollaborativeExportReceipt }
  | { readonly status: 'aborted'; readonly requestId: string; readonly state: XlsxCollaborationRoomState }

/** Server boundary. `authorizeAndCapture` must atomically authorize, quiesce a
 * fully-saved room head, and mint the immutable export unit. `completeExport`
 * records the exact output digest and resumes that unchanged head. */
export interface XlsxCollaborativeExportBoundary<DestinationRef> {
  authorizeAndCapture(context: {
    readonly protocol: typeof XLSX_COLLABORATIVE_EXPORT_PROTOCOL
    readonly requestId: string
    readonly room: string
    readonly expectedArtifactVersion: string
    readonly expectedLogEpoch: string
    readonly expectedHeadSequence: number
    readonly destination?: DestinationRef
    readonly signal: AbortSignal
  }): Promise<XlsxCollaborativeExportBeginResult>
  completeExport(context: {
    readonly protocol: typeof XLSX_COLLABORATIVE_EXPORT_PROTOCOL
    readonly requestId: string
    readonly transactionId: string
    readonly snapshot: XlsxCollaborativeExportSnapshot
    readonly file: XlsxCollaborativeExportFileBinding
    readonly signal: AbortSignal
  }): Promise<XlsxCollaborativeExportReceipt>
  /** Resolve an uncertain request idempotently. An uncommitted transaction must
   * be abandoned and its unchanged room resumed before `aborted` is returned. */
  resolveExport(context: {
    readonly protocol: typeof XLSX_COLLABORATIVE_EXPORT_PROTOCOL
    readonly requestId: string
    readonly room: string
    readonly signal: AbortSignal
  }): Promise<XlsxCollaborativeExportResolution>
}

export interface XlsxCollaborativeExportCoordinatorOptions {
  readonly idFactory: () => string
  readonly now?: () => number
  /** May only lower the exchange manager's pre-save output limit. */
  readonly maxFileBytes?: number
  readonly onRecoveryRequired?: (recovery: XlsxCollaborativeExportRecovery, cause: unknown) => void
}

export interface XlsxCollaborativeExportRecovery {
  readonly protocol: typeof XLSX_COLLABORATIVE_EXPORT_PROTOCOL
  readonly requestId: string
  readonly room: string
}

export interface XlsxCollaborativeExportResult {
  readonly receipt: XlsxCollaborativeExportReceipt
  readonly export: ExportResult
}

export type XlsxCollaborativeExportResyncResult =
  | { readonly status: 'committed'; readonly receipt: XlsxCollaborativeExportReceipt; readonly export?: ExportResult }
  | { readonly status: 'aborted'; readonly state: XlsxCollaborationRoomState }

export type XlsxCollaborativeExportErrorCode =
  | 'busy'
  | 'blocked'
  | 'canceled'
  | 'permission'
  | 'stale-head'
  | 'pending-outbound'
  | 'unsaved-head'
  | 'invalid-state'
  | 'invalid-request'
  | 'invalid-response'
  | 'uncertain-outcome'

export class XlsxCollaborativeExportError extends Error {
  override readonly name: string = 'XlsxCollaborativeExportError'

  constructor(readonly code: XlsxCollaborativeExportErrorCode, message: string, options?: ErrorOptions) {
    super(message, options)
  }
}

export class XlsxCollaborativeExportRecoveryError extends XlsxCollaborativeExportError {
  override readonly name = 'XlsxCollaborativeExportRecoveryError'

  constructor(readonly recovery: XlsxCollaborativeExportRecovery, cause: unknown) {
    super('uncertain-outcome', 'The collaborative export outcome is uncertain; resynchronization is required.', { cause })
  }
}

interface PendingRecovery<DestinationRef> {
  readonly token: XlsxCollaborativeExportRecovery
  readonly request: ValidatedRequest<DestinationRef>
  readonly lease?: XlsxCollaborativeExportLease
  readonly exported?: ExportResult
}

interface ValidatedRequest<DestinationRef> {
  readonly room: string
  readonly expectedArtifactVersion: string
  readonly expectedLogEpoch: string
  readonly expectedHeadSequence: number
  readonly destination?: DestinationRef
  readonly signal: AbortSignal
}

/** Coordinates a server-captured export. It never captures a browser snapshot:
 * only the immutable unit in the validated server lease reaches the exporter. */
export class XlsxCollaborativeExportCoordinator<Snapshot, SourceRef = string, DestinationRef = string> {
  private busy = false
  private recovery?: PendingRecovery<DestinationRef>
  private readonly now: () => number
  private readonly maxFileBytes: number

  constructor(
    private readonly exchange: XlsxExchangeManager<Snapshot, SourceRef, DestinationRef>,
    private readonly boundary: XlsxCollaborativeExportBoundary<DestinationRef>,
    private readonly options: XlsxCollaborativeExportCoordinatorOptions,
  ) {
    if (!options || typeof options.idFactory !== 'function') throw new TypeError('Collaborative export requires an idFactory.')
    this.now = options.now ?? Date.now
    this.maxFileBytes = options.maxFileBytes ?? exchange.maxOutputBytes
    if (!Number.isSafeInteger(this.maxFileBytes) || this.maxFileBytes < 1 || this.maxFileBytes > exchange.maxOutputBytes) {
      throw new TypeError('maxFileBytes must be a positive safe integer no greater than the exchange output limit.')
    }
  }

  get state(): { readonly busy: boolean; readonly blocked: boolean; readonly recovery?: XlsxCollaborativeExportRecovery } {
    return { busy: this.busy, blocked: !!this.recovery, recovery: this.recovery?.token }
  }

  async export(requestInput: XlsxCollaborativeExportRequest<DestinationRef>): Promise<XlsxCollaborativeExportResult> {
    if (this.busy) throw failure('busy', 'A collaborative export is already in progress.')
    if (this.recovery) throw failure('blocked', 'Collaborative export requires resynchronization.')
    const request = validateRequest(requestInput)
    throwIfAborted(request.signal)
    const requestId = this.nextId()
    const token = { protocol: XLSX_COLLABORATIVE_EXPORT_PROTOCOL, requestId, room: request.room } as const
    this.busy = true
    let lease: XlsxCollaborativeExportLease | undefined
    let exported: ExportResult | undefined

    try {
      let begin: XlsxCollaborativeExportBeginResult
      try {
        begin = await this.boundary.authorizeAndCapture({
          protocol: XLSX_COLLABORATIVE_EXPORT_PROTOCOL,
          requestId,
          room: request.room,
          expectedArtifactVersion: request.expectedArtifactVersion,
          expectedLogEpoch: request.expectedLogEpoch,
          expectedHeadSequence: request.expectedHeadSequence,
          destination: request.destination,
          signal: request.signal,
        })
      } catch (cause) {
        throw this.uncertain({ token, request }, cause)
      }

      try { validateBegin(begin) } catch (cause) { throw this.uncertain({ token, request }, cause) }
      if (begin.status !== 'granted') throw rejected(begin.status)
      try { validateLease(begin.lease, request, requestId, this.now()) } catch (cause) {
        throw this.uncertain({ token, request }, cause)
      }
      lease = begin.lease
      if (request.signal.aborted) throw this.uncertain({ token, request, lease }, request.signal.reason)

      const job = this.exchange.exportServerUnit({
        unitId: lease.snapshot.exportUnitId,
        destination: request.destination,
        expectedRevision: lease.snapshot.revision,
      })
      const cancel = () => job.cancel(abortMessage(request.signal))
      request.signal.addEventListener('abort', cancel, { once: true })
      try {
        exported = await job.result
      } catch (cause) {
        throw this.uncertain({ token, request, lease }, cause)
      } finally {
        request.signal.removeEventListener('abort', cancel)
      }

      let binding: XlsxCollaborativeExportFileBinding
      try { binding = validateExport(exported, lease, this.maxFileBytes) } catch (cause) {
        throw this.uncertain({ token, request, lease, exported }, cause)
      }

      let receipt: XlsxCollaborativeExportReceipt
      try {
        receipt = await this.boundary.completeExport({
          protocol: XLSX_COLLABORATIVE_EXPORT_PROTOCOL,
          requestId,
          transactionId: lease.transactionId,
          snapshot: lease.snapshot,
          file: binding,
          // Completion is authoritative once called; caller cancellation cannot
          // turn a valid commit response into a reported rollback.
          signal: request.signal,
        })
        validateReceipt(receipt, request, lease, binding, this.maxFileBytes)
      } catch (cause) {
        throw this.uncertain({ token, request, lease, exported }, cause)
      }

      return { receipt: cloneControl(receipt), export: copyExport(exported) }
    } finally {
      this.busy = false
    }
  }

  async resync(signal = new AbortController().signal): Promise<XlsxCollaborativeExportResyncResult> {
    if (this.busy) throw failure('busy', 'A collaborative export is already in progress.')
    const pending = this.recovery
    if (!pending) throw failure('invalid-state', 'There is no uncertain collaborative export to resynchronize.')
    throwIfAborted(signal)
    this.busy = true
    try {
      let resolution: XlsxCollaborativeExportResolution
      try {
        resolution = await this.boundary.resolveExport({ ...pending.token, signal })
        validateResolution(resolution, pending, this.maxFileBytes)
      } catch (cause) {
        this.notify(pending.token, cause)
        throw new XlsxCollaborativeExportRecoveryError(pending.token, cause)
      }
      this.recovery = undefined
      if (resolution.status === 'aborted') return { status: 'aborted', state: cloneControl(resolution.state) }
      return {
        status: 'committed',
        receipt: cloneControl(resolution.receipt),
        ...(pending.exported ? { export: copyExport(pending.exported) } : {}),
      }
    } finally {
      this.busy = false
    }
  }

  private nextId(): string {
    const id = this.options.idFactory()
    if (!validId(id)) throw new TypeError('Collaborative export idFactory returned an invalid request ID.')
    return id
  }

  private uncertain(pending: PendingRecovery<DestinationRef>, cause: unknown): XlsxCollaborativeExportRecoveryError {
    this.recovery = pending
    this.notify(pending.token, cause)
    return new XlsxCollaborativeExportRecoveryError(pending.token, cause)
  }

  private notify(token: XlsxCollaborativeExportRecovery, cause: unknown): void {
    try { this.options.onRecoveryRequired?.(token, cause) } catch { /* observer isolation */ }
  }
}

function validateRequest<DestinationRef>(input: XlsxCollaborativeExportRequest<DestinationRef>): ValidatedRequest<DestinationRef> {
  if (!plainObject(input) || !onlyKeys(input, ['room', 'expectedArtifactVersion', 'expectedLogEpoch', 'expectedHeadSequence', 'destination', 'signal'])
    || !validId(input.room) || !validId(input.expectedArtifactVersion) || !validId(input.expectedLogEpoch)
    || !validSequence(input.expectedHeadSequence)
    || input.signal !== undefined && !(input.signal instanceof AbortSignal)
    || input.destination !== undefined && !validJsonControl(input.destination)) {
    throw failure('invalid-request', 'Collaborative export requires a bounded room, artifact version, log epoch, and head sequence.')
  }
  assertRequestSize(input)
  return { ...input, signal: input.signal ?? new AbortController().signal }
}

function validateBegin(value: XlsxCollaborativeExportBeginResult): void {
  assertControlSize(value)
  if (!plainObject(value) || !validBeginStatus(value.status)) throw failure('invalid-response', 'Invalid collaborative export begin response.')
  if (value.status === 'granted') {
    if (!onlyKeys(value, ['status', 'lease']) || !('lease' in value)) throw failure('invalid-response', 'Granted export response requires one lease.')
  } else if (!onlyKeys(value, ['status'])) throw failure('invalid-response', 'Rejected export response contains unsupported fields.')
}

function validateLease<DestinationRef>(lease: XlsxCollaborativeExportLease, request: ValidatedRequest<DestinationRef>, requestId: string, now: number): void {
  assertControlSize(lease)
  if (!plainObject(lease) || !onlyExactKeys(lease, ['requestId', 'transactionId', 'expiresAt', 'state', 'snapshot'])
    || lease.requestId !== requestId || !validId(lease.transactionId) || !Number.isFinite(lease.expiresAt) || lease.expiresAt <= now) {
    throw failure('invalid-response', 'Collaborative export lease is invalid or expired.')
  }
  validateState(lease.state)
  validateSnapshot(lease.snapshot)
  if (lease.state.room !== request.room || lease.state.artifactVersion !== request.expectedArtifactVersion
    || lease.state.logEpoch !== request.expectedLogEpoch || lease.state.headSequence !== request.expectedHeadSequence
    || lease.state.phase !== 'quiesced' || lease.state.pendingOutbound !== 0 || lease.state.savedSequence !== lease.state.headSequence
    || lease.snapshot.room !== lease.state.room || lease.snapshot.sourceUnitId !== lease.state.unitId
    || lease.snapshot.artifactVersion !== lease.state.artifactVersion || lease.snapshot.logEpoch !== lease.state.logEpoch
    || lease.snapshot.headSequence !== lease.state.headSequence) {
    throw failure('invalid-response', 'Collaborative export lease is not bound to the requested quiesced saved head.')
  }
}

function validateSnapshot(value: XlsxCollaborativeExportSnapshot): void {
  if (!plainObject(value) || !onlyExactKeys(value, ['room', 'sourceUnitId', 'exportUnitId', 'artifactVersion', 'logEpoch', 'headSequence', 'revision'])
    || !validId(value.room) || !validId(value.sourceUnitId) || !validId(value.exportUnitId)
    || !validId(value.artifactVersion) || !validId(value.logEpoch) || !validSequence(value.headSequence)
    || !validOpaque(value.revision)) throw failure('invalid-response', 'Collaborative export snapshot binding is invalid.')
}

function validateExport(result: ExportResult, lease: XlsxCollaborativeExportLease, maxFileBytes: number): XlsxCollaborativeExportFileBinding {
  if (!plainObject(result) || !onlyKeys(result, ['file', 'location', 'metadata']) || !plainObject(result.file)
    || !onlyKeys(result.file, ['bytes', 'name', 'mediaType', 'revision', 'metadata'])) {
    throw failure('invalid-response', 'Collaborative exporter returned an invalid result envelope.')
  }
  const { file } = result
  if (!(file.bytes instanceof Uint8Array) || file.bytes.byteLength < 1 || file.bytes.byteLength > maxFileBytes
    || file.mediaType !== XLSX_MEDIA_TYPE || file.revision !== lease.snapshot.revision
    || file.name !== undefined && !validText(file.name, 1024)
    || !validMetadata(file.metadata) || result.location !== undefined && !validText(result.location, 4096)
    || !validMetadata(result.metadata)) throw failure('invalid-response', 'Collaborative export is not a bounded XLSX bound to the captured snapshot revision.')
  return {
    revision: file.revision,
    byteLength: file.bytes.byteLength,
    sha256: `sha256:${sha256HexBytes(file.bytes)}`,
    ...(result.location === undefined ? {} : { location: result.location }),
  }
}

function validateReceipt<DestinationRef>(
  receipt: XlsxCollaborativeExportReceipt,
  request: ValidatedRequest<DestinationRef>,
  lease: XlsxCollaborativeExportLease,
  binding: XlsxCollaborativeExportFileBinding,
  maxFileBytes: number,
): void {
  assertControlSize(receipt)
  if (!plainObject(receipt) || !onlyExactKeys(receipt, ['requestId', 'transactionId', 'snapshot', 'file', 'state'])
    || receipt.requestId !== lease.requestId || receipt.transactionId !== lease.transactionId) {
    throw failure('invalid-response', 'Collaborative export completion receipt is invalid.')
  }
  validateSnapshot(receipt.snapshot)
  validateFileBinding(receipt.file, maxFileBytes)
  validateState(receipt.state)
  if (!sameControl(receipt.snapshot, lease.snapshot) || !sameControl(receipt.file, binding)
    || receipt.file.revision !== receipt.snapshot.revision
    || receipt.state.room !== request.room || receipt.state.unitId !== lease.state.unitId
    || receipt.state.artifactVersion !== request.expectedArtifactVersion || receipt.state.logEpoch !== request.expectedLogEpoch
    || receipt.state.headSequence !== request.expectedHeadSequence || receipt.state.savedSequence !== receipt.state.headSequence
    || receipt.state.pendingOutbound !== 0 || receipt.state.phase !== 'active') {
    throw failure('invalid-response', 'Collaborative export receipt does not bind the exact exported head and bytes.')
  }
}

function validateResolution<DestinationRef>(value: XlsxCollaborativeExportResolution, pending: PendingRecovery<DestinationRef>, maxFileBytes: number): void {
  assertControlSize(value)
  if (!plainObject(value) || (value.status !== 'committed' && value.status !== 'aborted')) throw failure('invalid-response', 'Invalid collaborative export resync response.')
  if (value.status === 'aborted') {
    if (!onlyExactKeys(value, ['status', 'requestId', 'state']) || value.requestId !== pending.token.requestId) throw failure('invalid-response', 'Invalid aborted export resolution.')
    validateState(value.state)
    if (value.state.room !== pending.request.room || value.state.artifactVersion !== pending.request.expectedArtifactVersion
      || value.state.logEpoch !== pending.request.expectedLogEpoch || value.state.headSequence !== pending.request.expectedHeadSequence
      || pending.lease !== undefined && value.state.unitId !== pending.lease.state.unitId
      || value.state.savedSequence !== value.state.headSequence || value.state.pendingOutbound !== 0 || value.state.phase !== 'active') {
      throw failure('invalid-response', 'Aborted export resolution did not restore the exact requested head.')
    }
    return
  }
  if (!onlyExactKeys(value, ['status', 'receipt'])) throw failure('invalid-response', 'Invalid committed export resolution.')
  const lease = pending.lease
  if (!lease) {
    validateRecoveredReceipt(value.receipt, pending, maxFileBytes)
    return
  }
  const binding = pending.exported
    ? validateExport(pending.exported, lease, Number.MAX_SAFE_INTEGER)
    : value.receipt.file
  validateReceipt(value.receipt, pending.request, lease, binding, maxFileBytes)
}

function validateRecoveredReceipt<DestinationRef>(receipt: XlsxCollaborativeExportReceipt, pending: PendingRecovery<DestinationRef>, maxFileBytes: number): void {
  assertControlSize(receipt)
  if (!plainObject(receipt) || !onlyExactKeys(receipt, ['requestId', 'transactionId', 'snapshot', 'file', 'state'])
    || receipt.requestId !== pending.token.requestId || !validId(receipt.transactionId)) throw failure('invalid-response', 'Recovered export receipt is invalid.')
  validateSnapshot(receipt.snapshot)
  validateFileBinding(receipt.file, maxFileBytes)
  validateState(receipt.state)
  if (receipt.file.revision !== receipt.snapshot.revision
    || receipt.snapshot.room !== pending.request.room || receipt.snapshot.artifactVersion !== pending.request.expectedArtifactVersion
    || receipt.snapshot.logEpoch !== pending.request.expectedLogEpoch || receipt.snapshot.headSequence !== pending.request.expectedHeadSequence
    || receipt.state.room !== receipt.snapshot.room || receipt.state.unitId !== receipt.snapshot.sourceUnitId
    || receipt.state.artifactVersion !== receipt.snapshot.artifactVersion || receipt.state.logEpoch !== receipt.snapshot.logEpoch
    || receipt.state.headSequence !== receipt.snapshot.headSequence || receipt.state.savedSequence !== receipt.state.headSequence
    || receipt.state.pendingOutbound !== 0 || receipt.state.phase !== 'active') throw failure('invalid-response', 'Recovered export receipt does not bind the requested head.')
}

function validateFileBinding(value: XlsxCollaborativeExportFileBinding, maxFileBytes = 256 * 1024 * 1024): void {
  if (!plainObject(value) || !onlyKeys(value, ['revision', 'byteLength', 'sha256', 'location'])
    || !validOpaque(value.revision) || !Number.isSafeInteger(value.byteLength) || value.byteLength < 1 || value.byteLength > maxFileBytes
    || typeof value.sha256 !== 'string' || !/^sha256:[0-9a-f]{64}$/.test(value.sha256)
    || value.location !== undefined && !validText(value.location, 4096)) throw failure('invalid-response', 'Collaborative export file binding is invalid.')
}

function validateState(state: XlsxCollaborationRoomState): void {
  if (!plainObject(state) || !onlyExactKeys(state, ['room', 'unitId', 'artifactVersion', 'logEpoch', 'headSequence', 'savedSequence', 'pendingOutbound', 'phase'])
    || !validId(state.room) || !validId(state.unitId) || !validId(state.artifactVersion) || !validId(state.logEpoch)
    || !validSequence(state.headSequence) || !validSequence(state.savedSequence) || state.savedSequence > state.headSequence
    || !validSequence(state.pendingOutbound) || !['active', 'quiesced', 'recovering'].includes(state.phase)) {
    throw failure('invalid-response', 'Collaborative export room state is invalid.')
  }
}

function rejected(status: Exclude<XlsxCollaborativeExportBeginResult['status'], 'granted'>): XlsxCollaborativeExportError {
  const code = status === 'denied' ? 'permission' : status === 'busy' ? 'busy' : status
  return failure(code, `Collaborative export was rejected: ${status}.`)
}

function validBeginStatus(value: unknown): value is XlsxCollaborativeExportBeginResult['status'] {
  return typeof value === 'string' && ['granted', 'denied', 'busy', 'stale-head', 'pending-outbound', 'unsaved-head', 'invalid-state'].includes(value)
}

function copyExport(value: ExportResult): ExportResult {
  return {
    file: {
      ...value.file,
      bytes: value.file.bytes.slice(),
      ...(value.file.metadata ? { metadata: { ...value.file.metadata } } : {}),
    },
    ...(value.location === undefined ? {} : { location: value.location }),
    ...(value.metadata ? { metadata: { ...value.metadata } } : {}),
  }
}

function validMetadata(value: Readonly<Record<string, string>> | undefined): boolean {
  return value === undefined || plainObject(value) && Object.keys(value).length <= 32
    && Object.entries(value).every(([key, entry]) => validText(key, 128) && validText(entry, 2048))
}

function assertControlSize(value: unknown): void {
  let encoded: string
  try { encoded = JSON.stringify(value) } catch (cause) { throw failure('invalid-response', 'Collaborative export control payload must be JSON.', cause) }
  if (encoded === undefined || new TextEncoder().encode(encoded).byteLength > MAX_CONTROL_BYTES) {
    throw failure('invalid-response', 'Collaborative export control payload exceeds its size limit.')
  }
}

function assertRequestSize<DestinationRef>(value: XlsxCollaborativeExportRequest<DestinationRef>): void {
  try {
    const encoded = JSON.stringify({
      room: value.room,
      expectedArtifactVersion: value.expectedArtifactVersion,
      expectedLogEpoch: value.expectedLogEpoch,
      expectedHeadSequence: value.expectedHeadSequence,
      ...(value.destination === undefined ? {} : { destination: value.destination }),
    })
    if (new TextEncoder().encode(encoded).byteLength > MAX_CONTROL_BYTES) throw new Error('oversized')
  } catch (cause) {
    throw failure('invalid-request', 'Collaborative export request must be bounded JSON data.', cause)
  }
}

function cloneControl<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T
}

function sameControl(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}

function onlyKeys(value: object, keys: readonly string[]): boolean {
  return Object.keys(value).every((key) => keys.includes(key))
}

function onlyExactKeys(value: object, keys: readonly string[]): boolean {
  return Object.keys(value).length === keys.length && onlyKeys(value, keys)
}

function plainObject(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value) && Object.getPrototypeOf(value) === Object.prototype
}

function validJsonControl(value: unknown, depth = 0): boolean {
  if (depth > 16) return false
  if (value === null || typeof value === 'boolean' || typeof value === 'string') return true
  if (typeof value === 'number') return Number.isFinite(value)
  if (Array.isArray(value)) return value.length <= 1024 && value.every((entry) => validJsonControl(entry, depth + 1))
  if (!plainObject(value) || Object.keys(value).length > 128) return false
  return Object.entries(value).every(([key, entry]) => validText(key, 128) && validJsonControl(entry, depth + 1))
}

function validText(value: unknown, maximum: number): value is string {
  return typeof value === 'string' && value.length > 0 && value.length <= maximum && value.trim() === value && !/[\u0000-\u001f\u007f]/.test(value)
}

function validId(value: unknown): value is string {
  return validText(value, 256)
}

function validOpaque(value: unknown): value is string {
  return validText(value, 512)
}

function validSequence(value: unknown): value is number {
  return Number.isSafeInteger(value) && (value as number) >= 0
}

function throwIfAborted(signal: AbortSignal): void {
  if (signal.aborted) throw failure('canceled', abortMessage(signal), signal.reason)
}

function abortMessage(signal: AbortSignal): string {
  return signal.reason instanceof Error ? signal.reason.message : String(signal.reason || 'Collaborative export canceled.')
}

function failure(code: XlsxCollaborativeExportErrorCode, message: string, cause?: unknown): XlsxCollaborativeExportError {
  return new XlsxCollaborativeExportError(code, message, cause === undefined ? undefined : { cause })
}
