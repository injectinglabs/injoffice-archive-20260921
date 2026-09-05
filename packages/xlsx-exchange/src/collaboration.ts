import type { XlsxExchangeManager } from './manager.js'
import type {
  ImportServerUnitRequest,
  ImportServerUnitResult,
  XlsxExchangeJob,
} from './types.js'

export const XLSX_COLLABORATIVE_IMPORT_PROTOCOL = 'injoffice.xlsx-collaborative-import.v1' as const

export type XlsxCollaborativeImportTarget =
  | { readonly mode: 'new-room'; readonly requestedRoom?: string }
  | {
      readonly mode: 'replace-room'
      readonly room: string
      readonly expectedArtifactVersion: string
      readonly expectedHeadSequence: number
    }

export interface XlsxCollaborativeImportRequest<SourceRef> {
  readonly source: ImportServerUnitRequest<SourceRef>['source']
  readonly target: XlsxCollaborativeImportTarget
  readonly signal?: AbortSignal
}

export interface XlsxCollaborativeImportLease {
  readonly transactionId: string
  readonly expiresAt: number
}

export interface XlsxCollaborationRoomState {
  readonly room: string
  readonly unitId: string
  readonly artifactVersion: string
  readonly logEpoch: string
  readonly headSequence: number
  readonly savedSequence: number
  readonly pendingOutbound: number
  readonly phase: 'active' | 'quiesced' | 'recovering'
}

export interface XlsxCollaborativeImportReceipt extends XlsxCollaborationRoomState {
  readonly transactionId: string
  readonly importedUnitId: string
}

/**
 * Host/server boundary for attaching a converted server unit to collaboration.
 * `commitImport` must atomically validate the lease and replacement head, bind
 * the imported unit, reset the operation epoch, and return the durable result.
 */
export interface XlsxCollaborativeImportBoundary {
  authorizeAndBegin(context: {
    readonly protocol: typeof XLSX_COLLABORATIVE_IMPORT_PROTOCOL
    readonly target: XlsxCollaborativeImportTarget
    readonly signal: AbortSignal
  }): Promise<XlsxCollaborativeImportLease | undefined>
  quiesce(context: { readonly room: string; readonly signal: AbortSignal }): Promise<XlsxCollaborationRoomState>
  commitImport(context: {
    readonly protocol: typeof XLSX_COLLABORATIVE_IMPORT_PROTOCOL
    readonly transactionId: string
    readonly importedUnit: ImportServerUnitResult
    readonly target: XlsxCollaborativeImportTarget
    readonly previous?: XlsxCollaborationRoomState
    readonly signal: AbortSignal
  }): Promise<XlsxCollaborativeImportReceipt>
  /** Release staged resources for a transaction that did not commit. */
  abortImport(context: {
    readonly transactionId: string
    readonly importedUnitId?: string
    readonly cause: unknown
  }): Promise<void>
  /** Resume the unchanged room after a pre-commit replacement failure. */
  resume(context: { readonly room: string; readonly cause: unknown }): Promise<XlsxCollaborationRoomState>
  /** Activate the already-committed room locally. This call must be idempotent. */
  activate(context: { readonly receipt: XlsxCollaborativeImportReceipt }): Promise<void>
  onRecoveryRequired?(context: { readonly receipt: XlsxCollaborativeImportReceipt; readonly cause: unknown }): void
}

export type XlsxCollaborativeImportErrorCode =
  | 'busy'
  | 'canceled'
  | 'permission'
  | 'invalid-target'
  | 'invalid-lease'
  | 'invalid-state'
  | 'pending-outbound'
  | 'unsaved-head'
  | 'stale-head'
  | 'invalid-receipt'

export class XlsxCollaborativeImportError extends Error {
  override readonly name = 'XlsxCollaborativeImportError'

  constructor(readonly code: XlsxCollaborativeImportErrorCode, message: string, options?: ErrorOptions) {
    super(message, options)
  }
}

/** A durable commit succeeded, but the local editor did not activate it. */
export class XlsxCollaborativeImportRecoveryError extends Error {
  override readonly name = 'XlsxCollaborativeImportRecoveryError'

  constructor(readonly receipt: XlsxCollaborativeImportReceipt, cause: unknown) {
    super('The collaborative import committed, but local activation requires recovery.', { cause })
  }
}

/**
 * Serializes collaborative imports and separates the durable server commit from
 * local activation. It deliberately consumes `importServerUnit` without the
 * manager's local `load` option: only the committed collaboration receipt may
 * become the active editor unit.
 */
export class XlsxCollaborativeImportCoordinator<Snapshot, SourceRef = string, DestinationRef = string> {
  private busy = false

  constructor(
    private readonly exchange: XlsxExchangeManager<Snapshot, SourceRef, DestinationRef>,
    private readonly boundary: XlsxCollaborativeImportBoundary,
    private readonly now: () => number = Date.now,
  ) {}

  async import(request: XlsxCollaborativeImportRequest<SourceRef>): Promise<XlsxCollaborativeImportReceipt> {
    if (this.busy) throw conflict('busy', 'A collaborative import is already in progress.')
    this.busy = true
    let lease: XlsxCollaborativeImportLease | undefined
    let imported: ImportServerUnitResult | undefined
    let previous: XlsxCollaborationRoomState | undefined
    let job: XlsxExchangeJob<ImportServerUnitResult> | undefined
    let committed: XlsxCollaborativeImportReceipt | undefined
    const signal = request?.signal ?? new AbortController().signal

    try {
      validateRequest(request)
      throwIfAborted(signal)
      lease = await this.boundary.authorizeAndBegin({
        protocol: XLSX_COLLABORATIVE_IMPORT_PROTOCOL,
        target: request.target,
        signal,
      })
      throwIfAborted(signal)
      if (!lease) throw conflict('permission', 'The current user cannot import into this collaboration target.')
      validateLease(lease, this.now())

      job = this.exchange.importServerUnit({ source: request.source })
      const cancelJob = () => job?.cancel('Collaborative import canceled.')
      signal.addEventListener('abort', cancelJob, { once: true })
      try {
        imported = await job.result
      } finally {
        signal.removeEventListener('abort', cancelJob)
      }
      throwIfAborted(signal)

      if (request.target.mode === 'replace-room') {
        previous = await this.boundary.quiesce({ room: request.target.room, signal })
        validateReplacementState(previous, request.target)
      }
      throwIfAborted(signal)

      committed = await this.boundary.commitImport({
        protocol: XLSX_COLLABORATIVE_IMPORT_PROTOCOL,
        transactionId: lease.transactionId,
        importedUnit: imported,
        target: request.target,
        previous,
        signal,
      })
      validateReceipt(committed, lease, imported, request.target)

      try {
        // The durable commit is authoritative. Activation is attempted even if
        // the caller aborts after the commit response arrives.
        await this.boundary.activate({ receipt: committed })
      } catch (cause) {
        this.boundary.onRecoveryRequired?.({ receipt: committed, cause })
        throw new XlsxCollaborativeImportRecoveryError(committed, cause)
      }
      return committed
    } catch (cause) {
      if (!committed && lease) {
        job?.cancel('Collaborative import did not commit.')
        if (previous) {
          try {
            const resumed = await this.boundary.resume({ room: previous.room, cause })
            validateResumedState(resumed, previous)
          } catch {
            // Preserve the original failure. The host still has the transaction
            // abort hook and its own room recovery channel.
          }
        }
        try {
          await this.boundary.abortImport({
            transactionId: lease.transactionId,
            importedUnitId: imported?.unitId,
            cause,
          })
        } catch {
          // Cleanup failure cannot change whether the durable commit occurred.
        }
      }
      if (signal.aborted && !(cause instanceof XlsxCollaborativeImportRecoveryError)) {
        throw conflict('canceled', abortMessage(signal), cause)
      }
      throw cause
    } finally {
      this.busy = false
    }
  }
}

function validateRequest<SourceRef>(request: XlsxCollaborativeImportRequest<SourceRef>): void {
  if (!request || typeof request !== 'object' || !request.source || typeof request.source !== 'object') {
    throw conflict('invalid-target', 'Collaborative import requires a source and target.')
  }
  const target = request.target
  if (!target || typeof target !== 'object') throw conflict('invalid-target', 'Collaborative import target is required.')
  if (target.mode === 'new-room') {
    if (target.requestedRoom !== undefined && !validId(target.requestedRoom)) {
      throw conflict('invalid-target', 'requestedRoom must be a nonempty identifier.')
    }
    return
  }
  if (target.mode !== 'replace-room' || !validId(target.room) || !validId(target.expectedArtifactVersion)
    || !validSequence(target.expectedHeadSequence)) {
    throw conflict('invalid-target', 'Replacement requires room, artifact version, and a nonnegative expected head sequence.')
  }
}

function validateLease(lease: XlsxCollaborativeImportLease, now: number): void {
  if (!lease || typeof lease !== 'object' || !validId(lease.transactionId)
    || !Number.isFinite(lease.expiresAt) || lease.expiresAt <= now) {
    throw conflict('invalid-lease', 'The collaborative import lease is invalid or expired.')
  }
}

function validateReplacementState(state: XlsxCollaborationRoomState, target: Extract<XlsxCollaborativeImportTarget, { mode: 'replace-room' }>): void {
  validateState(state)
  if (state.room !== target.room || state.phase !== 'quiesced') {
    throw conflict('invalid-state', 'The replacement room did not enter the expected quiesced state.')
  }
  if (state.pendingOutbound !== 0) throw conflict('pending-outbound', 'Outbound collaboration edits remain pending.')
  if (state.savedSequence !== state.headSequence) throw conflict('unsaved-head', 'The collaboration head is not durably saved.')
  if (state.artifactVersion !== target.expectedArtifactVersion || state.headSequence !== target.expectedHeadSequence) {
    throw conflict('stale-head', 'The collaboration room no longer matches the requested replacement head.')
  }
}

function validateReceipt(
  receipt: XlsxCollaborativeImportReceipt,
  lease: XlsxCollaborativeImportLease,
  imported: ImportServerUnitResult,
  target: XlsxCollaborativeImportTarget,
): void {
  try { validateState(receipt) } catch (cause) {
    throw conflict('invalid-receipt', 'The server returned an invalid collaborative import receipt.', cause)
  }
  const expectedRoom = target.mode === 'replace-room' ? target.room : target.requestedRoom
  if (receipt.transactionId !== lease.transactionId || receipt.importedUnitId !== imported.unitId
    || receipt.unitId !== imported.unitId || receipt.phase !== 'active' || receipt.pendingOutbound !== 0
    || receipt.savedSequence !== receipt.headSequence || (expectedRoom !== undefined && receipt.room !== expectedRoom)) {
    throw conflict('invalid-receipt', 'The server receipt does not represent the committed imported unit.')
  }
}

function validateResumedState(resumed: XlsxCollaborationRoomState, previous: XlsxCollaborationRoomState): void {
  validateState(resumed)
  if (resumed.phase !== 'active' || resumed.room !== previous.room || resumed.unitId !== previous.unitId
    || resumed.artifactVersion !== previous.artifactVersion || resumed.logEpoch !== previous.logEpoch
    || resumed.headSequence !== previous.headSequence || resumed.savedSequence !== previous.savedSequence) {
    throw conflict('invalid-state', 'The resumed room does not represent the pre-import collaboration state.')
  }
}

function validateState(state: XlsxCollaborationRoomState): void {
  if (!state || typeof state !== 'object' || !validId(state.room) || !validId(state.unitId)
    || !validId(state.artifactVersion) || !validId(state.logEpoch)
    || !['active', 'quiesced', 'recovering'].includes(state.phase)
    || !validSequence(state.headSequence) || !validSequence(state.savedSequence)
    || !validSequence(state.pendingOutbound) || state.savedSequence > state.headSequence) {
    throw conflict('invalid-state', 'Collaboration room state is invalid.')
  }
}

function throwIfAborted(signal: AbortSignal): void {
  if (signal.aborted) throw conflict('canceled', abortMessage(signal))
}

function abortMessage(signal: AbortSignal): string {
  return signal.reason instanceof Error ? signal.reason.message : String(signal.reason || 'Collaborative import canceled.')
}

function conflict(code: XlsxCollaborativeImportErrorCode, message: string, cause?: unknown): XlsxCollaborativeImportError {
  return new XlsxCollaborativeImportError(code, message, cause === undefined ? undefined : { cause })
}

function validId(value: unknown): value is string {
  return typeof value === 'string' && value.trim().length > 0
}

function validSequence(value: unknown): value is number {
  return Number.isSafeInteger(value) && (value as number) >= 0
}
