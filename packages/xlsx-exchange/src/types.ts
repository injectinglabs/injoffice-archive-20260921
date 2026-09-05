export type XlsxExchangeJobKind =
  | 'import-snapshot'
  | 'import-server-unit'
  | 'export-snapshot'
  | 'export-server-unit'

export type XlsxExchangeJobState = 'queued' | 'running' | 'succeeded' | 'failed' | 'canceled'

export type XlsxExchangePhase =
  | 'queued'
  | 'reading'
  | 'transferring'
  | 'converting'
  | 'capturing'
  | 'loading'
  | 'writing'
  | 'complete'

export interface XlsxExchangeProgress {
  readonly phase: XlsxExchangePhase
  /** Monotonic normalized completion in the inclusive range 0..1. */
  readonly fraction: number
  readonly loadedBytes?: number
  readonly totalBytes?: number
  readonly message?: string
}

export type XlsxExchangeErrorCode =
  | 'INVALID_REQUEST'
  | 'UNSUPPORTED_OPERATION'
  | 'SOURCE_READ_FAILED'
  | 'IMPORT_FAILED'
  | 'LOAD_FAILED'
  | 'SNAPSHOT_CAPTURE_FAILED'
  | 'EXPORT_FAILED'
  | 'DESTINATION_WRITE_FAILED'
  | 'CANCELED'

export class XlsxExchangeError extends Error {
  override readonly name = 'XlsxExchangeError'

  constructor(
    readonly code: XlsxExchangeErrorCode,
    message: string,
    readonly details: {
      readonly jobId: string
      readonly kind: XlsxExchangeJobKind
      readonly phase: XlsxExchangePhase
      readonly retryable: boolean
      readonly cause?: unknown
    },
  ) {
    super(message, details.cause === undefined ? undefined : { cause: details.cause })
  }
}

export interface XlsxExchangeContext {
  readonly jobId: string
  readonly signal: AbortSignal
  report(progress: Omit<XlsxExchangeProgress, 'fraction'> & { fraction: number }): void
}

export type XlsxByteSource<SourceRef> =
  | { readonly kind: 'bytes'; readonly bytes: Uint8Array; readonly name?: string }
  | { readonly kind: 'storage'; readonly ref: SourceRef; readonly name?: string }

export interface XlsxReadResult {
  readonly bytes: Uint8Array
  readonly name?: string
  readonly mediaType?: string
}

export interface XlsxFileResult {
  readonly bytes: Uint8Array
  readonly name?: string
  readonly mediaType?: string
  readonly revision?: string
  readonly metadata?: Readonly<Record<string, string>>
}

export interface XlsxWriteResult {
  readonly location?: string
  readonly metadata?: Readonly<Record<string, string>>
}

export interface XlsxExchangeStorage<SourceRef, DestinationRef> {
  read(source: SourceRef, context: XlsxExchangeContext): Promise<XlsxReadResult>
  saveAs(file: XlsxFileResult, destination: DestinationRef | undefined, context: XlsxExchangeContext): Promise<XlsxWriteResult>
}

export interface XlsxSnapshotImportResult<Snapshot> {
  readonly snapshot: Snapshot
  readonly artifactId?: string
  readonly revision?: string
  readonly metadata?: Readonly<Record<string, string>>
}

export interface XlsxServerUnitImportResult {
  readonly unitId: string
  readonly revision?: string
  readonly metadata?: Readonly<Record<string, string>>
}

export interface XlsxExchangeCodec<Snapshot> {
  importSnapshot?(file: XlsxReadResult, context: XlsxExchangeContext): Promise<XlsxSnapshotImportResult<Snapshot>>
  importServerUnit?(file: XlsxReadResult, context: XlsxExchangeContext): Promise<XlsxServerUnitImportResult>
  exportSnapshot?(snapshot: Snapshot, context: XlsxExchangeContext): Promise<XlsxFileResult>
  exportServerUnit?(unitId: string, context: XlsxExchangeContext): Promise<XlsxFileResult>
}

export interface XlsxLoadOptions {
  readonly mode: 'new-unit' | 'replace'
  /** Required for replace so an asynchronously completed import cannot replace a different active unit. */
  readonly targetUnitId?: string
}

export interface XlsxLoadResult {
  readonly unitId: string
}

export interface XlsxExchangeWorkspace<Snapshot> {
  loadSnapshot(snapshot: Snapshot, options: XlsxLoadOptions & { readonly sourceName?: string }, context: XlsxExchangeContext): Promise<XlsxLoadResult>
  loadServerUnit(unitId: string, options: XlsxLoadOptions, context: XlsxExchangeContext): Promise<XlsxLoadResult>
  captureSnapshot(unitId: string, context: XlsxExchangeContext): Promise<Snapshot>
}

export interface ImportSnapshotRequest<SourceRef> {
  readonly source: XlsxByteSource<SourceRef>
  readonly load?: XlsxLoadOptions
}

export interface ImportServerUnitRequest<SourceRef> {
  readonly source: XlsxByteSource<SourceRef>
  readonly load?: XlsxLoadOptions
}

export interface ExportSnapshotRequest<Snapshot, DestinationRef> {
  readonly snapshot?: Snapshot
  readonly unitId?: string
  readonly destination?: DestinationRef
}

export interface ExportServerUnitRequest<DestinationRef> {
  readonly unitId: string
  readonly destination?: DestinationRef
  /** Optional immutable capture revision checked before saveAs. */
  readonly expectedRevision?: string
}

export interface ImportSnapshotResult<Snapshot> extends XlsxSnapshotImportResult<Snapshot> {
  readonly loadedUnitId?: string
  readonly sourceName?: string
}

export interface ImportServerUnitResult extends XlsxServerUnitImportResult {
  readonly loadedUnitId?: string
  readonly sourceName?: string
}

export interface ExportResult extends XlsxWriteResult {
  readonly file: XlsxFileResult
}

export interface XlsxExchangeJobSnapshot {
  readonly id: string
  readonly kind: XlsxExchangeJobKind
  readonly state: XlsxExchangeJobState
  readonly progress: XlsxExchangeProgress
  readonly createdAt: number
  readonly startedAt?: number
  readonly finishedAt?: number
  readonly error?: XlsxExchangeError
}

export interface XlsxExchangeJob<Result> {
  readonly id: string
  readonly kind: XlsxExchangeJobKind
  readonly result: Promise<Result>
  cancel(reason?: string): void
  snapshot(): XlsxExchangeJobSnapshot
}

export type XlsxExchangeEvent =
  | { readonly type: 'job-created'; readonly job: XlsxExchangeJobSnapshot }
  | { readonly type: 'job-updated'; readonly job: XlsxExchangeJobSnapshot }

export interface XlsxExchangeManagerOptions<Snapshot, SourceRef, DestinationRef> {
  readonly codec: XlsxExchangeCodec<Snapshot>
  readonly workspace?: XlsxExchangeWorkspace<Snapshot>
  readonly storage?: XlsxExchangeStorage<SourceRef, DestinationRef>
  readonly createJobId?: () => string
  readonly now?: () => number
  /** Maximum XLSX produced by a codec, checked before bytes cross saveAs. */
  readonly maxOutputBytes?: number
}
