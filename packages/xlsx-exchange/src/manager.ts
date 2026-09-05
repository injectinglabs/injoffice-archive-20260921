import {
  XlsxExchangeError,
  type ExportResult,
  type ExportServerUnitRequest,
  type ExportSnapshotRequest,
  type ImportServerUnitRequest,
  type ImportServerUnitResult,
  type ImportSnapshotRequest,
  type ImportSnapshotResult,
  type XlsxByteSource,
  type XlsxExchangeContext,
  type XlsxExchangeErrorCode,
  type XlsxExchangeEvent,
  type XlsxExchangeJob,
  type XlsxExchangeJobKind,
  type XlsxExchangeJobSnapshot,
  type XlsxExchangeManagerOptions,
  type XlsxExchangePhase,
  type XlsxExchangeProgress,
  type XlsxFileResult,
  type XlsxReadResult,
} from './types.js'

const XLSX_MEDIA_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
const DEFAULT_MAX_OUTPUT_BYTES = 256 * 1024 * 1024
const PHASES = new Set<XlsxExchangePhase>(['queued', 'reading', 'transferring', 'converting', 'capturing', 'loading', 'writing', 'complete'])

interface MutableJob {
  readonly id: string
  readonly kind: XlsxExchangeJobKind
  state: XlsxExchangeJobSnapshot['state']
  progress: XlsxExchangeProgress
  readonly createdAt: number
  startedAt?: number
  finishedAt?: number
  error?: XlsxExchangeError
}

export class XlsxExchangeManager<Snapshot, SourceRef = string, DestinationRef = string> {
  private readonly jobs = new Map<string, MutableJob>()
  private readonly listeners = new Set<(event: XlsxExchangeEvent) => void>()
  private sequence = 0
  readonly maxOutputBytes: number

  constructor(private readonly options: XlsxExchangeManagerOptions<Snapshot, SourceRef, DestinationRef>) {
    if (!options || typeof options !== 'object' || !options.codec) throw new TypeError('XLSX exchange requires a codec.')
    this.maxOutputBytes = options.maxOutputBytes ?? DEFAULT_MAX_OUTPUT_BYTES
    if (!Number.isSafeInteger(this.maxOutputBytes) || this.maxOutputBytes < 1 || this.maxOutputBytes > DEFAULT_MAX_OUTPUT_BYTES) {
      throw new TypeError(`maxOutputBytes must be a positive safe integer no greater than ${DEFAULT_MAX_OUTPUT_BYTES}.`)
    }
  }

  subscribe(listener: (event: XlsxExchangeEvent) => void): () => void {
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }

  listJobs(): XlsxExchangeJobSnapshot[] {
    return [...this.jobs.values()].map((job) => this.freeze(job))
  }

  getJob(id: string): XlsxExchangeJobSnapshot | undefined {
    const job = this.jobs.get(id)
    return job ? this.freeze(job) : undefined
  }

  importSnapshot(request: ImportSnapshotRequest<SourceRef>): XlsxExchangeJob<ImportSnapshotResult<Snapshot>> {
    request = snapshotImportRequest(request)
    return this.start('import-snapshot', async (context) => {
      this.validateSource(request?.source, context)
      this.validateLoad(request.load, context)
      const file = await this.read(request.source, context)
      const importer = this.options.codec.importSnapshot
      if (!importer) this.unsupported(context, 'Snapshot import is not configured.')
      context.report({ phase: 'converting', fraction: 0.35 })
      const imported = await this.call('IMPORT_FAILED', context, () => importer!(file, context))
      if (!imported || imported.snapshot === undefined) {
        throw this.error('IMPORT_FAILED', context, 'Snapshot importer returned no snapshot.', false)
      }
      this.throwIfAborted(context)
      let loadedUnitId: string | undefined
      if (request.load) {
        const workspace = this.requireWorkspace(context)
        context.report({ phase: 'loading', fraction: 0.8 })
        const loaded = await this.call('LOAD_FAILED', context, () => workspace.loadSnapshot(
          imported.snapshot,
          { ...request.load!, sourceName: file.name },
          context,
        ))
        if (!loaded || !validId(loaded.unitId)) throw this.error('LOAD_FAILED', context, 'Workspace returned no loaded unitId.', false)
        loadedUnitId = loaded.unitId
      }
      return { ...imported, loadedUnitId, sourceName: file.name }
    })
  }

  importServerUnit(request: ImportServerUnitRequest<SourceRef>): XlsxExchangeJob<ImportServerUnitResult> {
    request = snapshotImportRequest(request)
    return this.start('import-server-unit', async (context) => {
      this.validateSource(request?.source, context)
      this.validateLoad(request.load, context)
      const file = await this.read(request.source, context)
      const importer = this.options.codec.importServerUnit
      if (!importer) this.unsupported(context, 'Server-unit import is not configured.')
      context.report({ phase: 'transferring', fraction: 0.35 })
      const imported = await this.call('IMPORT_FAILED', context, () => importer!(file, context))
      if (!imported || !validId(imported.unitId)) {
        throw this.error('IMPORT_FAILED', context, 'Server-unit importer returned no unitId.', false)
      }
      this.throwIfAborted(context)
      let loadedUnitId: string | undefined
      if (request.load) {
        const workspace = this.requireWorkspace(context)
        context.report({ phase: 'loading', fraction: 0.8 })
        const loaded = await this.call('LOAD_FAILED', context, () => workspace.loadServerUnit(imported.unitId, request.load!, context))
        if (!loaded || !validId(loaded.unitId)) throw this.error('LOAD_FAILED', context, 'Workspace returned no loaded unitId.', false)
        loadedUnitId = loaded.unitId
      }
      return { ...imported, loadedUnitId, sourceName: file.name }
    })
  }

  exportSnapshot(request: ExportSnapshotRequest<Snapshot, DestinationRef>): XlsxExchangeJob<ExportResult> {
    return this.start('export-snapshot', async (context) => {
      if (!request || typeof request !== 'object') this.invalid(context, 'Snapshot export request must be an object.')
      if (request.snapshot === undefined && !validId(request.unitId)) {
        this.invalid(context, 'Snapshot export requires snapshot or unitId.')
      }
      let snapshot = request.snapshot
      if (snapshot === undefined) {
        const workspace = this.requireWorkspace(context)
        context.report({ phase: 'capturing', fraction: 0.15 })
        snapshot = await this.call('SNAPSHOT_CAPTURE_FAILED', context, () => workspace.captureSnapshot(request.unitId!, context))
      }
      const exporter = this.options.codec.exportSnapshot
      if (!exporter) this.unsupported(context, 'Snapshot export is not configured.')
      this.throwIfAborted(context)
      context.report({ phase: 'converting', fraction: 0.4 })
      const file = await this.call('EXPORT_FAILED', context, async () => normalizeFile(await exporter!(snapshot!, context), this.maxOutputBytes))
      return this.write(file, request.destination, context)
    })
  }

  exportServerUnit(request: ExportServerUnitRequest<DestinationRef>): XlsxExchangeJob<ExportResult> {
    return this.start('export-server-unit', async (context) => {
      if (!request || typeof request !== 'object') this.invalid(context, 'Server-unit export request must be an object.')
      if (!validId(request.unitId)) this.invalid(context, 'Server-unit export requires a nonempty unitId.')
      if (request.expectedRevision !== undefined && !validId(request.expectedRevision)) this.invalid(context, 'Server-unit export expectedRevision must be a nonempty string.')
      const exporter = this.options.codec.exportServerUnit
      if (!exporter) this.unsupported(context, 'Server-unit export is not configured.')
      context.report({ phase: 'transferring', fraction: 0.3 })
      const file = await this.call('EXPORT_FAILED', context, async () => normalizeFile(await exporter!(request.unitId, context), this.maxOutputBytes))
      if (request.expectedRevision !== undefined && file.revision !== request.expectedRevision) {
        throw this.error('EXPORT_FAILED', context, 'Server-unit exporter returned a different immutable revision.', false)
      }
      return this.write(file, request.destination, context)
    })
  }

  private start<Result>(kind: XlsxExchangeJobKind, operation: (context: XlsxExchangeContext) => Promise<Result>): XlsxExchangeJob<Result> {
    const id = this.nextId()
    const controller = new AbortController()
    const job: MutableJob = {
      id,
      kind,
      state: 'queued',
      progress: { phase: 'queued', fraction: 0 },
      createdAt: this.now(),
    }
    this.jobs.set(id, job)
    this.emit({ type: 'job-created', job: this.freeze(job) })
    const context: XlsxExchangeContext = {
      jobId: id,
      signal: controller.signal,
      report: (progress) => this.report(job, progress),
    }
    const result = Promise.resolve().then(async () => {
      job.state = 'running'
      job.startedAt = this.now()
      this.emitUpdated(job)
      try {
        this.throwIfAborted(context)
        const value = await operation(context)
        this.throwIfAborted(context)
        job.state = 'succeeded'
        job.progress = { phase: 'complete', fraction: 1 }
        job.finishedAt = this.now()
        this.emitUpdated(job)
        return value
      } catch (cause) {
        const error = this.toError(cause, context, controller.signal.aborted)
        job.state = error.code === 'CANCELED' ? 'canceled' : 'failed'
        job.error = error
        job.finishedAt = this.now()
        this.emitUpdated(job)
        throw error
      }
    })
    return Object.freeze({
      id,
      kind,
      result,
      cancel: (reason?: string) => controller.abort(reason ?? 'Canceled by caller.'),
      snapshot: () => this.freeze(job),
    })
  }

  private async read(source: XlsxByteSource<SourceRef>, context: XlsxExchangeContext): Promise<XlsxReadResult> {
    context.report({ phase: 'reading', fraction: 0.05 })
    let file: XlsxReadResult
    if (source.kind === 'bytes') {
      if (!(source.bytes instanceof Uint8Array) || source.bytes.byteLength === 0) this.invalid(context, 'XLSX source bytes must be nonempty.')
      file = { bytes: source.bytes, name: source.name, mediaType: XLSX_MEDIA_TYPE }
    } else {
      const storage = this.options.storage
      if (!storage) this.unsupported(context, 'Storage-backed source reading is not configured.')
      file = await this.call('SOURCE_READ_FAILED', context, () => storage!.read(source.ref, context))
      if (!(file.bytes instanceof Uint8Array) || file.bytes.byteLength === 0) {
        throw this.error('SOURCE_READ_FAILED', context, 'Storage returned empty or invalid XLSX bytes.', false)
      }
      file = { ...file, name: file.name ?? source.name }
    }
    this.throwIfAborted(context)
    return { ...file, bytes: copyBytes(file.bytes) }
  }

  private async write(file: XlsxFileResult, destination: DestinationRef | undefined, context: XlsxExchangeContext): Promise<ExportResult> {
    this.throwIfAborted(context)
    const storage = this.options.storage
    if (!storage) this.unsupported(context, 'Save-as storage is not configured.')
    context.report({ phase: 'writing', fraction: 0.85, totalBytes: file.bytes.byteLength })
    const safeFile = { ...file, bytes: copyBytes(file.bytes), mediaType: file.mediaType ?? XLSX_MEDIA_TYPE }
    const storageFile = { ...safeFile, bytes: copyBytes(safeFile.bytes) }
    const written = await this.call('DESTINATION_WRITE_FAILED', context, () => storage!.saveAs(storageFile, destination, context))
    return { ...written, file: safeFile }
  }

  private validateLoad(load: { mode: 'new-unit' | 'replace'; targetUnitId?: string } | undefined, context: XlsxExchangeContext): void {
    if (!load) return
    if (load.mode !== 'new-unit' && load.mode !== 'replace') this.invalid(context, 'Load mode must be new-unit or replace.')
    if (load.mode === 'replace' && !validId(load.targetUnitId)) {
      this.invalid(context, 'Replacing a workbook requires targetUnitId for stale-target protection.')
    }
  }

  private validateSource(source: XlsxByteSource<SourceRef> | undefined, context: XlsxExchangeContext): void {
    if (!source || typeof source !== 'object' || (source.kind !== 'bytes' && source.kind !== 'storage')) {
      this.invalid(context, 'XLSX source must be bytes or a storage reference.')
    }
  }

  private requireWorkspace(context: XlsxExchangeContext) {
    if (!this.options.workspace) this.unsupported(context, 'Workbook loading/capture is not configured.')
    return this.options.workspace!
  }

  private report(job: MutableJob, progress: XlsxExchangeProgress): void {
    if (job.state !== 'running') return
    if (!PHASES.has(progress.phase)) throw new TypeError('Unknown XLSX exchange progress phase.')
    if (!Number.isFinite(progress.fraction) || progress.fraction < 0 || progress.fraction > 1) {
      throw new RangeError('XLSX exchange progress fraction must be within 0..1.')
    }
    if (progress.loadedBytes !== undefined && (!Number.isSafeInteger(progress.loadedBytes) || progress.loadedBytes < 0)) {
      throw new RangeError('loadedBytes must be a nonnegative safe integer.')
    }
    if (progress.totalBytes !== undefined && (!Number.isSafeInteger(progress.totalBytes) || progress.totalBytes < 0)) {
      throw new RangeError('totalBytes must be a nonnegative safe integer.')
    }
    if (progress.loadedBytes !== undefined && progress.totalBytes !== undefined && progress.loadedBytes > progress.totalBytes) {
      throw new RangeError('loadedBytes cannot exceed totalBytes.')
    }
    job.progress = Object.freeze({ ...progress, fraction: Math.max(progress.fraction, job.progress.fraction) })
    this.emitUpdated(job)
  }

  private async call<Result>(code: XlsxExchangeErrorCode, context: XlsxExchangeContext, operation: () => Promise<Result>): Promise<Result> {
    try {
      const result = await operation()
      this.throwIfAborted(context)
      return result
    } catch (cause) {
      if (cause instanceof XlsxExchangeError) throw cause
      if (context.signal.aborted || isAbortError(cause)) throw this.canceled(context, cause)
      throw this.error(code, context, messageOf(cause), retryable(code), cause)
    }
  }

  private throwIfAborted(context: XlsxExchangeContext): void {
    if (context.signal.aborted) throw this.canceled(context, context.signal.reason)
  }

  private canceled(context: XlsxExchangeContext, cause?: unknown): XlsxExchangeError {
    return this.error('CANCELED', context, messageOf(cause) || 'XLSX exchange canceled.', false, cause)
  }

  private invalid(context: XlsxExchangeContext, message: string): never {
    throw this.error('INVALID_REQUEST', context, message, false)
  }

  private unsupported(context: XlsxExchangeContext, message: string): never {
    throw this.error('UNSUPPORTED_OPERATION', context, message, false)
  }

  private error(code: XlsxExchangeErrorCode, context: XlsxExchangeContext, message: string, canRetry: boolean, cause?: unknown): XlsxExchangeError {
    const job = this.jobs.get(context.jobId)!
    return new XlsxExchangeError(code, message || code, {
      jobId: context.jobId,
      kind: job.kind,
      phase: job.progress.phase,
      retryable: canRetry,
      cause,
    })
  }

  private toError(cause: unknown, context: XlsxExchangeContext, aborted: boolean): XlsxExchangeError {
    if (cause instanceof XlsxExchangeError) return cause
    if (aborted || isAbortError(cause)) return this.canceled(context, cause)
    return this.error('IMPORT_FAILED', context, messageOf(cause), false, cause)
  }

  private freeze(job: MutableJob): XlsxExchangeJobSnapshot {
    return Object.freeze({
      id: job.id,
      kind: job.kind,
      state: job.state,
      progress: Object.freeze({ ...job.progress }),
      createdAt: job.createdAt,
      startedAt: job.startedAt,
      finishedAt: job.finishedAt,
      error: job.error,
    })
  }

  private emitUpdated(job: MutableJob): void {
    this.emit({ type: 'job-updated', job: this.freeze(job) })
  }

  private emit(event: XlsxExchangeEvent): void {
    for (const listener of this.listeners) {
      try {
        listener(event)
      } catch {
        // Observers cannot change an exchange operation's outcome.
      }
    }
  }

  private nextId(): string {
    const id = this.options.createJobId?.() ?? `xlsx-job-${++this.sequence}`
    if (!validId(id) || this.jobs.has(id)) throw new Error('XLSX exchange job IDs must be unique nonempty strings.')
    return id
  }

  private now(): number {
    return this.options.now?.() ?? Date.now()
  }
}

function normalizeFile(file: XlsxFileResult, maxOutputBytes: number): XlsxFileResult {
  if (!file || !(file.bytes instanceof Uint8Array) || file.bytes.byteLength === 0) {
    throw new TypeError('XLSX exporter returned empty or invalid bytes.')
  }
  if (file.bytes.byteLength > maxOutputBytes) {
    throw new TypeError(`XLSX exporter exceeded the ${maxOutputBytes}-byte output limit.`)
  }
  if (file.mediaType !== undefined && file.mediaType !== XLSX_MEDIA_TYPE) {
    throw new TypeError('XLSX exporter returned a non-XLSX media type.')
  }
  if (file.name !== undefined && !boundedText(file.name, 1024)) {
    throw new TypeError('XLSX exporter returned an invalid or oversized file name.')
  }
  if (file.revision !== undefined && !boundedText(file.revision, 512)) {
    throw new TypeError('XLSX exporter returned an invalid or oversized revision.')
  }
  if (!boundedMetadata(file.metadata)) {
    throw new TypeError('XLSX exporter returned invalid or oversized metadata.')
  }
  return { ...file, bytes: copyBytes(file.bytes), mediaType: file.mediaType ?? XLSX_MEDIA_TYPE }
}

function boundedMetadata(value: Readonly<Record<string, string>> | undefined): boolean {
  return value === undefined || value !== null && typeof value === 'object' && !Array.isArray(value)
    && Object.getPrototypeOf(value) === Object.prototype && Object.keys(value).length <= 32
    && Object.entries(value).every(([key, entry]) => boundedText(key, 128) && boundedText(entry, 2048))
}

function boundedText(value: unknown, maximum: number): value is string {
  return typeof value === 'string' && value.length > 0 && value.length <= maximum && value.trim() === value
    && !/[\u0000-\u001f\u007f]/.test(value)
}

function copyBytes(bytes: Uint8Array): Uint8Array {
  return Uint8Array.from(bytes)
}

function snapshotImportRequest<SourceRef, Request extends ImportSnapshotRequest<SourceRef> | ImportServerUnitRequest<SourceRef>>(request: Request): Request {
  if (request?.source?.kind !== 'bytes' || !(request.source.bytes instanceof Uint8Array)) return request
  return {
    ...request,
    source: { ...request.source, bytes: copyBytes(request.source.bytes) },
  }
}

function validId(value: unknown): value is string {
  return typeof value === 'string' && value.trim().length > 0
}

function messageOf(value: unknown): string {
  if (typeof value === 'string') return value
  return value instanceof Error ? value.message : ''
}

function isAbortError(value: unknown): boolean {
  return value instanceof DOMException && value.name === 'AbortError'
}

function retryable(code: XlsxExchangeErrorCode): boolean {
  return code === 'SOURCE_READ_FAILED' || code === 'IMPORT_FAILED' || code === 'EXPORT_FAILED' || code === 'DESTINATION_WRITE_FAILED'
}
