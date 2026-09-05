export const CALCULATION_PROTOCOL_VERSION = 'injoffice.calculation.v1' as const

export type FormulaScalar = string | number | boolean | null

export interface FormulaCell {
  sheetId: string
  row: number
  column: number
  formula: string
}

export interface FormulaError {
  kind: 'error'
  code: '#NULL!' | '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A' | '#GETTING_DATA' | '#SPILL!' | '#CALC!' | '#CIRCULAR!' | 'ENGINE_ERROR'
  message?: string
}

export type SpillItem = FormulaScalar | FormulaError

export type FormulaCellValue =
  | { kind: 'value'; value: FormulaScalar }
  | FormulaError
  | { kind: 'spill'; values: ReadonlyArray<ReadonlyArray<SpillItem>> }

export interface FormulaCellResult {
  sheetId: string
  row: number
  column: number
  result: FormulaCellValue
}

export interface CalculationEngineOutput {
  cells: ReadonlyArray<FormulaCellResult>
  diagnostics?: ReadonlyArray<{ level: 'info' | 'warning'; code: string; message: string }>
}

export interface CalculationEngineContext<TSnapshot> {
  protocolVersion: typeof CALCULATION_PROTOCOL_VERSION
  jobId: string
  workbookId: string
  sourceRevision: string
  sourceFingerprint: string
  formulas: ReadonlyArray<Readonly<FormulaCell>>
  snapshot: TSnapshot
  metadata: Readonly<Record<string, string>>
  signal: AbortSignal
}

export interface CalculationEngine<TSnapshot> {
  id: string
  version: string
  /** Collaborative/server execution requires an engine that promises repeatable output for identical inputs. */
  deterministic: true
  calculate(context: Readonly<CalculationEngineContext<TSnapshot>>): CalculationEngineOutput | Promise<CalculationEngineOutput>
}

export interface CalculationRequest<TSnapshot> {
  jobId: string
  workbookId: string
  /** Authoritative document revision. Results may only be applied to this revision. */
  sourceRevision: string
  /** Host-provided fingerprint of the complete input snapshot, not merely the formula list. */
  sourceFingerprint: string
  snapshot: TSnapshot
  formulas: ReadonlyArray<FormulaCell>
  metadata?: Readonly<Record<string, string>>
  /** Reject execution unless the configured engine contract matches this fingerprint. */
  expectedEngineFingerprint?: string
}

export interface CalculationResult {
  protocolVersion: typeof CALCULATION_PROTOCOL_VERSION
  jobId: string
  workbookId: string
  sourceRevision: string
  sourceFingerprint: string
  requestFingerprint: string
  engineFingerprint: string
  cells: ReadonlyArray<FormulaCellResult>
  diagnostics: ReadonlyArray<{ level: 'info' | 'warning'; code: string; message: string }>
}

export type CalculationJobState = 'queued' | 'running' | 'completed' | 'canceled' | 'timed-out' | 'failed'

export interface CalculationJobSnapshot {
  id: string
  state: CalculationJobState
  workbookId: string
  sourceRevision: string
  requestFingerprint: string
}

export interface CalculationJob {
  readonly id: string
  readonly result: Promise<CalculationResult>
  cancel(reason?: unknown): boolean
  getState(): CalculationJobSnapshot
}

export type CalculationEvent =
  | { type: 'queued' | 'started'; job: CalculationJobSnapshot }
  | { type: 'completed'; job: CalculationJobSnapshot; result: CalculationResult }
  | { type: 'canceled' | 'timed-out' | 'failed'; job: CalculationJobSnapshot; error: CalculationError }

export type CalculationErrorCode =
  | 'ABORTED'
  | 'DUPLICATE_JOB'
  | 'ENGINE_FAILED'
  | 'ENGINE_MISMATCH'
  | 'INVALID_ENGINE'
  | 'INVALID_REQUEST'
  | 'INVALID_RESULT'
  | 'STALE_RESULT'
  | 'TIMEOUT'

export class CalculationError extends Error {
  constructor(public readonly code: CalculationErrorCode, message: string, options?: ErrorOptions) {
    super(message, options)
    this.name = 'CalculationError'
  }
}

export interface CalculationManagerOptions {
  timeoutMs?: number
  retainedJobs?: number
}

export interface CalculationSubmitOptions {
  signal?: AbortSignal
  timeoutMs?: number
}

const MAX_ROW = 1_048_575
const MAX_COLUMN = 16_383
const STABLE_ID = /^[A-Za-z0-9][A-Za-z0-9._:/-]*$/

function requireText(value: unknown, label: string, max = 512): asserts value is string {
  if (typeof value !== 'string' || value.length === 0 || value.length > max) throw new CalculationError('INVALID_REQUEST', `${label} must contain 1-${max} characters`)
}

function requireStableId(value: unknown, label: string): asserts value is string {
  requireText(value, label, 256)
  if (!STABLE_ID.test(value)) throw new CalculationError('INVALID_REQUEST', `${label} must be a stable identifier`)
}

function clone<T>(value: T, label: string, code: 'INVALID_REQUEST' | 'INVALID_RESULT' = 'INVALID_REQUEST'): T {
  try {
    return structuredClone(value)
  } catch {
    throw new CalculationError(code, `${label} must be structured-cloneable`)
  }
}

function assertCellAddress(input: { sheetId: unknown; row: unknown; column: unknown }, label: string, code: 'INVALID_REQUEST' | 'INVALID_RESULT'): void {
  if (typeof input.sheetId !== 'string' || input.sheetId.length === 0 || input.sheetId.length > 256) throw new CalculationError(code, `${label}.sheetId must contain 1-256 characters`)
  if (!Number.isSafeInteger(input.row) || (input.row as number) < 0 || (input.row as number) > MAX_ROW) throw new CalculationError(code, `${label}.row is outside worksheet bounds`)
  if (!Number.isSafeInteger(input.column) || (input.column as number) < 0 || (input.column as number) > MAX_COLUMN) throw new CalculationError(code, `${label}.column is outside worksheet bounds`)
}

function cellKey(cell: { sheetId: string; row: number; column: number }): string {
  return JSON.stringify([cell.sheetId, cell.row, cell.column])
}

function compareText(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0
}

function canonicalFormulas(input: ReadonlyArray<FormulaCell>): FormulaCell[] {
  if (!Array.isArray(input)) throw new CalculationError('INVALID_REQUEST', 'formulas must be an array')
  if (input.length > 1_000_000) throw new CalculationError('INVALID_REQUEST', 'formulas must not exceed 1000000 cells')
  const seen = new Set<string>()
  const formulas = input.map((cell, index) => {
    if (!cell || typeof cell !== 'object') throw new CalculationError('INVALID_REQUEST', `formulas[${index}] must be an object`)
    assertCellAddress(cell, `formulas[${index}]`, 'INVALID_REQUEST')
    requireText(cell.formula, `formulas[${index}].formula`, 32_768)
    if (!cell.formula.startsWith('=')) throw new CalculationError('INVALID_REQUEST', `formulas[${index}].formula must start with =`)
    const key = cellKey(cell)
    if (seen.has(key)) throw new CalculationError('INVALID_REQUEST', `formulas contains duplicate cell ${cell.sheetId}!R${cell.row + 1}C${cell.column + 1}`)
    seen.add(key)
    return { sheetId: cell.sheetId, row: cell.row, column: cell.column, formula: cell.formula }
  })
  return formulas.sort((left, right) => compareText(left.sheetId, right.sheetId) || left.row - right.row || left.column - right.column)
}

function canonicalMetadata(input: Readonly<Record<string, string>> | undefined): Record<string, string> {
  if (input === undefined) return {}
  if (!input || typeof input !== 'object' || Array.isArray(input)) throw new CalculationError('INVALID_REQUEST', 'metadata must be a string record')
  const result: Record<string, string> = Object.create(null) as Record<string, string>
  for (const key of Object.keys(input).sort()) {
    requireText(key, 'metadata key', 128)
    const value = input[key]
    if (typeof value !== 'string' || value.length > 2_048) throw new CalculationError('INVALID_REQUEST', `metadata.${key} must be a string no longer than 2048 characters`)
    result[key] = value
  }
  return result
}

// Browser-safe deterministic contract identifier; not a cryptographic content signature.
function contractHash(value: string): string {
  let hash = 0x811c9dc5
  for (let index = 0; index < value.length; index++) {
    hash ^= value.charCodeAt(index)
    hash = Math.imul(hash, 0x01000193)
  }
  return `fnv1a32:${(hash >>> 0).toString(16).padStart(8, '0')}`
}

function engineContract<T>(engine: CalculationEngine<T>): string {
  if (!engine || typeof engine !== 'object' || engine.deterministic !== true || typeof engine.calculate !== 'function') throw new CalculationError('INVALID_ENGINE', 'engine must expose calculate() and declare deterministic: true')
  if (typeof engine.id !== 'string' || engine.id.length === 0 || engine.id.length > 256 || !STABLE_ID.test(engine.id)) throw new CalculationError('INVALID_ENGINE', 'engine.id must be a stable identifier')
  if (typeof engine.version !== 'string' || engine.version.length === 0 || engine.version.length > 128) throw new CalculationError('INVALID_ENGINE', 'engine.version must contain 1-128 characters')
  return contractHash(JSON.stringify({ protocolVersion: CALCULATION_PROTOCOL_VERSION, id: engine.id, version: engine.version, deterministic: true }))
}

function validateScalar(value: unknown, label: string): asserts value is FormulaScalar {
  if (value !== null && typeof value !== 'string' && typeof value !== 'boolean' && (typeof value !== 'number' || !Number.isFinite(value))) throw new CalculationError('INVALID_RESULT', `${label} must be a finite number, string, boolean, or null`)
}

const ERROR_CODES = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!', '#CIRCULAR!', 'ENGINE_ERROR'])

function validateError(value: FormulaError, label: string): void {
  if (!ERROR_CODES.has(value.code)) throw new CalculationError('INVALID_RESULT', `${label}.code is not a supported formula error`)
  if (value.message !== undefined && (typeof value.message !== 'string' || value.message.length > 4_096)) throw new CalculationError('INVALID_RESULT', `${label}.message must not exceed 4096 characters`)
}

function validateCellValue(value: FormulaCellValue, cell: FormulaCellResult, label: string): void {
  if (!value || typeof value !== 'object') throw new CalculationError('INVALID_RESULT', `${label}.result must be an object`)
  if (value.kind === 'value') return validateScalar(value.value, `${label}.result.value`)
  if (value.kind === 'error') return validateError(value, `${label}.result`)
  if (value.kind !== 'spill' || !Array.isArray(value.values) || value.values.length === 0) throw new CalculationError('INVALID_RESULT', `${label}.result must be a value, error, or non-empty spill`)
  let width: number | undefined
  for (let row = 0; row < value.values.length; row++) {
    const items = value.values[row]
    if (!Array.isArray(items) || items.length === 0) throw new CalculationError('INVALID_RESULT', `${label}.result.values[${row}] must be a non-empty row`)
    if (width === undefined) width = items.length
    if (items.length !== width) throw new CalculationError('INVALID_RESULT', `${label}.result.values must be rectangular`)
    for (let column = 0; column < items.length; column++) {
      const item = items[column]
      if (item && typeof item === 'object') validateError(item as FormulaError, `${label}.result.values[${row}][${column}]`)
      else validateScalar(item, `${label}.result.values[${row}][${column}]`)
    }
  }
  if (cell.row + value.values.length - 1 > MAX_ROW || cell.column + (width ?? 0) - 1 > MAX_COLUMN) throw new CalculationError('INVALID_RESULT', `${label}.result spill exceeds worksheet bounds`)
}

function validateOutput(output: CalculationEngineOutput): CalculationEngineOutput {
  if (!output || typeof output !== 'object' || !Array.isArray(output.cells)) throw new CalculationError('INVALID_RESULT', 'engine output must contain a cells array')
  if (output.cells.length > 1_000_000) throw new CalculationError('INVALID_RESULT', 'engine output must not exceed 1000000 cells')
  const seen = new Set<string>()
  output.cells.forEach((cell, index) => {
    if (!cell || typeof cell !== 'object') throw new CalculationError('INVALID_RESULT', `cells[${index}] must be an object`)
    assertCellAddress(cell, `cells[${index}]`, 'INVALID_RESULT')
    const key = cellKey(cell)
    if (seen.has(key)) throw new CalculationError('INVALID_RESULT', `engine output contains duplicate cell ${cell.sheetId}!R${cell.row + 1}C${cell.column + 1}`)
    seen.add(key)
    validateCellValue(cell.result, cell, `cells[${index}]`)
  })
  if (output.diagnostics !== undefined) {
    if (!Array.isArray(output.diagnostics)) throw new CalculationError('INVALID_RESULT', 'diagnostics must be an array')
    if (output.diagnostics.length > 10_000) throw new CalculationError('INVALID_RESULT', 'diagnostics must not exceed 10000 entries')
    output.diagnostics.forEach((item, index) => {
      if (!item || !['info', 'warning'].includes(item.level) || typeof item.code !== 'string' || !item.code || item.code.length > 128 || typeof item.message !== 'string' || item.message.length > 4_096) throw new CalculationError('INVALID_RESULT', `diagnostics[${index}] is invalid`)
    })
  }
  return output
}

function timeoutValue(input: number | undefined, fallback: number): number {
  const value = input ?? fallback
  if (!Number.isSafeInteger(value) || value < 1 || value > 3_600_000) throw new CalculationError('INVALID_REQUEST', 'timeoutMs must be an integer from 1 through 3600000')
  return value
}

/** Fail closed before applying derived values to a workbook that may have advanced. */
export function assertCalculationResultCurrent(result: CalculationResult, current: { workbookId: string; revision: string; fingerprint: string }): void {
  if (result.workbookId !== current.workbookId || result.sourceRevision !== current.revision || result.sourceFingerprint !== current.fingerprint) {
    throw new CalculationError('STALE_RESULT', 'calculation result does not match the current workbook identity, revision, and fingerprint')
  }
}

/** Validate and detach a calculation result received across a trust boundary. */
export function validateCalculationResult(input: unknown): CalculationResult {
  if (!input || typeof input !== 'object') throw new CalculationError('INVALID_RESULT', 'calculation result must be an object')
  const result = input as Partial<CalculationResult>
  if (result.protocolVersion !== CALCULATION_PROTOCOL_VERSION) throw new CalculationError('INVALID_RESULT', 'calculation result protocol version is unsupported')
  try {
    requireStableId(result.jobId, 'jobId')
    requireStableId(result.workbookId, 'workbookId')
    requireText(result.sourceRevision, 'sourceRevision')
    requireText(result.sourceFingerprint, 'sourceFingerprint')
    requireText(result.requestFingerprint, 'requestFingerprint')
    requireText(result.engineFingerprint, 'engineFingerprint')
  } catch (cause) {
    throw new CalculationError('INVALID_RESULT', cause instanceof Error ? cause.message : 'calculation result identity is invalid', { cause })
  }
  const output = validateOutput({ cells: result.cells as FormulaCellResult[], diagnostics: result.diagnostics })
  return clone({
    protocolVersion: CALCULATION_PROTOCOL_VERSION,
    jobId: result.jobId,
    workbookId: result.workbookId,
    sourceRevision: result.sourceRevision,
    sourceFingerprint: result.sourceFingerprint,
    requestFingerprint: result.requestFingerprint,
    engineFingerprint: result.engineFingerprint,
    cells: [...output.cells].sort((left, right) => compareText(left.sheetId, right.sheetId) || left.row - right.row || left.column - right.column),
    diagnostics: output.diagnostics ?? [],
  }, 'calculation result', 'INVALID_RESULT')
}

/** Renderer- and transport-neutral lifecycle around a host-injected deterministic formula engine. */
export class CalculationManager<TSnapshot> {
  readonly engineFingerprint: string
  private readonly jobs = new Map<string, { snapshot: CalculationJobSnapshot; controller: AbortController }>()
  private readonly terminalOrder: string[] = []
  private readonly listeners = new Set<(event: Readonly<CalculationEvent>) => void>()
  private readonly defaultTimeoutMs: number
  private readonly retainedJobs: number

  constructor(private readonly engine: CalculationEngine<TSnapshot>, options: CalculationManagerOptions = {}) {
    this.engineFingerprint = engineContract(engine)
    this.defaultTimeoutMs = timeoutValue(options.timeoutMs, 30_000)
    this.retainedJobs = options.retainedJobs ?? 100
    if (!Number.isSafeInteger(this.retainedJobs) || this.retainedJobs < 0 || this.retainedJobs > 10_000) throw new CalculationError('INVALID_REQUEST', 'retainedJobs must be an integer from 0 through 10000')
  }

  onEvent(listener: (event: Readonly<CalculationEvent>) => void): () => void {
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }

  getJob(id: string): CalculationJobSnapshot | undefined {
    const job = this.jobs.get(id)
    return job ? clone(job.snapshot, 'job snapshot') : undefined
  }

  cancel(id: string, reason?: unknown): boolean {
    const job = this.jobs.get(id)
    if (!job || (job.snapshot.state !== 'queued' && job.snapshot.state !== 'running')) return false
    job.controller.abort(reason)
    return true
  }

  submit(request: CalculationRequest<TSnapshot>, options: CalculationSubmitOptions = {}): CalculationJob {
    requireStableId(request?.jobId, 'jobId')
    requireStableId(request.workbookId, 'workbookId')
    requireText(request.sourceRevision, 'sourceRevision')
    requireText(request.sourceFingerprint, 'sourceFingerprint')
    if (this.jobs.has(request.jobId)) throw new CalculationError('DUPLICATE_JOB', `job ${request.jobId} already exists`)
    if (request.expectedEngineFingerprint !== undefined && request.expectedEngineFingerprint !== this.engineFingerprint) throw new CalculationError('ENGINE_MISMATCH', `expected engine ${request.expectedEngineFingerprint}, received ${this.engineFingerprint}`)
    const formulas = canonicalFormulas(request.formulas)
    const metadata = canonicalMetadata(request.metadata)
    const snapshot = clone(request.snapshot, 'snapshot')
    const ms = timeoutValue(options.timeoutMs, this.defaultTimeoutMs)
    const requestFingerprint = contractHash(JSON.stringify({
      protocolVersion: CALCULATION_PROTOCOL_VERSION,
      workbookId: request.workbookId,
      sourceRevision: request.sourceRevision,
      sourceFingerprint: request.sourceFingerprint,
      engineFingerprint: this.engineFingerprint,
      formulas,
      metadata,
    }))
    const controller = new AbortController()
    const record = {
      snapshot: { id: request.jobId, state: 'queued' as CalculationJobState, workbookId: request.workbookId, sourceRevision: request.sourceRevision, requestFingerprint },
      controller,
    }
    this.jobs.set(request.jobId, record)
    this.emit({ type: 'queued', job: record.snapshot })
    const result = this.run(record, { request, snapshot, formulas, metadata }, options, ms)
    return {
      id: request.jobId,
      result,
      cancel: (reason?: unknown) => this.cancel(request.jobId, reason),
      getState: () => clone(record.snapshot, 'job snapshot'),
    }
  }

  private async run(
    record: { snapshot: CalculationJobSnapshot; controller: AbortController },
    input: { request: CalculationRequest<TSnapshot>; snapshot: TSnapshot; formulas: FormulaCell[]; metadata: Record<string, string> },
    options: CalculationSubmitOptions,
    ms: number,
  ): Promise<CalculationResult> {
    const { request } = input
    const externalAbort = () => record.controller.abort(options.signal?.reason)
    options.signal?.addEventListener('abort', externalAbort, { once: true })
    if (options.signal?.aborted) externalAbort()
    let timedOut = false
    const timer = setTimeout(() => {
      timedOut = true
      record.controller.abort(new CalculationError('TIMEOUT', `calculation exceeded ${ms}ms`))
    }, ms)
    try {
      if (record.controller.signal.aborted) throw new CalculationError('ABORTED', 'calculation was canceled before execution')
      this.transition(record, 'running')
      this.emit({ type: 'started', job: record.snapshot })
      const aborted = new Promise<never>((_, reject) => {
        record.controller.signal.addEventListener('abort', () => reject(new CalculationError(timedOut ? 'TIMEOUT' : 'ABORTED', timedOut ? `calculation exceeded ${ms}ms` : 'calculation was canceled')), { once: true })
      })
      let output: CalculationEngineOutput
      try {
        output = await Promise.race([
          Promise.resolve(this.engine.calculate(Object.freeze({
            protocolVersion: CALCULATION_PROTOCOL_VERSION,
            jobId: request.jobId,
            workbookId: request.workbookId,
            sourceRevision: request.sourceRevision,
            sourceFingerprint: request.sourceFingerprint,
            formulas: Object.freeze(input.formulas.map((formula) => Object.freeze(formula))),
            snapshot: input.snapshot,
            metadata: Object.freeze(input.metadata),
            signal: record.controller.signal,
          }))),
          aborted,
        ])
      } catch (cause) {
        if (cause instanceof CalculationError && (cause.code === 'ABORTED' || cause.code === 'TIMEOUT')) throw cause
        if (record.controller.signal.aborted) throw new CalculationError(timedOut ? 'TIMEOUT' : 'ABORTED', timedOut ? `calculation exceeded ${ms}ms` : 'calculation was canceled', { cause })
        throw new CalculationError('ENGINE_FAILED', `calculation engine ${this.engine.id}@${this.engine.version} failed`, { cause })
      }
      const checked = clone(validateOutput(output), 'engine output', 'INVALID_RESULT')
      const result = validateCalculationResult({
        protocolVersion: CALCULATION_PROTOCOL_VERSION,
        jobId: request.jobId,
        workbookId: request.workbookId,
        sourceRevision: request.sourceRevision,
        sourceFingerprint: request.sourceFingerprint,
        requestFingerprint: record.snapshot.requestFingerprint,
        engineFingerprint: this.engineFingerprint,
        cells: [...checked.cells].sort((left, right) => compareText(left.sheetId, right.sheetId) || left.row - right.row || left.column - right.column),
        diagnostics: checked.diagnostics ?? [],
      })
      this.transition(record, 'completed')
      this.emit({ type: 'completed', job: record.snapshot, result })
      return clone(result, 'calculation result')
    } catch (cause) {
      const error = cause instanceof CalculationError ? cause : new CalculationError('ENGINE_FAILED', 'calculation failed', { cause })
      const state = error.code === 'TIMEOUT' ? 'timed-out' : error.code === 'ABORTED' ? 'canceled' : 'failed'
      this.transition(record, state)
      this.emit({ type: state === 'timed-out' ? 'timed-out' : state === 'canceled' ? 'canceled' : 'failed', job: record.snapshot, error })
      throw error
    } finally {
      clearTimeout(timer)
      options.signal?.removeEventListener('abort', externalAbort)
      this.terminalOrder.push(record.snapshot.id)
      while (this.terminalOrder.length > this.retainedJobs) {
        const expired = this.terminalOrder.shift()
        if (expired !== undefined) this.jobs.delete(expired)
      }
    }
  }

  private transition(record: { snapshot: CalculationJobSnapshot }, state: CalculationJobState): void {
    record.snapshot = { ...record.snapshot, state }
  }

  private emit(event: CalculationEvent): void {
    for (const listener of this.listeners) {
      try {
        if ('error' in event) {
          listener({
            type: event.type,
            job: clone(event.job, 'calculation event'),
            error: new CalculationError(event.error.code, event.error.message),
          })
        } else {
          listener(clone(event, 'calculation event'))
        }
      } catch {
        // Lifecycle observers cannot change calculation semantics.
      }
    }
  }
}
