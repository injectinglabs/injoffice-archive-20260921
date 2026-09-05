import { sha256Hex } from './sha256'

export type RangeValue = string | number | boolean | null
export type RangeGrid = ReadonlyArray<ReadonlyArray<RangeValue>>

export interface RangeWindow {
  sheetId: string
  startRow: number
  startColumn: number
  rowCount: number
  columnCount: number
}

export interface RangePreprocessContext {
  range: RangeWindow
  /** Stable source revision used by collaborators to reject stale work. */
  revision: string
  mode?: 'local' | 'collaborative'
  metadata?: Readonly<Record<string, string>>
}

export interface RangePreprocessStageContext extends RangePreprocessContext {
  signal: AbortSignal
}

export interface RangePreprocessStage {
  id: string
  version: string
  /** Stable ordering key. IDs break ties, so registration order is irrelevant. */
  order: number
  /** Collaborative execution rejects stages that do not make this guarantee. */
  deterministic: boolean
  process(grid: RangeGrid, context: Readonly<RangePreprocessStageContext>): RangeGrid | Promise<RangeGrid>
}

export interface RangePreprocessStageDescriptor {
  id: string
  version: string
  order: number
  deterministic: boolean
}

export const RANGE_PREPROCESS_MANIFEST_VERSION = 1 as const
const RANGE_PREPROCESS_FINGERPRINT_DOMAIN = 'injoffice:range-preprocess-manifest:v1\n'

/** Portable, credential-free execution contract. Stage implementations remain
 * host code; this manifest can be persisted with workbook/server metadata and
 * checked before a refresh is allowed to run. */
export interface RangePreprocessManifestV1 {
  version: typeof RANGE_PREPROCESS_MANIFEST_VERSION
  stages: RangePreprocessStageDescriptor[]
  fingerprint: string
}

export type RangePreprocessEvent =
  | { type: 'started'; fingerprint: string }
  | { type: 'stage-started'; stageId: string }
  | { type: 'stage-completed'; stageId: string }
  | { type: 'stage-failed'; stageId: string; error: RangePreprocessError }
  | { type: 'canceled'; stageId?: string }
  | { type: 'completed'; fingerprint: string }

export interface RangePreprocessRunOptions {
  signal?: AbortSignal
  /** Require every peer to execute the same ordered stage/version contract. */
  expectedFingerprint?: string
  /** Stronger persisted contract check, including every ordered descriptor. */
  expectedManifest?: RangePreprocessManifestV1
  onEvent?: (event: RangePreprocessEvent) => void
}

export interface RangePreprocessResult {
  grid: RangeValue[][]
  fingerprint: string
  revision: string
  stages: RangePreprocessStageDescriptor[]
}

export type RangePreprocessErrorCode =
  | 'ABORTED'
  | 'DUPLICATE_STAGE'
  | 'FINGERPRINT_MISMATCH'
  | 'INVALID_CONTEXT'
  | 'INVALID_GRID'
  | 'INVALID_MANIFEST'
  | 'INVALID_STAGE'
  | 'NON_DETERMINISTIC_STAGE'
  | 'STAGE_FAILED'

export class RangePreprocessError extends Error {
  constructor(
    public readonly code: RangePreprocessErrorCode,
    message: string,
    public readonly stageId?: string,
    options?: ErrorOptions,
  ) {
    super(message, options)
    this.name = 'RangePreprocessError'
  }
}

const STAGE_ID = /^[a-z0-9]+(?:[._-][a-z0-9]+)*$/

function descriptor(stage: RangePreprocessStage): RangePreprocessStageDescriptor {
  return { id: stage.id, version: stage.version, order: stage.order, deterministic: stage.deterministic }
}

function validDescriptor(value: unknown): value is RangePreprocessStageDescriptor {
  if (!plainObject(value) || !exactKeys(value, ['deterministic', 'id', 'order', 'version'])) return false
  return typeof value.id === 'string' && STAGE_ID.test(value.id) && value.id.length <= 128
    && typeof value.version === 'string' && value.version.length > 0 && value.version.length <= 128
    && Number.isSafeInteger(value.order) && typeof value.deterministic === 'boolean'
}

function validateStage(stage: RangePreprocessStage): void {
  if (!STAGE_ID.test(stage.id) || stage.id.length > 128) throw new RangePreprocessError('INVALID_STAGE', 'stage id must be a stable lowercase identifier')
  if (!stage.version || stage.version.length > 128) throw new RangePreprocessError('INVALID_STAGE', `stage ${stage.id} must declare a version`, stage.id)
  if (!Number.isSafeInteger(stage.order)) throw new RangePreprocessError('INVALID_STAGE', `stage ${stage.id} order must be a safe integer`, stage.id)
  if (typeof stage.deterministic !== 'boolean' || typeof stage.process !== 'function') throw new RangePreprocessError('INVALID_STAGE', `stage ${stage.id} has an invalid execution contract`, stage.id)
}

function validateContext(context: RangePreprocessContext): void {
  const { range } = context
  if (!range || typeof range.sheetId !== 'string' || !range.sheetId || !context.revision) throw new RangePreprocessError('INVALID_CONTEXT', 'sheetId and revision are required')
  for (const value of [range.startRow, range.startColumn, range.rowCount, range.columnCount]) {
    if (!Number.isSafeInteger(value) || value < 0) throw new RangePreprocessError('INVALID_CONTEXT', 'range coordinates and dimensions must be non-negative integers')
  }
  if ((range.rowCount === 0) !== (range.columnCount === 0)) throw new RangePreprocessError('INVALID_CONTEXT', 'an empty range must have zero rows and zero columns')
  if (range.startRow + range.rowCount > 1_048_576 || range.startColumn + range.columnCount > 16_384) throw new RangePreprocessError('INVALID_CONTEXT', 'range exceeds XLSX worksheet bounds')
}

function cloneAndValidateGrid(grid: RangeGrid, expected?: RangeWindow): RangeValue[][] {
  if (!Array.isArray(grid)) throw new RangePreprocessError('INVALID_GRID', 'processor output must be a two-dimensional array')
  const rows: RangeValue[][] = []
  let width: number | undefined
  for (const row of grid) {
    if (!Array.isArray(row)) throw new RangePreprocessError('INVALID_GRID', 'every grid row must be an array')
    if (width === undefined) width = row.length
    if (row.length !== width) throw new RangePreprocessError('INVALID_GRID', 'grid rows must be rectangular')
    const next: RangeValue[] = []
    for (const cell of row) {
      if (cell !== null && typeof cell !== 'string' && typeof cell !== 'boolean' && (typeof cell !== 'number' || !Number.isFinite(cell))) throw new RangePreprocessError('INVALID_GRID', 'cells must be finite numbers, strings, booleans, or null')
      next.push(cell)
    }
    rows.push(next)
  }
  const columns = width ?? 0
  if (expected && (rows.length !== expected.rowCount || columns !== expected.columnCount)) throw new RangePreprocessError('INVALID_GRID', 'grid dimensions must match the declared range')
  return rows
}

function readonlyGrid(grid: RangeValue[][]): RangeGrid {
  for (const row of grid) Object.freeze(row)
  return Object.freeze(grid)
}

function plainObject(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value) && Object.getPrototypeOf(value) === Object.prototype
}

function exactKeys(value: object, expected: readonly string[]): boolean {
  const keys = Object.keys(value).sort()
  return JSON.stringify(keys) === JSON.stringify([...expected].sort())
}

export function fingerprintRangePreprocessStages(stages: readonly RangePreprocessStageDescriptor[]): string {
  // v1 hashes UTF-8 bytes for the domain above followed by compact JSON whose
  // object keys are exactly id, version, order, deterministic in that order.
  return `sha256:${sha256Hex(RANGE_PREPROCESS_FINGERPRINT_DOMAIN + JSON.stringify(canonicalDescriptors(stages)))}`
}

function canonicalDescriptors(stages: readonly RangePreprocessStageDescriptor[]): RangePreprocessStageDescriptor[] {
  return stages.map(({ id, version, order, deterministic }) => ({ id, version, order, deterministic }))
}

function assertManifestMatches(value: unknown, current: RangePreprocessManifestV1): asserts value is RangePreprocessManifestV1 {
  if (!isRangePreprocessManifest(value)) throw new RangePreprocessError('INVALID_MANIFEST', 'preprocessing manifest is malformed or has an invalid fingerprint')
  if (value.fingerprint !== current.fingerprint
    || JSON.stringify(canonicalDescriptors(value.stages)) !== JSON.stringify(canonicalDescriptors(current.stages))) {
    throw new RangePreprocessError('FINGERPRINT_MISMATCH', `expected ${value.fingerprint}, received ${current.fingerprint}`)
  }
}

export function isRangePreprocessManifest(value: unknown): value is RangePreprocessManifestV1 {
  if (!plainObject(value) || !exactKeys(value, ['fingerprint', 'stages', 'version'])
    || value.version !== RANGE_PREPROCESS_MANIFEST_VERSION || !Array.isArray(value.stages) || value.stages.length > 256
    || !value.stages.every(validDescriptor) || !/^sha256:[0-9a-f]{64}$/.test(String(value.fingerprint))) return false
  const stages = value.stages as RangePreprocessStageDescriptor[]
  if (new Set(stages.map(({ id }) => id)).size !== stages.length) return false
  const ordered = [...stages].sort((left, right) => left.order - right.order || left.id.localeCompare(right.id))
  return ordered.every((stage, index) => stage === stages[index]) && value.fingerprint === fingerprintRangePreprocessStages(stages)
}

/** Host-configured, renderer-neutral preprocessing with deterministic peer contracts. */
export class RangePreprocessPipeline {
  private readonly stages = new Map<string, RangePreprocessStage>()

  private orderedStages(): RangePreprocessStage[] {
    return [...this.stages.values()].sort((left, right) => left.order - right.order || left.id.localeCompare(right.id))
  }

  register(stage: RangePreprocessStage): () => void {
    validateStage(stage)
    const stageId = stage.id
    if (this.stages.has(stageId)) throw new RangePreprocessError('DUPLICATE_STAGE', `stage ${stageId} is already registered`, stageId)
    const registered = Object.freeze({
      id: stageId,
      version: stage.version,
      order: stage.order,
      deterministic: stage.deterministic,
      process: stage.process,
    })
    this.stages.set(stageId, registered)
    return () => { if (this.stages.get(stageId) === registered) this.stages.delete(stageId) }
  }

  list(): RangePreprocessStageDescriptor[] {
    return this.orderedStages().map(descriptor)
  }

  fingerprint(): string {
    return fingerprintRangePreprocessStages(this.list())
  }

  manifest(): RangePreprocessManifestV1 {
    const stages = this.list()
    return { version: RANGE_PREPROCESS_MANIFEST_VERSION, stages, fingerprint: fingerprintRangePreprocessStages(stages) }
  }

  assertCompatibleManifest(value: unknown): asserts value is RangePreprocessManifestV1 {
    assertManifestMatches(value, this.manifest())
  }

  async run(grid: RangeGrid, context: RangePreprocessContext, options: RangePreprocessRunOptions = {}): Promise<RangePreprocessResult> {
    validateContext(context)
    // Snapshot both metadata and implementations before validating the peer
    // contract. Registry changes made by lifecycle callbacks affect only the
    // next run, never the code executing under this run's fingerprint.
    const registeredStages = this.orderedStages()
    const stages = registeredStages.map(descriptor)
    const fingerprint = fingerprintRangePreprocessStages(stages)
    if (options.expectedManifest !== undefined) assertManifestMatches(options.expectedManifest, { version: RANGE_PREPROCESS_MANIFEST_VERSION, stages, fingerprint })
    if (options.expectedFingerprint !== undefined && options.expectedFingerprint !== fingerprint) throw new RangePreprocessError('FINGERPRINT_MISMATCH', `expected ${options.expectedFingerprint}, received ${fingerprint}`)
    if (context.mode === 'collaborative') {
      const unsafe = stages.find((stage) => !stage.deterministic)
      if (unsafe) throw new RangePreprocessError('NON_DETERMINISTIC_STAGE', `stage ${unsafe.id} is not safe for collaborative execution`, unsafe.id)
    }
    let current = cloneAndValidateGrid(grid, context.range)
    const controller = new AbortController()
    const abort = () => controller.abort(options.signal?.reason)
    options.signal?.addEventListener('abort', abort, { once: true })
    if (options.signal?.aborted) abort()
    options.onEvent?.({ type: 'started', fingerprint })
    try {
      for (let index = 0; index < stages.length; index++) {
        const stageInfo = stages[index]!
        if (controller.signal.aborted) {
          options.onEvent?.({ type: 'canceled', stageId: stageInfo.id })
          throw new RangePreprocessError('ABORTED', 'range preprocessing was canceled', stageInfo.id)
        }
        const stage = registeredStages[index]!
        options.onEvent?.({ type: 'stage-started', stageId: stage.id })
        try {
          const stageContext = Object.freeze({
            ...context,
            range: Object.freeze({ ...context.range }),
            metadata: context.metadata ? Object.freeze({ ...context.metadata }) : undefined,
            signal: controller.signal,
          })
          const output = await stage.process(readonlyGrid(current), stageContext)
          if (controller.signal.aborted) {
            options.onEvent?.({ type: 'canceled', stageId: stage.id })
            throw new RangePreprocessError('ABORTED', 'range preprocessing was canceled', stage.id)
          }
          current = cloneAndValidateGrid(output, context.range)
          options.onEvent?.({ type: 'stage-completed', stageId: stage.id })
        } catch (cause) {
          if (cause instanceof RangePreprocessError && cause.code === 'ABORTED') throw cause
          if (controller.signal.aborted) {
            options.onEvent?.({ type: 'canceled', stageId: stage.id })
            throw new RangePreprocessError('ABORTED', 'range preprocessing was canceled', stage.id, { cause })
          }
          const error = new RangePreprocessError('STAGE_FAILED', `stage ${stage.id} failed`, stage.id, { cause })
          options.onEvent?.({ type: 'stage-failed', stageId: stage.id, error })
          throw error
        }
      }
      options.onEvent?.({ type: 'completed', fingerprint })
      return { grid: current, fingerprint, revision: context.revision, stages }
    } finally {
      options.signal?.removeEventListener('abort', abort)
    }
  }
}
