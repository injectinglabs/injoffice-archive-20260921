import type { ChartSpec } from './types'

export type ChartImageFormat = 'png' | 'jpeg' | 'svg'

export interface ChartImageExportOptions {
  format?: ChartImageFormat
  /** Renderer scale multiplier. The bounded public range is 0.1 through 8. */
  pixelRatio?: number
  backgroundColor?: string
}

export interface ChartImageRenderRequest {
  readonly jobId: string
  readonly chart: Readonly<ChartSpec>
  readonly format: ChartImageFormat
  readonly pixelRatio: number
  readonly backgroundColor?: string
}

export interface ChartImageRenderContext {
  readonly signal: AbortSignal
}

export interface ChartImageExportArtifact {
  readonly jobId: string
  readonly chartId: string
  readonly format: ChartImageFormat
  readonly mediaType: 'image/png' | 'image/jpeg' | 'image/svg+xml'
  readonly bytes: Uint8Array
  readonly width: number
  readonly height: number
}

export interface ChartImageExportHost {
  render(request: ChartImageRenderRequest, context: ChartImageRenderContext): Promise<Omit<ChartImageExportArtifact, 'jobId' | 'chartId' | 'format'>>
}

export type ChartImageExportErrorCode =
  | 'UNKNOWN_CHART'
  | 'INVALID_REQUEST'
  | 'INVALID_RESULT'
  | 'CANCELED'
  | 'TIMED_OUT'
  | 'HOST_FAILED'

export class ChartImageExportError extends Error {
  constructor(public readonly code: ChartImageExportErrorCode, message: string, options?: ErrorOptions) {
    super(message, options)
    this.name = 'ChartImageExportError'
  }
}

export type ChartImageExportEvent =
  | { type: 'started'; jobId: string; chartId: string; format: ChartImageFormat }
  | { type: 'completed'; jobId: string; chartId: string; format: ChartImageFormat; mediaType: ChartImageExportArtifact['mediaType']; byteLength: number; width: number; height: number }
  | { type: 'canceled' | 'timed-out' | 'failed'; jobId: string; chartId: string; format: ChartImageFormat; error: ChartImageExportError }

export interface ChartImageExportJob {
  readonly jobId: string
  readonly chartId: string
  readonly result: Promise<ChartImageExportArtifact>
  cancel(reason?: string): boolean
}

export interface ChartImageExportManagerOptions {
  timeoutMs?: number
  maxBytes?: number
  createJobId?: () => string
}

const MEDIA_TYPES: Record<ChartImageFormat, ChartImageExportArtifact['mediaType']> = {
  png: 'image/png',
  jpeg: 'image/jpeg',
  svg: 'image/svg+xml',
}

let exportSequence = 0

function cloneChart(chart: ChartSpec): ChartSpec {
  return structuredClone(chart)
}

function deepFreeze<T>(value: T, seen = new WeakSet<object>()): T {
  if (!value || typeof value !== 'object' || seen.has(value)) return value
  seen.add(value)
  for (const child of Object.values(value)) deepFreeze(child, seen)
  return Object.freeze(value)
}

function normalizeOptions(options: ChartImageExportOptions): Required<Pick<ChartImageExportOptions, 'format' | 'pixelRatio'>> & Pick<ChartImageExportOptions, 'backgroundColor'> {
  if (!options || typeof options !== 'object' || Array.isArray(options)) throw new ChartImageExportError('INVALID_REQUEST', 'options must be an object')
  const format = options.format ?? 'png'
  const pixelRatio = options.pixelRatio ?? 1
  if (!['png', 'jpeg', 'svg'].includes(format)) throw new ChartImageExportError('INVALID_REQUEST', 'format must be png, jpeg, or svg')
  if (!Number.isFinite(pixelRatio) || pixelRatio < 0.1 || pixelRatio > 8) throw new ChartImageExportError('INVALID_REQUEST', 'pixelRatio must be between 0.1 and 8')
  if (options.backgroundColor !== undefined && (typeof options.backgroundColor !== 'string' || options.backgroundColor.length > 128 || /[\u0000-\u001f\u007f]/.test(options.backgroundColor))) {
    throw new ChartImageExportError('INVALID_REQUEST', 'backgroundColor must be a bounded printable string')
  }
  return { format, pixelRatio, backgroundColor: options.backgroundColor }
}

function isAbortSignal(value: unknown): value is AbortSignal {
  if (!value || typeof value !== 'object') return false
  const signal = value as Partial<AbortSignal>
  return typeof signal.aborted === 'boolean' && typeof signal.addEventListener === 'function' && typeof signal.removeEventListener === 'function'
}

function validateArtifact(value: unknown, expected: ChartImageExportArtifact['mediaType'], maxBytes: number): asserts value is Omit<ChartImageExportArtifact, 'jobId' | 'chartId' | 'format'> {
  if (!value || typeof value !== 'object') throw new ChartImageExportError('INVALID_RESULT', 'host returned no image artifact')
  const artifact = value as Partial<ChartImageExportArtifact>
  if (artifact.mediaType !== expected) throw new ChartImageExportError('INVALID_RESULT', `host mediaType must be ${expected}`)
  if (!(artifact.bytes instanceof Uint8Array) || artifact.bytes.byteLength === 0 || artifact.bytes.byteLength > maxBytes) throw new ChartImageExportError('INVALID_RESULT', `host bytes must contain 1 to ${maxBytes} bytes`)
  if (!Number.isSafeInteger(artifact.width) || artifact.width! <= 0 || artifact.width! > 100_000) throw new ChartImageExportError('INVALID_RESULT', 'host width must be a positive bounded integer')
  if (!Number.isSafeInteger(artifact.height) || artifact.height! <= 0 || artifact.height! > 100_000) throw new ChartImageExportError('INVALID_RESULT', 'host height must be a positive bounded integer')
}

/** Cancellable host-neutral image export jobs. This class validates lifecycle
 * and artifacts; the injected host remains responsible for actual rendering. */
export class ChartImageExportManager {
  private readonly listeners = new Set<(event: ChartImageExportEvent) => void>()
  private readonly active = new Set<string>()
  private readonly timeoutMs: number
  private readonly maxBytes: number
  private readonly createJobId: () => string

  constructor(
    private readonly getChart: (id: string) => ChartSpec | undefined,
    private readonly host: ChartImageExportHost,
    options: ChartImageExportManagerOptions = {},
  ) {
    if (!host || typeof host.render !== 'function') throw new TypeError('host.render must be a function')
    if (options.createJobId !== undefined && typeof options.createJobId !== 'function') throw new TypeError('createJobId must be a function')
    this.timeoutMs = options.timeoutMs ?? 30_000
    this.maxBytes = options.maxBytes ?? 50 * 1024 * 1024
    this.createJobId = options.createJobId ?? (() => `chart-export-${Date.now().toString(36)}-${(++exportSequence).toString(36)}`)
    if (!Number.isSafeInteger(this.timeoutMs) || this.timeoutMs < 1 || this.timeoutMs > 300_000) throw new TypeError('timeoutMs must be an integer from 1 to 300000')
    if (!Number.isSafeInteger(this.maxBytes) || this.maxBytes < 1 || this.maxBytes > 250 * 1024 * 1024) throw new TypeError('maxBytes must be an integer from 1 to 262144000')
  }

  onEvent(listener: (event: ChartImageExportEvent) => void): () => void {
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }

  export(chartId: string, options: ChartImageExportOptions = {}, signal?: AbortSignal): Promise<ChartImageExportArtifact> {
    return this.start(chartId, options, signal).result
  }

  start(chartId: string, options: ChartImageExportOptions = {}, externalSignal?: AbortSignal): ChartImageExportJob {
    if (typeof chartId !== 'string' || !chartId) throw new ChartImageExportError('INVALID_REQUEST', 'chartId is required')
    const chart = this.getChart(chartId)
    if (!chart) throw new ChartImageExportError('UNKNOWN_CHART', `chart ${chartId} does not exist`)
    const normalized = normalizeOptions(options)
    if (externalSignal !== undefined && !isAbortSignal(externalSignal)) throw new ChartImageExportError('INVALID_REQUEST', 'signal must implement AbortSignal')
    let jobId: string
    try {
      jobId = this.createJobId()
    } catch (cause) {
      throw new ChartImageExportError('INVALID_REQUEST', 'createJobId failed', { cause })
    }
    if (typeof jobId !== 'string' || !jobId || jobId.length > 128 || this.active.has(jobId)) throw new ChartImageExportError('INVALID_REQUEST', 'createJobId returned an invalid or active id')
    const controller = new AbortController()
    let terminal = false
    let timeout: ReturnType<typeof setTimeout> | undefined
    let externalAbort: (() => void) | undefined
    const finish = () => {
      terminal = true
      this.active.delete(jobId)
      if (timeout) clearTimeout(timeout)
      if (externalAbort) externalSignal?.removeEventListener('abort', externalAbort)
    }
    const cancel = (reason?: string): boolean => {
      if (terminal || controller.signal.aborted) return false
      controller.abort(new ChartImageExportError('CANCELED', reason || 'chart image export canceled'))
      return true
    }
    let chartSnapshot: ChartSpec
    try {
      chartSnapshot = cloneChart(chart)
    } catch (cause) {
      throw new ChartImageExportError('INVALID_REQUEST', 'chart snapshot must be structured-cloneable', { cause })
    }
    const request: ChartImageRenderRequest = Object.freeze({
      jobId,
      chart: deepFreeze(chartSnapshot),
      format: normalized.format,
      pixelRatio: normalized.pixelRatio,
      ...(normalized.backgroundColor === undefined ? {} : { backgroundColor: normalized.backgroundColor }),
    })
    this.active.add(jobId)
    this.emit({ type: 'started', jobId, chartId, format: normalized.format })

    const aborted = new Promise<never>((_resolve, reject) => {
      controller.signal.addEventListener('abort', () => reject(controller.signal.reason), { once: true })
    })
    timeout = setTimeout(() => {
      if (!controller.signal.aborted) controller.abort(new ChartImageExportError('TIMED_OUT', `chart image export exceeded ${this.timeoutMs}ms`))
    }, this.timeoutMs)
    if (externalSignal) {
      externalAbort = () => cancel('chart image export canceled by caller')
      if (externalSignal.aborted) externalAbort()
      else externalSignal.addEventListener('abort', externalAbort, { once: true })
    }

    const result = Promise.race([
      Promise.resolve().then(() => this.host.render(request, { signal: controller.signal })),
      aborted,
    ]).then((value): ChartImageExportArtifact => {
      validateArtifact(value, MEDIA_TYPES[normalized.format], this.maxBytes)
      return Object.freeze({
        jobId,
        chartId,
        format: normalized.format,
        mediaType: value.mediaType,
        bytes: new Uint8Array(value.bytes),
        width: value.width,
        height: value.height,
      })
    }).then((artifact): ChartImageExportArtifact => {
      finish()
      this.emit({ type: 'completed', jobId, chartId, format: normalized.format, mediaType: artifact.mediaType, byteLength: artifact.bytes.byteLength, width: artifact.width, height: artifact.height })
      return artifact
    }).catch((cause: unknown) => {
      const error = cause instanceof ChartImageExportError
        ? cause
        : new ChartImageExportError('HOST_FAILED', 'chart image export host failed', { cause })
      finish()
      this.emit({ type: error.code === 'CANCELED' ? 'canceled' : error.code === 'TIMED_OUT' ? 'timed-out' : 'failed', jobId, chartId, format: normalized.format, error })
      throw error
    })
    return Object.freeze({ jobId, chartId, result, cancel })
  }

  private emit(event: ChartImageExportEvent): void {
    for (const listener of this.listeners) {
      try { listener(event) } catch { /* observer failures cannot change the job */ }
    }
  }
}
