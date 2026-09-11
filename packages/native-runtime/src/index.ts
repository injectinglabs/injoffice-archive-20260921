export const NATIVE_WASM_WORKER_PROTOCOL = 'injoffice.native-wasm-worker' as const
export const NATIVE_WASM_WORKER_VERSION = 1 as const

export type NativeWasmFormat = 'xlsx' | 'docx' | 'pptx'
export type NativeWasmOperation = 'init' | 'extract' | 'inspect' | 'apply'

export interface NativeWasmAssets {
  wasmUrl: string
  goRuntimeUrl: string
}

export interface NativeWasmMessageEvent {
  data: unknown
}

export interface NativeWasmWorkerErrorEvent {
  message?: string
}

export interface NativeWasmAbortSignal {
  readonly aborted: boolean
  readonly reason?: unknown
  addEventListener(type: 'abort', listener: () => void, options?: { once?: boolean }): void
  removeEventListener(type: 'abort', listener: () => void): void
}

export interface NativeWasmWorker {
  postMessage(message: unknown, transfer: ArrayBuffer[]): void
  terminate(): void
  addEventListener(type: 'message', listener: (event: NativeWasmMessageEvent) => void): void
  addEventListener(type: 'error', listener: (event: NativeWasmWorkerErrorEvent) => void): void
  removeEventListener(type: 'message', listener: (event: NativeWasmMessageEvent) => void): void
  removeEventListener(type: 'error', listener: (event: NativeWasmWorkerErrorEvent) => void): void
}

export interface NativeWasmClientOptions {
  format: NativeWasmFormat
  workerFactory: () => NativeWasmWorker
  assets: NativeWasmAssets
  operationTimeoutMs?: number
}

export interface NativeWasmOperationOptions {
  signal?: NativeWasmAbortSignal
}

export interface NativeWasmClient {
  extract(bytes: Uint8Array, options?: NativeWasmOperationOptions): Promise<string>
  /** Optional read-only supplemental projection; unsupported engines may refuse. */
  inspect(bytes: Uint8Array, options?: NativeWasmOperationOptions): Promise<string>
  apply(
    original: Uint8Array,
    payload: string | Uint8Array,
    expectedRevision: string,
    options?: NativeWasmOperationOptions,
  ): Promise<Uint8Array>
  terminate(): void
}

export class NativeWasmError extends Error {
  readonly code: string
  readonly fatal: boolean

  constructor(code: string, message: string, fatal = false) {
    super(message)
    this.name = 'NativeWasmError'
    this.code = code
    this.fatal = fatal
  }
}

type RequestBase = {
  protocol: typeof NATIVE_WASM_WORKER_PROTOCOL
  version: typeof NATIVE_WASM_WORKER_VERSION
  id: string
  format: NativeWasmFormat
  op: NativeWasmOperation
}

export type NativeWasmWorkerRequest =
  | (RequestBase & { op: 'init'; assets: NativeWasmAssets })
  | (RequestBase & { op: 'extract'; bytes: ArrayBuffer })
  | (RequestBase & { op: 'inspect'; bytes: ArrayBuffer })
  | (RequestBase & {
    op: 'apply'
    original: ArrayBuffer
    payload: string | ArrayBuffer
    expectedRevision: string
  })

type ResponseBase = RequestBase & { ok: boolean }

export type NativeWasmWorkerResponse =
  | (ResponseBase & { op: 'init'; ok: true })
  | (ResponseBase & { op: 'extract'; ok: true; result: { contractJson: string } })
  | (ResponseBase & { op: 'inspect'; ok: true; result: { contractJson: string } })
  | (ResponseBase & { op: 'apply'; ok: true; result: { bytes: ArrayBuffer } })
  | (ResponseBase & { ok: false; error: { code: string; message: string; fatal: boolean } })

/** Unsolicited notification used when an initialized native runtime exits. */
export interface NativeWasmWorkerFatalNotification {
  protocol: typeof NATIVE_WASM_WORKER_PROTOCOL
  version: typeof NATIVE_WASM_WORKER_VERSION
  format: NativeWasmFormat
  op: 'fatal'
  error: { code: string; message: string; fatal: true }
}

const DEFAULT_OPERATION_TIMEOUT_MS = 30_000
const MAX_ERROR_LENGTH = 4_096

type Pending = {
  id: string
  op: NativeWasmOperation
  resolve: (response: NativeWasmWorkerResponse) => void
  reject: (error: Error) => void
  timer: ReturnType<typeof setTimeout>
  signal?: NativeWasmAbortSignal
  abort?: () => void
}

const isRecord = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value)

const hasExactKeys = (value: Record<string, unknown>, expected: readonly string[]) => {
  const actual = Object.keys(value).sort()
  const wanted = [...expected].sort()
  return actual.length === wanted.length && actual.every((key, index) => key === wanted[index])
}

const boundedMessage = (value: unknown, fallback: string) => {
  const message = typeof value === 'string' && value.length > 0 ? value : fallback
  return message.slice(0, MAX_ERROR_LENGTH)
}

function validateResponse(
  value: unknown,
  expected: Pick<RequestBase, 'id' | 'format' | 'op'>,
): NativeWasmWorkerResponse {
  if (!isRecord(value)) throw new NativeWasmError('MALFORMED_RESPONSE', 'Native worker returned a non-object response.', true)
  const common = value.ok === true
    ? expected.op === 'init'
      ? ['protocol', 'version', 'id', 'format', 'op', 'ok']
      : ['protocol', 'version', 'id', 'format', 'op', 'ok', 'result']
    : ['protocol', 'version', 'id', 'format', 'op', 'ok', 'error']
  if (!hasExactKeys(value, common)) throw new NativeWasmError('MALFORMED_RESPONSE', 'Native worker response keys are invalid.', true)
  if (value.protocol !== NATIVE_WASM_WORKER_PROTOCOL || value.version !== NATIVE_WASM_WORKER_VERSION) {
    throw new NativeWasmError('PROTOCOL_MISMATCH', 'Native worker protocol or version does not match.', true)
  }
  if (value.id !== expected.id || value.format !== expected.format || value.op !== expected.op) {
    throw new NativeWasmError('RESPONSE_MISMATCH', 'Native worker response identity does not match its request.', true)
  }
  if (value.ok === false) {
    if (!isRecord(value.error) || !hasExactKeys(value.error, ['code', 'message', 'fatal'])) {
      throw new NativeWasmError('MALFORMED_RESPONSE', 'Native worker returned an invalid error.', true)
    }
    if (typeof value.error.code !== 'string' || typeof value.error.message !== 'string' || typeof value.error.fatal !== 'boolean') {
      throw new NativeWasmError('MALFORMED_RESPONSE', 'Native worker error fields are invalid.', true)
    }
    return value as NativeWasmWorkerResponse
  }
  if (value.ok !== true) throw new NativeWasmError('MALFORMED_RESPONSE', 'Native worker response has an invalid status.', true)
  if (expected.op === 'extract' || expected.op === 'inspect') {
    if (!isRecord(value.result) || !hasExactKeys(value.result, ['contractJson']) || typeof value.result.contractJson !== 'string') {
      throw new NativeWasmError('MALFORMED_RESPONSE', 'Native worker returned invalid extraction JSON.', true)
    }
  } else if (expected.op === 'apply') {
    if (!isRecord(value.result) || !hasExactKeys(value.result, ['bytes']) || !(value.result.bytes instanceof ArrayBuffer)) {
      throw new NativeWasmError('MALFORMED_RESPONSE', 'Native worker returned invalid package bytes.', true)
    }
  }
  return value as NativeWasmWorkerResponse
}

function validateFatalNotification(value: unknown, format: NativeWasmFormat): NativeWasmWorkerFatalNotification {
  if (!isRecord(value) || !hasExactKeys(value, ['protocol', 'version', 'format', 'op', 'error'])) {
    throw new NativeWasmError('MALFORMED_RESPONSE', 'Native worker fatal notification keys are invalid.', true)
  }
  if (value.protocol !== NATIVE_WASM_WORKER_PROTOCOL || value.version !== NATIVE_WASM_WORKER_VERSION || value.format !== format || value.op !== 'fatal') {
    throw new NativeWasmError('PROTOCOL_MISMATCH', 'Native worker fatal notification identity does not match.', true)
  }
  if (!isRecord(value.error) || !hasExactKeys(value.error, ['code', 'message', 'fatal']) ||
      typeof value.error.code !== 'string' || typeof value.error.message !== 'string' || value.error.fatal !== true) {
    throw new NativeWasmError('MALFORMED_RESPONSE', 'Native worker fatal notification error is invalid.', true)
  }
  return value as unknown as NativeWasmWorkerFatalNotification
}

class NativeWasmClientImpl implements NativeWasmClient {
  private worker?: NativeWasmWorker
  private pending?: Pending
  private initialized = false
  private disposed = false
  private sequence = 0
  private resetError?: Error
  private readonly settledIds = new Set<string>()
  private queue: Promise<void> = Promise.resolve()
  private readonly timeoutMs: number

  constructor(private readonly options: NativeWasmClientOptions) {
    const timeout = options.operationTimeoutMs ?? DEFAULT_OPERATION_TIMEOUT_MS
    if (!Number.isFinite(timeout) || timeout <= 0) throw new RangeError('operationTimeoutMs must be a positive finite number')
    if (!options.assets.wasmUrl || !options.assets.goRuntimeUrl) throw new TypeError('wasmUrl and goRuntimeUrl are required')
    this.timeoutMs = timeout
  }

  async extract(bytes: Uint8Array, options: NativeWasmOperationOptions = {}): Promise<string> {
    return this.readJSON('extract', bytes, options)
  }

  async inspect(bytes: Uint8Array, options: NativeWasmOperationOptions = {}): Promise<string> {
    return this.readJSON('inspect', bytes, options)
  }

  private async readJSON(op: 'extract' | 'inspect', bytes: Uint8Array, options: NativeWasmOperationOptions): Promise<string> {
    if (this.disposed) throw new NativeWasmError('TERMINATED', 'Native WASM client was terminated.', true)
    if (options.signal?.aborted) throw abortError(options.signal)
    const copy = copyBytes(bytes)
    return this.enqueue(async () => {
      await this.ensureInitialized(options.signal)
      const response = await this.request({ op, bytes: copy.buffer }, [copy.buffer], options.signal)
      if (response.ok !== true || response.op !== op) throw new NativeWasmError('MALFORMED_RESPONSE', 'Native worker returned the wrong read-only response.', true)
      return response.result.contractJson
    }, options.signal)
  }

  async apply(
    original: Uint8Array,
    payload: string | Uint8Array,
    expectedRevision: string,
    options: NativeWasmOperationOptions = {},
  ): Promise<Uint8Array> {
    if (!expectedRevision) throw new TypeError('expectedRevision is required')
    if (this.disposed) throw new NativeWasmError('TERMINATED', 'Native WASM client was terminated.', true)
    if (options.signal?.aborted) throw abortError(options.signal)
    const originalCopy = copyBytes(original)
    let payloadValue: string | ArrayBuffer
    if (typeof payload === 'string') {
      payloadValue = payload
    } else {
      payloadValue = copyBytes(payload).buffer
    }
    return this.enqueue(async () => {
      await this.ensureInitialized(options.signal)
      const transfers: ArrayBuffer[] = [originalCopy.buffer]
      if (payloadValue instanceof ArrayBuffer) transfers.push(payloadValue)
      const response = await this.request({
        op: 'apply',
        original: originalCopy.buffer,
        payload: payloadValue,
        expectedRevision,
      }, transfers, options.signal)
      if (response.ok !== true || response.op !== 'apply') throw new NativeWasmError('MALFORMED_RESPONSE', 'Native worker returned the wrong apply response.', true)
      return new Uint8Array(response.result.bytes)
    }, options.signal)
  }

  terminate(): void {
    this.disposed = true
    this.reset(new NativeWasmError('TERMINATED', 'Native WASM client was terminated.', true))
  }

  private enqueue<T>(operation: () => Promise<T>, signal?: NativeWasmAbortSignal): Promise<T> {
    if (this.disposed) return Promise.reject(new NativeWasmError('TERMINATED', 'Native WASM client was terminated.', true))
    if (signal?.aborted) return Promise.reject(abortError(signal))
    const scheduled = this.queue.then(async () => {
      if (signal?.aborted) throw abortError(signal)
      return operation()
    }, async () => {
      if (signal?.aborted) throw abortError(signal)
      return operation()
    })
    this.queue = scheduled.then(() => undefined, () => undefined)
    if (!signal) return scheduled
    let abort: (() => void) | undefined
    const cancellation = new Promise<T>((_resolve, reject) => {
      abort = () => reject(abortError(signal))
      signal.addEventListener('abort', abort, { once: true })
    })
    return Promise.race([scheduled, cancellation]).finally(() => {
      if (abort) signal.removeEventListener('abort', abort)
    })
  }

  private async ensureInitialized(signal?: NativeWasmAbortSignal): Promise<void> {
    if (this.disposed) throw new NativeWasmError('TERMINATED', 'Native WASM client was terminated.', true)
    if (this.initialized) return
    if (signal?.aborted) throw abortError(signal)
    if (!this.worker) {
      this.resetError = undefined
      this.worker = this.options.workerFactory()
      this.worker.addEventListener('message', this.onMessage)
      this.worker.addEventListener('error', this.onError)
    }
    const initializingWorker = this.worker
    const response = await this.request({ op: 'init', assets: { ...this.options.assets } }, [], signal)
    if (this.worker !== initializingWorker) {
      throw this.resetError ?? new NativeWasmError('WORKER_UNAVAILABLE', 'Native worker exited during initialization.', true)
    }
    if (response.ok !== true || response.op !== 'init') throw new NativeWasmError('MALFORMED_RESPONSE', 'Native worker did not initialize.', true)
    this.initialized = true
  }

  private request(
    body: Omit<NativeWasmWorkerRequest, keyof RequestBase>,
    transfer: ArrayBuffer[],
    signal?: NativeWasmAbortSignal,
  ): Promise<NativeWasmWorkerResponse> {
    if (this.disposed) return Promise.reject(new NativeWasmError('TERMINATED', 'Native WASM client was terminated.', true))
    if (!this.worker) {
      return Promise.reject(this.resetError ?? new NativeWasmError('WORKER_UNAVAILABLE', 'Native worker is unavailable.', true))
    }
    if (signal?.aborted) return Promise.reject(abortError(signal))
    const id = `request-${++this.sequence}`
    const request = {
      protocol: NATIVE_WASM_WORKER_PROTOCOL,
      version: NATIVE_WASM_WORKER_VERSION,
      id,
      format: this.options.format,
      ...body,
    } as NativeWasmWorkerRequest
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        const error = new NativeWasmError('TIMEOUT', `Native ${request.op} exceeded ${this.timeoutMs}ms.`, true)
        this.reset(error)
      }, this.timeoutMs)
      const pending: Pending = { id, op: request.op, resolve, reject, timer, signal }
      if (signal) {
        pending.abort = () => this.reset(abortError(signal))
        signal.addEventListener('abort', pending.abort, { once: true })
      }
      this.pending = pending
      try {
        this.worker?.postMessage(request, transfer)
      } catch (reason) {
        this.reset(new NativeWasmError('POST_FAILED', boundedMessage(reason instanceof Error ? reason.message : reason, 'Could not post to native worker.'), true))
      }
    })
  }

  private readonly onMessage = (event: NativeWasmMessageEvent) => {
    if (isRecord(event.data) && event.data.op === 'fatal') {
      try {
        const notification = validateFatalNotification(event.data, this.options.format)
        this.reset(new NativeWasmError(
          notification.error.code,
          boundedMessage(notification.error.message, 'Native runtime exited.'),
          true,
        ))
      } catch (reason) {
        this.reset(reason instanceof NativeWasmError
          ? reason
          : new NativeWasmError('MALFORMED_RESPONSE', boundedMessage(reason, 'Native worker fatal notification was invalid.'), true))
      }
      return
    }
    const pending = this.pending
    if (!pending) return
    if (isRecord(event.data) && typeof event.data.id === 'string' && event.data.id !== pending.id && this.settledIds.has(event.data.id)) return
    try {
      const response = validateResponse(event.data, { id: pending.id, format: this.options.format, op: pending.op })
      this.rememberSettled(pending.id)
      this.clearPending()
      if (response.ok === false) {
        const error = new NativeWasmError(response.error.code, boundedMessage(response.error.message, 'Native worker refused the operation.'), response.error.fatal)
        if (error.fatal) this.reset()
        pending.reject(error)
        return
      }
      pending.resolve(response)
    } catch (reason) {
      const error = reason instanceof NativeWasmError
        ? reason
        : new NativeWasmError('MALFORMED_RESPONSE', boundedMessage(reason, 'Native worker response was invalid.'), true)
      this.clearPending()
      this.reset()
      pending.reject(error)
    }
  }

  private readonly onError = (event: NativeWasmWorkerErrorEvent) => {
    const error = new NativeWasmError('WORKER_ERROR', boundedMessage(event.message, 'Native worker failed.'), true)
    this.reset(error)
  }

  private clearPending(): void {
    if (!this.pending) return
    clearTimeout(this.pending.timer)
    if (this.pending.signal && this.pending.abort) this.pending.signal.removeEventListener('abort', this.pending.abort)
    this.pending = undefined
  }

  private rememberSettled(id: string): void {
    this.settledIds.add(id)
    if (this.settledIds.size > 32) this.settledIds.delete(this.settledIds.values().next().value as string)
  }

  private reset(error?: Error): void {
    this.resetError = error
    const pending = this.pending
    this.clearPending()
    if (this.worker) {
      this.worker.removeEventListener('message', this.onMessage)
      this.worker.removeEventListener('error', this.onError)
      this.worker.terminate()
    }
    this.worker = undefined
    this.initialized = false
    if (pending && error) pending.reject(error)
  }
}

function copyBytes(bytes: Uint8Array): Uint8Array<ArrayBuffer> {
  const copy = new Uint8Array(bytes.byteLength)
  copy.set(bytes)
  return copy
}

function abortError(signal: NativeWasmAbortSignal): Error {
  if (signal.reason instanceof Error) return signal.reason
  return new DOMException('Native WASM operation was aborted.', 'AbortError')
}

export function createNativeWasmClient(options: NativeWasmClientOptions): NativeWasmClient {
  return new NativeWasmClientImpl(options)
}
