import { describe, expect, it } from 'vitest'
import {
  NATIVE_WASM_WORKER_PROTOCOL,
  NATIVE_WASM_WORKER_VERSION,
  NativeWasmError,
  createNativeWasmClient,
  type NativeWasmWorker,
  type NativeWasmMessageEvent,
  type NativeWasmWorkerErrorEvent,
  type NativeWasmWorkerRequest,
  type NativeWasmWorkerResponse,
  type NativeWasmWorkerFatalNotification,
} from './index'

class FakeWorker implements NativeWasmWorker {
  readonly messages: Array<{ message: NativeWasmWorkerRequest; transfer: ArrayBuffer[] }> = []
  terminated = false
  private readonly messagesListeners = new Set<(event: NativeWasmMessageEvent) => void>()
  private readonly errorListeners = new Set<(event: NativeWasmWorkerErrorEvent) => void>()

  constructor(private readonly handler: (worker: FakeWorker, request: NativeWasmWorkerRequest) => void) {}

  postMessage(message: unknown, transfer: ArrayBuffer[]): void {
    const request = message as NativeWasmWorkerRequest
    this.messages.push({ message: request, transfer })
    this.handler(this, request)
  }

  terminate(): void {
    this.terminated = true
  }

  addEventListener(type: 'message' | 'error', listener: ((event: NativeWasmMessageEvent) => void) | ((event: NativeWasmWorkerErrorEvent) => void)): void {
    if (type === 'message') this.messagesListeners.add(listener as (event: NativeWasmMessageEvent) => void)
    else this.errorListeners.add(listener as (event: NativeWasmWorkerErrorEvent) => void)
  }

  removeEventListener(type: 'message' | 'error', listener: ((event: NativeWasmMessageEvent) => void) | ((event: NativeWasmWorkerErrorEvent) => void)): void {
    if (type === 'message') this.messagesListeners.delete(listener as (event: NativeWasmMessageEvent) => void)
    else this.errorListeners.delete(listener as (event: NativeWasmWorkerErrorEvent) => void)
  }

  respond(response: NativeWasmWorkerResponse | unknown): void {
    queueMicrotask(() => {
      for (const listener of this.messagesListeners) listener({ data: response })
    })
  }

  fail(notification: NativeWasmWorkerFatalNotification): void {
    this.respond(notification)
  }
}

const success = (request: NativeWasmWorkerRequest, result?: unknown): NativeWasmWorkerResponse => ({
  protocol: NATIVE_WASM_WORKER_PROTOCOL,
  version: NATIVE_WASM_WORKER_VERSION,
  id: request.id,
  format: request.format,
  op: request.op,
  ok: true,
  ...(request.op === 'init' ? {} : { result }),
} as NativeWasmWorkerResponse)

const createWorker = (handler?: (worker: FakeWorker, request: NativeWasmWorkerRequest) => void) =>
  new FakeWorker(handler ?? ((worker, request) => {
    if (request.op === 'init') worker.respond(success(request))
    else if (request.op === 'extract') worker.respond(success(request, { contractJson: '{"version":2}' }))
    else worker.respond(success(request, { bytes: new Uint8Array([7, 8, 9]).buffer }))
  }))

const options = (factory: () => NativeWasmWorker, operationTimeoutMs = 100) => ({
  format: 'xlsx' as const,
  workerFactory: factory,
  assets: { wasmUrl: 'https://example.test/xlsx.wasm', goRuntimeUrl: 'https://example.test/wasm_exec.js' },
  operationTimeoutMs,
})

if (false) {
  const browserWorker: NativeWasmWorker = new Worker('worker.js')
  void browserWorker
}

describe('native WASM client', () => {
  it('initializes once, serializes calls, and never transfers caller-owned buffers', async () => {
    const worker = createWorker()
    const original = new Uint8Array([1, 2, 3])
    const payload = new Uint8Array([4, 5])
    const client = createNativeWasmClient(options(() => worker))

    await expect(client.extract(original)).resolves.toBe('{"version":2}')
    await expect(client.apply(original, payload, 'sha256:test')).resolves.toEqual(new Uint8Array([7, 8, 9]))

    expect(original).toEqual(new Uint8Array([1, 2, 3]))
    expect(payload).toEqual(new Uint8Array([4, 5]))
    expect(worker.messages.map(({ message }) => message.op)).toEqual(['init', 'extract', 'apply'])
    const extract = worker.messages[1]
    const apply = worker.messages[2]
    expect(extract.transfer).toEqual([(extract.message as Extract<NativeWasmWorkerRequest, { op: 'extract' }>).bytes])
    expect(apply.transfer).toHaveLength(2)
    expect((apply.message as Extract<NativeWasmWorkerRequest, { op: 'apply' }>).original).not.toBe(original.buffer)
    expect((apply.message as Extract<NativeWasmWorkerRequest, { op: 'apply' }>).payload).not.toBe(payload.buffer)
  })

  it('snapshots extraction bytes when the operation is called', async () => {
    const worker = createWorker()
    const bytes = new Uint8Array([1, 2, 3])
    const client = createNativeWasmClient(options(() => worker))

    const extraction = client.extract(bytes)
    bytes.fill(9)

    await expect(extraction).resolves.toBe('{"version":2}')
    const request = worker.messages.find(({ message }) => message.op === 'extract')?.message
    expect(request?.op).toBe('extract')
    expect(new Uint8Array((request as Extract<NativeWasmWorkerRequest, { op: 'extract' }>).bytes)).toEqual(new Uint8Array([1, 2, 3]))
  })

  it('snapshots apply inputs while the operation waits in the queue', async () => {
    let releaseExtract: ((response: NativeWasmWorkerResponse) => void) | undefined
    let markExtractStarted: (() => void) | undefined
    const extractStarted = new Promise<void>((resolve) => {
      markExtractStarted = resolve
    })
    const worker = createWorker((current, request) => {
      if (request.op === 'init') current.respond(success(request))
      else if (request.op === 'extract') {
        releaseExtract = (response) => current.respond(response)
        markExtractStarted?.()
      } else current.respond(success(request, { bytes: new Uint8Array([7, 8, 9]).buffer }))
    })
    const client = createNativeWasmClient(options(() => worker))
    const active = client.extract(new Uint8Array([1]))
    await extractStarted

    const original = new Uint8Array([2, 3, 4])
    const payload = new Uint8Array([5, 6])
    const queued = client.apply(original, payload, 'sha256:test')
    original.fill(8)
    payload.fill(9)
    const extractRequest = worker.messages.find(({ message }) => message.op === 'extract')?.message
    releaseExtract?.(success(extractRequest as Extract<NativeWasmWorkerRequest, { op: 'extract' }>, { contractJson: '{"version":2}' }))

    await expect(active).resolves.toBe('{"version":2}')
    await expect(queued).resolves.toEqual(new Uint8Array([7, 8, 9]))
    const request = worker.messages.find(({ message }) => message.op === 'apply')?.message as Extract<NativeWasmWorkerRequest, { op: 'apply' }>
    expect(new Uint8Array(request.original)).toEqual(new Uint8Array([2, 3, 4]))
    expect(new Uint8Array(request.payload as ArrayBuffer)).toEqual(new Uint8Array([5, 6]))
  })

  it('preserves bounded native refusals without killing a healthy worker', async () => {
    const worker = createWorker((current, request) => {
      if (request.op === 'init') current.respond(success(request))
      else current.respond({
        protocol: NATIVE_WASM_WORKER_PROTOCOL,
        version: NATIVE_WASM_WORKER_VERSION,
        id: request.id,
        format: request.format,
        op: request.op,
        ok: false,
        error: { code: 'REFUSED', message: 'x'.repeat(5_000), fatal: false },
      })
    })
    const client = createNativeWasmClient(options(() => worker))

    const error = await client.extract(new Uint8Array([1])).catch((reason: unknown) => reason)
    expect(error).toBeInstanceOf(NativeWasmError)
    expect(error).toMatchObject({ code: 'REFUSED', fatal: false })
    expect((error as Error).message).toHaveLength(4_096)
    expect(worker.terminated).toBe(false)
  })

  it('terminates a timed-out worker and lazily starts a fresh one', async () => {
    const workers: FakeWorker[] = []
    const client = createNativeWasmClient(options(() => {
      const worker = workers.length === 0
        ? createWorker(() => {})
        : createWorker()
      workers.push(worker)
      return worker
    }, 10))

    await expect(client.extract(new Uint8Array([1]))).rejects.toMatchObject({ code: 'TIMEOUT', fatal: true })
    expect(workers[0].terminated).toBe(true)
    await expect(client.extract(new Uint8Array([2]))).resolves.toBe('{"version":2}')
    expect(workers).toHaveLength(2)
  })

  it('aborts the active operation and restarts on the next call', async () => {
    const workers: FakeWorker[] = []
    const controller = new AbortController()
    const client = createNativeWasmClient(options(() => {
      const worker = workers.length === 0 ? createWorker(() => {}) : createWorker()
      workers.push(worker)
      return worker
    }))
    const pending = client.extract(new Uint8Array([1]), { signal: controller.signal })
    await Promise.resolve()
    controller.abort()

    await expect(pending).rejects.toMatchObject({ name: 'AbortError' })
    expect(workers[0].terminated).toBe(true)
    await expect(client.extract(new Uint8Array([2]))).resolves.toBe('{"version":2}')
  })

  it('ignores stale messages and terminates on a malformed matching response', async () => {
    let initialized: NativeWasmWorkerRequest | undefined
    const worker = createWorker((current, request) => {
      if (request.op === 'init') {
        initialized = request
        current.respond(success(request))
        return
      }
      if (initialized) current.respond(success(initialized))
      current.respond({
        protocol: NATIVE_WASM_WORKER_PROTOCOL,
        version: NATIVE_WASM_WORKER_VERSION,
        id: request.id,
        format: request.format,
        op: request.op,
        ok: true,
        unexpected: true,
      })
    })
    const client = createNativeWasmClient(options(() => worker))

    await expect(client.extract(new Uint8Array([1]))).rejects.toMatchObject({ code: 'MALFORMED_RESPONSE', fatal: true })
    expect(worker.terminated).toBe(true)
  })

  it('fails fast on an unknown response id', async () => {
    const worker = createWorker((current, request) => {
      current.respond({
        protocol: NATIVE_WASM_WORKER_PROTOCOL,
        version: NATIVE_WASM_WORKER_VERSION,
        id: 'unknown-request',
        format: request.format,
        op: request.op,
        ok: true,
      })
    })
    const client = createNativeWasmClient(options(() => worker, 1_000))

    await expect(client.extract(new Uint8Array([1]))).rejects.toMatchObject({ code: 'RESPONSE_MISMATCH', fatal: true })
    expect(worker.terminated).toBe(true)
  })

  it('rejects an aborted queued call without waiting for the active timeout', async () => {
    const worker = createWorker((current, request) => {
      if (request.op === 'init') current.respond(success(request))
    })
    const client = createNativeWasmClient(options(() => worker, 1_000))
    const active = client.extract(new Uint8Array([1]))
    await Promise.resolve()
    await Promise.resolve()
    await Promise.resolve()
    const controller = new AbortController()
    const queued = client.extract(new Uint8Array([2]), { signal: controller.signal })
    controller.abort()

    await expect(queued).rejects.toMatchObject({ name: 'AbortError' })
    expect(worker.terminated).toBe(false)
    expect(worker.messages.map(({ message }) => message.op)).toEqual(['init', 'extract'])
    client.terminate()
    await expect(active).rejects.toMatchObject({ code: 'TERMINATED' })
  })

  it('permanently rejects calls after explicit termination', async () => {
    const worker = createWorker()
    const client = createNativeWasmClient(options(() => worker))
    await client.extract(new Uint8Array([1]))
    client.terminate()

    expect(worker.terminated).toBe(true)
    await expect(client.extract(new Uint8Array([2]))).rejects.toMatchObject({ code: 'TERMINATED' })
  })

  it('discards an initialized worker when its native runtime exits while idle', async () => {
    const workers: FakeWorker[] = []
    const client = createNativeWasmClient(options(() => {
      const worker = createWorker()
      workers.push(worker)
      return worker
    }))
    await client.extract(new Uint8Array([1]))

    workers[0].fail({
      protocol: NATIVE_WASM_WORKER_PROTOCOL,
      version: NATIVE_WASM_WORKER_VERSION,
      format: 'xlsx',
      op: 'fatal',
      error: { code: 'ENGINE_EXITED', message: 'Go runtime exited.', fatal: true },
    })
    await Promise.resolve()

    expect(workers[0].terminated).toBe(true)
    await expect(client.extract(new Uint8Array([2]))).resolves.toBe('{"version":2}')
    expect(workers).toHaveLength(2)
  })

  it('rejects active work and restarts queued work when the native runtime exits', async () => {
    const workers: FakeWorker[] = []
    const client = createNativeWasmClient(options(() => {
      const worker = workers.length === 0
        ? createWorker((current, request) => {
            if (request.op === 'init') current.respond(success(request))
          })
        : createWorker()
      workers.push(worker)
      return worker
    }))
    const active = client.extract(new Uint8Array([1]))
    const queued = client.extract(new Uint8Array([2]))
    await Promise.resolve()
    await Promise.resolve()

    workers[0].fail({
      protocol: NATIVE_WASM_WORKER_PROTOCOL,
      version: NATIVE_WASM_WORKER_VERSION,
      format: 'xlsx',
      op: 'fatal',
      error: { code: 'ENGINE_EXITED', message: 'Go runtime exited.', fatal: true },
    })

    await expect(active).rejects.toMatchObject({ code: 'ENGINE_EXITED', message: 'Go runtime exited.', fatal: true })
    await expect(queued).resolves.toBe('{"version":2}')
    expect(workers[0].terminated).toBe(true)
    expect(workers).toHaveLength(2)
  })

  it('terminates on a malformed fatal notification', async () => {
    const worker = createWorker()
    const client = createNativeWasmClient(options(() => worker))
    await client.extract(new Uint8Array([1]))

    worker.respond({
      protocol: NATIVE_WASM_WORKER_PROTOCOL,
      version: NATIVE_WASM_WORKER_VERSION,
      format: 'docx',
      op: 'fatal',
      error: { code: 'ENGINE_EXITED', message: 'wrong format', fatal: true },
    })
    await Promise.resolve()

    expect(worker.terminated).toBe(true)
  })

  it('does not start queued work after termination', async () => {
    const worker = createWorker((current, request) => {
      if (request.op === 'init') current.respond(success(request))
    })
    const client = createNativeWasmClient(options(() => worker))
    const active = client.extract(new Uint8Array([1]))
    const queued = client.extract(new Uint8Array([2]))
    await Promise.resolve()
    client.terminate()

    await expect(active).rejects.toMatchObject({ code: 'TERMINATED' })
    await expect(queued).rejects.toMatchObject({ code: 'TERMINATED' })
    expect(worker.messages.map(({ message }) => message.op)).toEqual(['init'])
  })
})
