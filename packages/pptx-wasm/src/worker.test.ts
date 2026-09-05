import { readFileSync } from 'node:fs'
import { createContext, runInContext } from 'node:vm'
import { describe, expect, it } from 'vitest'

const workerSource = readFileSync(new URL('../worker/pptxnative.worker.js', import.meta.url), 'utf8')

type Binding = {
  extract(bytes: Uint8Array): unknown
  apply(original: Uint8Array, payload: string | Uint8Array, expectedRevision: string): unknown
}

type WorkerMessage = { op?: string; ok?: boolean; error?: { code?: string; message?: string; fatal?: boolean }; result?: { contractJson?: string } }

function createWorkerHarness(binding: Binding, runResult?: () => Promise<unknown>) {
  const messages: WorkerMessage[] = []
  let closed = false
  const globals: Record<string, unknown> = {
    ArrayBuffer, Error, Promise, Uint8Array, clearTimeout, queueMicrotask, setTimeout,
    postMessage: (message: WorkerMessage) => messages.push(message),
    importScripts: () => undefined,
    fetch: async () => ({ ok: true, arrayBuffer: async () => new ArrayBuffer(0) }),
    WebAssembly: { instantiate: async () => ({ instance: {} }) },
  }
  globals.self = globals
  globals.close = () => { closed = true }
  globals.Go = class {
    readonly importObject = {}
    run(): Promise<unknown> {
      globals.pptxnative = binding
      queueMicrotask(() => (globals.pptxnativeOnReady as () => void)())
      return runResult ? runResult() : new Promise(() => undefined)
    }
  }
  const context = createContext(globals)
  runInContext(workerSource, context, { filename: 'pptxnative.worker.js' })
  const send = async (request: Record<string, unknown>) => {
    await (context.onmessage as (event: { data: unknown }) => Promise<void>)({ data: request })
    return messages.at(-1)
  }
  const base = { protocol: 'injoffice.native-wasm-worker', version: 1, format: 'pptx' }
  return {
    messages,
    isClosed: () => closed,
    init: () => send({ ...base, id: 'init-1', op: 'init', assets: { wasmUrl: '/engine.wasm', goRuntimeUrl: '/wasm_exec.js' } }),
    extract: (id: string) => send({ ...base, id, op: 'extract', bytes: new Uint8Array([1]).buffer }),
  }
}

describe('PPTX WASM worker binding envelopes', () => {
  it('closes after a fatal Go response', async () => {
    const worker = createWorkerHarness({
      extract: () => ({ ok: false, error: 'pptxnative panic: boom', fatal: true }),
      apply: () => ({ ok: true, value: new Uint8Array([1]) }),
    })
    await worker.init()
    await expect(worker.extract('extract-1')).resolves.toMatchObject({ ok: false, error: { code: 'NATIVE_FAILED', fatal: true } })
    expect(worker.isClosed()).toBe(true)
  })

  it('keeps ordinary and legacy binding refusals recoverable', async () => {
    let calls = 0
    const worker = createWorkerHarness({
      extract: () => ++calls === 1 ? { ok: false, error: 'invalid package', fatal: false } : ++calls === 3 ? { ok: false, error: 'stale package' } : { ok: true, value: '{}' },
      apply: () => ({ ok: true, value: new Uint8Array([1]) }),
    })
    await worker.init()
    await expect(worker.extract('extract-1')).resolves.toMatchObject({ ok: false, error: { code: 'NATIVE_REFUSED', fatal: false } })
    await expect(worker.extract('extract-2')).resolves.toMatchObject({ ok: false, error: { code: 'NATIVE_REFUSED', fatal: false } })
    await expect(worker.extract('extract-3')).resolves.toMatchObject({ ok: true })
    expect(worker.isClosed()).toBe(false)
  })

  it('fails closed on malformed success and reports a later engine exit', async () => {
    const malformed = createWorkerHarness({
      extract: () => ({ ok: true, value: 42 }),
      apply: () => ({ ok: true, value: new Uint8Array([1]) }),
    })
    await malformed.init()
    await expect(malformed.extract('extract-1')).resolves.toMatchObject({ ok: false, error: { fatal: true } })
    expect(malformed.isClosed()).toBe(true)

    let exit!: () => void
    const runtimeExit = new Promise<void>((resolve) => { exit = resolve })
    const exited = createWorkerHarness({ extract: () => ({ ok: true, value: '{}' }), apply: () => ({ ok: true, value: new Uint8Array([1]) }) }, () => runtimeExit)
    await exited.init()
    exit()
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(exited.messages.at(-1)).toMatchObject({ op: 'fatal', error: { code: 'ENGINE_EXITED', fatal: true } })
    expect(exited.isClosed()).toBe(true)
  })
})
