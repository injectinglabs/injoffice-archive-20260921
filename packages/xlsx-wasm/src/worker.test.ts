import { readFileSync } from 'node:fs'
import { createContext, runInContext } from 'node:vm'
import { describe, expect, it } from 'vitest'

const workerSource = readFileSync(new URL('../worker/xlsxnative.worker.js', import.meta.url), 'utf8')

type Binding = {
  extract(bytes: Uint8Array): unknown
  inspect?(bytes: Uint8Array): unknown
  apply(original: Uint8Array, payload: string | Uint8Array, expectedRevision: string): unknown
}

type WorkerMessage = {
  op?: string
  ok?: boolean
  error?: { code?: string; message?: string; fatal?: boolean }
  result?: { contractJson?: string }
}

function createWorkerHarness(binding: Binding) {
  const messages: WorkerMessage[] = []
  let closed = false
  const globals: Record<string, unknown> = {
    ArrayBuffer,
    Error,
    Promise,
    Uint8Array,
    clearTimeout,
    queueMicrotask,
    setTimeout,
    postMessage: (message: WorkerMessage) => messages.push(message),
    importScripts: () => undefined,
    fetch: async () => ({ ok: true, arrayBuffer: async () => new ArrayBuffer(0) }),
    WebAssembly: { instantiate: async () => ({ instance: {} }) },
  }
  globals.self = globals
  globals.close = () => { closed = true }
  globals.Go = class {
    readonly importObject = {}

    run(): Promise<never> {
      globals.xlsxnative = binding
      queueMicrotask(() => (globals.xlsxnativeOnReady as () => void)())
      return new Promise<never>(() => undefined)
    }
  }
  const context = createContext(globals)
  runInContext(workerSource, context, { filename: 'xlsxnative.worker.js' })
  const send = async (request: Record<string, unknown>) => {
    await (context.onmessage as (event: { data: unknown }) => Promise<void>)({ data: request })
    return messages.at(-1)
  }
  const base = {
    protocol: 'injoffice.native-wasm-worker',
    version: 1,
    format: 'xlsx',
  }
  return {
    messages,
    isClosed: () => closed,
    init: () => send({ ...base, id: 'init-1', op: 'init', assets: { wasmUrl: '/engine.wasm', goRuntimeUrl: '/wasm_exec.js' } }),
    extract: (id: string) => send({ ...base, id, op: 'extract', bytes: new Uint8Array([1]).buffer }),
    inspect: (id: string) => send({ ...base, id, op: 'inspect', bytes: new Uint8Array([1]).buffer }),
  }
}

describe('XLSX WASM worker binding envelopes', () => {
  it('inspects without extraction/mutation calls and refuses old engines recoverably', async () => {
    const binding={extract:()=>({ok:true,value:'{}'}),apply:()=>{throw new Error('must not mutate')}}
    const old=createWorkerHarness(binding);await old.init();expect(await old.inspect('old')).toMatchObject({ok:false,error:{fatal:false}});expect(old.isClosed()).toBe(false)
    const current=createWorkerHarness({...binding,inspect:()=>({ok:true,value:'{"version":1}'})});await current.init();expect(await current.inspect('new')).toMatchObject({ok:true,result:{contractJson:'{"version":1}'}})
  })
  it('propagates a recovered Go panic as fatal and closes the worker', async () => {
    const worker = createWorkerHarness({
      extract: () => ({ ok: false, error: 'xlsxnative panic: boom', fatal: true }),
      apply: () => ({ ok: true, value: new Uint8Array([1]) }),
    })
    await worker.init()

    await expect(worker.extract('extract-1')).resolves.toMatchObject({
      ok: false,
      error: { code: 'NATIVE_FAILED', message: 'xlsxnative panic: boom', fatal: true },
    })
    expect(worker.isClosed()).toBe(true)
  })

  it('keeps ordinary and legacy Go refusals recoverable', async () => {
    let calls = 0
    const worker = createWorkerHarness({
      extract: () => {
        calls++
        return calls === 1
          ? { ok: false, error: 'xlsxpatch: package is empty', fatal: false }
          : calls === 2
            ? { ok: false, error: 'xlsxpatch: stale revision' }
            : { ok: true, value: '{"version":2}' }
      },
      apply: () => ({ ok: true, value: new Uint8Array([1]) }),
    })
    await worker.init()

    await expect(worker.extract('extract-1')).resolves.toMatchObject({ ok: false, error: { code: 'NATIVE_REFUSED', fatal: false } })
    await expect(worker.extract('extract-2')).resolves.toMatchObject({ ok: false, error: { code: 'NATIVE_REFUSED', fatal: false } })
    await expect(worker.extract('extract-3')).resolves.toMatchObject({ ok: true, result: { contractJson: '{"version":2}' } })
    expect(worker.isClosed()).toBe(false)
  })

  it('fails closed on a malformed successful binding envelope', async () => {
    const worker = createWorkerHarness({
      extract: () => ({ ok: true, value: 42 }),
      apply: () => ({ ok: true, value: new Uint8Array([1]) }),
    })
    await worker.init()

    await expect(worker.extract('extract-1')).resolves.toMatchObject({ ok: false, error: { code: 'NATIVE_FAILED', fatal: true } })
    expect(worker.isClosed()).toBe(true)
  })
})
