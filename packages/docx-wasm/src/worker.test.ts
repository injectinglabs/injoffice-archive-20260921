import { readFileSync } from 'node:fs'
import { createContext, runInContext } from 'node:vm'
import { describe, expect, it } from 'vitest'

const workerSource = readFileSync(new URL('../worker/docxnative.worker.js', import.meta.url), 'utf8')

type Binding = {
  inspect?(bytes:Uint8Array):unknown
  extract(bytes: Uint8Array): unknown
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
    run(): Promise<never> {
      globals.docxnative = binding
      queueMicrotask(() => (globals.docxnativeOnReady as () => void)())
      return new Promise<never>(() => undefined)
    }
  }
  const context = createContext(globals)
  runInContext(workerSource, context, { filename: 'docxnative.worker.js' })
  const send = async (request: Record<string, unknown>) => {
    await (context.onmessage as (event: { data: unknown }) => Promise<void>)({ data: request })
    return messages.at(-1)
  }
  const base = { protocol: 'injoffice.native-wasm-worker', version: 1, format: 'docx' }
  return {
    isClosed: () => closed,
    init: () => send({ ...base, id: 'init-1', op: 'init', assets: { wasmUrl: '/engine.wasm', goRuntimeUrl: '/wasm_exec.js' } }),
    extract: (id: string) => send({ ...base, id, op: 'extract', bytes: new Uint8Array([1]).buffer }),
    inspect: (id:string)=>send({...base,id,op:'inspect',bytes:new Uint8Array([1]).buffer}),
  }
}

describe('DOCX WASM worker binding envelopes', () => {
  it('keeps inspection a separate read-only operation and refuses missing matching engines',async()=>{
    const bindings={extract:()=>({ok:true,value:'extract'}),apply:()=>({ok:true,value:new Uint8Array([1])})}
    const missing=createWorkerHarness(bindings);await missing.init()
    await expect(missing.inspect('inspect-missing')).resolves.toMatchObject({op:'inspect',ok:false,error:{fatal:true}})
    const worker=createWorkerHarness({...bindings,inspect:()=>({ok:true,value:'{"protocol":"injoffice.docx.partial-source"}'})});await worker.init()
    await expect(worker.inspect('inspect-1')).resolves.toMatchObject({op:'inspect',ok:true,result:{contractJson:'{"protocol":"injoffice.docx.partial-source"}'}})
    await expect(worker.extract('extract-1')).resolves.toMatchObject({op:'extract',ok:true,result:{contractJson:'extract'}})
  })
  it('propagates a recovered Go panic as fatal and closes the worker', async () => {
    const worker = createWorkerHarness({
      extract: () => ({ ok: false, error: 'docxnative panic: boom', fatal: true }),
      apply: () => ({ ok: true, value: new Uint8Array([1]) }),
    })
    await worker.init()
    await expect(worker.extract('extract-1')).resolves.toMatchObject({ ok: false, error: { code: 'NATIVE_FAILED', fatal: true } })
    expect(worker.isClosed()).toBe(true)
  })

  it('keeps ordinary and legacy Go refusals recoverable', async () => {
    let calls = 0
    const worker = createWorkerHarness({
      extract: () => {
        calls++
        return calls === 1
          ? { ok: false, error: 'docxpatch: empty', fatal: false }
          : calls === 2
            ? { ok: false, error: 'legacy refusal' }
            : { ok: true, value: '{"version":1}' }
      },
      apply: () => ({ ok: true, value: new Uint8Array([1]) }),
    })
    await worker.init()
    await expect(worker.extract('extract-1')).resolves.toMatchObject({ ok: false, error: { code: 'NATIVE_REFUSED', fatal: false } })
    await expect(worker.extract('extract-2')).resolves.toMatchObject({ ok: false, error: { code: 'NATIVE_REFUSED', fatal: false } })
    await expect(worker.extract('extract-3')).resolves.toMatchObject({ ok: true, result: { contractJson: '{"version":1}' } })
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
