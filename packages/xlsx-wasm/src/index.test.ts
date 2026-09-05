import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import {
  NATIVE_WASM_WORKER_PROTOCOL,
  NATIVE_WASM_WORKER_VERSION,
  type NativeWasmMessageEvent,
  type NativeWasmWorker,
  type NativeWasmWorkerErrorEvent,
  type NativeWasmWorkerRequest,
  type NativeWasmWorkerResponse,
} from '@injoffice/native-runtime'
import {
  NativeWorkbookV2ValidationError,
  WORKBOOK_MUTATION_PROTOCOL,
  WORKBOOK_MUTATION_VERSION,
  decodeNativeWorkbookV2,
  type WorkbookMutationBatchV1,
} from '@injoffice/sheets/browser'
import {
  XLSX_WASM_NATIVE_MAX_PACKAGE_BYTES,
  adaptWorkbookMutationBatchV1,
  createXlsxWasmClient,
  resolveXlsxWasmAssetUrls,
} from './index'

const fixtureJson = readFileSync(
  new URL('../../../go/xlsxpatch/testdata/native-xlsx-v2/valid/excel-authored-happy-tree.json', import.meta.url),
  'utf8',
).trim()
const fixtureDecoded = decodeNativeWorkbookV2(fixtureJson)
if (!fixtureDecoded.ok) throw new Error('XLSX WASM test fixture must be valid')
const fixtureWorkbook = fixtureDecoded.value

const clearBatch = (revision = fixtureWorkbook.source.package_sha256): WorkbookMutationBatchV1 => ({
  protocol: WORKBOOK_MUTATION_PROTOCOL,
  version: WORKBOOK_MUTATION_VERSION,
  batch_id: 'save-1',
  expected_revision: revision,
  operations: [{
    operation_id: 'clear-1',
    sheet_id: fixtureWorkbook.sheets[0].id,
    kind: 'cell.clear_value',
    cell: { row: 0, column: 0 },
  }],
})

class FakeWorker implements NativeWasmWorker {
  readonly requests: NativeWasmWorkerRequest[] = []
  terminated = false
  private readonly messageListeners = new Set<(event: NativeWasmMessageEvent) => void>()
  private readonly errorListeners = new Set<(event: NativeWasmWorkerErrorEvent) => void>()

  constructor(private readonly extractJson = fixtureJson) {}

  postMessage(value: unknown, _transfer: ArrayBuffer[]): void {
    const request = value as NativeWasmWorkerRequest
    this.requests.push(request)
    if (request.op === 'init') this.respond(success(request))
    else if (request.op === 'extract') this.respond(success(request, { contractJson: this.extractJson }))
    else this.respond(success(request, { bytes: new Uint8Array([4, 5, 6]).buffer }))
  }

  terminate(): void {
    this.terminated = true
  }

  addEventListener(type: 'message' | 'error', listener: ((event: NativeWasmMessageEvent) => void) | ((event: NativeWasmWorkerErrorEvent) => void)): void {
    if (type === 'message') this.messageListeners.add(listener as (event: NativeWasmMessageEvent) => void)
    else this.errorListeners.add(listener as (event: NativeWasmWorkerErrorEvent) => void)
  }

  removeEventListener(type: 'message' | 'error', listener: ((event: NativeWasmMessageEvent) => void) | ((event: NativeWasmWorkerErrorEvent) => void)): void {
    if (type === 'message') this.messageListeners.delete(listener as (event: NativeWasmMessageEvent) => void)
    else this.errorListeners.delete(listener as (event: NativeWasmWorkerErrorEvent) => void)
  }

  private respond(response: NativeWasmWorkerResponse): void {
    queueMicrotask(() => {
      for (const listener of this.messageListeners) listener({ data: response })
    })
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

describe('XLSX WASM package client', () => {
  it('resolves package-relative defaults and exact explicit overrides', () => {
    const defaults = resolveXlsxWasmAssetUrls()
    expect(defaults.workerUrl).toMatch(/\/xlsxnative\.worker\.js$/)
    expect(defaults.wasmUrl).toMatch(/\/xlsxnative\.wasm$/)
    expect(defaults.goRuntimeUrl).toMatch(/\/wasm_exec\.js$/)

    expect(resolveXlsxWasmAssetUrls({
      workerUrl: new URL('https://cdn.example/worker.js'),
      wasmUrl: 'https://cdn.example/engine.wasm',
      goRuntimeUrl: 'https://cdn.example/go.js',
    })).toEqual({
      workerUrl: 'https://cdn.example/worker.js',
      wasmUrl: 'https://cdn.example/engine.wasm',
      goRuntimeUrl: 'https://cdn.example/go.js',
    })
  })

  it('is lazy, validates extract v2, and sends a validated CAS-bound transaction', async () => {
    const worker = new FakeWorker()
    let calls = 0
    const client = createXlsxWasmClient({
      workerUrl: '/worker.js',
      wasmUrl: '/engine.wasm',
      goRuntimeUrl: '/go.js',
      workerFactory: (url) => {
        calls++
        expect(url).toBe('/worker.js')
        return worker
      },
    })

    expect(calls).toBe(0)
    const workbook = await client.extract(new Uint8Array([1, 2, 3]))
    expect(workbook.version).toBe(2)
    expect(calls).toBe(1)
    expect(worker.requests[0]).toMatchObject({
      op: 'init',
      assets: { wasmUrl: '/engine.wasm', goRuntimeUrl: '/go.js' },
    })
    const transaction = adaptWorkbookMutationBatchV1(workbook, clearBatch())
    await expect(client.apply(new Uint8Array([1]), workbook, transaction)).resolves.toEqual(new Uint8Array([4, 5, 6]))
    expect(worker.requests.map(({ op }) => op)).toEqual(['init', 'extract', 'apply'])
    expect(worker.requests[2]).toMatchObject({
      expectedRevision: workbook.source.package_sha256,
      payload: JSON.stringify(transaction),
    })
  })

  it('refuses extraction JSON that fails the public sheets contract', async () => {
    const worker = new FakeWorker('{"version":2}')
    const client = createXlsxWasmClient({ workerFactory: () => worker })

    await expect(client.extract(new Uint8Array([1]))).rejects.toBeInstanceOf(NativeWorkbookV2ValidationError)
    expect(worker.terminated).toBe(true)
    await expect(client.extract(new Uint8Array([1]))).rejects.toMatchObject({ code: 'TERMINATED', fatal: true })
  })

  it('throws synchronously on package and UTF-8 mutation limits before worker copies', () => {
    let calls = 0
    const packageClient = createXlsxWasmClient({
      maxPackageBytes: 2,
      workerFactory: () => {
        calls++
        return new FakeWorker()
      },
    })
    expect(() => packageClient.extract(new Uint8Array([1, 2, 3]))).toThrow(/maxPackageBytes is 2/)
    expect(calls).toBe(0)

    const payloadClient = createXlsxWasmClient({
      maxMutationPayloadBytes: 128,
      workerFactory: () => {
        calls++
        return new FakeWorker()
      },
    })
    const transaction = adaptWorkbookMutationBatchV1(fixtureWorkbook, {
      ...clearBatch(),
      operations: [{
        operation_id: 'unicode-1',
        sheet_id: fixtureWorkbook.sheets[0].id,
        kind: 'cell.set_value',
        cell: { row: 0, column: 0 },
        value: '🚀'.repeat(40),
      }],
    })
    expect(() => payloadClient.apply(new Uint8Array([1]), fixtureWorkbook, transaction)).toThrow(/UTF-8 bytes/)
    expect(calls).toBe(0)
    expect(() => createXlsxWasmClient({ maxPackageBytes: XLSX_WASM_NATIVE_MAX_PACKAGE_BYTES + 1 })).toThrow(/maxPackageBytes/)
  })

  it('adapts only the native subset and enforces outer and inner CAS values', () => {
    const transaction = adaptWorkbookMutationBatchV1(fixtureWorkbook, clearBatch())
    expect(transaction).toEqual({
      expected_revision: fixtureWorkbook.revision,
      cells: clearBatch().operations,
    })
    expect(() => adaptWorkbookMutationBatchV1(fixtureWorkbook, clearBatch('sha256:' + '0'.repeat(64))))
      .toThrow(/outer expected_revision/)

    const merge = {
      ...clearBatch(),
      operations: [{
        operation_id: 'merge-1',
        sheet_id: fixtureWorkbook.sheets[0].id,
        kind: 'range.merge',
        range: { row: 0, column: 0, end_row: 0, end_column: 1 },
      }],
    } as WorkbookMutationBatchV1
    expect(() => adaptWorkbookMutationBatchV1(fixtureWorkbook, merge)).toThrow(/not supported/)

    const client = createXlsxWasmClient({ workerFactory: () => new FakeWorker() })
    expect(() => client.apply(new Uint8Array([1]), fixtureWorkbook, {
      ...transaction,
      expected_revision: 'rev:' + '0'.repeat(64),
    })).toThrow(/workbook\.revision/)
  })

  it('refuses mutation inputs that cannot round-trip through the native contract', () => {
    const styleThenCell: WorkbookMutationBatchV1 = {
      ...clearBatch(),
      operations: [{
        operation_id: 'style-1',
        sheet_id: fixtureWorkbook.sheets[0].id,
        kind: 'style.patch',
        range: { row: 0, column: 0, end_row: 0, end_column: 0 },
        style: { bold: true },
      }, clearBatch().operations[0]],
    }
    expect(() => adaptWorkbookMutationBatchV1(fixtureWorkbook, styleThenCell)).toThrow(/cell operations before style operations/)

    const loneSurrogate: WorkbookMutationBatchV1 = {
      ...clearBatch(),
      operations: [{
        operation_id: 'unicode-1',
        sheet_id: fixtureWorkbook.sheets[0].id,
        kind: 'cell.set_value',
        cell: { row: 0, column: 0 },
        value: '\ud800',
      }],
    }
    expect(() => adaptWorkbookMutationBatchV1(fixtureWorkbook, loneSurrogate)).toThrow(/unpaired UTF-16/)

    const emptyFormula: WorkbookMutationBatchV1 = {
      ...clearBatch(),
      operations: [{
        operation_id: 'formula-1',
        sheet_id: fixtureWorkbook.sheets[0].id,
        kind: 'cell.set_formula',
        cell: { row: 0, column: 0 },
        formula: '=',
      }],
    }
    expect(() => adaptWorkbookMutationBatchV1(fixtureWorkbook, emptyFormula)).toThrow(/must contain content/)
  })

  it('keeps module construction SSR-safe and reports missing Worker on use', async () => {
    const client = createXlsxWasmClient()
    await expect(client.extract(new Uint8Array([1]))).rejects.toMatchObject({ code: 'WORKER_UNAVAILABLE', fatal: true })
  })
})
