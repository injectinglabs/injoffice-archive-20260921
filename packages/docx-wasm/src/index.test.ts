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
  OFFICE_MUTATION_PROTOCOL,
  OFFICE_MUTATION_VERSION,
  NativeDocxValidationError,
  decodeNativeDocxDocument,
  type NativeDocxDocumentV1,
  type NativeDocxOfficeMutationEnvelopeV1,
} from '@injoffice/docs'
import {
  DOCX_WASM_NATIVE_MAX_PACKAGE_BYTES,
  createDocxWasmClient,
  resolveDocxWasmAssetUrls,
} from './index'

const fixtureRecord = JSON.parse(readFileSync(
  new URL('../../../go/officecompat/corpus/generated/expected/docx-strict-relocated.json', import.meta.url),
  'utf8',
)) as { native: unknown }
const fixtureDecoded = decodeNativeDocxDocument(fixtureRecord.native)
if (!fixtureDecoded.ok) throw new Error('DOCX WASM test fixture must be valid')
const fixtureDocument = fixtureDecoded.value
const fixtureRun = fixtureDocument.body.blocks[0].paragraph!.runs[0]
const fixtureParagraph = fixtureDocument.body.blocks[0].paragraph!
const multiRunRecord = JSON.parse(readFileSync(
  new URL('../../../go/officecompat/corpus/generated/expected/docx-inline-png-page-paint.json', import.meta.url),
  'utf8',
)) as { native: unknown }
const multiRunDecoded = decodeNativeDocxDocument(multiRunRecord.native)
if (!multiRunDecoded.ok) throw new Error('DOCX WASM multi-run fixture must be valid')
const multiRunDocument = multiRunDecoded.value
const multiRunParagraph = multiRunDocument.body.blocks.map((block) => block.paragraph).find((paragraph) => paragraph && paragraph.runs.length > 1)!

const envelope = (document = fixtureDocument): NativeDocxOfficeMutationEnvelopeV1 => ({
  protocol: OFFICE_MUTATION_PROTOCOL,
  version: OFFICE_MUTATION_VERSION,
  format: 'docx',
  mutation_id: 'save-1',
  expected_revision: document.source.package_sha256,
  payload: {
    mutations: [{
      target_kind: 'run',
      target_id: fixtureRun.id,
      expected_xml_sha256: fixtureRun.anchor.xml_sha256,
      text: 'Strict native',
    }],
  },
})

class FakeWorker implements NativeWasmWorker {
  readonly requests: NativeWasmWorkerRequest[] = []
  terminated = false
  private readonly messageListeners = new Set<(event: NativeWasmMessageEvent) => void>()
  private readonly errorListeners = new Set<(event: NativeWasmWorkerErrorEvent) => void>()

  constructor(private readonly extractJson = JSON.stringify(fixtureDocument)) {}

  postMessage(value: unknown, _transfer: ArrayBuffer[]): void {
    const request = value as NativeWasmWorkerRequest
    this.requests.push(request)
    if (request.op === 'init') this.respond(success(request))
    else if (request.op === 'extract') this.respond(success(request, { contractJson: this.extractJson }))
    else this.respond(success(request, { bytes: new Uint8Array([4, 5, 6]).buffer }))
  }

  terminate(): void { this.terminated = true }

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

describe('DOCX WASM package client', () => {
  it('resolves package-relative defaults and exact explicit overrides', () => {
    const defaults = resolveDocxWasmAssetUrls()
    expect(defaults.workerUrl).toMatch(/\/docxnative\.worker\.js$/)
    expect(defaults.wasmUrl).toMatch(/\/docxnative\.wasm$/)
    expect(defaults.goRuntimeUrl).toMatch(/\/wasm_exec\.js$/)
    expect(resolveDocxWasmAssetUrls({
      workerUrl: new URL('https://cdn.example/worker.js'),
      wasmUrl: 'https://cdn.example/engine.wasm',
      goRuntimeUrl: 'https://cdn.example/go.js',
    })).toEqual({
      workerUrl: 'https://cdn.example/worker.js',
      wasmUrl: 'https://cdn.example/engine.wasm',
      goRuntimeUrl: 'https://cdn.example/go.js',
    })
  })

  it('is lazy, validates extraction, and sends the exact CAS-bound payload', async () => {
    const worker = new FakeWorker()
    let calls = 0
    const client = createDocxWasmClient({
      workerUrl: '/worker.js', wasmUrl: '/engine.wasm', goRuntimeUrl: '/go.js',
      workerFactory: (url) => { calls++; expect(url).toBe('/worker.js'); return worker },
    })
    expect(calls).toBe(0)
    const document = await client.extract(new Uint8Array([1, 2, 3]))
    expect(document.version).toBe(1)
    expect(calls).toBe(1)
    await expect(client.apply(new Uint8Array([1]), document, envelope(document))).resolves.toEqual(new Uint8Array([4, 5, 6]))
    expect(worker.requests.map(({ op }) => op)).toEqual(['init', 'extract', 'apply'])
    expect(worker.requests[2]).toMatchObject({
      expectedRevision: document.source.package_sha256,
      payload: JSON.stringify(envelope(document).payload),
    })
  })

  it('terminates on extraction JSON that fails the public Docs contract', async () => {
    const worker = new FakeWorker('{"version":1}')
    const client = createDocxWasmClient({ workerFactory: () => worker })
    await expect(client.extract(new Uint8Array([1]))).rejects.toBeInstanceOf(NativeDocxValidationError)
    expect(worker.terminated).toBe(true)
    await expect(client.extract(new Uint8Array([1]))).rejects.toMatchObject({ code: 'TERMINATED', fatal: true })
  })

  it('throws synchronously on byte limits before creating or copying into a worker', () => {
    let calls = 0
    const packageClient = createDocxWasmClient({
      maxPackageBytes: 2,
      workerFactory: () => { calls++; return new FakeWorker() },
    })
    expect(() => packageClient.extract(new Uint8Array([1, 2, 3]))).toThrow(/maxPackageBytes is 2/)
    const payloadClient = createDocxWasmClient({ maxMutationPayloadBytes: 128, workerFactory: () => { calls++; return new FakeWorker() } })
    const large = envelope()
    large.payload.mutations[0].text = '🚀'.repeat(40)
    expect(() => payloadClient.apply(new Uint8Array([1]), fixtureDocument, large)).toThrow(/UTF-8 bytes/)
    expect(calls).toBe(0)
    expect(() => createDocxWasmClient({ maxPackageBytes: DOCX_WASM_NATIVE_MAX_PACKAGE_BYTES + 1 })).toThrow(/maxPackageBytes/)
  })

  it('strictly checks the public envelope, document CAS, and run anchor', () => {
    const client = createDocxWasmClient({ workerFactory: () => new FakeWorker() })
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, { ...envelope(), expected_revision: 'sha256:' + '0'.repeat(64) })).toThrow(/expected_revision/)
    const stale = envelope()
    stale.payload.mutations[0].expected_xml_sha256 = 'sha256:' + '0'.repeat(64)
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, stale)).toThrow(/run text anchor/)
    const unknown = { ...envelope(), extra: true } as unknown as NativeDocxOfficeMutationEnvelopeV1
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, unknown)).toThrow(/fields are invalid/)
    const accessor = envelope() as unknown as Record<string, unknown>
    Object.defineProperty(accessor, 'payload', { get: () => envelope().payload, enumerable: true })
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, accessor as unknown as NativeDocxOfficeMutationEnvelopeV1)).toThrow(/accessor/)
    const invalidDocument = { ...fixtureDocument, protocol: 'wrong' } as unknown as NativeDocxDocumentV1
    expect(() => client.apply(new Uint8Array([1]), invalidDocument, envelope())).toThrow(NativeDocxValidationError)
  })

  it('accepts the Go-proven single-text-run paragraph target with its paragraph anchor', async () => {
    const worker = new FakeWorker()
    const client = createDocxWasmClient({ workerFactory: () => worker })
    const paragraphEnvelope = envelope()
    paragraphEnvelope.payload.mutations[0] = {
      target_kind: 'paragraph',
      target_id: fixtureParagraph.id,
      expected_xml_sha256: fixtureParagraph.anchor.xml_sha256,
      text: 'Strict paragraph',
    }

    await expect(client.apply(new Uint8Array([1]), fixtureDocument, paragraphEnvelope)).resolves.toEqual(new Uint8Array([4, 5, 6]))
    expect(worker.requests[1]).toMatchObject({ payload: JSON.stringify(paragraphEnvelope.payload) })
  })

  it('refuses paragraph targets with multiple runs and paragraph/run overlap', () => {
    const client = createDocxWasmClient({ workerFactory: () => new FakeWorker() })
    const unsupported = envelope(multiRunDocument)
    unsupported.payload.mutations[0] = {
      target_kind: 'paragraph',
      target_id: multiRunParagraph.id,
      expected_xml_sha256: multiRunParagraph.anchor.xml_sha256,
      text: 'Flattened',
    }
    expect(() => client.apply(new Uint8Array([1]), multiRunDocument, unsupported)).toThrow(/paragraph text anchor/)

    const overlap = envelope()
    overlap.payload.mutations.unshift({
      target_kind: 'paragraph',
      target_id: fixtureParagraph.id,
      expected_xml_sha256: fixtureParagraph.anchor.xml_sha256,
      text: 'Paragraph text',
    })
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, overlap)).toThrow(/overlaps another paragraph or run target/)
  })

  it('refuses duplicate targets and invalid XML text', () => {
    const client = createDocxWasmClient({ workerFactory: () => new FakeWorker() })
    const duplicate = envelope()
    duplicate.payload.mutations.push({ ...duplicate.payload.mutations[0] })
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, duplicate)).toThrow(/same native target/)
    const control = envelope()
    control.payload.mutations[0].text = 'bad\u0000text'
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, control)).toThrow(/XML 1.0/)
    const carriageReturn = envelope()
    carriageReturn.payload.mutations[0].text = 'normalized\rtext'
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, carriageReturn)).toThrow(/XML 1.0/)
  })

  it('keeps module construction SSR-safe and reports missing Worker on use', async () => {
    const client = createDocxWasmClient()
    await expect(client.extract(new Uint8Array([1]))).rejects.toMatchObject({ code: 'WORKER_UNAVAILABLE', fatal: true })
  })
})
