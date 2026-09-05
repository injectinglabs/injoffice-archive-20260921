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
import type { NativeElement, NativePptxDeck } from '@injoffice/pptx-native'
import {
  PPTX_WASM_NATIVE_MAX_PACKAGE_BYTES,
  PptxNativeContractError,
  createPptxWasmClient,
  resolvePptxWasmAssetUrls,
  type PptxNativeExactParagraphV1,
} from './index'

const fixtureJson = readFileSync(new URL('../../../go/pptxpatch/testdata/native-contract/valid/parsed-full.json', import.meta.url), 'utf8').trim()
const fixtureDeck = JSON.parse(fixtureJson) as NativePptxDeck
fixtureDeck.sourceRevision = `rev-${'a'.repeat(64)}`
const runtimeFixtureJson = JSON.stringify(fixtureDeck)

class FakeWorker implements NativeWasmWorker {
  readonly requests: NativeWasmWorkerRequest[] = []
  terminated = false
  private readonly messageListeners = new Set<(event: NativeWasmMessageEvent) => void>()
  private readonly errorListeners = new Set<(event: NativeWasmWorkerErrorEvent) => void>()

  constructor(private readonly extractJson = runtimeFixtureJson) {}

  postMessage(value: unknown, _transfer: ArrayBuffer[]): void {
    const request = value as NativeWasmWorkerRequest
    this.requests.push(request)
    if (request.op === 'init') this.respond(success(request))
    else if (request.op === 'extract') this.respond(success(request, { contractJson: this.extractJson }))
    else this.respond(success(request, { bytes: new Uint8Array([7, 8, 9]).buffer }))
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
    queueMicrotask(() => this.messageListeners.forEach((listener) => listener({ data: response })))
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

function visit(elements: ReadonlyArray<NativeElement>, predicate: (element: NativeElement) => boolean): NativeElement | undefined {
  for (const element of elements) {
    if (predicate(element)) return element
    if (element.kind === 'group') {
      const child = visit(element.children, predicate)
      if (child) return child
    }
  }
}

function textMutation(deck = fixtureDeck) {
  const target = deck.slides.flatMap((slide) => slide.elements).map((element) => visit([element], (item) => (item.kind === 'text' || item.kind === 'shape') && item.source !== undefined && item.paragraphs !== undefined)).find(Boolean)
  if (!target || (target.kind !== 'text' && target.kind !== 'shape') || !target.source || !target.paragraphs) throw new Error('fixture has no editable exact text')
  const paragraphs: PptxNativeExactParagraphV1[] = target.paragraphs.map((paragraph) => ({
    align: paragraph.align ?? 'left',
    level: paragraph.level ?? 0,
    bullet: false,
    runs: paragraph.runs.map((run) => ({
      text: run.text,
      bold: run.bold ?? false,
      italic: run.italic ?? false,
      fontSizeHundredthPt: run.fontSizeHundredthPt ?? 1800,
      color: run.color ?? '000000',
      fontFamily: run.fontFamily ?? 'Arial',
    })),
  }))
  paragraphs[0]!.runs[0]!.text += ' edited'
  return {
    expectedSourceRevision: deck.sourceRevision!,
    operations: [{
      operationId: 'replace-title',
      kind: 'text.replace' as const,
      elementId: target.id,
      expectedFingerprintSha256: target.source.fingerprintSha256,
      paragraphs,
    }],
  }
}

describe('PPTX WASM package client', () => {
  it('resolves package-relative defaults and exact overrides', () => {
    expect(resolvePptxWasmAssetUrls()).toMatchObject({
      workerUrl: expect.stringMatching(/\/pptxnative\.worker\.js$/),
      wasmUrl: expect.stringMatching(/\/pptxnative\.wasm$/),
      goRuntimeUrl: expect.stringMatching(/\/wasm_exec\.js$/),
    })
    expect(resolvePptxWasmAssetUrls({
      workerUrl: new URL('https://cdn.example/worker.js'),
      wasmUrl: 'https://cdn.example/engine.wasm',
      goRuntimeUrl: 'https://cdn.example/go.js',
    })).toEqual({ workerUrl: 'https://cdn.example/worker.js', wasmUrl: 'https://cdn.example/engine.wasm', goRuntimeUrl: 'https://cdn.example/go.js' })
  })

  it('is lazy, validates extraction, and sends a typed CAS-bound mutation', async () => {
    const worker = new FakeWorker()
    let calls = 0
    const client = createPptxWasmClient({
      workerUrl: '/worker.js', wasmUrl: '/engine.wasm', goRuntimeUrl: '/go.js',
      workerFactory: (url) => { calls++; expect(url).toBe('/worker.js'); return worker },
    })
    expect(calls).toBe(0)
    const deck = await client.extract(new Uint8Array([1, 2, 3]))
    expect(deck.contractVersion).toBe('pptx-native/v1')
    expect(calls).toBe(1)
    await expect(client.apply(new Uint8Array([1]), deck, textMutation(deck))).resolves.toEqual(new Uint8Array([7, 8, 9]))
    expect(worker.requests.map(({ op }) => op)).toEqual(['init', 'extract', 'apply'])
    expect(worker.requests[2]).toMatchObject({
      expectedRevision: `sha256:${deck.sourceRevision!.slice(4)}`,
    })
    const applyRequest = worker.requests[2]
    if (applyRequest?.op !== 'apply' || typeof applyRequest.payload !== 'string') throw new Error('missing apply request')
    expect(JSON.parse(applyRequest.payload)).toEqual(textMutation(deck))
  })

  it('terminates on malformed or schema-invalid extraction JSON', async () => {
    for (const invalid of ['{', '{"contractVersion":"pptx-native/v1"}']) {
      const worker = new FakeWorker(invalid)
      const client = createPptxWasmClient({ workerFactory: () => worker })
      await expect(client.extract(new Uint8Array([1]))).rejects.toBeInstanceOf(PptxNativeContractError)
      expect(worker.terminated).toBe(true)
    }
  })

  it('refuses unsupported, stale, and structurally invalid mutations before worker use', () => {
    let calls = 0
    const client = createPptxWasmClient({ workerFactory: () => { calls++; return new FakeWorker() } })
    const mutation = textMutation()
    expect(() => client.apply(new Uint8Array([1]), fixtureDeck, { ...mutation, expectedSourceRevision: `rev-${'0'.repeat(64)}` })).toThrow(/sourceRevision/)
    expect(() => client.apply(new Uint8Array([1]), fixtureDeck, {
      ...mutation,
      operations: [{ ...mutation.operations[0], kind: 'slide.insert' } as never],
    })).toThrow(/Unsupported PPTX native mutation kind/)
    expect(() => client.apply(new Uint8Array([1]), fixtureDeck, {
      ...mutation,
      operations: [{ ...mutation.operations[0], unexpected: true } as never],
    })).toThrow(/unknown field/)
    expect(calls).toBe(0)
  })

  it('enforces package and UTF-8 payload limits synchronously', () => {
    let calls = 0
    const packageClient = createPptxWasmClient({ maxPackageBytes: 2, workerFactory: () => { calls++; return new FakeWorker() } })
    expect(() => packageClient.extract(new Uint8Array([1, 2, 3]))).toThrow(/maxPackageBytes is 2/)
    const payloadClient = createPptxWasmClient({ maxMutationPayloadBytes: 256, workerFactory: () => { calls++; return new FakeWorker() } })
    const mutation = textMutation()
    mutation.operations[0].paragraphs[0]!.runs[0]!.text = '🚀'.repeat(100)
    expect(() => payloadClient.apply(new Uint8Array([1]), fixtureDeck, mutation)).toThrow(/UTF-8 bytes/)
    expect(() => createPptxWasmClient({ maxPackageBytes: PPTX_WASM_NATIVE_MAX_PACKAGE_BYTES + 1 })).toThrow(/maxPackageBytes/)
    expect(calls).toBe(0)
  })

  it('keeps construction SSR-safe and reports missing Worker on use', async () => {
    const client = createPptxWasmClient()
    await expect(client.extract(new Uint8Array([1]))).rejects.toMatchObject({ code: 'WORKER_UNAVAILABLE', fatal: true })
  })
})
