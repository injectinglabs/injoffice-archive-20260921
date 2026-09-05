import {
  NativeWasmError,
  createNativeWasmClient,
  type NativeWasmAbortSignal,
  type NativeWasmClient,
  type NativeWasmWorker,
} from '@injoffice/native-runtime'
import {
  PPTX_NATIVE_RESOURCE_LIMITS,
  assertNativePptx,
  validateNativePptx,
  type NativeElement,
  type NativePptxDeck,
  type NativeTransform,
} from '@injoffice/pptx-native'

export const PPTX_WASM_NATIVE_MAX_PACKAGE_BYTES = 512 * 1024 * 1024
export const PPTX_WASM_NATIVE_MAX_MUTATION_PAYLOAD_BYTES = 3 * 1024 * 1024
export const DEFAULT_PPTX_WASM_MAX_PACKAGE_BYTES = 32 * 1024 * 1024
export const DEFAULT_PPTX_WASM_MAX_MUTATION_PAYLOAD_BYTES = 1024 * 1024

const MAX_OPERATIONS = 10_000
const MAX_MUTATION_PARAGRAPHS = 10_000
const MAX_MUTATION_RUNS = 100_000
const NATIVE_ID = /^[A-Za-z0-9][A-Za-z0-9._~:/-]{0,255}$/
const SHA256 = /^[0-9a-f]{64}$/
const COLOR = /^[0-9A-F]{6}$/

export interface PptxWasmAssetUrls {
  workerUrl: string
  wasmUrl: string
  goRuntimeUrl: string
}

export interface PptxWasmClientOptions {
  workerUrl?: string | URL
  wasmUrl?: string | URL
  goRuntimeUrl?: string | URL
  operationTimeoutMs?: number
  /** Browser-side ceiling, capped by the native 512 MiB package limit. */
  maxPackageBytes?: number
  /** UTF-8 JSON ceiling, capped by the native 3 MiB mutation limit. */
  maxMutationPayloadBytes?: number
  workerFactory?: (workerUrl: string) => NativeWasmWorker
}

export interface PptxWasmOperationOptions { signal?: NativeWasmAbortSignal }

export interface PptxNativeExactTextRunV1 {
  text: string
  bold: boolean
  italic: boolean
  fontSizeHundredthPt: number
  color: string
  fontFamily: string
}

export interface PptxNativeExactParagraphV1 {
  runs: ReadonlyArray<PptxNativeExactTextRunV1>
  align: 'left' | 'center' | 'right'
  level: number
  bullet: false
}

export interface PptxNativeTextReplaceMutationV1 {
  operationId: string
  kind: 'text.replace'
  elementId: string
  expectedFingerprintSha256: string
  paragraphs: ReadonlyArray<PptxNativeExactParagraphV1>
}

export interface PptxNativeExactStrokeV1 {
  color: string
  widthEmu: number
  cap: 'flat' | 'round' | 'square'
  join: 'round' | 'bevel' | 'miter'
  dash: 'solid'
  miterLimit?: number
}

export interface PptxNativeExactAutoShapeV1 {
  transform: NativeTransform
  preset: 'rect' | 'ellipse' | 'triangle' | 'diamond'
  /** Omission encodes explicit no-fill. */
  fill?: string
  /** Omission encodes an explicit no-fill outline. */
  stroke?: PptxNativeExactStrokeV1
}

export interface PptxNativeAutoShapeUpdateMutationV1 {
  operationId: string
  kind: 'autoshape.update'
  elementId: string
  expectedFingerprintSha256: string
  autoShape: PptxNativeExactAutoShapeV1
}

export type PptxNativeMutationV1 = PptxNativeTextReplaceMutationV1 | PptxNativeAutoShapeUpdateMutationV1

export interface PptxNativeMutationRequestV1 {
  expectedSourceRevision: string
  operations: ReadonlyArray<PptxNativeMutationV1>
}

export class PptxNativeContractError extends TypeError {
  readonly issues: ReadonlyArray<{ path: string; code: string; message: string }>

  constructor(issues: ReadonlyArray<{ path: string; code: string; message: string }>) {
    super(`Invalid native PPTX contract: ${issues.slice(0, 8).map((issue) => `${issue.path}: ${issue.message}`).join('; ')}`)
    this.name = 'PptxNativeContractError'
    this.issues = issues
  }
}

export interface PptxWasmClient {
  extract(bytes: Uint8Array, options?: PptxWasmOperationOptions): Promise<NativePptxDeck>
  apply(
    original: Uint8Array,
    deck: NativePptxDeck,
    mutation: PptxNativeMutationRequestV1,
    options?: PptxWasmOperationOptions,
  ): Promise<Uint8Array>
  terminate(): void
}

export function resolvePptxWasmAssetUrls(options: Pick<PptxWasmClientOptions, 'workerUrl' | 'wasmUrl' | 'goRuntimeUrl'> = {}): PptxWasmAssetUrls {
  return {
    workerUrl: toUrl(options.workerUrl, new URL('./pptxnative.worker.js', import.meta.url)),
    wasmUrl: toUrl(options.wasmUrl, new URL('./pptxnative.wasm', import.meta.url)),
    goRuntimeUrl: toUrl(options.goRuntimeUrl, new URL('./wasm_exec.js', import.meta.url)),
  }
}

export function createPptxWasmClient(options: PptxWasmClientOptions = {}): PptxWasmClient {
  const assets = resolvePptxWasmAssetUrls(options)
  const workerFactory = options.workerFactory ?? createBrowserWorker
  const maxPackageBytes = boundedLimit(options.maxPackageBytes, DEFAULT_PPTX_WASM_MAX_PACKAGE_BYTES, PPTX_WASM_NATIVE_MAX_PACKAGE_BYTES, 'maxPackageBytes')
  const maxMutationPayloadBytes = boundedLimit(options.maxMutationPayloadBytes, DEFAULT_PPTX_WASM_MAX_MUTATION_PAYLOAD_BYTES, PPTX_WASM_NATIVE_MAX_MUTATION_PAYLOAD_BYTES, 'maxMutationPayloadBytes')
  const native = createNativeWasmClient({
    format: 'pptx',
    workerFactory: () => workerFactory(assets.workerUrl),
    assets: { wasmUrl: assets.wasmUrl, goRuntimeUrl: assets.goRuntimeUrl },
    operationTimeoutMs: options.operationTimeoutMs,
  })
  return new PptxWasmClientImpl(native, maxPackageBytes, maxMutationPayloadBytes)
}

class PptxWasmClientImpl implements PptxWasmClient {
  constructor(
    private readonly native: NativeWasmClient,
    private readonly maxPackageBytes: number,
    private readonly maxMutationPayloadBytes: number,
  ) {}

  extract(bytes: Uint8Array, options: PptxWasmOperationOptions = {}): Promise<NativePptxDeck> {
    assertPackageSize(bytes, this.maxPackageBytes)
    return this.extractValidated(bytes, options)
  }

  apply(
    original: Uint8Array,
    deck: NativePptxDeck,
    mutation: PptxNativeMutationRequestV1,
    options: PptxWasmOperationOptions = {},
  ): Promise<Uint8Array> {
    assertPackageSize(original, this.maxPackageBytes)
    assertNativePptx(deck)
    const normalized = validateMutation(deck, mutation)
    const payload = JSON.stringify(normalized)
    const payloadBytes = new TextEncoder().encode(payload).byteLength
    if (payloadBytes > this.maxMutationPayloadBytes) {
      throw new RangeError(`PPTX mutation payload is ${payloadBytes} UTF-8 bytes; maxMutationPayloadBytes is ${this.maxMutationPayloadBytes}.`)
    }
    return this.native.apply(original, payload, outerRevision(normalized.expectedSourceRevision), options)
  }

  terminate(): void { this.native.terminate() }

  private async extractValidated(bytes: Uint8Array, options: PptxWasmOperationOptions): Promise<NativePptxDeck> {
    const contractJson = await this.native.extract(bytes, options)
    let input: unknown
    try {
      input = JSON.parse(contractJson)
    } catch {
      this.native.terminate()
      throw new PptxNativeContractError([{ path: '$', code: 'json.parse', message: 'native engine returned invalid JSON' }])
    }
    const result = validateNativePptx(input)
    if (!result.ok) {
      this.native.terminate()
      throw new PptxNativeContractError(result.issues)
    }
    return result.value
  }
}

function validateMutation(deck: NativePptxDeck, input: PptxNativeMutationRequestV1): PptxNativeMutationRequestV1 {
  if (!isRecord(input)) throw new TypeError('PPTX native mutation must be an object.')
  rejectUnknownKeys(input, ['expectedSourceRevision', 'operations'], 'mutation')
  const expectedSourceRevision = input.expectedSourceRevision
  const requestedOperations = input.operations
  if (deck.origin !== 'parsed' || !deck.sourceRevision || !/^rev-[0-9a-f]{64}$/.test(deck.sourceRevision)) {
    throw new NativeWasmError('UNSUPPORTED_SOURCE', 'PPTX mutation requires a parsed deck with an exact source revision.')
  }
  if (expectedSourceRevision !== deck.sourceRevision) {
    throw new NativeWasmError('STALE_REVISION', 'Mutation expectedSourceRevision must equal the extracted deck sourceRevision.')
  }
  if (!Array.isArray(requestedOperations) || requestedOperations.length === 0 || requestedOperations.length > MAX_OPERATIONS) {
    throw new RangeError(`PPTX mutation operations must contain 1..${MAX_OPERATIONS} items.`)
  }
  const elements = indexElements(deck)
  const operationIds = new Set<string>()
  const elementIds = new Set<string>()
  let paragraphs = 0
  let runs = 0
  let textCodeUnits = 0
  const operations = requestedOperations.map((operation, index) => {
    if (!isRecord(operation)) throw new TypeError(`PPTX mutation operation ${index} must be an object.`)
    validateId(operation.operationId, `operation ${index} operationId`)
    validateId(operation.elementId, `operation ${index} elementId`)
    if (operationIds.has(operation.operationId)) throw new TypeError(`PPTX mutation operationId ${JSON.stringify(operation.operationId)} is duplicated.`)
    if (elementIds.has(operation.elementId)) throw new TypeError(`PPTX mutation elementId ${JSON.stringify(operation.elementId)} is targeted more than once.`)
    operationIds.add(operation.operationId)
    elementIds.add(operation.elementId)
    if (typeof operation.expectedFingerprintSha256 !== 'string' || !SHA256.test(operation.expectedFingerprintSha256)) {
      throw new TypeError(`PPTX mutation operation ${index} has an invalid expectedFingerprintSha256.`)
    }
    const target = elements.get(operation.elementId)
    if (!target?.source) throw new NativeWasmError('UNSUPPORTED_TARGET', `PPTX mutation target ${JSON.stringify(operation.elementId)} is not a parsed source element.`)
    if (target.source.fingerprintSha256 !== operation.expectedFingerprintSha256) throw new NativeWasmError('STALE_FINGERPRINT', `PPTX mutation target ${JSON.stringify(operation.elementId)} has a stale fingerprint.`)
    if (target.compatibility.status === 'refused') throw new NativeWasmError('UNSUPPORTED_TARGET', `PPTX mutation target ${JSON.stringify(operation.elementId)} is refused.`)
    if (operation.kind === 'text.replace') {
      rejectUnknownKeys(operation, ['operationId', 'kind', 'elementId', 'expectedFingerprintSha256', 'paragraphs'], `operation ${index}`)
      if (target.kind !== 'text' && target.kind !== 'shape') throw new NativeWasmError('UNSUPPORTED_OPERATION', 'text.replace supports only extracted text and shape elements.')
      if (!Array.isArray(operation.paragraphs) || operation.paragraphs.length === 0) throw new TypeError(`PPTX mutation operation ${index} requires paragraphs.`)
      paragraphs += operation.paragraphs.length
      if (paragraphs > MAX_MUTATION_PARAGRAPHS) throw new RangeError(`PPTX mutation exceeds ${MAX_MUTATION_PARAGRAPHS} paragraphs.`)
      operation.paragraphs.forEach((paragraph, paragraphIndex) => {
        const counts = validateParagraph(paragraph, `operation ${index} paragraph ${paragraphIndex}`)
        runs += counts.runs
        textCodeUnits += counts.textCodeUnits
      })
      if (runs > MAX_MUTATION_RUNS || textCodeUnits > PPTX_NATIVE_RESOURCE_LIMITS.maxTotalTextCodeUnits) throw new RangeError('PPTX mutation text resource budget exceeded.')
      return {
        operationId: operation.operationId,
        kind: 'text.replace',
        elementId: operation.elementId,
        expectedFingerprintSha256: operation.expectedFingerprintSha256,
        paragraphs: operation.paragraphs.map(normalizeParagraph),
      } satisfies PptxNativeTextReplaceMutationV1
    }
    if (operation.kind === 'autoshape.update') {
      rejectUnknownKeys(operation, ['operationId', 'kind', 'elementId', 'expectedFingerprintSha256', 'autoShape'], `operation ${index}`)
      if (target.kind !== 'shape' || target.preset === undefined) throw new NativeWasmError('UNSUPPORTED_OPERATION', 'autoshape.update supports only exact extracted AutoShape elements.')
      validateAutoShape(operation.autoShape, `operation ${index} autoShape`)
      return {
        operationId: operation.operationId,
        kind: 'autoshape.update',
        elementId: operation.elementId,
        expectedFingerprintSha256: operation.expectedFingerprintSha256,
        autoShape: normalizeAutoShape(operation.autoShape),
      } satisfies PptxNativeAutoShapeUpdateMutationV1
    }
    throw new NativeWasmError('UNSUPPORTED_OPERATION', `Unsupported PPTX native mutation kind ${JSON.stringify(operation.kind)}.`)
  })
  return { expectedSourceRevision, operations }
}

function normalizeParagraph(value: unknown): PptxNativeExactParagraphV1 {
  const paragraph = value as Record<string, unknown>
  return {
    runs: (paragraph.runs as Array<Record<string, unknown>>).map((run) => ({
      text: run.text as string,
      bold: run.bold as boolean,
      italic: run.italic as boolean,
      fontSizeHundredthPt: run.fontSizeHundredthPt as number,
      color: run.color as string,
      fontFamily: run.fontFamily as string,
    })),
    align: paragraph.align as PptxNativeExactParagraphV1['align'],
    level: paragraph.level as number,
    bullet: false,
  }
}

function normalizeAutoShape(value: unknown): PptxNativeExactAutoShapeV1 {
  const shape = value as Record<string, unknown>
  const transform = shape.transform as Record<string, unknown>
  const stroke = shape.stroke as Record<string, unknown> | undefined
  return {
    transform: { x: transform.x as number, y: transform.y as number, cx: transform.cx as number, cy: transform.cy as number },
    preset: shape.preset as PptxNativeExactAutoShapeV1['preset'],
    ...(shape.fill === undefined ? {} : { fill: shape.fill as string }),
    ...(stroke === undefined ? {} : {
      stroke: {
        color: stroke.color as string,
        widthEmu: stroke.widthEmu as number,
        cap: stroke.cap as PptxNativeExactStrokeV1['cap'],
        join: stroke.join as PptxNativeExactStrokeV1['join'],
        dash: 'solid',
        ...(stroke.miterLimit === undefined ? {} : { miterLimit: stroke.miterLimit as number }),
      },
    }),
  }
}

function validateParagraph(value: unknown, path: string): { runs: number; textCodeUnits: number } {
  if (!isRecord(value)) throw new TypeError(`${path} must be an object.`)
  rejectUnknownKeys(value, ['runs', 'align', 'level', 'bullet'], path)
  if (!Array.isArray(value.runs) || value.runs.length > PPTX_NATIVE_RESOURCE_LIMITS.maxRunsPerParagraph) throw new RangeError(`${path} has an invalid run count.`)
  if (value.align !== 'left' && value.align !== 'center' && value.align !== 'right') throw new TypeError(`${path} has an unsupported align value.`)
  if (!Number.isSafeInteger(value.level) || (value.level as number) < 0 || (value.level as number) > 8 || value.bullet !== false) throw new TypeError(`${path} must have level 0..8 and bullet false.`)
  let textCodeUnits = 0
  value.runs.forEach((run, index) => {
    const runPath = `${path} run ${index}`
    if (!isRecord(run)) throw new TypeError(`${runPath} must be an object.`)
    rejectUnknownKeys(run, ['text', 'bold', 'italic', 'fontSizeHundredthPt', 'color', 'fontFamily'], runPath)
    if (typeof run.text !== 'string' || /[\t\r\n]/.test(run.text) || hasInvalidXmlOrSurrogate(run.text)) throw new TypeError(`${runPath} contains unsupported text.`)
    if (typeof run.bold !== 'boolean' || typeof run.italic !== 'boolean') throw new TypeError(`${runPath} requires exact bold and italic values.`)
    if (!Number.isSafeInteger(run.fontSizeHundredthPt) || (run.fontSizeHundredthPt as number) < 1 || (run.fontSizeHundredthPt as number) > 400_000) throw new RangeError(`${runPath} has an invalid font size.`)
    if (typeof run.color !== 'string' || !COLOR.test(run.color)) throw new TypeError(`${runPath} color must be canonical uppercase sRGB.`)
    if (typeof run.fontFamily !== 'string' || run.fontFamily.length === 0 || run.fontFamily.length > 256 || hasInvalidXmlOrSurrogate(run.fontFamily)) throw new TypeError(`${runPath} has an invalid font family.`)
    if (run.text.length > PPTX_NATIVE_RESOURCE_LIMITS.maxTextCodeUnits) throw new RangeError(`${runPath} exceeds the text resource limit.`)
    textCodeUnits += run.text.length
  })
  return { runs: value.runs.length, textCodeUnits }
}

function validateAutoShape(value: unknown, path: string): void {
  if (!isRecord(value)) throw new TypeError(`${path} must be an object.`)
  rejectUnknownKeys(value, ['transform', 'preset', 'fill', 'stroke'], path)
  if (!isRecord(value.transform)) throw new TypeError(`${path}.transform must be an object.`)
  rejectUnknownKeys(value.transform, ['x', 'y', 'cx', 'cy'], `${path}.transform`)
  for (const name of ['x', 'y', 'cx', 'cy'] as const) {
    const number = value.transform[name]
    if (!Number.isSafeInteger(number) || Object.is(number, -0) || ((name === 'cx' || name === 'cy') && (number as number) < 1)) throw new RangeError(`${path}.transform.${name} is outside the exact native range.`)
  }
  if (value.preset !== 'rect' && value.preset !== 'ellipse' && value.preset !== 'triangle' && value.preset !== 'diamond') throw new NativeWasmError('UNSUPPORTED_OPERATION', `${path}.preset is outside the native mutation subset.`)
  if (value.fill !== undefined && (typeof value.fill !== 'string' || !COLOR.test(value.fill))) throw new TypeError(`${path}.fill must be canonical uppercase sRGB.`)
  if (value.stroke === undefined) return
  if (!isRecord(value.stroke)) throw new TypeError(`${path}.stroke must be an object.`)
  rejectUnknownKeys(value.stroke, ['color', 'widthEmu', 'cap', 'join', 'dash', 'miterLimit'], `${path}.stroke`)
  if (typeof value.stroke.color !== 'string' || !COLOR.test(value.stroke.color) || !Number.isSafeInteger(value.stroke.widthEmu) || (value.stroke.widthEmu as number) < 0 || (value.stroke.widthEmu as number) > PPTX_NATIVE_RESOURCE_LIMITS.maxLineWidthEmu) throw new TypeError(`${path}.stroke has invalid color or width.`)
  if (!['flat', 'round', 'square'].includes(value.stroke.cap as string) || !['round', 'bevel', 'miter'].includes(value.stroke.join as string) || value.stroke.dash !== 'solid') throw new TypeError(`${path}.stroke is not complete exact solid-line metadata.`)
  if (value.stroke.join === 'miter') {
    if (!Number.isSafeInteger(value.stroke.miterLimit) || (value.stroke.miterLimit as number) < 0 || (value.stroke.miterLimit as number) > PPTX_NATIVE_RESOURCE_LIMITS.maxDrawingPercentage) throw new TypeError(`${path}.stroke miter join requires a bounded miterLimit.`)
  } else if (value.stroke.miterLimit !== undefined) throw new TypeError(`${path}.stroke non-miter join must omit miterLimit.`)
}

function indexElements(deck: NativePptxDeck): Map<string, NativeElement> {
  const result = new Map<string, NativeElement>()
  const visit = (elements: ReadonlyArray<NativeElement>) => elements.forEach((element) => {
    result.set(element.id, element)
    if (element.kind === 'group') visit(element.children)
  })
  deck.slides.forEach((slide) => visit(slide.elements))
  return result
}

function createBrowserWorker(workerUrl: string): NativeWasmWorker {
  if (typeof globalThis.Worker !== 'function') throw new NativeWasmError('WORKER_UNAVAILABLE', 'PPTX WASM requires a browser Worker; provide workerFactory in non-browser runtimes.', true)
  return new globalThis.Worker(workerUrl) as unknown as NativeWasmWorker
}

function assertPackageSize(bytes: Uint8Array, limit: number): void {
  if (!(bytes instanceof Uint8Array)) throw new TypeError('PPTX package must be a Uint8Array.')
  if (bytes.byteLength === 0) throw new RangeError('PPTX package must not be empty.')
  if (bytes.byteLength > limit) throw new RangeError(`PPTX package is ${bytes.byteLength} bytes; maxPackageBytes is ${limit}.`)
}

function boundedLimit(value: number | undefined, fallback: number, hardMaximum: number, name: string): number {
  const resolved = value ?? fallback
  if (!Number.isSafeInteger(resolved) || resolved <= 0 || resolved > hardMaximum) throw new RangeError(`${name} must be a positive safe integer no greater than ${hardMaximum}.`)
  return resolved
}

function outerRevision(nativeRevision: string): string { return `sha256:${nativeRevision.slice('rev-'.length)}` }

function validateId(value: unknown, label: string): asserts value is string {
  if (typeof value !== 'string' || !NATIVE_ID.test(value)) throw new TypeError(`PPTX mutation ${label} is invalid.`)
}

function isRecord(value: unknown): value is Record<string, unknown> { return value !== null && typeof value === 'object' && !Array.isArray(value) }

function rejectUnknownKeys(value: Record<string, unknown>, allowed: readonly string[], path: string): void {
  const expected = new Set(allowed)
  const unknown = Object.keys(value).find((key) => !expected.has(key))
  if (unknown !== undefined) throw new TypeError(`PPTX native ${path} contains unknown field ${JSON.stringify(unknown)}.`)
}

function hasInvalidXmlOrSurrogate(value: string): boolean {
  for (let index = 0; index < value.length; index++) {
    const code = value.charCodeAt(index)
    if (code === 0x9 || code === 0xa || code === 0xd || code >= 0x20 && code <= 0xd7ff || code >= 0xe000 && code <= 0xfffd) continue
    if (code >= 0xd800 && code <= 0xdbff && index + 1 < value.length) {
      const next = value.charCodeAt(index + 1)
      if (next >= 0xdc00 && next <= 0xdfff) { index++; continue }
    }
    return true
  }
  return false
}

function toUrl(value: string | URL | undefined, fallback: URL): string {
  if (value === undefined) return fallback.href
  const resolved = typeof value === 'string' ? value : value.href
  if (resolved.length === 0) throw new TypeError('PPTX WASM asset URLs must not be empty')
  return resolved
}
