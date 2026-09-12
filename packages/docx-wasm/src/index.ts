import {
  NativeWasmError,
  createNativeWasmClient,
  type NativeWasmAbortSignal,
  type NativeWasmClient,
  type NativeWasmWorker,
} from '@injoffice/native-runtime'
import {
  OFFICE_MUTATION_PROTOCOL,
  OFFICE_MUTATION_VERSION,
  NativeDocxValidationError,
  decodeNativeDocxDocument,
  decodeNativeDocxJson,
  decodeNativeDocxResolvedLayout,
  createNativeDocxEquationPreviewsV1,
  type NativeDocxEquationPreviewV1,
  decodeNativeDocxEquationContextNoticesV1,
  type NativeDocxEquationContextNoticeV1,
  decodeNativeDocxNestedTableOmissionsV1,
  type NativeDocxNestedTableOmissionsV1,
  type NativeDocxResolvedLayoutInputV1,
  type NativeDocxDocumentV1,
  type NativeDocxOfficeMutationEnvelopeV1,
  type NativeDocxTextMutationPayloadV1,
} from '@injoffice/docs/native-docx'

export const DOCX_WASM_NATIVE_MAX_PACKAGE_BYTES = 128 * 1024 * 1024
export const DOCX_WASM_NATIVE_MAX_MUTATION_PAYLOAD_BYTES = 8 * 1024 * 1024
export const DEFAULT_DOCX_WASM_MAX_PACKAGE_BYTES = 32 * 1024 * 1024
export const DEFAULT_DOCX_WASM_MAX_MUTATION_PAYLOAD_BYTES = 3 * 1024 * 1024
export const DOCX_WASM_NATIVE_MAX_MUTATIONS = 10_000
export const DOCX_WASM_NATIVE_MAX_TEXT_CODE_UNITS = 1_048_576

export interface DocxWasmAssetUrls {
  workerUrl: string
  wasmUrl: string
  goRuntimeUrl: string
}

export interface DocxWasmClientOptions {
  workerUrl?: string | URL
  wasmUrl?: string | URL
  goRuntimeUrl?: string | URL
  operationTimeoutMs?: number
  /** Browser-side ceiling, capped by the native 128 MiB package limit. */
  maxPackageBytes?: number
  /** UTF-8 payload ceiling, capped by the native 8 MiB mutation limit. */
  maxMutationPayloadBytes?: number
  workerFactory?: (workerUrl: string) => NativeWasmWorker
}

export interface DocxWasmOperationOptions {
  signal?: NativeWasmAbortSignal
}

export interface DocxWasmClient {
  inspectPartialContent(bytes:Uint8Array,options?:DocxWasmOperationOptions):Promise<{document:NativeDocxDocumentV1;resolved_layout:NativeDocxResolvedLayoutInputV1;equations?:NativeDocxEquationPreviewV1[];equation_context_notices?:NativeDocxEquationContextNoticeV1[];nested_table_omissions?:NativeDocxNestedTableOmissionsV1}>
  extract(bytes: Uint8Array, options?: DocxWasmOperationOptions): Promise<NativeDocxDocumentV1>
  apply(
    original: Uint8Array,
    document: NativeDocxDocumentV1,
    envelope: NativeDocxOfficeMutationEnvelopeV1,
    options?: DocxWasmOperationOptions,
  ): Promise<Uint8Array>
  terminate(): void
}

export function resolveDocxWasmAssetUrls(options: Pick<DocxWasmClientOptions, 'workerUrl' | 'wasmUrl' | 'goRuntimeUrl'> = {}): DocxWasmAssetUrls {
  return {
    workerUrl: toUrl(options.workerUrl, new URL('./docxnative.worker.js', import.meta.url)),
    wasmUrl: toUrl(options.wasmUrl, new URL('./docxnative.wasm', import.meta.url)),
    goRuntimeUrl: toUrl(options.goRuntimeUrl, new URL('./wasm_exec.js', import.meta.url)),
  }
}

export function createDocxWasmClient(options: DocxWasmClientOptions = {}): DocxWasmClient {
  const assets = resolveDocxWasmAssetUrls(options)
  const workerFactory = options.workerFactory ?? createBrowserWorker
  const maxPackageBytes = boundedLimit(options.maxPackageBytes, DEFAULT_DOCX_WASM_MAX_PACKAGE_BYTES, DOCX_WASM_NATIVE_MAX_PACKAGE_BYTES, 'maxPackageBytes')
  const maxMutationPayloadBytes = boundedLimit(options.maxMutationPayloadBytes, DEFAULT_DOCX_WASM_MAX_MUTATION_PAYLOAD_BYTES, DOCX_WASM_NATIVE_MAX_MUTATION_PAYLOAD_BYTES, 'maxMutationPayloadBytes')
  const native = createNativeWasmClient({
    format: 'docx',
    workerFactory: () => workerFactory(assets.workerUrl),
    assets: { wasmUrl: assets.wasmUrl, goRuntimeUrl: assets.goRuntimeUrl },
    operationTimeoutMs: options.operationTimeoutMs,
  })
  return new DocxWasmClientImpl(native, maxPackageBytes, maxMutationPayloadBytes)
}

class DocxWasmClientImpl implements DocxWasmClient {
  constructor(
    private readonly native: NativeWasmClient,
    private readonly maxPackageBytes: number,
    private readonly maxMutationPayloadBytes: number,
  ) {}

  extract(bytes: Uint8Array, options: DocxWasmOperationOptions = {}): Promise<NativeDocxDocumentV1> {
    assertPackageSize(bytes, this.maxPackageBytes)
    return this.extractValidated(bytes, options)
  }

  async inspectPartialContent(bytes:Uint8Array,options:DocxWasmOperationOptions={}):Promise<{document:NativeDocxDocumentV1;resolved_layout:NativeDocxResolvedLayoutInputV1;equations?:NativeDocxEquationPreviewV1[];equation_context_notices?:NativeDocxEquationContextNoticeV1[];nested_table_omissions?:NativeDocxNestedTableOmissionsV1}>{
    assertPackageSize(bytes,this.maxPackageBytes)
    if(options.signal?.aborted){const error=new Error('Partial inspection aborted');error.name='AbortError';throw error}
    const snapshot=new Uint8Array(bytes)
    const hash='sha256:'+Array.from(new Uint8Array(await globalThis.crypto.subtle.digest('SHA-256',snapshot)),b=>b.toString(16).padStart(2,'0')).join('')
    const json=await this.native.inspect(snapshot,options)
    try{
      if(json.length>16*1024*1024||new TextEncoder().encode(json).byteLength>16*1024*1024)throw new TypeError('Partial source exceeds response budget')
      const value=JSON.parse(json) as Record<string,unknown>
      if(!value||!['document,package_sha256,protocol,resolved_layout,version','document,equations,package_sha256,protocol,resolved_layout,version','document,equation_context_notices,equations,package_sha256,protocol,resolved_layout,version'].includes(Object.keys(value).filter(k=>k!=='nested_table_omissions').sort().join(','))||value.protocol!=='injoffice.docx.partial-source'||value.version!==1||value.package_sha256!==hash)throw new TypeError('Partial source package identity mismatch')
      const document=decodeNativeDocxDocument(value.document),layout=decodeNativeDocxResolvedLayout(value.resolved_layout)
      if(!document.ok||!layout.ok||document.value.source.package_sha256!==hash||document.value.document_id!==layout.value.document_id||document.value.revision!==layout.value.revision||document.value.source.main_part!==layout.value.source_parts.main_part)throw new TypeError('Invalid joined partial source models')
      const notices=decodeNativeDocxEquationContextNoticesV1(document.value,layout.value,value.equation_context_notices)
      const nested=decodeNativeDocxNestedTableOmissionsV1(document.value,value.nested_table_omissions)
      return {document:document.value,resolved_layout:layout.value,...(value.equations===undefined?{}:{equations:createNativeDocxEquationPreviewsV1(document.value,layout.value,value.equations,notices)}),...(value.equation_context_notices===undefined?{}:{equation_context_notices:notices}),...(value.nested_table_omissions===undefined?{}:{nested_table_omissions:nested})}
    }catch(error){this.native.terminate();throw error}
  }

  apply(
    original: Uint8Array,
    document: NativeDocxDocumentV1,
    envelope: NativeDocxOfficeMutationEnvelopeV1,
    options: DocxWasmOperationOptions = {},
  ): Promise<Uint8Array> {
    assertPackageSize(original, this.maxPackageBytes)
    const decodedDocument = decodeNativeDocxDocument(document)
    if (!decodedDocument.ok) throw new NativeDocxValidationError(decodedDocument.issues)
    const payload = validateEnvelope(decodedDocument.value, envelope)
    const encoded = JSON.stringify(payload)
    const payloadBytes = new TextEncoder().encode(encoded).byteLength
    if (payloadBytes > this.maxMutationPayloadBytes) {
      throw new RangeError(`DOCX mutation payload is ${payloadBytes} UTF-8 bytes; maxMutationPayloadBytes is ${this.maxMutationPayloadBytes}.`)
    }
    return this.native.apply(original, encoded, document.source.package_sha256, options)
  }

  terminate(): void {
    this.native.terminate()
  }

  private async extractValidated(bytes: Uint8Array, options: DocxWasmOperationOptions): Promise<NativeDocxDocumentV1> {
    const contractJson = await this.native.extract(bytes, options)
    const decoded = decodeNativeDocxJson(contractJson)
    if (!decoded.ok) {
      this.native.terminate()
      throw new NativeDocxValidationError(decoded.issues)
    }
    return decoded.value
  }
}

function validateEnvelope(document: NativeDocxDocumentV1, value: NativeDocxOfficeMutationEnvelopeV1): NativeDocxTextMutationPayloadV1 {
  const envelope = plainObject(value, 'DOCX mutation envelope')
  exactKeys(envelope, ['protocol', 'version', 'format', 'mutation_id', 'expected_revision', 'payload'], 'DOCX mutation envelope')
  if (envelope.protocol !== OFFICE_MUTATION_PROTOCOL || envelope.version !== OFFICE_MUTATION_VERSION || envelope.format !== 'docx') {
    throw new TypeError('DOCX mutation envelope protocol, version, or format does not match.')
  }
  if (typeof envelope.mutation_id !== 'string' || !/^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/.test(envelope.mutation_id)) {
    throw new TypeError('DOCX mutation envelope mutation_id is invalid.')
  }
  if (envelope.expected_revision !== document.source.package_sha256) {
    throw new NativeWasmError('STALE_REVISION', 'DOCX mutation envelope expected_revision must equal the extracted source.package_sha256.')
  }
  const payload = plainObject(envelope.payload, 'DOCX mutation payload')
  exactKeys(payload, ['mutations'], 'DOCX mutation payload')
  if (!Array.isArray(payload.mutations) || payload.mutations.length < 1 || payload.mutations.length > DOCX_WASM_NATIVE_MAX_MUTATIONS) {
    throw new RangeError(`DOCX mutation count must be 1..${DOCX_WASM_NATIVE_MAX_MUTATIONS}.`)
  }
  const targets = documentTargets(document)
  const seenTargets = new Set<string>()
  const seenTextNodes = new Set<string>()
  const mutations: NativeDocxTextMutationPayloadV1['mutations'] = payload.mutations.map((value, index) => {
    const mutation = plainObject(value, `DOCX mutation ${index}`)
    exactKeys(mutation, ['target_kind', 'target_id', 'expected_xml_sha256', 'text'], `DOCX mutation ${index}`)
    const targetKind = mutation.target_kind
    if ((targetKind !== 'paragraph' && targetKind !== 'run') || typeof mutation.target_id !== 'string' || typeof mutation.expected_xml_sha256 !== 'string' || typeof mutation.text !== 'string') {
      throw new TypeError(`DOCX mutation ${index} has invalid field types.`)
    }
    const targetKey = `${targetKind}\0${mutation.target_id}`
    const target = targets.get(targetKey)
    if (!target || target.anchorSha256 !== mutation.expected_xml_sha256) {
      throw new NativeWasmError('STALE_TARGET', `DOCX mutation ${index} does not match an extracted ${targetKind} text anchor.`)
    }
    if (seenTargets.has(targetKey)) throw new NativeWasmError('DUPLICATE_TARGET', `DOCX mutation ${index} repeats the same native target.`)
    if (seenTextNodes.has(target.textNodeKey)) throw new NativeWasmError('OVERLAPPING_TARGETS', `DOCX mutation ${index} overlaps another paragraph or run target.`)
    seenTargets.add(targetKey)
    seenTextNodes.add(target.textNodeKey)
    if (mutation.text.length > DOCX_WASM_NATIVE_MAX_TEXT_CODE_UNITS || !xmlTextValid(mutation.text)) {
      throw new TypeError(`DOCX mutation ${index} text is not bounded XML 1.0 character data.`)
    }
    return {
      target_kind: targetKind,
      target_id: mutation.target_id,
      expected_xml_sha256: mutation.expected_xml_sha256,
      text: mutation.text,
    }
  })
  return { mutations }
}

type DocxTextTarget = { anchorSha256: string; textNodeKey: string }

function documentTargets(document: NativeDocxDocumentV1): Map<string, DocxTextTarget> {
  const targets = new Map<string, DocxTextTarget>()
  const stories = [document.body, ...document.headers, ...document.footers, ...document.notes, ...document.comment_stories]
  for (const story of stories) {
    for (const block of story.blocks) {
      const paragraphs = block.paragraph ? [block.paragraph] : block.table?.rows.flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs)) ?? []
      for (const paragraph of paragraphs) {
        const textRuns = paragraph.runs.filter((run) => run.kind === 'text')
        for (const run of textRuns) {
          targets.set(`run\0${run.id}`, { anchorSha256: run.anchor.xml_sha256, textNodeKey: run.id })
        }
        if (paragraph.runs.length === 1 && textRuns.length === 1) {
          targets.set(`paragraph\0${paragraph.id}`, { anchorSha256: paragraph.anchor.xml_sha256, textNodeKey: textRuns[0]!.id })
        }
      }
    }
  }
  return targets
}

function xmlTextValid(value: string): boolean {
  for (const scalar of value) {
    const code = scalar.codePointAt(0)!
    if (code === 0x9 || code === 0xa || (code >= 0x20 && code <= 0xd7ff) || (code >= 0xe000 && code <= 0xfffd) || (code >= 0x10000 && code <= 0x10ffff)) continue
    return false
  }
  return true
}

function createBrowserWorker(workerUrl: string): NativeWasmWorker {
  if (typeof globalThis.Worker !== 'function') {
    throw new NativeWasmError('WORKER_UNAVAILABLE', 'DOCX WASM requires a browser Worker; provide workerFactory in non-browser runtimes.', true)
  }
  return new globalThis.Worker(workerUrl) as unknown as NativeWasmWorker
}

function assertPackageSize(bytes: Uint8Array, limit: number): void {
  if (!(bytes instanceof Uint8Array)) throw new TypeError('DOCX package must be a Uint8Array.')
  if (bytes.byteLength === 0) throw new RangeError('DOCX package must not be empty.')
  if (bytes.byteLength > limit) throw new RangeError(`DOCX package is ${bytes.byteLength} bytes; maxPackageBytes is ${limit}.`)
}

function boundedLimit(value: number | undefined, fallback: number, hardMaximum: number, name: string): number {
  const resolved = value ?? fallback
  if (!Number.isSafeInteger(resolved) || resolved <= 0 || resolved > hardMaximum) {
    throw new RangeError(`${name} must be a positive safe integer no greater than ${hardMaximum}.`)
  }
  return resolved
}

function plainObject(value: unknown, name: string): Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value) || Object.getPrototypeOf(value) !== Object.prototype) {
    throw new TypeError(`${name} must be a plain object.`)
  }
  if (Object.getOwnPropertySymbols(value).length > 0) throw new TypeError(`${name} fields are invalid.`)
  for (const descriptor of Object.values(Object.getOwnPropertyDescriptors(value))) {
    if (!('value' in descriptor)) throw new TypeError(`${name} must not contain accessor properties.`)
  }
  return value as Record<string, unknown>
}

function exactKeys(value: Record<string, unknown>, wanted: readonly string[], name: string): void {
  const actual = Object.getOwnPropertyNames(value).sort()
  const expected = [...wanted].sort()
  if (actual.length !== expected.length || actual.some((key, index) => key !== expected[index])) throw new TypeError(`${name} fields are invalid.`)
}

function toUrl(value: string | URL | undefined, fallback: URL): string {
  if (value === undefined) return fallback.href
  const resolved = typeof value === 'string' ? value : value.href
  if (resolved.length === 0) throw new TypeError('DOCX WASM asset URLs must not be empty')
  return resolved
}
