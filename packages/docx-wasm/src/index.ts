import { validateInsert, type DocxInsertPayload } from './insert.js'
import { validateSectionPageMutation } from './sectionPage.js'
import {decodeNativeDocxTextboxGeometryV1,type NativeDocxTextboxGeometryEvidenceV1} from '@injoffice/docs/native-docx'
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
  decodeNativeDocxReviewEvidenceV1,
  decodeNativeDocxTextboxEvidenceV1,
  type NativeDocxTextboxEvidenceV1,
  decodeNativeDocxNestedTextEvidenceV1,
  type NativeDocxNestedTextEvidenceV1,
  decodeNativeDocxPartialTableTextContextsV1,
  type NativeDocxPartialTableTextContextV1,
  type NativeDocxReviewEvidenceV1,
  type NativeDocxEquationPreviewV1,
  decodeNativeDocxEquationContextNoticesV1,
  type NativeDocxEquationContextNoticeV1,
  decodeNativeDocxNestedTableOmissionsV1,
  type NativeDocxNestedTableOmissionsV1,
  type NativeDocxResolvedLayoutInputV1,
  type NativeDocxDocumentV1,
  type NativeDocxOfficeMutationEnvelopeV1,
  type NativeDocxRunFormatPayloadV1,
  type NativeDocxParagraphPropertyPatchV1,
  type NativeDocxRunPropertyPatchV1,
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
  inspectPartialContent(bytes:Uint8Array,options?:DocxWasmOperationOptions):Promise<{document:NativeDocxDocumentV1;resolved_layout:NativeDocxResolvedLayoutInputV1;equations?:NativeDocxEquationPreviewV1[];equation_context_notices?:NativeDocxEquationContextNoticeV1[];nested_table_omissions?:NativeDocxNestedTableOmissionsV1;review_changes?:NativeDocxReviewEvidenceV1;table_text_contexts?:NativeDocxPartialTableTextContextV1[];nested_text?:NativeDocxNestedTextEvidenceV1[];textbox_inventory?:NativeDocxTextboxEvidenceV1;textbox_geometry?:NativeDocxTextboxGeometryEvidenceV1}>
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

  async inspectPartialContent(bytes:Uint8Array,options:DocxWasmOperationOptions={}):Promise<{document:NativeDocxDocumentV1;resolved_layout:NativeDocxResolvedLayoutInputV1;equations?:NativeDocxEquationPreviewV1[];equation_context_notices?:NativeDocxEquationContextNoticeV1[];nested_table_omissions?:NativeDocxNestedTableOmissionsV1;review_changes?:NativeDocxReviewEvidenceV1;table_text_contexts?:NativeDocxPartialTableTextContextV1[];nested_text?:NativeDocxNestedTextEvidenceV1[];textbox_inventory?:NativeDocxTextboxEvidenceV1;textbox_geometry?:NativeDocxTextboxGeometryEvidenceV1}>{
    assertPackageSize(bytes,this.maxPackageBytes)
    if(options.signal?.aborted){const error=new Error('Partial inspection aborted');error.name='AbortError';throw error}
    const snapshot=new Uint8Array(bytes)
    const hash='sha256:'+Array.from(new Uint8Array(await globalThis.crypto.subtle.digest('SHA-256',snapshot)),b=>b.toString(16).padStart(2,'0')).join('')
    const json=await this.native.inspect(snapshot,options)
    try{
      if(json.length>16*1024*1024||new TextEncoder().encode(json).byteLength>16*1024*1024)throw new TypeError('Partial source exceeds response budget')
      const value=JSON.parse(json) as Record<string,unknown>
      if(!value||!['document,package_sha256,protocol,resolved_layout,version','document,equations,package_sha256,protocol,resolved_layout,version','document,equation_context_notices,equations,package_sha256,protocol,resolved_layout,version'].includes(Object.keys(value).filter(k=>k!=='nested_table_omissions'&&k!=='review_changes'&&k!=='table_text_contexts'&&k!=='textbox_geometry'&&k!=='textbox_inventory'&&k!=='nested_text').sort().join(','))||value.protocol!=='injoffice.docx.partial-source'||value.version!==1||value.package_sha256!==hash)throw new TypeError('Partial source package identity mismatch')
      const document=decodeNativeDocxDocument(value.document),layout=decodeNativeDocxResolvedLayout(value.resolved_layout)
      if(!document.ok||!layout.ok||document.value.source.package_sha256!==hash||document.value.document_id!==layout.value.document_id||document.value.revision!==layout.value.revision||document.value.source.main_part!==layout.value.source_parts.main_part)throw new TypeError('Invalid joined partial source models')
      const notices=decodeNativeDocxEquationContextNoticesV1(document.value,layout.value,value.equation_context_notices)
      const nested=decodeNativeDocxNestedTableOmissionsV1(document.value,value.nested_table_omissions)
      const tableContexts=decodeNativeDocxPartialTableTextContextsV1(document.value,layout.value,value.table_text_contexts)
      const textboxes=decodeNativeDocxTextboxEvidenceV1(document.value,value.textbox_inventory)
      const nestedText=decodeNativeDocxNestedTextEvidenceV1(document.value,layout.value,nested,tableContexts,value.nested_text)
      const review=decodeNativeDocxReviewEvidenceV1(document.value,value.review_changes)
      const geometry=decodeNativeDocxTextboxGeometryV1(document.value,value.textbox_geometry)
      return {...(value.textbox_geometry===undefined?{}:{textbox_geometry:geometry}),...(value.nested_text===undefined?{}:{nested_text:nestedText}),...(value.textbox_inventory===undefined?{}:{textbox_inventory:textboxes}),...(value.table_text_contexts===undefined?{}:{table_text_contexts:tableContexts}),...(value.review_changes===undefined?{}:{review_changes:review}),document:document.value,resolved_layout:layout.value,...(value.equations===undefined?{}:{equations:createNativeDocxEquationPreviewsV1(document.value,layout.value,value.equations,notices)}),...(value.equation_context_notices===undefined?{}:{equation_context_notices:notices}),...(value.nested_table_omissions===undefined?{}:{nested_table_omissions:nested})}
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

function validateEnvelope(document: NativeDocxDocumentV1, value: NativeDocxOfficeMutationEnvelopeV1): NativeDocxOfficeMutationEnvelopeV1['payload'] {
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
  if (payload.mutations.some(value => { const m=plainObject(value, 'DOCX mutation'); return m.operation==='page_break.insert'||m.image!==undefined })) return validateInsert(document, payload.mutations)
  if (payload.mutations.some(value => plainObject(value, 'DOCX mutation').operation !== undefined)) {
    if (payload.mutations.length !== 1) throw new TypeError('Structural edits require one mutation.')
    const mutation = plainObject(payload.mutations[0], 'DOCX structural mutation')
    if (mutation.operation === 'hyperlink.set') {
      allowedKeys(mutation, ['target_kind', 'target_id', 'expected_xml_sha256', 'operation', 'hyperlink', 'range'], 'DOCX hyperlink mutation')
      const paragraph = document.body.blocks.find(block => block.paragraph && (mutation.target_kind === 'paragraph' ? block.paragraph.id === mutation.target_id : block.paragraph.runs.some(run => run.id === mutation.target_id)))?.paragraph
      const run = mutation.target_kind === 'run' ? paragraph?.runs.find(run => run.id === mutation.target_id) : undefined
      const target = mutation.target_kind === 'paragraph' ? paragraph : run
      if (!paragraph || !target || target.anchor.xml_sha256 !== mutation.expected_xml_sha256) throw new NativeWasmError('STALE_TARGET', 'The hyperlink target anchor changed.')
      if (!paragraph.edit_policy.allowed_operations.includes('hyperlink.set') || (run && !run.can_edit_hyperlink)) throw new TypeError('This text cannot be linked safely.')
      const link = plainObject(mutation.hyperlink, 'DOCX hyperlink')
      allowedKeys(link, ['url', 'expected_xml_sha256'], 'DOCX hyperlink')
      if (link.url !== null) {
        if (typeof link.url !== 'string' || link.url.length > 2048 || /[\r\n\t \\<>]/.test(link.url)) throw new TypeError('Enter a web or email link.')
        let parsed: URL
        try { parsed = new URL(link.url) } catch { throw new TypeError('Enter a web or email link.') }
        if (!['http:', 'https:', 'mailto:'].includes(parsed.protocol) || parsed.username || parsed.password || (parsed.protocol === 'mailto:' && !parsed.pathname)) throw new TypeError('Enter a web or email link.')
      }
      if (link.expected_xml_sha256 !== run?.hyperlink?.anchor.xml_sha256) throw new NativeWasmError('STALE_TARGET', 'The hyperlink wrapper anchor changed.')
      const text = run ? run.text ?? '' : paragraph.runs.map(run => run.text ?? '').join('')
      let range: {start_utf16: number; end_utf16: number} | undefined
      if (mutation.range !== undefined) {
        const span = plainObject(mutation.range, 'DOCX hyperlink range')
        exactKeys(span, ['start_utf16', 'end_utf16'], 'DOCX hyperlink range')
        const start = span.start_utf16 as number, end = span.end_utf16 as number
        if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end <= start || end > text.length || !xmlTextValid(text.slice(0, start)) || !xmlTextValid(text.slice(0, end))) throw new TypeError('Select text on UTF-16 character boundaries.')
        range = {start_utf16: start, end_utf16: end}
      }
      if (!text.length) throw new TypeError('Select nonempty text to link.')
      return {mutations: [{target_kind: run ? 'run' : 'paragraph', target_id: target.id, expected_xml_sha256: target.anchor.xml_sha256, operation: 'hyperlink.set', hyperlink: {url: link.url as string | null, ...(run?.hyperlink ? {expected_xml_sha256: run.hyperlink.anchor.xml_sha256} : {})}, ...(range ? {range} : {})}]}
    }
    if (mutation.operation === 'section.page.patch') return validateSectionPageMutation(document, mutation)
    const split = mutation.operation === 'paragraph.split'
    exactKeys(mutation, ['target_kind', 'target_id', 'expected_xml_sha256', 'operation', split ? 'split' : 'text'], 'DOCX structural mutation')
    if (mutation.target_kind !== 'paragraph' || (!split && (mutation.operation !== 'block.insert_after' || mutation.text !== ''))) throw new TypeError('Unsupported DOCX structural operation.')
    const paragraph = document.body.blocks.find(block => block.paragraph?.id === mutation.target_id)?.paragraph
    if (!paragraph || paragraph.anchor.xml_sha256 !== mutation.expected_xml_sha256) throw new NativeWasmError('STALE_TARGET', 'The body paragraph anchor changed.')
    if (!paragraph.edit_policy.allowed_operations.includes(mutation.operation as 'paragraph.split' | 'block.insert_after')) throw new TypeError('This paragraph does not allow the structural operation.')
    const anchor = {target_kind: 'paragraph' as const, target_id: paragraph.id, expected_xml_sha256: paragraph.anchor.xml_sha256}
    if (!split) return {mutations: [{...anchor, operation: 'block.insert_after', text: ''}]}
    const selector = plainObject(mutation.split, 'DOCX split')
    exactKeys(selector, ['run_id', 'offset_utf16'], 'DOCX split')
    const run = paragraph.runs.find(run => run.id === selector.run_id)
    const offset = selector.offset_utf16
    if (!run || typeof run.text !== 'string' || !Number.isSafeInteger(offset) || (offset as number) < 0 || (offset as number) > run.text.length || !xmlTextValid(run.text.slice(0, offset as number)) || !xmlTextValid(run.text.slice(offset as number))) throw new TypeError('Split must be a UTF-16 character boundary in a text run.')
    return {mutations: [{...anchor, operation: 'paragraph.split', split: {run_id: run.id, offset_utf16: offset as number}}]}
  }
  const targets = documentTargets(document)
  const seenTargets = new Set<string>()
  const seenTextNodes = new Set<string>()
  const formatting = payload.mutations.filter((value) => plainObject(value, 'DOCX mutation').properties !== undefined)
  if (formatting.length !== 0 && formatting.length !== payload.mutations.length) {
    throw new TypeError('A DOCX transaction may not mix text replacements with run-property patches.')
  }
  if (formatting.length !== 0) return { mutations: payload.mutations.map((value, index) => validateFormatMutation(document, targets, seenTargets, value, index)) }
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

const DOCX_ALIGNMENT_VALUES = ['left', 'center', 'right', 'both', 'distribute']
const DOCX_UNDERLINE_VALUES = ['none', 'single', 'double', 'words']
const DOCX_HIGHLIGHT_VALUES = ['none', 'black', 'blue', 'cyan', 'darkBlue', 'darkCyan', 'darkGray', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'green', 'lightGray', 'magenta', 'red', 'white', 'yellow']

/**
 * Run formatting is validated against the extracted contract before it leaves
 * the browser: the target must be one the extraction issued, its anchor must
 * be the one it issued, and every property must be inside the subset native
 * extraction reads back. The engine re-proves all of this and owns the
 * refusal; this boundary only keeps an obviously invalid request local.
 */
function validateFormatMutation(document: NativeDocxDocumentV1, targets: Map<string, DocxTextTarget>, seenTargets: Set<string>, value: unknown, index: number): NativeDocxRunFormatPayloadV1['mutations'][number] {
  const mutation = plainObject(value, `DOCX mutation ${index}`)
  allowedKeys(mutation, ['target_kind', 'target_id', 'expected_xml_sha256', 'properties', 'range'], `DOCX mutation ${index}`)
  const targetKind = mutation.target_kind
  if ((targetKind !== 'paragraph' && targetKind !== 'run') || typeof mutation.target_id !== 'string' || typeof mutation.expected_xml_sha256 !== 'string') {
    throw new TypeError(`DOCX mutation ${index} has invalid field types.`)
  }
  const targetKey = `${targetKind}\0${mutation.target_id}`
  const anchorSha256 = targetKind === 'paragraph' ? paragraphAnchors(document).get(mutation.target_id) : targets.get(targetKey)?.anchorSha256
  if (!anchorSha256 || anchorSha256 !== mutation.expected_xml_sha256) {
    throw new NativeWasmError('STALE_TARGET', `DOCX mutation ${index} does not match an extracted ${targetKind} anchor.`)
  }
  if (seenTargets.has(targetKey)) throw new NativeWasmError('DUPLICATE_TARGET', `DOCX mutation ${index} repeats the same native target.`)
  seenTargets.add(targetKey)
  const properties = plainObject(mutation.properties, `DOCX mutation ${index} properties`)
  allowedKeys(properties, ['bold', 'italic', 'underline', 'font_family', 'font_size_half_points', 'color', 'highlight', 'alignment', 'spacing_before_twips', 'spacing_after_twips', 'indent_left_twips', 'indent_right_twips', 'first_line_twips', 'hanging_twips', 'line_spacing', 'line_rule'], `DOCX mutation ${index} properties`)
  const paragraphKeys = ['alignment', 'spacing_before_twips', 'spacing_after_twips', 'indent_left_twips', 'indent_right_twips', 'first_line_twips', 'hanging_twips', 'line_spacing', 'line_rule']
  if (Object.keys(properties).some(key => paragraphKeys.includes(key))) {
    allowedKeys(properties, paragraphKeys, `DOCX mutation ${index} paragraph properties`)
    if (targetKind !== 'paragraph' || mutation.range !== undefined) throw new TypeError('Paragraph properties need a whole paragraph target.')
    if (properties.alignment !== undefined && !DOCX_ALIGNMENT_VALUES.includes(properties.alignment as string)) throw new TypeError('Invalid paragraph alignment.')
    for (const key of paragraphKeys.slice(1)) {
      const value = properties[key]
      if (value === undefined || value === null) continue
      if (key === 'line_rule') {
        if (!['auto', 'exact', 'atLeast'].includes(value as string)) throw new TypeError('Invalid line rule.')
      } else if (!Number.isSafeInteger(value) || (value as number) < (key.startsWith('indent_') ? -31680 : 0) || (value as number) > 31680) throw new TypeError(`Invalid paragraph ${key}.`)
    }
    if (properties.first_line_twips != null && properties.hanging_twips != null) throw new TypeError('First line and hanging are mutually exclusive.')
    return { target_kind: targetKind, target_id: mutation.target_id, expected_xml_sha256: mutation.expected_xml_sha256, properties: properties as NativeDocxParagraphPropertyPatchV1 }
  }
  const patch: NativeDocxRunPropertyPatchV1 = {}
  for (const name of ['bold', 'italic'] as const) {
    if (properties[name] === undefined) continue
    if (typeof properties[name] !== 'boolean') throw new TypeError(`DOCX mutation ${index} ${name} must be a boolean.`)
    patch[name] = properties[name]
  }
  if (properties.underline !== undefined) {
    if (!DOCX_UNDERLINE_VALUES.includes(properties.underline as string)) throw new TypeError(`DOCX mutation ${index} underline is outside the written subset.`)
    patch.underline = properties.underline as NativeDocxRunPropertyPatchV1['underline']
  }
  if (properties.highlight !== undefined) {
    if (!DOCX_HIGHLIGHT_VALUES.includes(properties.highlight as string)) throw new TypeError(`DOCX mutation ${index} highlight is outside the written subset.`)
    patch.highlight = properties.highlight as string
  }
  if (properties.font_family !== undefined) {
    const family = properties.font_family
    if (typeof family !== 'string' || family.length === 0 || family.length > 64 || /[<>&"'\u0000-\u001f\u007f]/.test(family)) throw new TypeError(`DOCX mutation ${index} font_family is not a bounded font name.`)
    patch.font_family = family
  }
  if (properties.font_size_half_points !== undefined) {
    const size = properties.font_size_half_points
    if (!Number.isSafeInteger(size) || (size as number) < 2 || (size as number) > 3276) throw new TypeError(`DOCX mutation ${index} font_size_half_points must be a whole 2..3276 value.`)
    patch.font_size_half_points = size as number
  }
  if (properties.color !== undefined) {
    if (typeof properties.color !== 'string' || !/^[0-9A-Fa-f]{6}$/.test(properties.color)) throw new TypeError(`DOCX mutation ${index} color must be six hex digits.`)
    patch.color = properties.color.toUpperCase()
  }
  if (Object.keys(patch).length === 0) throw new TypeError(`DOCX mutation ${index} sets no run property.`)
  if (mutation.range === undefined) return { target_kind: targetKind, target_id: mutation.target_id, expected_xml_sha256: mutation.expected_xml_sha256, properties: patch }
  const range = plainObject(mutation.range, `DOCX mutation ${index} range`)
  exactKeys(range, ['start_utf16', 'end_utf16'], `DOCX mutation ${index} range`)
  const { start_utf16: start, end_utf16: end } = range as { start_utf16: unknown; end_utf16: unknown }
  if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || (start as number) < 0 || (end as number) <= (start as number) || (end as number) > DOCX_WASM_NATIVE_MAX_TEXT_CODE_UNITS) {
    throw new TypeError(`DOCX mutation ${index} range must be a non-empty UTF-16 span.`)
  }
  return { target_kind: targetKind, target_id: mutation.target_id, expected_xml_sha256: mutation.expected_xml_sha256, properties: patch, range: { start_utf16: start as number, end_utf16: end as number } }
}

function paragraphAnchors(document: NativeDocxDocumentV1): Map<string, string> {
  const anchors = new Map<string, string>()
  for (const story of [document.body, ...document.headers, ...document.footers, ...document.notes, ...document.comment_stories]) {
    for (const block of story.blocks) {
      const paragraphs = block.paragraph ? [block.paragraph] : block.table?.rows.flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs)) ?? []
      for (const paragraph of paragraphs) anchors.set(paragraph.id, paragraph.anchor.xml_sha256)
    }
  }
  return anchors
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

/** Every key must be one of `wanted`; unlike exactKeys none of them is required. */
function allowedKeys(value: Record<string, unknown>, wanted: readonly string[], name: string): void {
  for (const key of Object.getOwnPropertyNames(value)) if (!wanted.includes(key)) throw new TypeError(`${name} carries unknown key ${JSON.stringify(key)}.`)
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
