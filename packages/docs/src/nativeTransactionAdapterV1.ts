import {
  DOCX_NATIVE_LIMITS,
  decodeNativeDocxDocument,
  type NativeDocxDocumentV1,
  type NativeDocxRunPropertiesV1,
} from './nativeContract.js'

export const DOCX_PROSEMIRROR_TRANSACTION_PROTOCOL = 'injoffice.docx.prosemirror-transaction'
export const DOCX_PROSEMIRROR_TRANSACTION_VERSION = 1 as const
export const DOCX_PROSEMIRROR_PROJECTION_SCHEMA = 'injoffice.docx.prosemirror-text'
export const DOCX_PROSEMIRROR_PROJECTION_SCHEMA_VERSION = 1 as const
export const OFFICE_MUTATION_PROTOCOL = 'injoffice.office.mutations'
export const OFFICE_MUTATION_VERSION = 1 as const

export const DOCX_TRANSACTION_ADAPTER_LIMITS = {
  maxJsonBytes: 4 * 1024 * 1024,
  maxPayloadBytes: 3 * 1024 * 1024,
  maxSteps: 5,
  maxMarksPerRun: 64,
  maxMarkLength: 256,
  maxMutationIdLength: 128,
} as const

export interface NativeDocxProseMirrorRunProjectionV1 {
  run_id: string
  text_from: number
  text_to: number
  text: string
  marks: string[]
  native_properties: NativeDocxRunPropertiesV1
}

export interface NativeDocxProseMirrorParagraphProjectionV1 {
  paragraph_id: string
  node_from: number
  node_to: number
  runs: NativeDocxProseMirrorRunProjectionV1[]
}

export interface NativeDocxProseMirrorDocumentProjectionV1 {
  doc_size: number
  paragraphs: NativeDocxProseMirrorParagraphProjectionV1[]
}

export interface NativeDocxProseMirrorTextSliceV1 {
  open_start: 0
  open_end: 0
  text: string
  marks: string[]
}

export interface NativeDocxProseMirrorReplaceStepV1 {
  step_type: 'replace'
  step_index: number
  from: number
  to: number
  paragraph_id: string
  run_id: string
  expected_xml_sha256: string
  deleted_text: string
  before_text: string
  after_text: string
  slice: NativeDocxProseMirrorTextSliceV1
}

export interface NativeDocxProseMirrorTransactionV1 {
  protocol: typeof DOCX_PROSEMIRROR_TRANSACTION_PROTOCOL
  version: typeof DOCX_PROSEMIRROR_TRANSACTION_VERSION
  schema_id: typeof DOCX_PROSEMIRROR_PROJECTION_SCHEMA
  schema_version: typeof DOCX_PROSEMIRROR_PROJECTION_SCHEMA_VERSION
  mutation_id: string
  document_id: string
  expected_revision: string
  before: NativeDocxProseMirrorDocumentProjectionV1
  steps: NativeDocxProseMirrorReplaceStepV1[]
  after: NativeDocxProseMirrorDocumentProjectionV1
}

export interface NativeDocxTextMutationPayloadV1 {
  mutations: Array<{
    /** Paragraph is allowed only when it contains exactly one run and that run is text. */
    target_kind: 'paragraph' | 'run'
    target_id: string
    expected_xml_sha256: string
    text: string
  }>
}

export interface NativeDocxOfficeMutationEnvelopeV1 {
  protocol: typeof OFFICE_MUTATION_PROTOCOL
  version: typeof OFFICE_MUTATION_VERSION
  format: 'docx'
  mutation_id: string
  expected_revision: string
  payload: NativeDocxTextMutationPayloadV1
}

export type NativeDocxTransactionAdapterIssueCode =
  | 'INVALID_INPUT'
  | 'UNKNOWN_FIELD'
  | 'REQUIRED'
  | 'INVALID_TYPE'
  | 'INVALID_VALUE'
  | 'LIMIT_EXCEEDED'
  | 'STALE_REVISION'
  | 'STALE_TARGET'
  | 'SCHEMA_DRIFT'
  | 'PARTIAL_COVERAGE'
  | 'UNSUPPORTED_STRUCTURE'
  | 'AMBIGUOUS_EDIT'
  | 'OVERLAPPING_STEPS'
  | 'REORDERED_STEPS'
  | 'MAPPING_DRIFT'
  | 'SEMANTIC_NO_OP'

export interface NativeDocxTransactionAdapterIssueV1 {
  code: NativeDocxTransactionAdapterIssueCode
  path: string
  message: string
}

export type NativeDocxTransactionAdapterResultV1 =
  | {
      ok: true
      value: {
        envelope: NativeDocxOfficeMutationEnvelopeV1
        encoded_envelope: string
      }
    }
  | { ok: false; issues: NativeDocxTransactionAdapterIssueV1[] }

/**
 * Optional host seam. The host may inspect an actual ProseMirror Transaction,
 * but must return the complete strict v1 projection. The adapter never imports
 * ProseMirror and never treats its document as OOXML or layout authority.
 */
export interface NativeDocxProseMirrorTransactionProjectorV1<Transaction = unknown> {
  project(transaction: Transaction, document: NativeDocxDocumentV1): unknown
}

type JsonObject = Record<string, unknown>

interface MutableRun {
  paragraphId: string
  runId: string
  order: number
  anchorSHA: string
  originalText: string
  text: string
  marks: string[]
  nativeProperties: NativeDocxRunPropertiesV1
  textFrom: number
  textTo: number
  origins: number[]
}

interface MutableParagraph {
  paragraphId: string
  nodeFrom: number
  nodeTo: number
  runs: MutableRun[]
}

const transactionFields = ['protocol', 'version', 'schema_id', 'schema_version', 'mutation_id', 'document_id', 'expected_revision', 'before', 'steps', 'after'] as const
const projectionFields = ['doc_size', 'paragraphs'] as const
const paragraphFields = ['paragraph_id', 'node_from', 'node_to', 'runs'] as const
const runFields = ['run_id', 'text_from', 'text_to', 'text', 'marks', 'native_properties'] as const
const stepFields = ['step_type', 'step_index', 'from', 'to', 'paragraph_id', 'run_id', 'expected_xml_sha256', 'deleted_text', 'before_text', 'after_text', 'slice'] as const
const sliceFields = ['open_start', 'open_end', 'text', 'marks'] as const
const nativePropertyFields = ['character_style_id', 'font_family', 'font_size_half_points', 'bold', 'italic', 'underline', 'color', 'highlight', 'language', 'rtl', 'hidden'] as const
const mutationID = /^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/
const sha256 = /^sha256:[0-9a-f]{64}$/

function issue(code: NativeDocxTransactionAdapterIssueCode, path: string, message: string): NativeDocxTransactionAdapterIssueV1 {
  return { code, path, message }
}

function fail(code: NativeDocxTransactionAdapterIssueCode, path: string, message: string): NativeDocxTransactionAdapterResultV1 {
  return { ok: false, issues: [issue(code, path, message)] }
}

function pointer(value: string): string {
  return value.replace(/~/g, '~0').replace(/\//g, '~1')
}

function plainObject(value: unknown, path: string, fields: readonly string[]): JsonObject {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) throw issue(value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, value === undefined ? 'field is required' : 'must be a plain object')
  if (Object.getPrototypeOf(value) !== Object.prototype) throw issue('INVALID_INPUT', path, 'must have the ordinary object prototype')
  if (Object.getOwnPropertySymbols(value).length > 0) throw issue('UNKNOWN_FIELD', path, 'symbol keys are not allowed')
  const descriptors = Object.getOwnPropertyDescriptors(value)
  const allowed = new Set(fields)
  for (const key of Object.keys(descriptors).sort()) {
    const descriptor = descriptors[key]!
    if (!('value' in descriptor)) throw issue('INVALID_INPUT', `${path}/${pointer(key)}`, 'accessor properties are not allowed')
    if (!allowed.has(key)) throw issue('UNKNOWN_FIELD', `${path}/${pointer(key)}`, `unknown field ${JSON.stringify(key)}`)
  }
  return value as JsonObject
}

function required(object: JsonObject, key: string, path: string): unknown {
  if (!Object.prototype.hasOwnProperty.call(object, key)) throw issue('REQUIRED', `${path}/${key}`, 'field is required')
  return object[key]
}

function stringValue(value: unknown, path: string, maxLength: number = DOCX_NATIVE_LIMITS.maxTextLength): string {
  if (typeof value !== 'string') throw issue(value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, value === undefined ? 'field is required' : 'must be a string')
  if (value.length > maxLength) throw issue('LIMIT_EXCEEDED', path, `string exceeds ${maxLength} UTF-16 code units`)
  if (!wellFormedUnicode(value)) throw issue('INVALID_VALUE', path, 'must not contain unpaired UTF-16 surrogates')
  return value
}

function integer(value: unknown, path: string, minimum = 0): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value)) throw issue('INVALID_TYPE', path, 'must be a safe integer')
  if (Object.is(value, -0) || value < minimum) throw issue('INVALID_VALUE', path, `must be an integer at least ${minimum}`)
  return value
}

function arrayValue(value: unknown, path: string, maximum: number = DOCX_NATIVE_LIMITS.maxCollectionItems): unknown[] {
  if (!Array.isArray(value)) throw issue(value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, value === undefined ? 'field is required' : 'must be an array')
  if (Object.getPrototypeOf(value) !== Array.prototype) throw issue('INVALID_INPUT', path, 'must have the ordinary array prototype')
  if (value.length > maximum) throw issue('LIMIT_EXCEEDED', path, `array exceeds ${maximum} items`)
  return value
}

function marksValue(value: unknown, path: string): string[] {
  const marks = arrayValue(value, path, DOCX_TRANSACTION_ADAPTER_LIMITS.maxMarksPerRun).map((entry, index) =>
    stringValue(entry, `${path}/${index}`, DOCX_TRANSACTION_ADAPTER_LIMITS.maxMarkLength),
  )
  if (new Set(marks).size !== marks.length) throw issue('INVALID_VALUE', path, 'marks must be duplicate-free')
  return [...marks]
}

function nativePropertiesValue(value: unknown, path: string): NativeDocxRunPropertiesV1 {
  const object = plainObject(value, path, nativePropertyFields)
  const result: Record<string, string | number | boolean> = {}
  for (const key of Object.keys(object)) {
    const entry = object[key]
    if (key === 'font_size_half_points') result[key] = integer(entry, `${path}/${key}`, 1)
    else if (['bold', 'italic', 'rtl', 'hidden'].includes(key)) {
      if (typeof entry !== 'boolean') throw issue('INVALID_TYPE', `${path}/${key}`, 'must be a boolean')
      result[key] = entry
    } else result[key] = stringValue(entry, `${path}/${key}`, 256)
  }
  return result as NativeDocxRunPropertiesV1
}

function preflight(value: unknown): void {
  let nodes = 0
  const visit = (entry: unknown, path: string, depth: number): void => {
    nodes += 1
    if (nodes > DOCX_NATIVE_LIMITS.maxNodes) throw issue('LIMIT_EXCEEDED', path, `value graph exceeds ${DOCX_NATIVE_LIMITS.maxNodes} nodes`)
    if (depth > DOCX_NATIVE_LIMITS.maxDepth) throw issue('LIMIT_EXCEEDED', path, `value graph exceeds depth ${DOCX_NATIVE_LIMITS.maxDepth}`)
    if (entry === null) throw issue('INVALID_VALUE', path, 'null is not allowed')
    if (typeof entry === 'number' && (!Number.isFinite(entry) || Object.is(entry, -0))) throw issue('INVALID_VALUE', path, 'non-finite numbers and negative zero are not allowed')
    if (Array.isArray(entry)) {
      if (Object.getPrototypeOf(entry) !== Array.prototype) throw issue('INVALID_INPUT', path, 'must have the ordinary array prototype')
      if (entry.length > DOCX_NATIVE_LIMITS.maxCollectionItems) throw issue('LIMIT_EXCEEDED', path, `array exceeds ${DOCX_NATIVE_LIMITS.maxCollectionItems} items`)
      for (let index = 0; index < entry.length; index += 1) visit(entry[index], `${path}/${index}`, depth + 1)
      return
    }
    if (typeof entry === 'object') {
      if (Object.getPrototypeOf(entry) !== Object.prototype || Object.getOwnPropertySymbols(entry).length > 0) throw issue('INVALID_INPUT', path, 'only ordinary string-keyed objects are allowed')
      const descriptors = Object.getOwnPropertyDescriptors(entry)
      for (const key of Object.keys(descriptors).sort()) {
        const descriptor = descriptors[key]!
        if (!('value' in descriptor)) throw issue('INVALID_INPUT', `${path}/${pointer(key)}`, 'accessor properties are not allowed')
        visit(descriptor.value, `${path}/${pointer(key)}`, depth + 1)
      }
    }
  }
  visit(value, '', 0)
}

function decodeProjection(value: unknown, path: string): NativeDocxProseMirrorDocumentProjectionV1 {
  const object = plainObject(value, path, projectionFields)
  const paragraphs = arrayValue(required(object, 'paragraphs', path), `${path}/paragraphs`).map((paragraphValue, paragraphIndex) => {
    const paragraphPath = `${path}/paragraphs/${paragraphIndex}`
    const paragraph = plainObject(paragraphValue, paragraphPath, paragraphFields)
    const runs = arrayValue(required(paragraph, 'runs', paragraphPath), `${paragraphPath}/runs`).map((runValue, runIndex) => {
      const runPath = `${paragraphPath}/runs/${runIndex}`
      const run = plainObject(runValue, runPath, runFields)
      return {
        run_id: stringValue(required(run, 'run_id', runPath), `${runPath}/run_id`, 256),
        text_from: integer(required(run, 'text_from', runPath), `${runPath}/text_from`),
        text_to: integer(required(run, 'text_to', runPath), `${runPath}/text_to`),
        text: stringValue(required(run, 'text', runPath), `${runPath}/text`),
        marks: marksValue(required(run, 'marks', runPath), `${runPath}/marks`),
        native_properties: nativePropertiesValue(required(run, 'native_properties', runPath), `${runPath}/native_properties`),
      }
    })
    return {
      paragraph_id: stringValue(required(paragraph, 'paragraph_id', paragraphPath), `${paragraphPath}/paragraph_id`, 256),
      node_from: integer(required(paragraph, 'node_from', paragraphPath), `${paragraphPath}/node_from`),
      node_to: integer(required(paragraph, 'node_to', paragraphPath), `${paragraphPath}/node_to`),
      runs,
    }
  })
  return { doc_size: integer(required(object, 'doc_size', path), `${path}/doc_size`), paragraphs }
}

function decodeStep(value: unknown, path: string, expectedIndex: number): NativeDocxProseMirrorReplaceStepV1 {
  const object = plainObject(value, path, stepFields)
  if (required(object, 'step_type', path) !== 'replace') throw issue('UNSUPPORTED_STRUCTURE', `${path}/step_type`, 'v1 accepts only exact text replace steps')
  const stepIndex = integer(required(object, 'step_index', path), `${path}/step_index`)
  if (stepIndex !== expectedIndex) throw issue('REORDERED_STEPS', `${path}/step_index`, `expected contiguous step_index ${expectedIndex}`)
  const sliceObject = plainObject(required(object, 'slice', path), `${path}/slice`, sliceFields)
  const openStart = integer(required(sliceObject, 'open_start', `${path}/slice`), `${path}/slice/open_start`)
  const openEnd = integer(required(sliceObject, 'open_end', `${path}/slice`), `${path}/slice/open_end`)
  if (openStart !== 0 || openEnd !== 0) throw issue('UNSUPPORTED_STRUCTURE', `${path}/slice`, 'open or structural slices are not supported')
  return {
    step_type: 'replace',
    step_index: stepIndex,
    from: integer(required(object, 'from', path), `${path}/from`),
    to: integer(required(object, 'to', path), `${path}/to`),
    paragraph_id: stringValue(required(object, 'paragraph_id', path), `${path}/paragraph_id`, 256),
    run_id: stringValue(required(object, 'run_id', path), `${path}/run_id`, 256),
    expected_xml_sha256: stringValue(required(object, 'expected_xml_sha256', path), `${path}/expected_xml_sha256`, 71),
    deleted_text: stringValue(required(object, 'deleted_text', path), `${path}/deleted_text`),
    before_text: stringValue(required(object, 'before_text', path), `${path}/before_text`),
    after_text: stringValue(required(object, 'after_text', path), `${path}/after_text`),
    slice: {
      open_start: 0,
      open_end: 0,
      text: stringValue(required(sliceObject, 'text', `${path}/slice`), `${path}/slice/text`),
      marks: marksValue(required(sliceObject, 'marks', `${path}/slice`), `${path}/slice/marks`),
    },
  }
}

export function decodeNativeDocxProseMirrorTransactionV1(value: unknown): NativeDocxProseMirrorTransactionV1 | NativeDocxTransactionAdapterIssueV1[] {
  try {
    preflight(value)
    const object = plainObject(value, '', transactionFields)
    if (required(object, 'protocol', '') !== DOCX_PROSEMIRROR_TRANSACTION_PROTOCOL) throw issue('SCHEMA_DRIFT', '/protocol', 'unsupported transaction protocol')
    if (required(object, 'version', '') !== DOCX_PROSEMIRROR_TRANSACTION_VERSION) throw issue('SCHEMA_DRIFT', '/version', 'unsupported transaction version')
    if (required(object, 'schema_id', '') !== DOCX_PROSEMIRROR_PROJECTION_SCHEMA) throw issue('SCHEMA_DRIFT', '/schema_id', 'unsupported projection schema')
    if (required(object, 'schema_version', '') !== DOCX_PROSEMIRROR_PROJECTION_SCHEMA_VERSION) throw issue('SCHEMA_DRIFT', '/schema_version', 'unsupported projection schema version')
    const stepsValue = arrayValue(required(object, 'steps', ''), '/steps', DOCX_TRANSACTION_ADAPTER_LIMITS.maxSteps)
    if (stepsValue.length === 0) throw issue('SEMANTIC_NO_OP', '/steps', 'transaction must contain at least one step')
    return {
      protocol: DOCX_PROSEMIRROR_TRANSACTION_PROTOCOL,
      version: DOCX_PROSEMIRROR_TRANSACTION_VERSION,
      schema_id: DOCX_PROSEMIRROR_PROJECTION_SCHEMA,
      schema_version: DOCX_PROSEMIRROR_PROJECTION_SCHEMA_VERSION,
      mutation_id: stringValue(required(object, 'mutation_id', ''), '/mutation_id', DOCX_TRANSACTION_ADAPTER_LIMITS.maxMutationIdLength),
      document_id: stringValue(required(object, 'document_id', ''), '/document_id', 256),
      expected_revision: stringValue(required(object, 'expected_revision', ''), '/expected_revision', 71),
      before: decodeProjection(required(object, 'before', ''), '/before'),
      steps: stepsValue.map((entry, index) => decodeStep(entry, `/steps/${index}`, index)),
      after: decodeProjection(required(object, 'after', ''), '/after'),
    }
  } catch (error) {
    if (isIssue(error)) return [error]
    return [issue('INVALID_INPUT', '', 'input could not be inspected safely')]
  }
}

export function decodeNativeDocxProseMirrorTransactionJsonV1(json: string): NativeDocxProseMirrorTransactionV1 | NativeDocxTransactionAdapterIssueV1[] {
  if (typeof json !== 'string') return [issue('INVALID_TYPE', '', 'JSON input must be a string')]
  const byteLength = new TextEncoder().encode(json).byteLength
  if (byteLength === 0 || byteLength > DOCX_TRANSACTION_ADAPTER_LIMITS.maxJsonBytes) return [issue('LIMIT_EXCEEDED', '', `encoded transaction size must be 1..${DOCX_TRANSACTION_ADAPTER_LIMITS.maxJsonBytes} bytes`)]
  try {
    const decoded = decodeNativeDocxProseMirrorTransactionV1(JSON.parse(json))
    if (Array.isArray(decoded)) return decoded
    const duplicate = duplicateJsonObjectMemberPath(json)
    if (duplicate !== undefined) return [issue('UNKNOWN_FIELD', duplicate, 'duplicate JSON object member is not allowed')]
    return decoded
  } catch {
    return [issue('INVALID_INPUT', '', 'transaction must be one valid JSON value')]
  }
}

export function adaptNativeDocxProseMirrorTransactionV1(documentValue: NativeDocxDocumentV1, transactionValue: unknown): NativeDocxTransactionAdapterResultV1 {
  try {
    preflight(documentValue)
  } catch (error) {
    return { ok: false, issues: [isIssue(error) ? { ...error, path: `/document${error.path}` } : issue('INVALID_INPUT', '/document', 'native document could not be inspected safely')] }
  }
  const document = decodeNativeDocxDocument(documentValue)
  if (!document.ok) return { ok: false, issues: document.issues.slice(0, DOCX_NATIVE_LIMITS.maxIssues).map((entry) => issue('INVALID_INPUT', `/document${entry.path}`, `${entry.code}: ${entry.message}`)) }
  const unsupported = unsupportedDocumentReason(document.value)
  if (unsupported) return fail('UNSUPPORTED_STRUCTURE', unsupported.path, unsupported.message)
  const decoded = decodeNativeDocxProseMirrorTransactionV1(transactionValue)
  if (Array.isArray(decoded)) return { ok: false, issues: decoded }
  return adaptDecoded(document.value, decoded)
}

export function adaptNativeDocxProseMirrorTransactionWithHostV1<Transaction>(
  document: NativeDocxDocumentV1,
  transaction: Transaction,
  projector: NativeDocxProseMirrorTransactionProjectorV1<Transaction>,
): NativeDocxTransactionAdapterResultV1 {
  if (!projector || typeof projector.project !== 'function') return fail('INVALID_INPUT', '/projector', 'host projector must provide project(transaction, document)')
  let authoritativeDocument: NativeDocxDocumentV1
  try {
    preflight(document)
    const decodedDocument = decodeNativeDocxDocument(document)
    if (!decodedDocument.ok) return { ok: false, issues: decodedDocument.issues.slice(0, DOCX_NATIVE_LIMITS.maxIssues).map((entry) => issue('INVALID_INPUT', `/document${entry.path}`, `${entry.code}: ${entry.message}`)) }
    authoritativeDocument = structuredClone(decodedDocument.value)
  } catch {
    return fail('INVALID_INPUT', '/document', 'native document could not be snapshotted safely')
  }
  try {
    const projection = projector.project(transaction, structuredClone(authoritativeDocument))
    return adaptNativeDocxProseMirrorTransactionV1(authoritativeDocument, projection)
  } catch {
    return fail('INVALID_INPUT', '/projector', 'host projector threw before producing a complete projection')
  }
}

function adaptDecoded(document: NativeDocxDocumentV1, transaction: NativeDocxProseMirrorTransactionV1): NativeDocxTransactionAdapterResultV1 {
  if (!mutationID.test(transaction.mutation_id)) return fail('INVALID_VALUE', '/mutation_id', 'must use 1..128 restricted ASCII identifier characters')
  if (transaction.document_id !== document.document_id) return fail('STALE_TARGET', '/document_id', 'transaction targets a different native document')
  if (!sha256.test(transaction.expected_revision) || transaction.expected_revision !== document.source.package_sha256) return fail('STALE_REVISION', '/expected_revision', 'must equal the exact NativeDocumentV1 source.package_sha256')
  const unsupported = unsupportedDocumentReason(document)
  if (unsupported) return fail('UNSUPPORTED_STRUCTURE', unsupported.path, unsupported.message)

  const initialized = initializeProjection(document, transaction.before, '/before')
  if (Array.isArray(initialized)) return { ok: false, issues: initialized }
  const { paragraphs, runs } = initialized
  const touched = new Map<string, MutableRun>()
  let previous: { order: number; start: number; end: number } | undefined

  for (const step of transaction.steps) {
    const path = `/steps/${step.step_index}`
    const run = runs.get(step.run_id)
    if (!run || run.paragraphId !== step.paragraph_id) return fail('STALE_TARGET', `${path}/run_id`, 'step target is absent from the exact native projection')
    if (step.expected_xml_sha256 !== run.anchorSHA) return fail('STALE_TARGET', `${path}/expected_xml_sha256`, 'step target anchor fingerprint is stale')
    if (step.before_text !== run.text) return fail('MAPPING_DRIFT', `${path}/before_text`, 'step does not bind the exact pre-step run text')
    if (step.from > step.to || step.from < run.textFrom || step.to > run.textTo) return fail('AMBIGUOUS_EDIT', path, 'step range must be wholly inside its declared native text run')
    const candidates = [...runs.values()].filter((candidate) => step.from >= candidate.textFrom && step.to <= candidate.textTo)
    if (candidates.length !== 1 || candidates[0]?.runId !== run.runId) return fail('AMBIGUOUS_EDIT', path, 'step range is ambiguous at a native run boundary')
    const localFrom = step.from - run.textFrom
    const localTo = step.to - run.textFrom
    if (splitsSurrogate(run.text, localFrom) || splitsSurrogate(run.text, localTo)) return fail('AMBIGUOUS_EDIT', path, 'step range bisects a UTF-16 surrogate pair')
    if (run.text.slice(localFrom, localTo) !== step.deleted_text) return fail('MAPPING_DRIFT', `${path}/deleted_text`, 'deleted_text does not equal the exact UTF-16 source range')
    if (step.slice.text !== '' && !equalStrings(step.slice.marks, run.marks)) return fail('AMBIGUOUS_EDIT', `${path}/slice/marks`, 'inserted text marks must exactly match the owning native run')
    if (step.slice.text === '' && step.slice.marks.length !== 0) return fail('AMBIGUOUS_EDIT', `${path}/slice/marks`, 'an empty text slice must not invent marks')
    if (!xmlTextValid(step.slice.text)) return fail('INVALID_VALUE', `${path}/slice/text`, 'inserted text is not valid XML 1.0 character data')
    const nextText = run.text.slice(0, localFrom) + step.slice.text + run.text.slice(localTo)
    if (nextText !== step.after_text) return fail('MAPPING_DRIFT', `${path}/after_text`, 'after_text is not the exact result of this step')
    if (hasUnattestedEdgeWhitespace(run.originalText, nextText)) return fail('AMBIGUOUS_EDIT', `${path}/after_text`, 'leading or trailing XML whitespace cannot be attested by the v1 native projection')
    const originalStart = run.origins[localFrom]!
    const originalEnd = run.origins[localTo]!
    if (previous) {
      if (run.order < previous.order || (run.order === previous.order && originalStart < previous.start)) return fail('REORDERED_STEPS', path, 'steps must remain in native source order')
      if (run.order === previous.order && (originalStart < previous.end || (previous.start === previous.end && originalStart === previous.start))) return fail('OVERLAPPING_STEPS', path, 'steps must not overlap or revisit the same source insertion point')
    }
    previous = { order: run.order, start: originalStart, end: originalEnd }
    const oldLength = run.text.length
    run.text = nextText
    const replacementOrigins = step.slice.text.length === 0
      ? [originalEnd]
      : [originalStart, ...Array.from({ length: step.slice.text.length - 1 }, () => originalStart), originalEnd]
    run.origins = [...run.origins.slice(0, localFrom), ...replacementOrigins, ...run.origins.slice(localTo + 1)]
    const delta = run.text.length - oldLength
    run.textTo += delta
    shiftProjection(paragraphs, runs, run, delta)
    touched.set(run.runId, run)
  }

  const finalProjection = snapshotProjection(paragraphs)
  if (canonicalJSON(finalProjection) !== canonicalJSON(transaction.after)) return fail('PARTIAL_COVERAGE', '/after', 'after projection must exactly cover the complete simulated body with unchanged identities, marks, and native properties')
  const mutations = [...touched.values()].sort((left, right) => left.order - right.order).filter((run) => run.text !== run.originalText).map((run) => ({
    target_kind: 'run' as const,
    target_id: run.runId,
    expected_xml_sha256: run.anchorSHA,
    text: run.text,
  }))
  if (mutations.length === 0) return fail('SEMANTIC_NO_OP', '/steps', 'transaction produces no native text change')
  const payload: NativeDocxTextMutationPayloadV1 = { mutations }
  const envelope: NativeDocxOfficeMutationEnvelopeV1 = {
    protocol: OFFICE_MUTATION_PROTOCOL,
    version: OFFICE_MUTATION_VERSION,
    format: 'docx',
    mutation_id: transaction.mutation_id,
    expected_revision: transaction.expected_revision,
    payload,
  }
  const encodedPayload = canonicalJSON(payload)
  if (new TextEncoder().encode(encodedPayload).byteLength > DOCX_TRANSACTION_ADAPTER_LIMITS.maxPayloadBytes) return fail('LIMIT_EXCEEDED', '/payload', `encoded native payload exceeds ${DOCX_TRANSACTION_ADAPTER_LIMITS.maxPayloadBytes} bytes`)
  const encodedEnvelope = canonicalJSON(envelope)
  if (new TextEncoder().encode(encodedEnvelope).byteLength > DOCX_TRANSACTION_ADAPTER_LIMITS.maxJsonBytes) return fail('LIMIT_EXCEEDED', '', `encoded mutation envelope exceeds ${DOCX_TRANSACTION_ADAPTER_LIMITS.maxJsonBytes} bytes`)
  return { ok: true, value: { envelope, encoded_envelope: encodedEnvelope } }
}

function initializeProjection(document: NativeDocxDocumentV1, projection: NativeDocxProseMirrorDocumentProjectionV1, path: string): { paragraphs: MutableParagraph[]; runs: Map<string, MutableRun> } | NativeDocxTransactionAdapterIssueV1[] {
  const nativeParagraphs = document.body.blocks.map((block) => block.paragraph!)
  if (projection.paragraphs.length !== nativeParagraphs.length) return [issue('PARTIAL_COVERAGE', `${path}/paragraphs`, 'projection must cover every body paragraph exactly once')]
  const paragraphs: MutableParagraph[] = []
  const runs = new Map<string, MutableRun>()
  let cursor = 0
  let order = 0
  for (let paragraphIndex = 0; paragraphIndex < nativeParagraphs.length; paragraphIndex += 1) {
    const native = nativeParagraphs[paragraphIndex]!
    const projected = projection.paragraphs[paragraphIndex]!
    const paragraphPath = `${path}/paragraphs/${paragraphIndex}`
    if (projected.paragraph_id !== native.id) return [issue('STALE_TARGET', `${paragraphPath}/paragraph_id`, 'paragraph identity or native order changed')]
    if (projected.node_from !== cursor) return [issue('MAPPING_DRIFT', `${paragraphPath}/node_from`, `expected ProseMirror node start ${cursor}`)]
    if (projected.runs.length !== native.runs.length) return [issue('PARTIAL_COVERAGE', `${paragraphPath}/runs`, 'projection must cover every native inline run exactly once')]
    const mutableRuns: MutableRun[] = []
    let textCursor = cursor + 1
    for (let runIndex = 0; runIndex < native.runs.length; runIndex += 1) {
      const nativeRun = native.runs[runIndex]!
      const projectedRun = projected.runs[runIndex]!
      const runPath = `${paragraphPath}/runs/${runIndex}`
      if (nativeRun.kind !== 'text' || nativeRun.text === undefined) return [issue('UNSUPPORTED_STRUCTURE', runPath, `inline ${nativeRun.kind} nodes are not mutable in v1`)]
      if (projectedRun.run_id !== nativeRun.id) return [issue('STALE_TARGET', `${runPath}/run_id`, 'run identity or native order changed')]
      if (projectedRun.text !== nativeRun.text) return [issue('STALE_TARGET', `${runPath}/text`, 'projected text does not equal the native source text')]
      if (canonicalJSON(projectedRun.native_properties) !== canonicalJSON(nativeRun.properties ?? {})) return [issue('SCHEMA_DRIFT', `${runPath}/native_properties`, 'projected native properties do not equal the source run properties')]
      if (projectedRun.text_from !== textCursor || projectedRun.text_to !== textCursor + nativeRun.text.length) return [issue('MAPPING_DRIFT', runPath, 'run positions do not use exact ProseMirror UTF-16 text coordinates')]
      const mutable: MutableRun = {
        paragraphId: native.id,
        runId: nativeRun.id,
        order: order++,
        anchorSHA: nativeRun.anchor.xml_sha256,
        originalText: nativeRun.text,
        text: nativeRun.text,
        marks: [...projectedRun.marks],
        nativeProperties: { ...(nativeRun.properties ?? {}) },
        textFrom: projectedRun.text_from,
        textTo: projectedRun.text_to,
        origins: Array.from({ length: nativeRun.text.length + 1 }, (_, index) => index),
      }
      mutableRuns.push(mutable)
      runs.set(mutable.runId, mutable)
      textCursor = mutable.textTo
    }
    if (projected.node_to !== textCursor + 1) return [issue('MAPPING_DRIFT', `${paragraphPath}/node_to`, 'paragraph node_to must include both ProseMirror boundary tokens')]
    paragraphs.push({ paragraphId: native.id, nodeFrom: projected.node_from, nodeTo: projected.node_to, runs: mutableRuns })
    cursor = projected.node_to
  }
  if (projection.doc_size !== cursor) return [issue('PARTIAL_COVERAGE', `${path}/doc_size`, 'doc_size must exactly cover all projected body nodes')]
  return { paragraphs, runs }
}

function unsupportedDocumentReason(document: NativeDocxDocumentV1): { path: string; message: string } | undefined {
  if (document.body.blocks.some((block) => block.kind !== 'paragraph')) return { path: '/document/body/blocks', message: 'body tables and structural blocks are not supported' }
  if (document.notes.length > 0 || document.comment_stories.length > 0 || document.comments.length > 0) return { path: '/document/notes', message: 'notes and comments require a different native editing contract' }
  if (document.unsupported.length > 0) return { path: '/document/unsupported', message: 'unsupported, field, tracked-change, or partial native semantics refuse the complete transaction' }
  const stories = [document.body, ...document.headers, ...document.footers, ...document.notes, ...document.comment_stories]
  for (const story of stories) {
    for (const block of story.blocks) {
      if (block.kind !== 'paragraph') return { path: '/document', message: 'tables and non-paragraph story blocks are outside the v1 transaction adapter' }
      const unsupportedRun = block.paragraph?.runs.find((run) => run.kind !== 'text' || run.page_field !== undefined || run.properties?.vertical_alignment && run.properties.vertical_alignment !== 'baseline')
      if (unsupportedRun) return { path: '/document', message: `${unsupportedRun.kind} inline nodes are outside the v1 transaction adapter` }
    }
  }
  for (const block of document.body.blocks) {
    const paragraph = block.paragraph!
    if (paragraph.edit_policy.mode !== 'read-write' || !paragraph.edit_policy.allowed_operations.includes('text.replace')) return { path: '/document/body/blocks', message: 'every projected body paragraph must allow exact native text replacement' }
  }
  return undefined
}

function shiftProjection(paragraphs: MutableParagraph[], runs: Map<string, MutableRun>, changed: MutableRun, delta: number): void {
  if (delta === 0) return
  for (const run of runs.values()) {
    if (run.order > changed.order) {
      run.textFrom += delta
      run.textTo += delta
    }
  }
  let found = false
  for (const paragraph of paragraphs) {
    if (paragraph.runs.some((run) => run.runId === changed.runId)) {
      paragraph.nodeTo += delta
      found = true
    } else if (found) {
      paragraph.nodeFrom += delta
      paragraph.nodeTo += delta
    }
  }
}

function snapshotProjection(paragraphs: MutableParagraph[]): NativeDocxProseMirrorDocumentProjectionV1 {
  const projected = paragraphs.map((paragraph) => ({
    paragraph_id: paragraph.paragraphId,
    node_from: paragraph.nodeFrom,
    node_to: paragraph.nodeTo,
    runs: paragraph.runs.map((run) => ({
      run_id: run.runId,
      text_from: run.textFrom,
      text_to: run.textTo,
      text: run.text,
      marks: [...run.marks],
      native_properties: { ...run.nativeProperties },
    })),
  }))
  return { doc_size: projected.at(-1)?.node_to ?? 0, paragraphs: projected }
}

function canonicalJSON(value: unknown): string {
  const canonical = (entry: unknown): unknown => {
    if (Array.isArray(entry)) return entry.map(canonical)
    if (entry === null || typeof entry !== 'object') return entry
    const result: Record<string, unknown> = {}
    for (const key of Object.keys(entry).sort(compareUTF16)) result[key] = canonical((entry as Record<string, unknown>)[key])
    return result
  }
  return JSON.stringify(canonical(value))
}

function compareUTF16(left: string, right: string): number {
  const length = Math.min(left.length, right.length)
  for (let index = 0; index < length; index += 1) {
    const difference = left.charCodeAt(index) - right.charCodeAt(index)
    if (difference !== 0) return difference
  }
  return left.length - right.length
}

function equalStrings(left: readonly string[], right: readonly string[]): boolean {
  return left.length === right.length && left.every((value, index) => value === right[index])
}

function splitsSurrogate(text: string, offset: number): boolean {
  if (offset <= 0 || offset >= text.length) return false
  const before = text.charCodeAt(offset - 1)
  const after = text.charCodeAt(offset)
  return before >= 0xd800 && before <= 0xdbff && after >= 0xdc00 && after <= 0xdfff
}

function wellFormedUnicode(value: string): boolean {
  for (let index = 0; index < value.length; index += 1) {
    const unit = value.charCodeAt(index)
    if (unit >= 0xd800 && unit <= 0xdbff) {
      const next = value.charCodeAt(index + 1)
      if (!(next >= 0xdc00 && next <= 0xdfff)) return false
      index += 1
    } else if (unit >= 0xdc00 && unit <= 0xdfff) return false
  }
  return true
}

function xmlTextValid(value: string): boolean {
  if (!wellFormedUnicode(value)) return false
  for (let index = 0; index < value.length; index += 1) {
    const unit = value.charCodeAt(index)
    if (unit === 0xfffe || unit === 0xffff) return false
    if (unit === 0x9 || unit === 0xa || unit >= 0x20) continue
    return false
  }
  return true
}

function duplicateJsonObjectMemberPath(json: string): string | undefined {
  let offset = 0
  const whitespace = (): void => {
    while (offset < json.length && /[\t\n\r ]/.test(json[offset]!)) offset += 1
  }
  const jsonString = (): string => {
    const start = offset++
    while (offset < json.length) {
      if (json[offset] === '\\') {
        offset += 2
      } else if (json[offset++] === '"') {
        return JSON.parse(json.slice(start, offset)) as string
      }
    }
    throw new Error('unterminated JSON string')
  }
  const value = (path: string): string | undefined => {
    whitespace()
    if (json[offset] === '{') {
      offset += 1
      whitespace()
      const keys = new Set<string>()
      if (json[offset] === '}') {
        offset += 1
        return undefined
      }
      while (offset < json.length) {
        whitespace()
        const key = jsonString()
        const keyPath = `${path}/${pointer(key)}`
        if (keys.has(key)) return keyPath
        keys.add(key)
        whitespace()
        offset += 1 // colon; JSON.parse already proved the grammar.
        const duplicate = value(keyPath)
        if (duplicate !== undefined) return duplicate
        whitespace()
        if (json[offset++] === '}') return undefined
      }
    } else if (json[offset] === '[') {
      offset += 1
      whitespace()
      if (json[offset] === ']') {
        offset += 1
        return undefined
      }
      let index = 0
      while (offset < json.length) {
        const duplicate = value(`${path}/${index++}`)
        if (duplicate !== undefined) return duplicate
        whitespace()
        if (json[offset++] === ']') return undefined
      }
    } else if (json[offset] === '"') {
      jsonString()
    } else {
      while (offset < json.length && !/[\t\n\r ,}\]]/.test(json[offset]!)) offset += 1
    }
    return undefined
  }
  return value('')
}

function hasUnattestedEdgeWhitespace(before: string, after: string): boolean {
  const edge = (value: string): [boolean, boolean] => [/^[\t\n\r ]/.test(value), /[\t\n\r ]$/.test(value)]
  const [beforeLeading, beforeTrailing] = edge(before)
  const [afterLeading, afterTrailing] = edge(after)
  return afterLeading || afterTrailing || (beforeLeading && !afterLeading) || (beforeTrailing && !afterTrailing)
}

function isIssue(value: unknown): value is NativeDocxTransactionAdapterIssueV1 {
  return value !== null && typeof value === 'object' && typeof (value as { code?: unknown }).code === 'string' && typeof (value as { path?: unknown }).path === 'string'
}
