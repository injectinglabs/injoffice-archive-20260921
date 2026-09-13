import {decodeNativeDocxDocument, type NativeDocxSourceAnchorV1} from './nativeContract.js'

export interface NativeDocxTextboxV1 {
  package_sha256: string
  part_sha256: string
  paragraph_id: string
  diagnostic_id: string
  anchor: NativeDocxSourceAnchorV1
  kind: 'drawingml' | 'vml'
  status: 'supported' | 'omitted'
  paragraphs: string[]
  reason: string
}
export interface NativeDocxTextboxEvidenceV1 {items: NativeDocxTextboxV1[]; omitted_count: number}

function record(value: unknown, keys: string[]): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Object.getPrototypeOf(value) !== Object.prototype ||
      Reflect.ownKeys(value).some(key => typeof key !== 'string') ||
      Reflect.ownKeys(value).sort().join(',') !== [...keys].sort().join(',')) throw new TypeError('Invalid textbox evidence record')
  const result: Record<string, unknown> = {}
  for (const key of keys) {
    const property = Object.getOwnPropertyDescriptor(value, key)
    if (!property || !('value' in property)) throw new TypeError('Textbox accessors are forbidden')
    result[key] = property.value
  }
  return result
}
function array(value: unknown, max: number): unknown[] {
  if (!Array.isArray(value) || Object.getPrototypeOf(value) !== Array.prototype || value.length > max ||
      Reflect.ownKeys(value).length !== value.length + 1) throw new TypeError('Invalid textbox evidence array')
  return Array.from({length: value.length}, (_, index) => {
    const property = Object.getOwnPropertyDescriptor(value, String(index))
    if (!property || !('value' in property)) throw new TypeError('Textbox arrays must be dense plain data')
    return property.value
  })
}

/** Validates evidence from same-byte native inspection. Hash joins are not a
 * substitute for the trusted XML producer: arbitrary supplied text is not proof.
 * This inventory supplies no native layout or mutation capabilities. */
export function decodeNativeDocxTextboxEvidenceV1(source: unknown, input: unknown): NativeDocxTextboxEvidenceV1 {
  const decoded = decodeNativeDocxDocument(source)
  if (!decoded.ok) throw new TypeError('Invalid textbox source')
  if (input === undefined) return {items: [], omitted_count: 0}
  const document = decoded.value, outer = record(input, ['items', 'omitted_count'])
  if (!Number.isSafeInteger(outer.omitted_count) || Number(outer.omitted_count) < 0 || Number(outer.omitted_count) > 100000) throw new TypeError('Textbox omission budget exceeded')
  const paragraphs = new Map(document.body.blocks.flatMap(block => block.paragraph ? [[block.paragraph.id, block.paragraph] as const] : []))
  const seen = new Set<string>(), partHashes = new Set<string>()
  let units = 0, previousStart = -1
  const items = array(outer.items, 64).map(raw => {
    const item = record(raw, ['package_sha256', 'part_sha256', 'paragraph_id', 'diagnostic_id', 'anchor', 'kind', 'status', 'paragraphs', 'reason'])
    const anchor = record(item.anchor, ['part_name', 'path', 'start_byte', 'end_byte', 'xml_sha256'])
    if (item.package_sha256 !== document.source.package_sha256 || typeof item.part_sha256 !== 'string' || !/^sha256:[0-9a-f]{64}$/.test(item.part_sha256) ||
        typeof item.paragraph_id !== 'string' || typeof item.diagnostic_id !== 'string' || seen.has(item.diagnostic_id) || (item.kind !== 'drawingml' && item.kind !== 'vml')) throw new TypeError('Invalid textbox identity')
    const paragraph = paragraphs.get(item.paragraph_id), diagnostic = document.unsupported.find(d => d.id === item.diagnostic_id)
    if (!paragraph || !diagnostic?.anchor || diagnostic.scope_id !== paragraph.id || diagnostic.capability !== 'drawings' || diagnostic.preservation !== 'refuse-mutation' ||
        Object.entries(diagnostic.anchor).some(([key, value]) => anchor[key] !== value) || anchor.part_name !== document.source.main_part ||
        typeof anchor.path !== 'string' || !anchor.path.startsWith(paragraph.anchor.path + '/') || Number(anchor.start_byte) < paragraph.anchor.start_byte || Number(anchor.end_byte) > paragraph.anchor.end_byte || Number(anchor.start_byte) <= previousStart) throw new TypeError('Textbox source anchor does not join')
    const suffix = anchor.path.slice(paragraph.anchor.path.length)
    if (item.kind === 'vml' ? diagnostic.code !== 'UNMODELED_DRAWING' || !/^\/w:r\[[1-9][0-9]*\]\/w:pict\[[1-9][0-9]*\]$/.test(suffix)
      : !/^\/w:r\[[1-9][0-9]*\]\/w:drawing\[[1-9][0-9]*\](?:\/(?:wp|ns[0-9a-f]{8}):(?:inline|anchor)\[[1-9][0-9]*\](?:\/(?:a|ns[0-9a-f]{8}):graphic\[1\](?:\/(?:a|ns[0-9a-f]{8}):graphicData\[1\])?)?)?$/.test(suffix)) throw new TypeError('Textbox must join a direct body drawing')
    const text = array(item.paragraphs, 64)
    for (const paragraphText of text) {
      if (typeof paragraphText !== 'string' || paragraphText.length > 4096) throw new TypeError('Invalid textbox paragraph text')
      units += paragraphText.length
    }
    if (units > 100000 || typeof item.reason !== 'string' || item.reason.length > 256 ||
        (item.status === 'supported' ? !text.length || item.reason !== '' : item.status !== 'omitted' || text.length !== 0 || !item.reason.length)) throw new TypeError('Invalid textbox status or text budget')
    seen.add(item.diagnostic_id); partHashes.add(item.part_sha256); previousStart = Number(anchor.start_byte)
    return {...item, anchor, paragraphs: text} as unknown as NativeDocxTextboxV1
  })
  if (partHashes.size > 1) throw new TypeError('Inconsistent textbox main-part hashes')
  return {items, omitted_count: Number(outer.omitted_count)}
}

export function createNativeDocxTextboxInventoryV1(source: unknown, options: {policy: 'source-textbox-inventory-v1'; read_only: true}, evidence: unknown) {
  const policy = record(options, ['policy', 'read_only'])
  if (policy.policy !== 'source-textbox-inventory-v1' || policy.read_only !== true) throw new TypeError('Explicit read-only textbox policy required')
  const document = decodeNativeDocxDocument(source)
  if (!document.ok) throw new TypeError('Invalid textbox source')
  return {
    policy: 'source-textbox-inventory-v1' as const, read_only: true as const,
    source: {document_id: document.value.document_id, revision: document.value.revision, package_sha256: document.value.source.package_sha256},
    ...decodeNativeDocxTextboxEvidenceV1(document.value, evidence),
    source_diagnostics: structuredClone(document.value.unsupported),
  }
}
export type NativeDocxTextboxInventoryV1 = ReturnType<typeof createNativeDocxTextboxInventoryV1>
