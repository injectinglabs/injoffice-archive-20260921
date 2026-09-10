/** Browser-safe strict native DOCX page-paint output decoder and shared wire primitives.
 * This module never loads shaping, pagination, bidi or Node font providers. */
import { sha256 } from '@noble/hashes/sha2.js'
import { bytesToHex } from '@noble/hashes/utils.js'
import { DOCX_NATIVE_LIMITS, type NativeDocxValidationIssue } from './nativeContract.js'
import { decodeNativeDocxPaginationSettings } from './nativePaginationSettings.js'
import { decodeNativeDocxPagePaintResourceListV1, type NativeDocxPagePaintMediaAssetV1 } from './nativeImagePagePaintV1.js'
import { asciiLowerNative, compareNativeValidationIssues } from './nativeDeterminism.js'
import type { DecodeNativeDocxPagePaintV1Result, NativeDocxContentAddressedFaceV1, NativeDocxPagePaintDiagnosticCode, NativeDocxPagePaintRefusedV1, NativeDocxPagePaintSuccessV1 } from './nativePagePaintV1.js'

export const DOCX_PAGE_PAINT_REQUEST_PROTOCOL = 'injoffice.docx.page-paint-request'
export const DOCX_PAGE_PAINT_REQUEST_VERSION = 1 as const
export const DOCX_PAGE_PAINT_PROTOCOL = 'injoffice.docx.page-paint'
export const DOCX_PAGE_PAINT_VERSION = 1 as const

export const DOCX_PAGE_PAINT_LIMITS = Object.freeze({
  maxPages: 2_048,
  maxLines: 100_000,
  maxGlyphs: 500_000,
  maxUniqueGlyphOutlines: 100_000,
  maxProviderCalls: 100_000,
  maxPathCommandsPerGlyph: 65_536,
  maxPathCommands: 4_000_000,
  maxOutputNodes: 5_000_000,
  maxDesignCoordinate: 1_000_000_000,
  maxPaintCoordinateMilliPoints: 1_000_000_000_000,
  maxUnitsPerEm: 1_000_000,
  maxProviderMessageLength: 4_096,
})


export const DOCX_PAGE_PAINT_REQUEST_V1_BINDING_FIELDS = {
  IntegrityV1: ['font_manifest_sha256', 'shaped_lines_sha256', 'paginated_layout_sha256', 'table_projection_sha256', 'media_assets_sha256'],
  OutlineProviderV1: ['provider_id', 'provider_revision'],
  RequestV1: ['protocol', 'version', 'pagination_request', 'paginated_layout', 'font_manifest', 'media_assets', 'integrity', 'outline_provider'],
} as const

export const DOCX_PAGE_PAINT_V1_BINDING_FIELDS = {
  FaceV1: ['face_id', 'content_digest', 'collection_index'],
  PathMoveV1: ['kind', 'x_millipoints', 'y_millipoints'],
  PathLineV1: ['kind', 'x_millipoints', 'y_millipoints'],
  PathQuadraticV1: ['kind', 'control_x_millipoints', 'control_y_millipoints', 'x_millipoints', 'y_millipoints'],
  PathCubicV1: ['kind', 'control_1_x_millipoints', 'control_1_y_millipoints', 'control_2_x_millipoints', 'control_2_y_millipoints', 'x_millipoints', 'y_millipoints'],
  PathCloseV1: ['kind'],
  GlyphCommandV1: ['kind', 'id', 'line_id', 'fragment_id', 'source_id', 'glyph_index', 'face', 'glyph_id', 'font_size_millipoints', 'fill_rgb', 'fill_rule', 'outline_kind', 'path'],
  CellFillCommandV1: ['kind', 'id', 'table_id', 'row_id', 'cell_id', 'x_millipoints', 'y_millipoints', 'width_millipoints', 'height_millipoints', 'fill_rgb'],
  BorderCommandV1: ['kind', 'id', 'table_id', 'row_id', 'cell_id', 'edge', 'x1_millipoints', 'y1_millipoints', 'x2_millipoints', 'y2_millipoints', 'width_millipoints', 'stroke_rgb'],
  NoteSeparatorCommandV1: ['kind', 'id', 'line_id', 'story_id', 'x1_millipoints', 'y1_millipoints', 'x2_millipoints', 'y2_millipoints', 'width_millipoints', 'stroke_rgb'],
  ImageCropV1: ['left', 'top', 'right', 'bottom', 'unit'],
  ImageTransformV1: ['rotation_degrees', 'flip_horizontal', 'flip_vertical'],
  ImageCommandV1: ['kind', 'id', 'line_id', 'fragment_id', 'source_id', 'drawing_id', 'asset_id', 'x_millipoints', 'y_millipoints', 'width_millipoints', 'height_millipoints', 'source_crop', 'transform'],
  BodyLineV1: ['placed_line_id', 'line_id', 'paragraph_id', 'region', 'section_id', 'column_id', 'column_ordinal', 'source_line_ordinal', 'x_millipoints', 'y_millipoints', 'width_millipoints', 'height_millipoints', 'baseline_y_millipoints', 'command_ids'],
  HeaderFooterLineV1: ['placed_line_id', 'line_id', 'paragraph_id', 'region', 'section_id', 'source_line_ordinal', 'x_millipoints', 'y_millipoints', 'width_millipoints', 'height_millipoints', 'baseline_y_millipoints', 'command_ids'],
  BodyBoxV1: ['x_millipoints', 'y_millipoints', 'width_millipoints', 'height_millipoints'],
  ClipBoxV1: ['x_millipoints', 'y_millipoints', 'width_millipoints', 'height_millipoints'],
  ColumnV1: ['id', 'section_id', 'ordinal', 'x_millipoints', 'y_millipoints', 'width_millipoints', 'height_millipoints'],
  PageV1: ['id', 'ordinal', 'section_id', 'section_ids', 'kind', 'width_millipoints', 'height_millipoints', 'body_box', 'columns', 'background_rgb', 'clip_box', 'lines', 'commands'],
  ShapedSourceV1: ['protocol', 'version', 'available_width_millipoints', 'tab_interval_millipoints', 'sha256'],
  PaginatedSourceV1: ['protocol', 'version', 'sha256'],
  HeaderFooterSourceV1: ['protocol', 'version', 'sha256'],
  TableProjectionV1: ['sha256'],
  ManifestV1: ['manifest_id', 'revision', 'sha256'],
  MediaSourceV1: ['sha256'],
  ProvidersV1: ['resolver_id', 'resolver_revision', 'shaper_id', 'shaper_revision', 'bidi_id', 'bidi_revision', 'bidi_unicode_version', 'unicode13_revision', 'outline_id', 'outline_revision'],
  NumberingSourceV1: ['relationships_part', 'relationships_sha256', 'relationship_id', 'relationship_type', 'relationship_target', 'part_name', 'content_type', 'part_sha256', 'model_sha256'],
  ProvenanceV1: ['document_id', 'revision', 'package_sha256', 'main_part', 'body_story_id', 'numbering_source', 'pagination_settings', 'shaped_lines', 'paginated_layout', 'header_footer_layout', 'table_projection', 'font_manifest', 'media_assets', 'providers'],
  DiagnosticV1: ['code', 'severity', 'scope_id', 'message'],
  OutputV1: ['protocol', 'version', 'status', 'provenance', 'diagnostics', 'resources', 'pages'],
} as const

export type JsonObject = Record<string, unknown>
export const ID = /^[A-Za-z0-9][A-Za-z0-9._:/-]{0,1023}$/
export const SHORT_ID = /^[A-Za-z0-9][A-Za-z0-9._:/-]{0,255}$/
export const PROVIDER_ID = /^[A-Za-z0-9][A-Za-z0-9._:/-]{0,127}$/
export const SHA256 = /^sha256:[0-9a-f]{64}$/
export const RGB = /^[0-9A-F]{6}$/
export const ABSOLUTE_URI = /^[A-Za-z][A-Za-z0-9+.-]*:[^\u0000\r\n]{1,4090}$/
export const PART_SEGMENT = /^(?:[A-Za-z0-9._~!$&'()*+,;=@-]|%[0-9A-F]{2})+$/

export function canonicalJsonValue(value: unknown, active = new WeakSet<object>(), depth = 0, state = { nodes: 0 }): unknown {
  state.nodes += 1
  if (state.nodes > DOCX_PAGE_PAINT_LIMITS.maxOutputNodes || depth > DOCX_NATIVE_LIMITS.maxDepth) throw new RangeError('canonical hash input exceeds the bounded traversal budget')
  if (typeof value === 'number' && (!Number.isFinite(value) || Object.is(value, -0))) throw new TypeError('canonical hash input contains a non-canonical number')
  if (value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return value
  if (typeof value !== 'object') throw new TypeError('canonical hash input must contain only JSON wire values')
  if (active.has(value)) throw new TypeError('canonical hash input must be acyclic')
  const prototype = Object.getPrototypeOf(value)
  if (!Array.isArray(value) && prototype !== Object.prototype && prototype !== null) throw new TypeError('canonical hash input must contain only plain JSON objects and arrays')
  active.add(value)
  const output = Array.isArray(value)
    ? value.map((entry) => canonicalJsonValue(entry, active, depth + 1, state))
    : Object.fromEntries(Object.keys(value).sort().map((key) => [key, canonicalJsonValue((value as Record<string, unknown>)[key], active, depth + 1, state)]))
  active.delete(value)
  return output
}

export function canonicalWireSha256(value: unknown): string {
  return `sha256:${bytesToHex(sha256(new TextEncoder().encode(JSON.stringify(canonicalJsonValue(value)))))}`
}


export function issue(code: NativeDocxValidationIssue['code'], path: string, message: string): NativeDocxValidationIssue {
  return { code, path, message }
}

export function add(issues: NativeDocxValidationIssue[], code: NativeDocxValidationIssue['code'], path: string, message: string): void {
  if (issues.length >= DOCX_NATIVE_LIMITS.maxIssues) return
  if (!issues.some((entry) => entry.code === code && entry.path === path && entry.message === message)) issues.push(issue(code, path, message))
}

export function isObject(value: unknown): value is JsonObject {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

export function pointer(value: string): string {
  return value.replace(/~/g, '~0').replace(/\//g, '~1')
}

export function exactObject(value: unknown, path: string, fields: readonly string[], issues: NativeDocxValidationIssue[]): JsonObject | undefined {
  if (!isObject(value)) {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, 'must be an object')
    return undefined
  }
  const allowed = new Set(fields)
  for (const key of Object.keys(value).sort()) if (!allowed.has(key)) add(issues, 'UNKNOWN_FIELD', `${path}/${pointer(key)}`, `unknown field ${JSON.stringify(key)}`)
  for (const key of fields) if (!(key in value)) add(issues, 'REQUIRED', `${path}/${pointer(key)}`, 'field is required')
  return value
}

export function exactUnionObject(value: unknown, path: string, fields: readonly string[], issues: NativeDocxValidationIssue[]): JsonObject | undefined {
  if (!isObject(value)) {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, 'must be an object')
    return undefined
  }
  const allowed = new Set(fields)
  for (const key of Object.keys(value).sort()) if (!allowed.has(key)) add(issues, 'UNKNOWN_FIELD', `${path}/${pointer(key)}`, `unknown field ${JSON.stringify(key)}`)
  return value
}

export function stringValue(value: unknown, path: string, issues: NativeDocxValidationIssue[], pattern = SHORT_ID, max = 256): string | undefined {
  if (typeof value !== 'string') {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, 'must be a string')
    return undefined
  }
  if (value.length === 0 || value.length > max || !pattern.test(value) || /[\u0000\r\n]/.test(value)) {
    add(issues, 'INVALID_VALUE', path, 'contains an invalid or unbounded string')
    return undefined
  }
  return value
}

export function integer(value: unknown, path: string, issues: NativeDocxValidationIssue[], min: number, max: number): number | undefined {
  if (!Number.isSafeInteger(value) || Object.is(value, -0) || (value as number) < min || (value as number) > max) {
    add(issues, value === undefined ? 'REQUIRED' : 'OUT_OF_RANGE', path, `must be a safe integer from ${min} through ${max}`)
    return undefined
  }
  return value as number
}

export function safeClone<T>(value: T): T | undefined {
  try { return structuredClone(value) as T } catch { return undefined }
}

export function preflightWire(value: unknown, label: string, maxNodes: number = DOCX_PAGE_PAINT_LIMITS.maxOutputNodes, maxArray: number = DOCX_PAGE_PAINT_LIMITS.maxGlyphs): NativeDocxValidationIssue[] {
  const issues: NativeDocxValidationIssue[] = []
  const active = new WeakSet<object>()
  const stack: Array<{ value: unknown; path: string; depth: number; leaving?: boolean }> = [{ value, path: '', depth: 0 }]
  let nodes = 0
  try {
    while (stack.length > 0 && issues.length < DOCX_NATIVE_LIMITS.maxIssues) {
      const item = stack.pop()!
      if (item.leaving) { if (typeof item.value === 'object' && item.value !== null) active.delete(item.value); continue }
      nodes += 1
      if (nodes > maxNodes) { add(issues, 'LIMIT_EXCEEDED', item.path, `${label} exceeds ${maxNodes} traversed values`); break }
      if (item.depth > DOCX_NATIVE_LIMITS.maxDepth) { add(issues, 'LIMIT_EXCEEDED', item.path, `${label} exceeds ${DOCX_NATIVE_LIMITS.maxDepth} nesting levels`); break }
      const current = item.value
      if (current === null || current === undefined) { add(issues, 'INVALID_VALUE', item.path, 'JSON null and undefined are not permitted'); continue }
      if (typeof current === 'number' && (!Number.isFinite(current) || Object.is(current, -0))) { add(issues, 'INVALID_VALUE', item.path, 'non-finite numbers and negative zero are not permitted'); continue }
      if (typeof current !== 'object') {
        if (!['string', 'number', 'boolean'].includes(typeof current)) add(issues, 'INVALID_TYPE', item.path, 'must contain only JSON wire values')
        continue
      }
      if (active.has(current)) { add(issues, 'INVALID_VALUE', item.path, 'cyclic values are not valid JSON wire data'); continue }
      active.add(current)
      stack.push({ ...item, leaving: true })
      if (Array.isArray(current)) {
        if (current.length > maxArray) { add(issues, 'LIMIT_EXCEEDED', item.path, `array exceeds ${maxArray} items`); continue }
        for (let index = current.length - 1; index >= 0; index -= 1) stack.push({ value: current[index], path: `${item.path}/${index}`, depth: item.depth + 1 })
      } else {
        const prototype = Object.getPrototypeOf(current)
        if (prototype !== Object.prototype && prototype !== null) { add(issues, 'INVALID_TYPE', item.path, 'must contain only plain JSON objects and arrays'); continue }
        let keyCount = 0
        for (const key in current) {
          if (!Object.hasOwn(current, key)) continue
          keyCount += 1
          if (keyCount > 1_024) { add(issues, 'LIMIT_EXCEEDED', item.path, 'object exceeds 1024 fields'); break }
          stack.push({ value: (current as JsonObject)[key], path: `${item.path}/${pointer(key)}`, depth: item.depth + 1 })
        }
      }
    }
  } catch {
    add(issues, 'INVALID_VALUE', '', `${label} could not be safely inspected`)
  }
  return issues
}

export function sameWire(left: unknown, right: unknown): boolean {
  if (left === right) return true
  if (Array.isArray(left) || Array.isArray(right)) return Array.isArray(left) && Array.isArray(right) && left.length === right.length && left.every((entry, index) => sameWire(entry, right[index]))
  if (!isObject(left) || !isObject(right)) return false
  const leftKeys = Object.keys(left).sort()
  const rightKeys = Object.keys(right).sort()
  return leftKeys.length === rightKeys.length && leftKeys.every((key, index) => key === rightKeys[index] && sameWire(left[key], right[key]))
}

export function canonicalPart(value: string): string {
  try { return asciiLowerNative(value.split('/').map((segment) => decodeURIComponent(segment)).join('/')) } catch { return value }
}

export function validPartName(value: string): boolean {
  if (value.startsWith('/') || value.includes('\\') || value.includes('//')) return false
  return value.split('/').every((segment) => {
    if (segment === '.' || segment === '..' || !PART_SEGMENT.test(segment)) return false
    try {
      const decoded = decodeURIComponent(segment)
      return decoded !== '.' && decoded !== '..' && !decoded.endsWith('.') && !/[\\/?#%]/.test(decoded) && ![...decoded].some((character) => {
        const code = character.codePointAt(0) ?? 0
        return code < 0x20 || code === 0x7f
      })
    } catch { return false }
  })
}

export function validateNumberingSource(value: unknown, path: string, issues: NativeDocxValidationIssue[]): void {
  const source = exactObject(value, path, DOCX_PAGE_PAINT_V1_BINDING_FIELDS.NumberingSourceV1, issues)
  if (!source) return
  for (const key of ['relationships_part', 'part_name'] as const) {
    const part = stringValue(source[key], `${path}/${key}`, issues, /[^\u0000\r\n]+/, 4096)
    if (part && !validPartName(part)) add(issues, 'INVALID_VALUE', `${path}/${key}`, 'must be a canonical OPC part name')
  }
  for (const key of ['relationships_sha256', 'part_sha256', 'model_sha256'] as const) stringValue(source[key], `${path}/${key}`, issues, SHA256, 71)
  stringValue(source.relationship_id, `${path}/relationship_id`, issues)
  const relationshipType = stringValue(source.relationship_type, `${path}/relationship_type`, issues, ABSOLUTE_URI, 4096)
  if (relationshipType && relationshipType !== 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering' && relationshipType !== 'http://purl.oclc.org/ooxml/officeDocument/relationships/numbering') add(issues, 'INVALID_VALUE', `${path}/relationship_type`, 'must be the Strict or Transitional numbering relationship type')
  stringValue(source.relationship_target, `${path}/relationship_target`, issues, /[^\u0000\r\n]+/, 4096)
  if (source.content_type !== 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml') add(issues, 'INVALID_VALUE', `${path}/content_type`, 'must be the WordprocessingML numbering content type')
}


export function paintCommandID(placedLineID: string, fragmentID: string, glyphIndex: number): string {
  return `paint:${placedLineID}:${fragmentID}:${glyphIndex}`
}

export function paintImageCommandID(placedLineID: string, fragmentID: string): string {
  return `paint:${placedLineID}:${fragmentID}:image`
}

export function paintNoteSeparatorCommandID(placedLineID: string): string {
  return `paint:${placedLineID}:separator`
}


export function placedLineIDMatches(placedLineID: string, shapedLineID: string): boolean {
  if (placedLineID === `placed:${shapedLineID}`) return true
  const suffix = `:${shapedLineID}`
  if (!placedLineID.endsWith(suffix)) return false
  const prefix = placedLineID.slice(0, -suffix.length)
  return /^placed-note:[0-9]+:[0-9]+$/.test(prefix) || prefix.startsWith('placed:header:') || prefix.startsWith('placed:footer:')
}


export function validateFace(value: unknown, path: string, issues: NativeDocxValidationIssue[]): NativeDocxContentAddressedFaceV1 | undefined {
  const entry = exactUnionObject(value, path, DOCX_PAGE_PAINT_V1_BINDING_FIELDS.FaceV1, issues)
  if (!entry) return undefined
  const faceID = stringValue(entry.face_id, `${path}/face_id`, issues)
  const digest = stringValue(entry.content_digest, `${path}/content_digest`, issues, SHA256, 71)
  let collection: number | undefined
  if (entry.collection_index !== undefined) collection = integer(entry.collection_index, `${path}/collection_index`, issues, 0, 65_535)
  return faceID && digest ? { face_id: faceID, content_digest: digest, ...(collection !== undefined ? { collection_index: collection } : {}) } : undefined
}

export function validatePaintPath(value: unknown, path: string, outlineKind: string | undefined, issues: NativeDocxValidationIssue[], state: { pathCommands: number; nodes: number }): void {
  if (!Array.isArray(value)) { add(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, 'must be an array'); return }
  if ((outlineKind === 'path' && value.length === 0) || (outlineKind === 'empty' && value.length !== 0)) add(issues, 'INVALID_UNION', path, 'path presence must exactly match outline_kind')
  if (value.length > DOCX_PAGE_PAINT_LIMITS.maxPathCommandsPerGlyph) add(issues, 'LIMIT_EXCEEDED', path, `must contain at most ${DOCX_PAGE_PAINT_LIMITS.maxPathCommandsPerGlyph} commands`)
  state.pathCommands += value.length
  if (state.pathCommands > DOCX_PAGE_PAINT_LIMITS.maxPathCommands) add(issues, 'LIMIT_EXCEEDED', path, `paint paths exceed ${DOCX_PAGE_PAINT_LIMITS.maxPathCommands} commands`)
  let open = false
  let drawn = false
  let minX = Infinity
  let maxX = -Infinity
  let minY = Infinity
  let maxY = -Infinity
  value.slice(0, DOCX_PAGE_PAINT_LIMITS.maxPathCommandsPerGlyph).forEach((commandValue, index) => {
    const commandPath = `${path}/${index}`
    if (!isObject(commandValue) || typeof commandValue.kind !== 'string') { add(issues, 'INVALID_TYPE', commandPath, 'must be a path command object'); return }
    const fields = commandValue.kind === 'move_to' ? DOCX_PAGE_PAINT_V1_BINDING_FIELDS.PathMoveV1 : commandValue.kind === 'line_to' ? DOCX_PAGE_PAINT_V1_BINDING_FIELDS.PathLineV1 : commandValue.kind === 'quadratic_to' ? DOCX_PAGE_PAINT_V1_BINDING_FIELDS.PathQuadraticV1 : commandValue.kind === 'cubic_to' ? DOCX_PAGE_PAINT_V1_BINDING_FIELDS.PathCubicV1 : commandValue.kind === 'close_path' ? DOCX_PAGE_PAINT_V1_BINDING_FIELDS.PathCloseV1 : undefined
    if (!fields) { add(issues, 'INVALID_VALUE', `${commandPath}/kind`, 'unknown path command kind'); return }
    const command = exactObject(commandValue, commandPath, fields, issues)
    if (!command) return
    for (const key of fields) if (key !== 'kind') integer(command[key], `${commandPath}/${key}`, issues, -DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
    for (const [key, coordinateValue] of Object.entries(command)) if (key.endsWith('x_millipoints') && typeof coordinateValue === 'number') { minX = Math.min(minX, coordinateValue); maxX = Math.max(maxX, coordinateValue) }
    for (const [key, coordinateValue] of Object.entries(command)) if (key.endsWith('y_millipoints') && typeof coordinateValue === 'number') { minY = Math.min(minY, coordinateValue); maxY = Math.max(maxY, coordinateValue) }
    if (commandValue.kind === 'move_to') { if (open) add(issues, 'INVALID_VALUE', commandPath, 'previous contour must close before move_to'); open = true; drawn = false }
    else if (commandValue.kind === 'close_path') { if (!open || !drawn) add(issues, 'INVALID_VALUE', commandPath, 'close_path requires a non-empty open contour'); open = false; drawn = false }
    else { if (!open) add(issues, 'INVALID_VALUE', commandPath, 'drawing command requires a preceding move_to'); drawn = true }
  })
  if (open) add(issues, 'INVALID_VALUE', path, 'every glyph contour must close')
  if (outlineKind === 'path' && !(maxX > minX && maxY > minY)) add(issues, 'INVALID_VALUE', path, 'outlined glyph path must have non-degenerate two-dimensional bounds')
}

/** Strict structural decoder for stored or transported paint output. */
export function decodeNativeDocxPagePaintV1(value: unknown): DecodeNativeDocxPagePaintV1Result {
  const preflight = preflightWire(value, 'page-paint output')
  if (preflight.length > 0) return { ok: false, issues: preflight }
  const snapshot = safeClone(value)
  if (snapshot === undefined) return { ok: false, issues: [issue('INVALID_VALUE', '', 'page-paint output must be cloneable JSON wire data')] }
  const issues: NativeDocxValidationIssue[] = []
  const root = exactObject(snapshot, '', DOCX_PAGE_PAINT_V1_BINDING_FIELDS.OutputV1, issues)
  if (!root) return { ok: false, issues }
  if (root.protocol !== DOCX_PAGE_PAINT_PROTOCOL) add(issues, 'UNSUPPORTED_PROTOCOL', '/protocol', `must equal ${DOCX_PAGE_PAINT_PROTOCOL}`)
  if (root.version !== DOCX_PAGE_PAINT_VERSION) add(issues, 'UNSUPPORTED_VERSION', '/version', `must equal ${DOCX_PAGE_PAINT_VERSION}`)
  const status = root.status === 'painted' || root.status === 'refused' ? root.status : undefined
  if (!status) add(issues, 'INVALID_VALUE', '/status', 'must be painted or refused')
  const provenance = exactUnionObject(root.provenance, '/provenance', DOCX_PAGE_PAINT_V1_BINDING_FIELDS.ProvenanceV1, issues)
  if (provenance) {
    for (const key of DOCX_PAGE_PAINT_V1_BINDING_FIELDS.ProvenanceV1) if (key !== 'numbering_source' && !(key in provenance)) add(issues, 'REQUIRED', `/provenance/${key}`, 'field is required')
    stringValue(provenance.document_id, '/provenance/document_id', issues)
    stringValue(provenance.revision, '/provenance/revision', issues)
    stringValue(provenance.package_sha256, '/provenance/package_sha256', issues, SHA256, 71)
    const mainPart = stringValue(provenance.main_part, '/provenance/main_part', issues, /[^\u0000\r\n]+/, 4_096)
    if (mainPart && !validPartName(mainPart)) add(issues, 'INVALID_VALUE', '/provenance/main_part', 'must be a canonical OPC part name')
    stringValue(provenance.body_story_id, '/provenance/body_story_id', issues)
    if (provenance.numbering_source !== undefined) validateNumberingSource(provenance.numbering_source, '/provenance/numbering_source', issues)
    const shaped = exactObject(provenance.shaped_lines, '/provenance/shaped_lines', DOCX_PAGE_PAINT_V1_BINDING_FIELDS.ShapedSourceV1, issues)
    if (shaped) { if (shaped.protocol !== 'injoffice.docx.shaped-lines') add(issues, 'UNSUPPORTED_PROTOCOL', '/provenance/shaped_lines/protocol', 'must equal injoffice.docx.shaped-lines'); integer(shaped.version, '/provenance/shaped_lines/version', issues, 1, 1); integer(shaped.available_width_millipoints, '/provenance/shaped_lines/available_width_millipoints', issues, 1, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints); integer(shaped.tab_interval_millipoints, '/provenance/shaped_lines/tab_interval_millipoints', issues, 1, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints); stringValue(shaped.sha256, '/provenance/shaped_lines/sha256', issues, SHA256, 71) }
    const paginated = exactObject(provenance.paginated_layout, '/provenance/paginated_layout', DOCX_PAGE_PAINT_V1_BINDING_FIELDS.PaginatedSourceV1, issues)
    if (paginated) { if (paginated.protocol !== 'injoffice.docx.paginated-layout') add(issues, 'UNSUPPORTED_PROTOCOL', '/provenance/paginated_layout/protocol', 'must equal injoffice.docx.paginated-layout'); integer(paginated.version, '/provenance/paginated_layout/version', issues, 1, 1); stringValue(paginated.sha256, '/provenance/paginated_layout/sha256', issues, SHA256, 71) }
    const headerFooter = exactObject(provenance.header_footer_layout, '/provenance/header_footer_layout', DOCX_PAGE_PAINT_V1_BINDING_FIELDS.HeaderFooterSourceV1, issues)
    if (headerFooter) { if (headerFooter.protocol !== 'injoffice.docx.header-footer-layout') add(issues, 'UNSUPPORTED_PROTOCOL', '/provenance/header_footer_layout/protocol', `must equal ${'injoffice.docx.header-footer-layout'}`); integer(headerFooter.version, '/provenance/header_footer_layout/version', issues, 1, 1); stringValue(headerFooter.sha256, '/provenance/header_footer_layout/sha256', issues, SHA256, 71) }
    const tableProjection = exactObject(provenance.table_projection, '/provenance/table_projection', DOCX_PAGE_PAINT_V1_BINDING_FIELDS.TableProjectionV1, issues)
    if (tableProjection) stringValue(tableProjection.sha256, '/provenance/table_projection/sha256', issues, SHA256, 71)
    const manifest = exactObject(provenance.font_manifest, '/provenance/font_manifest', DOCX_PAGE_PAINT_V1_BINDING_FIELDS.ManifestV1, issues)
    if (manifest) { stringValue(manifest.manifest_id, '/provenance/font_manifest/manifest_id', issues); stringValue(manifest.revision, '/provenance/font_manifest/revision', issues); stringValue(manifest.sha256, '/provenance/font_manifest/sha256', issues, SHA256, 71) }
    const media = exactObject(provenance.media_assets, '/provenance/media_assets', DOCX_PAGE_PAINT_V1_BINDING_FIELDS.MediaSourceV1, issues)
    if (media) stringValue(media.sha256, '/provenance/media_assets/sha256', issues, SHA256, 71)
    const providers = exactObject(provenance.providers, '/provenance/providers', DOCX_PAGE_PAINT_V1_BINDING_FIELDS.ProvidersV1, issues)
    if (providers) {
      for (const key of ['resolver_id', 'resolver_revision', 'shaper_id', 'shaper_revision', 'bidi_id', 'bidi_revision', 'bidi_unicode_version', 'unicode13_revision'] as const) stringValue(providers[key], `/provenance/providers/${key}`, issues, SHORT_ID, 256)
      for (const key of ['outline_id', 'outline_revision'] as const) stringValue(providers[key], `/provenance/providers/${key}`, issues, PROVIDER_ID, 128)
    }
    const settings = decodeNativeDocxPaginationSettings(provenance.pagination_settings)
    if (!settings.ok) issues.push(...settings.issues.map((entry) => ({ ...entry, path: `/provenance/pagination_settings${entry.path}` })))
    else if (mainPart && canonicalPart(mainPart) !== canonicalPart(settings.value.main_part)) add(issues, 'BROKEN_REFERENCE', '/provenance/main_part', 'main part must match pagination settings provenance')
  }
  const diagnostics = Array.isArray(root.diagnostics) ? root.diagnostics : []
  if (!Array.isArray(root.diagnostics)) add(issues, root.diagnostics === undefined ? 'REQUIRED' : 'INVALID_TYPE', '/diagnostics', 'must be an array')
  if (diagnostics.length > DOCX_NATIVE_LIMITS.maxIssues) add(issues, 'LIMIT_EXCEEDED', '/diagnostics', `must contain at most ${DOCX_NATIVE_LIMITS.maxIssues} entries`)
  diagnostics.slice(0, DOCX_NATIVE_LIMITS.maxIssues).forEach((value, index) => {
    const path = `/diagnostics/${index}`
    const entry = exactObject(value, path, DOCX_PAGE_PAINT_V1_BINDING_FIELDS.DiagnosticV1, issues)
    if (!entry) return
    const codes: readonly NativeDocxPagePaintDiagnosticCode[] = ['upstream-refused', 'unsupported-diagnostic', 'unsupported-source', 'missing-font', 'missing-glyph', 'provider-refusal', 'provider-mismatch', 'provider-failure', 'invalid-provider-output', 'invalid-path', 'resource-limit', 'identity-mismatch', 'incomplete-page']
    if (typeof entry.code !== 'string' || !codes.includes(entry.code as NativeDocxPagePaintDiagnosticCode)) add(issues, 'INVALID_VALUE', `${path}/code`, 'must be an enumerated page-paint diagnostic code')
    if (entry.severity !== 'unsupported') add(issues, 'INVALID_VALUE', `${path}/severity`, 'must equal unsupported')
    stringValue(entry.scope_id, `${path}/scope_id`, issues, ID, 1024)
    stringValue(entry.message, `${path}/message`, issues, /[^\u0000\r\n]+/, DOCX_PAGE_PAINT_LIMITS.maxProviderMessageLength)
  })
  let resources: NativeDocxPagePaintMediaAssetV1[] = []
  try { resources = decodeNativeDocxPagePaintResourceListV1(root.resources) } catch (error) {
    add(issues, 'INVALID_VALUE', '/resources', error instanceof Error ? error.message : 'media resources are invalid')
  }
  if (status === 'refused' && resources.length !== 0) add(issues, 'INVALID_UNION', '/resources', 'refused page paint cannot expose partial media resources')
  const pages = Array.isArray(root.pages) ? root.pages : []
  if (!Array.isArray(root.pages)) add(issues, root.pages === undefined ? 'REQUIRED' : 'INVALID_TYPE', '/pages', 'must be an array')
  if (pages.length > DOCX_PAGE_PAINT_LIMITS.maxPages) add(issues, 'LIMIT_EXCEEDED', '/pages', `must contain at most ${DOCX_PAGE_PAINT_LIMITS.maxPages} pages`)
  const state = { pathCommands: 0, nodes: 0 }
  let glyphs = 0
  const pageIDs = new Set<string>()
  const commandIDs = new Set<string>()
  const placedLineIDs = new Set<string>()
  let lineCount = 0
  pages.slice(0, DOCX_PAGE_PAINT_LIMITS.maxPages).forEach((value, pageIndex) => {
    const path = `/pages/${pageIndex}`
    const page = exactObject(value, path, DOCX_PAGE_PAINT_V1_BINDING_FIELDS.PageV1, issues)
    if (!page) return
    const pageID = stringValue(page.id, `${path}/id`, issues, ID, 1024)
    if (pageID && pageIDs.has(pageID)) add(issues, 'DUPLICATE_ID', `${path}/id`, 'page id is duplicated')
    if (pageID) pageIDs.add(pageID)
    const ordinal = integer(page.ordinal, `${path}/ordinal`, issues, 0, DOCX_PAGE_PAINT_LIMITS.maxPages - 1)
    if (ordinal !== undefined && ordinal !== pageIndex) add(issues, 'INVALID_VALUE', `${path}/ordinal`, `must equal source-order ordinal ${pageIndex}`)
    const pageSectionID = stringValue(page.section_id, `${path}/section_id`, issues)
    const pageSectionIDs = Array.isArray(page.section_ids) ? page.section_ids.map((value, index) => stringValue(value, `${path}/section_ids/${index}`, issues)).filter((value): value is string => value !== undefined) : []
    if (!Array.isArray(page.section_ids)) add(issues, page.section_ids === undefined ? 'REQUIRED' : 'INVALID_TYPE', `${path}/section_ids`, 'must be an array')
    if (pageSectionID && pageSectionIDs[0] !== pageSectionID) add(issues, 'INVALID_VALUE', `${path}/section_ids`, 'first section identity must equal page section_id')
    if (new Set(pageSectionIDs).size !== pageSectionIDs.length) add(issues, 'DUPLICATE_ID', `${path}/section_ids`, 'page section identities must be unique')
    if (page.kind !== 'content' && page.kind !== 'parity-blank') add(issues, 'INVALID_VALUE', `${path}/kind`, 'must be content or parity-blank')
    const pageWidth = integer(page.width_millipoints, `${path}/width_millipoints`, issues, 1, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
    const pageHeight = integer(page.height_millipoints, `${path}/height_millipoints`, issues, 1, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
    const box = exactObject(page.body_box, `${path}/body_box`, DOCX_PAGE_PAINT_V1_BINDING_FIELDS.BodyBoxV1, issues)
    if (box) {
      const x = integer(box.x_millipoints, `${path}/body_box/x_millipoints`, issues, 0, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
      const y = integer(box.y_millipoints, `${path}/body_box/y_millipoints`, issues, 0, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
      const width = integer(box.width_millipoints, `${path}/body_box/width_millipoints`, issues, 1, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
      const height = integer(box.height_millipoints, `${path}/body_box/height_millipoints`, issues, 1, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
      if (x !== undefined && width !== undefined && pageWidth !== undefined && x + width > pageWidth) add(issues, 'OUT_OF_RANGE', `${path}/body_box`, 'body box exceeds page width')
      if (y !== undefined && height !== undefined && pageHeight !== undefined && y + height > pageHeight) add(issues, 'OUT_OF_RANGE', `${path}/body_box`, 'body box exceeds page height')
    }
    const pageColumns = new Map<string, { sectionID?: string; ordinal?: number; x?: number; y?: number; width?: number; height?: number }>()
    if (!Array.isArray(page.columns)) add(issues, page.columns === undefined ? 'REQUIRED' : 'INVALID_TYPE', `${path}/columns`, 'must be an array')
    else page.columns.forEach((columnValue, columnIndex) => {
      const columnPath = `${path}/columns/${columnIndex}`
      const column = exactObject(columnValue, columnPath, DOCX_PAGE_PAINT_V1_BINDING_FIELDS.ColumnV1, issues)
      if (!column) return
      const id = stringValue(column.id, `${columnPath}/id`, issues, ID, 1024)
      if (id && pageColumns.has(id)) add(issues, 'DUPLICATE_ID', `${columnPath}/id`, 'page column id is duplicated')
      const columnSectionID = stringValue(column.section_id, `${columnPath}/section_id`, issues)
      const columnOrdinal = integer(column.ordinal, `${columnPath}/ordinal`, issues, 0, 44)
      const columnX = integer(column.x_millipoints, `${columnPath}/x_millipoints`, issues, 0, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
      const columnY = integer(column.y_millipoints, `${columnPath}/y_millipoints`, issues, 0, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
      const columnWidth = integer(column.width_millipoints, `${columnPath}/width_millipoints`, issues, 1, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
      const columnHeight = integer(column.height_millipoints, `${columnPath}/height_millipoints`, issues, 1, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
      if (id) pageColumns.set(id, { sectionID: columnSectionID, ordinal: columnOrdinal, x: columnX, y: columnY, width: columnWidth, height: columnHeight })
      if (columnSectionID && !pageSectionIDs.includes(columnSectionID)) add(issues, 'BROKEN_REFERENCE', `${columnPath}/section_id`, 'paint column section must occur on the page')
      if (pageWidth !== undefined && columnX !== undefined && columnWidth !== undefined && columnX + columnWidth > pageWidth) add(issues, 'OUT_OF_RANGE', columnPath, 'paint column exceeds page width')
      if (pageHeight !== undefined && columnY !== undefined && columnHeight !== undefined && columnY + columnHeight > pageHeight) add(issues, 'OUT_OF_RANGE', columnPath, 'paint column exceeds page height')
    })
    if (page.background_rgb !== 'FFFFFF') add(issues, 'INVALID_VALUE', `${path}/background_rgb`, 'must equal explicit white page background FFFFFF')
    const clip = exactObject(page.clip_box, `${path}/clip_box`, DOCX_PAGE_PAINT_V1_BINDING_FIELDS.ClipBoxV1, issues)
    if (clip) {
      const clipX = integer(clip.x_millipoints, `${path}/clip_box/x_millipoints`, issues, 0, 0)
      const clipY = integer(clip.y_millipoints, `${path}/clip_box/y_millipoints`, issues, 0, 0)
      const clipWidth = integer(clip.width_millipoints, `${path}/clip_box/width_millipoints`, issues, 1, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
      const clipHeight = integer(clip.height_millipoints, `${path}/clip_box/height_millipoints`, issues, 1, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
      if (clipX !== 0 || clipY !== 0 || clipWidth !== pageWidth || clipHeight !== pageHeight) add(issues, 'INVALID_VALUE', `${path}/clip_box`, 'clip box must exactly equal page geometry')
    }
    const lines = Array.isArray(page.lines) ? page.lines : []
    if (!Array.isArray(page.lines)) add(issues, page.lines === undefined ? 'REQUIRED' : 'INVALID_TYPE', `${path}/lines`, 'must be an array')
    if (lines.length > DOCX_PAGE_PAINT_LIMITS.maxLines) add(issues, 'LIMIT_EXCEEDED', `${path}/lines`, `must contain at most ${DOCX_PAGE_PAINT_LIMITS.maxLines} lines`)
    const pageLineIDs = new Set<string>()
    const referencedCommandIDs: string[] = []
    const commandOwners = new Map<string, string>()
    const commandPlacementOwners = new Map<string, string>()
    lines.slice(0, DOCX_PAGE_PAINT_LIMITS.maxLines).forEach((lineValue, lineIndex) => {
      lineCount += 1
      const linePath = `${path}/lines/${lineIndex}`
      const lineFields = isObject(lineValue) && (lineValue.region === 'header' || lineValue.region === 'footer') ? DOCX_PAGE_PAINT_V1_BINDING_FIELDS.HeaderFooterLineV1 : DOCX_PAGE_PAINT_V1_BINDING_FIELDS.BodyLineV1
      const line = exactObject(lineValue, linePath, lineFields, issues)
      if (!line) return
      const placedID = stringValue(line.placed_line_id, `${linePath}/placed_line_id`, issues, ID, 1024)
      const lineID = stringValue(line.line_id, `${linePath}/line_id`, issues, ID, 1024)
      if (placedID && placedLineIDs.has(placedID)) add(issues, 'DUPLICATE_ID', `${linePath}/placed_line_id`, 'placed line id is duplicated')
      if (placedID) placedLineIDs.add(placedID)
      if (lineID && pageLineIDs.has(lineID)) add(issues, 'DUPLICATE_ID', `${linePath}/line_id`, 'line id is duplicated on the page')
      if (lineID) pageLineIDs.add(lineID)
      if (placedID && lineID && !placedLineIDMatches(placedID, lineID)) add(issues, 'INVALID_VALUE', `${linePath}/placed_line_id`, 'placed line id must derive from its body, header/footer, or note placement and shaped line id')
      stringValue(line.paragraph_id, `${linePath}/paragraph_id`, issues)
      const lineSectionID = stringValue(line.section_id, `${linePath}/section_id`, issues)
      const region = line.region === 'body' || line.region === 'header' || line.region === 'footer' || line.region === 'footnote' || line.region === 'endnote' ? line.region : undefined
      if (!region) add(issues, 'INVALID_VALUE', `${linePath}/region`, 'must be body, header, footer, footnote, or endnote')
      if (lineSectionID && !pageSectionIDs.includes(lineSectionID)) add(issues, 'BROKEN_REFERENCE', `${linePath}/section_id`, 'paint line section must occur on the page')
      const isColumnFlow = region === 'body' || region === 'footnote' || region === 'endnote'
      const columnID = isColumnFlow ? stringValue(line.column_id, `${linePath}/column_id`, issues, ID, 1024) : undefined
      const lineColumnOrdinal = isColumnFlow ? integer(line.column_ordinal, `${linePath}/column_ordinal`, issues, 0, 44) : undefined
      if (!isColumnFlow && (line.column_id !== undefined || line.column_ordinal !== undefined)) add(issues, 'INVALID_UNION', linePath, 'header/footer paint lines cannot claim a body or note flow column')
      const owningColumn = columnID ? pageColumns.get(columnID) : undefined
      if (isColumnFlow && (!owningColumn || owningColumn.sectionID !== lineSectionID || owningColumn.ordinal !== lineColumnOrdinal)) add(issues, 'BROKEN_REFERENCE', `${linePath}/column_id`, 'body and note paint lines must reference their exact page section column')
      integer(line.source_line_ordinal, `${linePath}/source_line_ordinal`, issues, 0, DOCX_PAGE_PAINT_LIMITS.maxLines)
      const lineX = integer(line.x_millipoints, `${linePath}/x_millipoints`, issues, 0, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
      const lineY = integer(line.y_millipoints, `${linePath}/y_millipoints`, issues, 0, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
      const lineWidth = integer(line.width_millipoints, `${linePath}/width_millipoints`, issues, 0, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
      const lineHeight = integer(line.height_millipoints, `${linePath}/height_millipoints`, issues, 1, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
      if (isColumnFlow && owningColumn?.x !== undefined && owningColumn.width !== undefined && lineX !== undefined && lineWidth !== undefined && (lineX < owningColumn.x || lineX + lineWidth > owningColumn.x + owningColumn.width)) add(issues, 'OUT_OF_RANGE', linePath, 'paint line exceeds its referenced column width')
      if (isColumnFlow && owningColumn?.y !== undefined && owningColumn.height !== undefined && lineY !== undefined && lineHeight !== undefined && (lineY < owningColumn.y || lineY + lineHeight > owningColumn.y + owningColumn.height)) add(issues, 'OUT_OF_RANGE', linePath, 'paint line exceeds its referenced column height')
      integer(line.baseline_y_millipoints, `${linePath}/baseline_y_millipoints`, issues, 0, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
      if (!Array.isArray(line.command_ids)) add(issues, line.command_ids === undefined ? 'REQUIRED' : 'INVALID_TYPE', `${linePath}/command_ids`, 'must be an array')
      else line.command_ids.forEach((id, index) => {
        const commandID = stringValue(id, `${linePath}/command_ids/${index}`, issues, ID, 1024)
        if (commandID) {
          if (commandOwners.has(commandID)) add(issues, 'DUPLICATE_ID', `${linePath}/command_ids/${index}`, 'paint command is referenced by more than one line')
          if (lineID) commandOwners.set(commandID, lineID)
          if (placedID) commandPlacementOwners.set(commandID, placedID)
          referencedCommandIDs.push(commandID)
        }
      })
    })
    const commands = Array.isArray(page.commands) ? page.commands : []
    if (!Array.isArray(page.commands)) add(issues, page.commands === undefined ? 'REQUIRED' : 'INVALID_TYPE', `${path}/commands`, 'must be an array')
    commands.slice(0, DOCX_PAGE_PAINT_LIMITS.maxGlyphs).forEach((commandValue, commandIndex) => {
      const commandPath = `${path}/commands/${commandIndex}`
      const kind = isObject(commandValue) ? commandValue.kind : undefined
      const fields = kind === 'fill_glyph_path' ? DOCX_PAGE_PAINT_V1_BINDING_FIELDS.GlyphCommandV1
        : kind === 'paint_inline_image' ? DOCX_PAGE_PAINT_V1_BINDING_FIELDS.ImageCommandV1
          : kind === 'fill_table_cell' ? DOCX_PAGE_PAINT_V1_BINDING_FIELDS.CellFillCommandV1
            : kind === 'stroke_table_border' ? DOCX_PAGE_PAINT_V1_BINDING_FIELDS.BorderCommandV1
              : kind === 'stroke_note_separator' ? DOCX_PAGE_PAINT_V1_BINDING_FIELDS.NoteSeparatorCommandV1 : undefined
      if (!fields) { add(issues, 'INVALID_VALUE', `${commandPath}/kind`, 'must be fill_glyph_path, paint_inline_image, fill_table_cell, stroke_table_border, or stroke_note_separator'); return }
      const command = exactObject(commandValue, commandPath, fields, issues)
      if (!command) return
      const commandID = stringValue(command.id, `${commandPath}/id`, issues, ID, 1024)
      if (commandID && commandIDs.has(commandID)) add(issues, 'DUPLICATE_ID', `${commandPath}/id`, 'paint command id is duplicated')
      if (commandID) commandIDs.add(commandID)
      if (command.kind === 'fill_table_cell') {
        for (const key of ['table_id', 'row_id', 'cell_id'] as const) stringValue(command[key], `${commandPath}/${key}`, issues)
        for (const key of ['x_millipoints', 'y_millipoints'] as const) integer(command[key], `${commandPath}/${key}`, issues, 0, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
        for (const key of ['width_millipoints', 'height_millipoints'] as const) integer(command[key], `${commandPath}/${key}`, issues, 1, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
        stringValue(command.fill_rgb, `${commandPath}/fill_rgb`, issues, RGB, 6)
        return
      }
      if (command.kind === 'stroke_table_border') {
        for (const key of ['table_id', 'row_id', 'cell_id'] as const) stringValue(command[key], `${commandPath}/${key}`, issues)
        if (!['top', 'right', 'bottom', 'left'].includes(command.edge as string)) add(issues, 'INVALID_VALUE', `${commandPath}/edge`, 'must be a physical table-cell edge')
        for (const key of ['x1_millipoints', 'y1_millipoints', 'x2_millipoints', 'y2_millipoints'] as const) integer(command[key], `${commandPath}/${key}`, issues, 0, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
        integer(command.width_millipoints, `${commandPath}/width_millipoints`, issues, 1, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
        stringValue(command.stroke_rgb, `${commandPath}/stroke_rgb`, issues, RGB, 6)
        return
      }
      const commandLineID = stringValue(command.line_id, `${commandPath}/line_id`, issues, ID, 1024)
      const commandPlacementID = commandID ? commandPlacementOwners.get(commandID) : undefined
      if (commandID && commandLineID && commandOwners.get(commandID) !== commandLineID) add(issues, 'BROKEN_REFERENCE', `${commandPath}/line_id`, 'paint command line id must match the line that owns its command id')
      if (command.kind === 'stroke_note_separator') {
        stringValue(command.story_id, `${commandPath}/story_id`, issues, ID, 1024)
        const x1 = integer(command.x1_millipoints, `${commandPath}/x1_millipoints`, issues, 0, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
        const y1 = integer(command.y1_millipoints, `${commandPath}/y1_millipoints`, issues, 0, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
        const x2 = integer(command.x2_millipoints, `${commandPath}/x2_millipoints`, issues, 0, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
        const y2 = integer(command.y2_millipoints, `${commandPath}/y2_millipoints`, issues, 0, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
        integer(command.width_millipoints, `${commandPath}/width_millipoints`, issues, 1, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
        stringValue(command.stroke_rgb, `${commandPath}/stroke_rgb`, issues, RGB, 6)
        if (x1 !== undefined && x2 !== undefined && x2 <= x1 || y1 !== undefined && y2 !== undefined && y1 !== y2 || command.stroke_rgb !== '000000') add(issues, 'INVALID_VALUE', commandPath, 'note separator must be one positive-width horizontal black rule')
        if (commandID && commandPlacementID && commandID !== paintNoteSeparatorCommandID(commandPlacementID)) add(issues, 'INVALID_VALUE', `${commandPath}/id`, 'note separator id must derive from its placed line')
        return
      }
      const fragmentID = stringValue(command.fragment_id, `${commandPath}/fragment_id`, issues, ID, 1024)
      stringValue(command.source_id, `${commandPath}/source_id`, issues)
      if (kind === 'fill_glyph_path') {
        glyphs += 1
        const glyphIndex = integer(command.glyph_index, `${commandPath}/glyph_index`, issues, 0, DOCX_PAGE_PAINT_LIMITS.maxGlyphs)
        if (commandID && fragmentID && glyphIndex !== undefined && commandPlacementID && commandID !== paintCommandID(commandPlacementID, fragmentID, glyphIndex)) add(issues, 'INVALID_VALUE', `${commandPath}/id`, 'paint command id must derive from placed line, fragment id, and glyph index')
        validateFace(command.face, `${commandPath}/face`, issues)
        integer(command.glyph_id, `${commandPath}/glyph_id`, issues, 0, 0xffffffff)
        integer(command.font_size_millipoints, `${commandPath}/font_size_millipoints`, issues, 1, 1_638_000)
        stringValue(command.fill_rgb, `${commandPath}/fill_rgb`, issues, RGB, 6)
        if (command.fill_rule !== 'nonzero') add(issues, 'INVALID_VALUE', `${commandPath}/fill_rule`, 'must equal nonzero')
        const outlineKind = command.outline_kind === 'path' || command.outline_kind === 'empty' ? command.outline_kind : undefined
        if (!outlineKind) add(issues, 'INVALID_VALUE', `${commandPath}/outline_kind`, 'must be path or empty')
        validatePaintPath(command.path, `${commandPath}/path`, outlineKind, issues, state)
      } else {
        stringValue(command.drawing_id, `${commandPath}/drawing_id`, issues, ID, 1024)
        const assetID = stringValue(command.asset_id, `${commandPath}/asset_id`, issues, ID, 1024)
        if (assetID && !resources.some((resource) => resource.id === assetID)) add(issues, 'BROKEN_REFERENCE', `${commandPath}/asset_id`, 'image command must reference one canonical output media resource')
        if (commandID && fragmentID && commandPlacementID && commandID !== paintImageCommandID(commandPlacementID, fragmentID)) add(issues, 'INVALID_VALUE', `${commandPath}/id`, 'image paint command id must derive from placed line and fragment id')
        integer(command.x_millipoints, `${commandPath}/x_millipoints`, issues, 0, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
        integer(command.y_millipoints, `${commandPath}/y_millipoints`, issues, 0, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
        integer(command.width_millipoints, `${commandPath}/width_millipoints`, issues, 1, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
        integer(command.height_millipoints, `${commandPath}/height_millipoints`, issues, 1, DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints)
        const crop = exactObject(command.source_crop, `${commandPath}/source_crop`, DOCX_PAGE_PAINT_V1_BINDING_FIELDS.ImageCropV1, issues)
        if (crop && (crop.left !== 0 || crop.top !== 0 || crop.right !== 0 || crop.bottom !== 0 || crop.unit !== 'one-hundred-thousandth')) add(issues, 'INVALID_VALUE', `${commandPath}/source_crop`, 'v1 image crop must be the explicit full source rectangle')
        const transform = exactObject(command.transform, `${commandPath}/transform`, DOCX_PAGE_PAINT_V1_BINDING_FIELDS.ImageTransformV1, issues)
        if (transform && (transform.rotation_degrees !== 0 || transform.flip_horizontal !== false || transform.flip_vertical !== false)) add(issues, 'INVALID_VALUE', `${commandPath}/transform`, 'v1 image transform must be explicit identity')
      }
    })
    if (commands.length > DOCX_PAGE_PAINT_LIMITS.maxGlyphs) add(issues, 'LIMIT_EXCEEDED', `${path}/commands`, `commands exceed ${DOCX_PAGE_PAINT_LIMITS.maxGlyphs}`)
    const actualCommandIDs = commands.flatMap((command) => isObject(command) && (command.kind === 'fill_glyph_path' || command.kind === 'paint_inline_image' || command.kind === 'stroke_note_separator') && typeof command.id === 'string' ? [command.id] : [])
    if (actualCommandIDs.length !== referencedCommandIDs.length || actualCommandIDs.some((id, index) => id !== referencedCommandIDs[index])) add(issues, 'BROKEN_REFERENCE', `${path}/lines`, 'line command ids must exactly cover page commands in replay order')
    if (page.kind === 'parity-blank' && (commands.length > 0 || lines.length > 0)) add(issues, 'INVALID_UNION', path, 'parity-blank pages cannot contain lines or paint commands')
    if (!pageID) return
  })
  if (glyphs > DOCX_PAGE_PAINT_LIMITS.maxGlyphs) add(issues, 'LIMIT_EXCEEDED', '/pages', `glyph commands exceed ${DOCX_PAGE_PAINT_LIMITS.maxGlyphs}`)
  if (lineCount > DOCX_PAGE_PAINT_LIMITS.maxLines) add(issues, 'LIMIT_EXCEEDED', '/pages', `placed lines exceed ${DOCX_PAGE_PAINT_LIMITS.maxLines}`)
  if (status === 'refused' && (pages.length > 0 || diagnostics.length === 0)) add(issues, 'INVALID_UNION', '', 'refused output requires diagnostics and no pages')
  if (status === 'painted' && (pages.length === 0 || diagnostics.length > 0)) add(issues, 'INVALID_UNION', '', 'painted output requires complete pages and no diagnostics')
  issues.sort(compareNativeValidationIssues)
  if (issues.length > 0) return { ok: false, issues: issues.slice(0, DOCX_NATIVE_LIMITS.maxIssues) }
  return status === 'refused'
    ? { ok: true, value: { ...(snapshot as NativeDocxPagePaintRefusedV1), resources: [] } }
    : { ok: true, value: { ...(snapshot as NativeDocxPagePaintSuccessV1), resources } }
}

