/** Strict decoder for the renderer-neutral DOCX paginated-layout v1 output. */

import { DOCX_MAX_TWIPS_FOR_MILLIPOINTS, DOCX_NATIVE_LIMITS, type NativeDocxIssueCode, type NativeDocxValidationIssue } from './nativeContract.js'
import {
  DOCX_PAGINATED_LAYOUT_PROTOCOL,
  DOCX_PAGINATED_LAYOUT_VERSION,
  DOCX_PAGINATION_DIAGNOSTIC_CODES,
  DOCX_PAGINATION_LIMITS,
  decodeNativeDocxPaginationRequestV1,
  validateNativeDocxPaginatedLayoutSourceV1,
  type NativeDocxPaginatedLayoutV1,
} from './nativePaginationV1.js'
import { decodeNativeDocxPaginationSettings } from './nativePaginationSettings.js'
import { DOCX_SHAPED_LINES_PROTOCOL, DOCX_SHAPED_LINES_VERSION } from './nativeShapingLines.js'
import { asciiLowerNative, compareNativeValidationIssues } from './nativeDeterminism.js'

export type DecodeNativeDocxPaginatedLayoutResult =
  | { ok: true; value: NativeDocxPaginatedLayoutV1 }
  | { ok: false; issues: NativeDocxValidationIssue[] }

export const DOCX_PAGINATED_LAYOUT_V1_BINDING_FIELDS = {
  ShapedSourceV1: ['protocol', 'version', 'available_width_millipoints', 'tab_interval_millipoints'],
  FontManifestV1: ['manifest_id', 'revision'],
  ProvidersV1: ['resolver_id', 'resolver_revision', 'shaper_id', 'shaper_revision', 'bidi_id', 'bidi_revision', 'bidi_unicode_version', 'unicode13_revision'],
  NumberingSourceV1: ['relationships_part', 'relationships_sha256', 'relationship_id', 'relationship_type', 'relationship_target', 'part_name', 'content_type', 'part_sha256', 'model_sha256'],
  ProvenanceV1: ['document_id', 'revision', 'package_sha256', 'main_part', 'body_story_id', 'numbering_source', 'pagination_settings', 'shaped_lines', 'font_manifest', 'providers'],
  DiagnosticV1: ['code', 'severity', 'scope_id', 'source_code', 'source_message', 'message'],
  BodyBoxV1: ['x_millipoints', 'y_millipoints', 'width_millipoints', 'height_millipoints'],
  ColumnV1: ['id', 'section_id', 'ordinal', 'x_millipoints', 'y_millipoints', 'width_millipoints', 'height_millipoints'],
  HeaderFooterReferenceV1: ['kind', 'story_id', 'relationship_id'],
  PlacedLineV1: ['id', 'line_id', 'paragraph_id', 'section_id', 'column_id', 'column_ordinal', 'source_line_ordinal', 'x_millipoints', 'y_millipoints', 'width_millipoints', 'height_millipoints', 'repeated_table_header', 'table_cell_id'],
  PlacedNoteStoryV1: ['id', 'story_id', 'story_kind', 'note_role', 'native_story_id', 'relationship_id', 'ordinal', 'reference_run_id', 'number', 'section_id', 'column_id', 'column_ordinal', 'top_millipoints', 'height_millipoints', 'lines'],
  ParagraphSliceV1: ['id', 'paragraph_id', 'section_id', 'column_id', 'column_ordinal', 'slice_ordinal', 'first_line_ordinal', 'last_line_ordinal', 'line_ids', 'top_millipoints', 'height_millipoints', 'space_before_millipoints', 'continued_from_previous_page', 'continues_on_next_page', 'continued_from_previous_column', 'continues_in_next_column', 'repeated_table_header', 'table_cell_id'],
  PageV1: ['id', 'ordinal', 'section_id', 'section_ids', 'section_page_ordinal', 'kind', 'parity_reason', 'parity_before_section_id', 'width_millipoints', 'height_millipoints', 'body_box', 'columns', 'header_refs', 'footer_refs', 'paragraph_slices', 'lines', 'note_stories', 'table_rows'],
  TableRowFragmentV1: ['id','table_id','row_id','row_ordinal','fragment_ordinal','section_id','column_id','column_ordinal','x_millipoints','y_millipoints','width_millipoints','height_millipoints','source_y_millipoints','source_height_millipoints'],
  SectionV1: ['section_id', 'starts_at_block_id', 'break_type', 'column_ids', 'page_ids'],
  PaginatedLayoutV1: ['protocol', 'version', 'status', 'provenance', 'diagnostics', 'sections', 'pages'],
} as const

type JsonObject = Record<string, unknown>
const ID = /^[A-Za-z0-9][A-Za-z0-9._:/-]{0,1023}$/
const SHORT_ID = /^[A-Za-z0-9][A-Za-z0-9._:/-]{0,255}$/
const SHA256 = /^sha256:[0-9a-f]{64}$/
const PART_SEGMENT = /^(?:[A-Za-z0-9._~!$&'()*+,;=@-]|%[0-9A-F]{2})+$/
const ABSOLUTE_URI = /^[A-Za-z][A-Za-z0-9+.-]*:[^\u0000\r\n]{1,4090}$/
const MAX_MESSAGE = 4_096

function add(issues: NativeDocxValidationIssue[], code: NativeDocxIssueCode, path: string, message: string): void {
  if (issues.length >= DOCX_NATIVE_LIMITS.maxIssues) return
  if (!issues.some((entry) => entry.code === code && entry.path === path && entry.message === message)) issues.push({ code, path, message })
}

function isObject(value: unknown): value is JsonObject {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

function pointer(value: string): string {
  return value.replace(/~/g, '~0').replace(/\//g, '~1')
}

function canonicalPartKey(value: string): string {
  return asciiLowerNative(value.split('/').map((segment) => decodeURIComponent(segment)).join('/'))
}

function object(value: unknown, path: string, fields: readonly string[], issues: NativeDocxValidationIssue[]): JsonObject | undefined {
  if (!isObject(value)) {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, 'must be an object')
    return undefined
  }
  const allowed = new Set<string>(fields)
  for (const key of Object.keys(value).sort()) if (!allowed.has(key)) add(issues, 'UNKNOWN_FIELD', `${path}/${pointer(key)}`, `unknown field ${JSON.stringify(key)}`)
  return value
}

function validateNumberingSource(value: unknown, path: string, issues: NativeDocxValidationIssue[]): void {
  const source = object(value, path, DOCX_PAGINATED_LAYOUT_V1_BINDING_FIELDS.NumberingSourceV1, issues)
  if (!source) return
  partName(source.relationships_part, `${path}/relationships_part`, issues)
  stringValue(source.relationships_sha256, `${path}/relationships_sha256`, issues, SHA256, 71)
  stringValue(source.relationship_id, `${path}/relationship_id`, issues)
  const relationshipType = stringValue(source.relationship_type, `${path}/relationship_type`, issues, ABSOLUTE_URI, 4096)
  if (relationshipType && relationshipType !== 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering' && relationshipType !== 'http://purl.oclc.org/ooxml/officeDocument/relationships/numbering') add(issues, 'INVALID_VALUE', `${path}/relationship_type`, 'must be the Strict or Transitional numbering relationship type')
  stringValue(source.relationship_target, `${path}/relationship_target`, issues, undefined, 4096)
  partName(source.part_name, `${path}/part_name`, issues)
  if (source.content_type !== 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml') add(issues, 'INVALID_VALUE', `${path}/content_type`, 'must be the WordprocessingML numbering content type')
  stringValue(source.part_sha256, `${path}/part_sha256`, issues, SHA256, 71)
  stringValue(source.model_sha256, `${path}/model_sha256`, issues, SHA256, 71)
}

function array(value: unknown, path: string, max: number, issues: NativeDocxValidationIssue[]): unknown[] {
  if (!Array.isArray(value)) {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, 'must be an array')
    return []
  }
  if (value.length > max) add(issues, 'LIMIT_EXCEEDED', path, `must contain at most ${max} items`)
  return value.slice(0, max)
}

function stringValue(value: unknown, path: string, issues: NativeDocxValidationIssue[], pattern = SHORT_ID, max = 256): string | undefined {
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

function optionalMessage(value: unknown, path: string, issues: NativeDocxValidationIssue[]): void {
  if (value === undefined) return
  stringValue(value, path, issues, /[\s\S]+/, MAX_MESSAGE)
}

function partName(value: unknown, path: string, issues: NativeDocxValidationIssue[]): string | undefined {
  const part = stringValue(value, path, issues, /[^\u0000\r\n]+/, 4_096)
  if (!part) return undefined
  if (part.startsWith('/') || part.includes('\\') || part.includes('//') || !part.split('/').every((segment) => {
    if (segment === '.' || segment === '..' || !PART_SEGMENT.test(segment)) return false
    try {
      const decoded = decodeURIComponent(segment)
      return decoded !== '.' && decoded !== '..' && !decoded.endsWith('.') && !/[\\/?#%]/.test(decoded) && ![...decoded].some((character) => {
        const code = character.codePointAt(0) ?? 0
        return code < 0x20 || code === 0x7f
      })
    } catch { return false }
  })) {
    add(issues, 'INVALID_VALUE', path, 'must be a canonical OPC part name')
    return undefined
  }
  return part
}

function integer(value: unknown, path: string, min: number, max: number, issues: NativeDocxValidationIssue[]): number | undefined {
  if (!Number.isSafeInteger(value) || Object.is(value, -0) || (value as number) < min || (value as number) > max) {
    add(issues, value === undefined ? 'REQUIRED' : 'OUT_OF_RANGE', path, `must be a safe integer from ${min} through ${max}`)
    return undefined
  }
  return value as number
}

function booleanValue(value: unknown, path: string, issues: NativeDocxValidationIssue[]): boolean | undefined {
  if (typeof value !== 'boolean') {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, 'must be a boolean')
    return undefined
  }
  return value
}

function enumValue(value: unknown, path: string, allowed: readonly string[], issues: NativeDocxValidationIssue[]): string | undefined {
  if (typeof value !== 'string' || !allowed.includes(value)) {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_VALUE', path, `must be one of ${allowed.join(', ')}`)
    return undefined
  }
  return value
}

function preflight(value: unknown, path: string, depth: number, active: WeakSet<object>, state: { nodes: number; bounded: boolean }, issues: NativeDocxValidationIssue[]): void {
  if (!state.bounded) return
  if (issues.length >= DOCX_NATIVE_LIMITS.maxIssues) {
    state.bounded = false
    return
  }
  state.nodes += 1
  if (state.nodes > DOCX_PAGINATION_LIMITS.maxOutputNodes) {
    add(issues, 'LIMIT_EXCEEDED', path, `paginated layout exceeds ${DOCX_PAGINATION_LIMITS.maxOutputNodes} traversed values`)
    state.bounded = false
    return
  }
  if (depth > DOCX_NATIVE_LIMITS.maxDepth) {
    add(issues, 'LIMIT_EXCEEDED', path, `paginated layout exceeds ${DOCX_NATIVE_LIMITS.maxDepth} nesting levels`)
    state.bounded = false
    return
  }
  if (value === null || value === undefined) {
    add(issues, 'INVALID_VALUE', path, 'JSON null and undefined are not permitted')
    return
  }
  if (typeof value === 'number' && (!Number.isFinite(value) || Object.is(value, -0))) add(issues, 'INVALID_VALUE', path, 'non-finite numbers and negative zero are not permitted')
  if (typeof value !== 'object') {
    if (!['string', 'number', 'boolean'].includes(typeof value)) add(issues, 'INVALID_TYPE', path, 'must contain only JSON wire values')
    return
  }
  if (active.has(value)) {
    add(issues, 'INVALID_VALUE', path, 'cyclic values are not valid JSON wire output')
    return
  }
  active.add(value)
  if (Array.isArray(value)) {
    const max = Math.max(DOCX_PAGINATION_LIMITS.maxLinePlacements, DOCX_PAGINATION_LIMITS.maxParagraphSlices)
    if (value.length > max) {
      add(issues, 'LIMIT_EXCEEDED', path, `array exceeds ${max} items`)
      state.bounded = false
    }
    for (let index = 0; state.bounded && index < Math.min(value.length, max); index += 1) preflight(value[index], `${path}/${index}`, depth + 1, active, state, issues)
  } else {
    for (const key of Object.keys(value).sort()) {
      if (!state.bounded) break
      preflight((value as JsonObject)[key], `${path}/${pointer(key)}`, depth + 1, active, state, issues)
    }
  }
  active.delete(value)
}

function validateReference(value: unknown, path: string, issues: NativeDocxValidationIssue[]): void {
  const entry = object(value, path, DOCX_PAGINATED_LAYOUT_V1_BINDING_FIELDS.HeaderFooterReferenceV1, issues)
  if (!entry) return
  enumValue(entry.kind, `${path}/kind`, ['default', 'first', 'even'], issues)
  stringValue(entry.story_id, `${path}/story_id`, issues)
  stringValue(entry.relationship_id, `${path}/relationship_id`, issues)
}

interface PageValidation {
  id?: string
  sectionID?: string
  sectionIDs: string[]
  sectionOrdinal?: number
  parityBeforeSectionID?: string
  parityReason?: string
  kind?: string
  columns: Array<{ id: string; sectionID?: string }>
}

interface SliceValidation {
  repeated?: boolean
  paragraphID: string
  sliceOrdinal?: number
  first?: number
  last?: number
  continued?: boolean
  continues?: boolean
  continuedColumn?: boolean
  continuesColumn?: boolean
  columnOrdinal?: number
  pageOrdinal: number
  path: string
}

function validatePage(
  value: unknown,
  path: string,
  expectedOrdinal: number,
  issues: NativeDocxValidationIssue[],
  pageIDs: Set<string>,
  lineIDs: Set<string>,
  noteSourceLineIDs: Set<string>,
  sliceIDs: Set<string>,
  referencedLineIDs: Set<string>,
  paragraphSlices: Map<string, SliceValidation[]>,
): PageValidation {
  const entry = object(value, path, DOCX_PAGINATED_LAYOUT_V1_BINDING_FIELDS.PageV1, issues)
  if (!entry) return { sectionIDs: [], columns: [] }
  const id = stringValue(entry.id, `${path}/id`, issues, ID, 1024)
  if (id && pageIDs.has(id)) add(issues, 'DUPLICATE_ID', `${path}/id`, 'page id is duplicated')
  if (id) pageIDs.add(id)
  const ordinal = integer(entry.ordinal, `${path}/ordinal`, 0, DOCX_PAGINATION_LIMITS.maxPages - 1, issues)
  if (ordinal !== undefined && ordinal !== expectedOrdinal) add(issues, 'INVALID_VALUE', `${path}/ordinal`, `must equal physical source-order ordinal ${expectedOrdinal}`)
  const sectionID = stringValue(entry.section_id, `${path}/section_id`, issues)
  const sectionIDs = array(entry.section_ids, `${path}/section_ids`, DOCX_NATIVE_LIMITS.maxCollectionItems, issues).map((value, index) => stringValue(value, `${path}/section_ids/${index}`, issues)).filter((value): value is string => value !== undefined)
  if (sectionID && sectionIDs[0] !== sectionID) add(issues, 'INVALID_VALUE', `${path}/section_ids`, 'first page section identity must equal the owning section_id')
  if (new Set(sectionIDs).size !== sectionIDs.length) add(issues, 'DUPLICATE_ID', `${path}/section_ids`, 'page section identities must be unique')
  const sectionOrdinal = integer(entry.section_page_ordinal, `${path}/section_page_ordinal`, 0, DOCX_PAGINATION_LIMITS.maxPages - 1, issues)
  const kind = enumValue(entry.kind, `${path}/kind`, ['content', 'parity-blank'], issues)
  const reason = entry.parity_reason === undefined ? undefined : enumValue(entry.parity_reason, `${path}/parity_reason`, ['odd-page-section', 'even-page-section'], issues)
  const parityBeforeSectionID = entry.parity_before_section_id === undefined ? undefined : stringValue(entry.parity_before_section_id, `${path}/parity_before_section_id`, issues)
  if ((kind === 'parity-blank') !== (reason !== undefined && parityBeforeSectionID !== undefined)) add(issues, 'INVALID_UNION', `${path}/parity_reason`, 'parity reason and target section must appear exactly on parity-blank pages')
  const expectedID = kind === 'parity-blank' ? `page:parity-before:${parityBeforeSectionID}` : `page:${sectionID}:${sectionOrdinal}`
  if (id && sectionID && sectionOrdinal !== undefined && id !== expectedID) add(issues, 'INVALID_VALUE', `${path}/id`, 'page id must equal its deterministic section/kind/ordinal identity')
  const width = integer(entry.width_millipoints, `${path}/width_millipoints`, 1, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
  const height = integer(entry.height_millipoints, `${path}/height_millipoints`, 1, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
  const box = object(entry.body_box, `${path}/body_box`, DOCX_PAGINATED_LAYOUT_V1_BINDING_FIELDS.BodyBoxV1, issues)
  let boxX: number | undefined
  let boxY: number | undefined
  let boxWidth: number | undefined
  let boxHeight: number | undefined
  if (box) {
    boxX = integer(box.x_millipoints, `${path}/body_box/x_millipoints`, 0, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
    boxY = integer(box.y_millipoints, `${path}/body_box/y_millipoints`, 0, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
    boxWidth = integer(box.width_millipoints, `${path}/body_box/width_millipoints`, 1, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
    boxHeight = integer(box.height_millipoints, `${path}/body_box/height_millipoints`, 1, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
    if (width !== undefined && boxX !== undefined && boxWidth !== undefined && boxX + boxWidth > width) add(issues, 'OUT_OF_RANGE', `${path}/body_box`, 'body box exceeds page width')
    if (height !== undefined && boxY !== undefined && boxHeight !== undefined && boxY + boxHeight > height) add(issues, 'OUT_OF_RANGE', `${path}/body_box`, 'body box exceeds page height')
  }
  const pageColumns = new Map<string, { sectionID?: string; ordinal?: number; x?: number; y?: number; width?: number; height?: number }>()
  array(entry.columns, `${path}/columns`, DOCX_NATIVE_LIMITS.maxCollectionItems, issues).forEach((columnValue, index) => {
    const columnPath = `${path}/columns/${index}`
    const column = object(columnValue, columnPath, DOCX_PAGINATED_LAYOUT_V1_BINDING_FIELDS.ColumnV1, issues)
    if (!column) return
    const columnID = stringValue(column.id, `${columnPath}/id`, issues, ID, 1024)
    const columnSectionID = stringValue(column.section_id, `${columnPath}/section_id`, issues)
    const columnOrdinal = integer(column.ordinal, `${columnPath}/ordinal`, 0, 44, issues)
    const x = integer(column.x_millipoints, `${columnPath}/x_millipoints`, 0, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
    const y = integer(column.y_millipoints, `${columnPath}/y_millipoints`, 0, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
    const columnWidth = integer(column.width_millipoints, `${columnPath}/width_millipoints`, 1, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
    const columnHeight = integer(column.height_millipoints, `${columnPath}/height_millipoints`, 1, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
    if (columnID && pageColumns.has(columnID)) add(issues, 'DUPLICATE_ID', `${columnPath}/id`, 'page column identity is duplicated')
    if (columnID) pageColumns.set(columnID, { sectionID: columnSectionID, ordinal: columnOrdinal, x, y, width: columnWidth, height: columnHeight })
    if (columnSectionID && !sectionIDs.includes(columnSectionID)) add(issues, 'BROKEN_REFERENCE', `${columnPath}/section_id`, 'column section must occur on this page')
    if (width !== undefined && x !== undefined && columnWidth !== undefined && x + columnWidth > width) add(issues, 'OUT_OF_RANGE', columnPath, 'column exceeds page width')
    if (height !== undefined && y !== undefined && columnHeight !== undefined && y + columnHeight > height) add(issues, 'OUT_OF_RANGE', columnPath, 'column exceeds page height')
  })
  for (const key of ['header_refs', 'footer_refs'] as const) array(entry[key], `${path}/${key}`, DOCX_NATIVE_LIMITS.maxCollectionItems, issues).forEach((reference, index) => validateReference(reference, `${path}/${key}/${index}`, issues))
  const pageLines = new Map<string, { cellID?: string; repeated?: boolean; paragraphID?: string; sectionID?: string; columnID?: string; columnOrdinal?: number; ordinal?: number; y?: number; height?: number }>()
  const pageLineOrder: string[] = []
  const previousLineBottoms = new Map<string, number>()
  array(entry.lines, `${path}/lines`, DOCX_PAGINATION_LIMITS.maxLinePlacements, issues).forEach((lineValue, index) => {
    const linePath = `${path}/lines/${index}`
    const line = object(lineValue, linePath, DOCX_PAGINATED_LAYOUT_V1_BINDING_FIELDS.PlacedLineV1, issues)
    if (!line) return
    const placedID = stringValue(line.id, `${linePath}/id`, issues, ID, 1024)
    const lineID = stringValue(line.line_id, `${linePath}/line_id`, issues, ID, 1024)
    const paragraphID = stringValue(line.paragraph_id, `${linePath}/paragraph_id`, issues)
    const lineSectionID = stringValue(line.section_id, `${linePath}/section_id`, issues)
    const columnID = stringValue(line.column_id, `${linePath}/column_id`, issues, ID, 1024)
    const columnOrdinal = integer(line.column_ordinal, `${linePath}/column_ordinal`, 0, 44, issues)
    const sourceOrdinal = integer(line.source_line_ordinal, `${linePath}/source_line_ordinal`, 0, DOCX_PAGINATION_LIMITS.maxLinePlacements, issues)
    const x = integer(line.x_millipoints, `${linePath}/x_millipoints`, 0, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
    const y = integer(line.y_millipoints, `${linePath}/y_millipoints`, 0, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
    const lineWidth = integer(line.width_millipoints, `${linePath}/width_millipoints`, 0, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
    const lineHeight = integer(line.height_millipoints, `${linePath}/height_millipoints`, 1, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
    const repeated = line.repeated_table_header === true
    const cellID = line.table_cell_id === undefined ? undefined : stringValue(line.table_cell_id, `${linePath}/table_cell_id`, issues, ID, 1024)
    if (repeated && !cellID) add(issues, 'REQUIRED', `${linePath}/table_cell_id`, 'repeated table header requires a cell identity')
    if (line.repeated_table_header !== undefined && !repeated) add(issues, 'INVALID_VALUE', `${linePath}/repeated_table_header`, 'when present must equal true')
    if (placedID && lineID && placedID !== (repeated ? `placed:${lineID}:table-header:${id}` : `placed:${lineID}`)) add(issues, 'INVALID_VALUE', `${linePath}/id`, 'placed line id must derive from the shaped line and repeated-header page identity')
    if (placedID && lineIDs.has(placedID)) add(issues, 'DUPLICATE_ID', `${linePath}/id`, 'placed line id is duplicated')
    if (placedID) lineIDs.add(placedID)
    if (lineIDs.size > DOCX_PAGINATION_LIMITS.maxLinePlacements) add(issues, 'LIMIT_EXCEEDED', `${linePath}/id`, `placed lines exceed ${DOCX_PAGINATION_LIMITS.maxLinePlacements}`)
    if (lineID && pageLines.has(lineID)) add(issues, 'DUPLICATE_ID', `${linePath}/line_id`, 'shaped line is placed more than once on this page')
    if (lineID) {
      pageLines.set(lineID, { cellID, repeated, paragraphID, sectionID: lineSectionID, columnID, columnOrdinal, ordinal: sourceOrdinal, y, height: lineHeight })
      pageLineOrder.push(lineID)
    }
    if (y !== undefined && lineHeight !== undefined) {
      const flowID = columnID ? JSON.stringify([columnID, cellID ?? null]) : undefined
      const previousLineBottom = flowID ? previousLineBottoms.get(flowID) : undefined
      if (previousLineBottom !== undefined && y < previousLineBottom) add(issues, 'INVALID_VALUE', `${linePath}/y_millipoints`, 'column lines must be vertically source-ordered without overlap')
      if (flowID) previousLineBottoms.set(flowID, y + lineHeight)
    }
    const owningColumn = columnID ? pageColumns.get(columnID) : undefined
    if (!owningColumn || owningColumn.sectionID !== lineSectionID || owningColumn.ordinal !== columnOrdinal) add(issues, 'BROKEN_REFERENCE', `${linePath}/column_id`, 'placed line must reference its exact page section column')
    if (owningColumn?.x !== undefined && owningColumn.width !== undefined && x !== undefined && lineWidth !== undefined && (x < owningColumn.x || x + lineWidth > owningColumn.x + owningColumn.width)) add(issues, 'OUT_OF_RANGE', linePath, 'line exceeds its page column width')
    if (owningColumn?.y !== undefined && owningColumn.height !== undefined && y !== undefined && lineHeight !== undefined && (y < owningColumn.y || y + lineHeight > owningColumn.y + owningColumn.height)) add(issues, 'OUT_OF_RANGE', linePath, 'line exceeds its page column height')
  })
  const slices = array(entry.paragraph_slices, `${path}/paragraph_slices`, DOCX_PAGINATION_LIMITS.maxParagraphSlices, issues)
  const pageSliceLineOrder: string[] = []
  const previousSliceBottoms = new Map<string, number>()
  slices.forEach((sliceValue, index) => {
    const slicePath = `${path}/paragraph_slices/${index}`
    const slice = object(sliceValue, slicePath, DOCX_PAGINATED_LAYOUT_V1_BINDING_FIELDS.ParagraphSliceV1, issues)
    if (!slice) return
    const paragraphID = stringValue(slice.paragraph_id, `${slicePath}/paragraph_id`, issues)
    const repeated = slice.repeated_table_header === true
    const cellID = slice.table_cell_id === undefined ? undefined : stringValue(slice.table_cell_id, `${slicePath}/table_cell_id`, issues, ID, 1024)
    if (slice.repeated_table_header !== undefined && !repeated) add(issues, 'INVALID_VALUE', `${slicePath}/repeated_table_header`, 'when present must equal true')
    const sliceSectionID = stringValue(slice.section_id, `${slicePath}/section_id`, issues)
    const sliceColumnID = stringValue(slice.column_id, `${slicePath}/column_id`, issues, ID, 1024)
    const sliceColumnOrdinal = integer(slice.column_ordinal, `${slicePath}/column_ordinal`, 0, 44, issues)
    const sliceOrdinal = integer(slice.slice_ordinal, `${slicePath}/slice_ordinal`, 0, DOCX_PAGINATION_LIMITS.maxParagraphSlices, issues)
    const sliceID = stringValue(slice.id, `${slicePath}/id`, issues, ID, 1024)
    if (sliceID && paragraphID && sliceOrdinal !== undefined && sliceID !== `slice:${paragraphID}:${sliceOrdinal}`) add(issues, 'INVALID_VALUE', `${slicePath}/id`, 'slice id must derive from paragraph id and slice ordinal')
    if (sliceID && sliceIDs.has(sliceID)) add(issues, 'DUPLICATE_ID', `${slicePath}/id`, 'paragraph slice id is duplicated')
    if (sliceID) sliceIDs.add(sliceID)
    if (sliceIDs.size > DOCX_PAGINATION_LIMITS.maxParagraphSlices) add(issues, 'LIMIT_EXCEEDED', `${slicePath}/id`, `paragraph slices exceed ${DOCX_PAGINATION_LIMITS.maxParagraphSlices}`)
    const first = integer(slice.first_line_ordinal, `${slicePath}/first_line_ordinal`, 0, DOCX_PAGINATION_LIMITS.maxLinePlacements, issues)
    const last = integer(slice.last_line_ordinal, `${slicePath}/last_line_ordinal`, 0, DOCX_PAGINATION_LIMITS.maxLinePlacements, issues)
    const top = integer(slice.top_millipoints, `${slicePath}/top_millipoints`, 0, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
    const sliceHeight = integer(slice.height_millipoints, `${slicePath}/height_millipoints`, 1, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
    integer(slice.space_before_millipoints, `${slicePath}/space_before_millipoints`, 0, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
    const continued = booleanValue(slice.continued_from_previous_page, `${slicePath}/continued_from_previous_page`, issues)
    const continues = booleanValue(slice.continues_on_next_page, `${slicePath}/continues_on_next_page`, issues)
    const continuedColumn = booleanValue(slice.continued_from_previous_column, `${slicePath}/continued_from_previous_column`, issues)
    const continuesColumn = booleanValue(slice.continues_in_next_column, `${slicePath}/continues_in_next_column`, issues)
    if (continued === continuedColumn && continued === true) add(issues, 'INVALID_UNION', slicePath, 'a slice cannot continue from both a page and a column')
    if (continues === continuesColumn && continues === true) add(issues, 'INVALID_UNION', slicePath, 'a slice cannot continue to both a page and a column')
    const ids = array(slice.line_ids, `${slicePath}/line_ids`, DOCX_PAGINATION_LIMITS.maxLinePlacements, issues).map((lineID, lineIndex) => stringValue(lineID, `${slicePath}/line_ids/${lineIndex}`, issues, ID, 1024))
    if (first !== undefined && last !== undefined && (last < first || ids.length !== last - first + 1)) add(issues, 'INVALID_VALUE', `${slicePath}/line_ids`, 'line ids must exactly cover the inclusive source ordinal range')
    let expectedY = top
    let summedHeight = 0
    ids.forEach((lineID, lineIndex) => {
      if (!lineID) return
      pageSliceLineOrder.push(lineID)
      if (repeated ? !referencedLineIDs.has(lineID) : referencedLineIDs.has(lineID)) add(issues, 'DUPLICATE_ID', `${slicePath}/line_ids/${lineIndex}`, 'a repeated header must follow one original shaped-line placement; other lines cannot repeat')
      if (!repeated) referencedLineIDs.add(lineID)
      const line = pageLines.get(lineID)
      if (line && line.repeated !== repeated) add(issues, 'BROKEN_REFERENCE', slicePath, 'slice and placed line must agree on repeated-header identity')
      if (line && line.cellID !== cellID) add(issues, 'BROKEN_REFERENCE', slicePath, 'slice and placed line must agree on their table cell identity')
      if (!line) add(issues, 'BROKEN_REFERENCE', `${slicePath}/line_ids/${lineIndex}`, 'slice line id must reference a line on the same page')
      if (line && paragraphID && line.paragraphID !== paragraphID) add(issues, 'BROKEN_REFERENCE', `${slicePath}/line_ids/${lineIndex}`, 'slice line must belong to its paragraph')
      if (line && (line.sectionID !== sliceSectionID || line.columnID !== sliceColumnID || line.columnOrdinal !== sliceColumnOrdinal)) add(issues, 'BROKEN_REFERENCE', `${slicePath}/line_ids/${lineIndex}`, 'slice lines must share the exact section and column identity')
      if (line && first !== undefined && line.ordinal !== first + lineIndex) add(issues, 'INVALID_VALUE', `${slicePath}/line_ids/${lineIndex}`, 'slice lines must be in contiguous shaped-line order')
      if (line && expectedY !== undefined && line.y !== expectedY) add(issues, 'INVALID_VALUE', `${slicePath}/line_ids/${lineIndex}`, 'slice lines must be vertically contiguous from slice top')
      if (line?.height !== undefined) {
        summedHeight += line.height
        expectedY = expectedY === undefined ? undefined : expectedY + line.height
      }
    })
    if (sliceHeight !== undefined && sliceHeight !== summedHeight) add(issues, 'INVALID_VALUE', `${slicePath}/height_millipoints`, 'slice height must equal its placed-line heights')
    if (top !== undefined && sliceHeight !== undefined) {
      const flowID = sliceColumnID ? JSON.stringify([sliceColumnID, cellID ?? null]) : undefined
      const previousSliceBottom = flowID ? previousSliceBottoms.get(flowID) : undefined
      if (previousSliceBottom !== undefined && top < previousSliceBottom) add(issues, 'INVALID_VALUE', `${slicePath}/top_millipoints`, 'paragraph slices must be vertically source-ordered without overlap')
      if (flowID) previousSliceBottoms.set(flowID, top + sliceHeight)
    }
    if (paragraphID) {
      const records = paragraphSlices.get(paragraphID) ?? []
      records.push({ repeated, paragraphID, sliceOrdinal, first, last, continued, continues, continuedColumn, continuesColumn, columnOrdinal: sliceColumnOrdinal, pageOrdinal: expectedOrdinal, path: slicePath })
      paragraphSlices.set(paragraphID, records)
    }
  })
  if (pageLineOrder.length !== pageSliceLineOrder.length || pageLineOrder.some((lineID, index) => lineID !== pageSliceLineOrder[index])) add(issues, 'BROKEN_REFERENCE', `${path}/paragraph_slices`, 'page slices must cover every placed line exactly once and in page line/source order')
  const notes = entry.note_stories === undefined ? [] : array(entry.note_stories, `${path}/note_stories`, DOCX_NATIVE_LIMITS.maxCollectionItems, issues)
  let previousNoteBottom: number | undefined
  let noteGroupKind: 'footnote' | 'endnote' | undefined
  let noteGroupColumnID: string | undefined
  notes.forEach((noteValue, noteIndex) => {
    const notePath = `${path}/note_stories/${noteIndex}`
    const note = object(noteValue, notePath, DOCX_PAGINATED_LAYOUT_V1_BINDING_FIELDS.PlacedNoteStoryV1, issues)
    if (!note) return
    const storyID = stringValue(note.story_id, `${notePath}/story_id`, issues)
    const noteID = stringValue(note.id, `${notePath}/id`, issues, ID, 1024)
    if (noteID && storyID && noteID !== `note-placement:${expectedOrdinal}:${storyID}`) add(issues, 'INVALID_VALUE', `${notePath}/id`, 'note placement id must derive from page ordinal and story id')
    const storyKind = enumValue(note.story_kind, `${notePath}/story_kind`, ['footnote', 'endnote'], issues) as 'footnote' | 'endnote' | undefined
    if (storyKind !== undefined && noteGroupKind !== undefined && storyKind !== noteGroupKind) add(issues, 'INVALID_VALUE', `${notePath}/story_kind`, 'one page note group cannot mix footnotes and endnotes')
    if (storyKind !== undefined) noteGroupKind ??= storyKind
    const role = enumValue(note.note_role, `${notePath}/note_role`, ['content', 'separator'], issues)
    if (typeof note.native_story_id !== 'string' || (role === 'separator' ? note.native_story_id !== '-1' : !/^[1-9][0-9]{0,18}$/.test(note.native_story_id))) add(issues, 'INVALID_VALUE', `${notePath}/native_story_id`, 'native note id must match its exact placed role')
    if ((noteIndex === 0 && role !== 'separator') || (noteIndex > 0 && role !== 'content')) add(issues, 'INVALID_VALUE', `${notePath}/note_role`, 'each page note group must contain one leading separator followed by content stories')
    stringValue(note.relationship_id, `${notePath}/relationship_id`, issues)
    const noteOrdinal = integer(note.ordinal, `${notePath}/ordinal`, 0, DOCX_NATIVE_LIMITS.maxCollectionItems, issues)
    if (noteOrdinal !== undefined && noteOrdinal !== noteIndex) add(issues, 'INVALID_VALUE', `${notePath}/ordinal`, 'note placement ordinal must equal page note order')
    if (role === 'content') {
      stringValue(note.reference_run_id, `${notePath}/reference_run_id`, issues)
      integer(note.number, `${notePath}/number`, 1, DOCX_NATIVE_LIMITS.maxCollectionItems, issues)
    } else if (note.reference_run_id !== undefined || note.number !== undefined) add(issues, 'INVALID_UNION', notePath, 'separator placement cannot carry reference or number')
    const noteSectionID = stringValue(note.section_id, `${notePath}/section_id`, issues)
    const noteColumnID = stringValue(note.column_id, `${notePath}/column_id`, issues, ID, 1024)
    const noteColumnOrdinal = integer(note.column_ordinal, `${notePath}/column_ordinal`, 0, 44, issues)
    const noteColumn = noteColumnID ? pageColumns.get(noteColumnID) : undefined
    if (!noteColumn || noteColumn.sectionID !== noteSectionID || noteColumn.ordinal !== noteColumnOrdinal) add(issues, 'BROKEN_REFERENCE', `${notePath}/column_id`, 'note placement must reference its exact page section column')
    if (noteColumnID !== undefined && noteGroupColumnID !== undefined && noteColumnID !== noteGroupColumnID) add(issues, 'BROKEN_REFERENCE', `${notePath}/column_id`, 'one page note group cannot cross columns')
    noteGroupColumnID ??= noteColumnID
    const bodyLineBottom = [...pageLines.values()].reduce((bottom, line) => line.columnID === noteColumnID && line.y !== undefined && line.height !== undefined ? Math.max(bottom, line.y + line.height) : bottom, noteColumn?.y ?? 0)
    const noteTop = integer(note.top_millipoints, `${notePath}/top_millipoints`, 0, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
    const noteHeight = integer(note.height_millipoints, `${notePath}/height_millipoints`, 1, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
    if (noteTop !== undefined && noteHeight !== undefined) {
      if (previousNoteBottom !== undefined && noteTop !== previousNoteBottom) add(issues, 'INVALID_VALUE', `${notePath}/top_millipoints`, 'note stories must be vertically contiguous in reference order')
      if (noteIndex === 0 && noteTop < bodyLineBottom) add(issues, 'INVALID_VALUE', `${notePath}/top_millipoints`, 'note stories must not overlap placed body lines')
      previousNoteBottom = noteTop + noteHeight
      if (noteColumn?.y !== undefined && noteColumn.height !== undefined && (noteTop < noteColumn.y || previousNoteBottom > noteColumn.y + noteColumn.height)) add(issues, 'OUT_OF_RANGE', notePath, 'note story exceeds its exact page column')
    }
    const noteLines = array(note.lines, `${notePath}/lines`, DOCX_PAGINATION_LIMITS.maxLinePlacements, issues)
    let summedHeight = 0
    noteLines.forEach((lineValue, lineIndex) => {
      const linePath = `${notePath}/lines/${lineIndex}`
      const line = object(lineValue, linePath, DOCX_PAGINATED_LAYOUT_V1_BINDING_FIELDS.PlacedLineV1, issues)
      if (!line) return
      const placedID = stringValue(line.id, `${linePath}/id`, issues, ID, 1024)
      const lineID = stringValue(line.line_id, `${linePath}/line_id`, issues, ID, 1024)
      stringValue(line.paragraph_id, `${linePath}/paragraph_id`, issues)
      const lineSectionID = stringValue(line.section_id, `${linePath}/section_id`, issues)
      const lineColumnID = stringValue(line.column_id, `${linePath}/column_id`, issues, ID, 1024)
      const lineColumnOrdinal = integer(line.column_ordinal, `${linePath}/column_ordinal`, 0, 44, issues)
      if (lineSectionID !== noteSectionID || lineColumnID !== noteColumnID || lineColumnOrdinal !== noteColumnOrdinal) add(issues, 'BROKEN_REFERENCE', linePath, 'note lines must preserve their owning note section and column identity')
      integer(line.source_line_ordinal, `${linePath}/source_line_ordinal`, 0, DOCX_PAGINATION_LIMITS.maxLinePlacements, issues)
      const x = integer(line.x_millipoints, `${linePath}/x_millipoints`, 0, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
      const y = integer(line.y_millipoints, `${linePath}/y_millipoints`, 0, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
      const lineWidth = integer(line.width_millipoints, `${linePath}/width_millipoints`, 0, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
      const lineHeight = integer(line.height_millipoints, `${linePath}/height_millipoints`, 1, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
      if (placedID && lineID && placedID !== `placed-note:${expectedOrdinal}:${noteIndex}:${lineID}`) add(issues, 'INVALID_VALUE', `${linePath}/id`, 'note placed line id must derive from page, story ordinal, and shaped line id')
      if (placedID && lineIDs.has(placedID)) add(issues, 'DUPLICATE_ID', `${linePath}/id`, 'placed line id is duplicated')
      if (placedID) lineIDs.add(placedID)
      if (lineIDs.size > DOCX_PAGINATION_LIMITS.maxLinePlacements) add(issues, 'LIMIT_EXCEEDED', `${linePath}/id`, `placed lines exceed ${DOCX_PAGINATION_LIMITS.maxLinePlacements}`)
      if (role === 'content' && lineID && noteSourceLineIDs.has(lineID)) add(issues, 'DUPLICATE_ID', `${linePath}/line_id`, 'content note shaped line may be placed only once')
      if (role === 'content' && lineID) noteSourceLineIDs.add(lineID)
      if (noteColumn?.x !== undefined && noteColumn.width !== undefined && x !== undefined && lineWidth !== undefined && (x < noteColumn.x || x + lineWidth > noteColumn.x + noteColumn.width)) add(issues, 'OUT_OF_RANGE', linePath, 'note line exceeds its exact page column width')
      if (noteTop !== undefined && noteHeight !== undefined && y !== undefined && lineHeight !== undefined && (y < noteTop || y + lineHeight > noteTop + noteHeight)) add(issues, 'OUT_OF_RANGE', linePath, 'note line exceeds its story placement')
      if (lineHeight !== undefined) summedHeight += lineHeight
    })
    if (noteHeight !== undefined && summedHeight > noteHeight) add(issues, 'OUT_OF_RANGE', `${notePath}/height_millipoints`, 'note height cannot be smaller than its placed lines')
    if (storyKind === undefined) return
  })
  const noteGroupColumn = noteGroupColumnID ? pageColumns.get(noteGroupColumnID) : undefined
  if (noteGroupKind === 'footnote' && previousNoteBottom !== undefined && noteGroupColumn?.y !== undefined && noteGroupColumn.height !== undefined && previousNoteBottom !== noteGroupColumn.y + noteGroupColumn.height) add(issues, 'INVALID_VALUE', `${path}/note_stories`, 'footnote group must be bottom-aligned to its exact page column')
  if (kind === 'parity-blank' && (pageLines.size > 0 || slices.length > 0 || notes.length > 0 || (Array.isArray(entry.header_refs) && entry.header_refs.length > 0) || (Array.isArray(entry.footer_refs) && entry.footer_refs.length > 0))) add(issues, 'INVALID_UNION', path, 'parity-blank pages must contain no body, note, header, or footer placements/references')
  return {
    id, sectionID, sectionIDs, sectionOrdinal, parityBeforeSectionID, parityReason: reason, kind,
    columns: [...pageColumns.entries()].map(([columnID, column]) => ({ id: columnID, sectionID: column.sectionID })),
  }
}

export function decodeNativeDocxPaginatedLayout(value: unknown): DecodeNativeDocxPaginatedLayoutResult {
  const issues: NativeDocxValidationIssue[] = []
  const preflightState = { nodes: 0, bounded: true }
  preflight(value, '', 0, new WeakSet(), preflightState, issues)
  if (!preflightState.bounded || issues.some((entry) => entry.code === 'LIMIT_EXCEEDED')) return { ok: false, issues }
  const root = object(value, '', DOCX_PAGINATED_LAYOUT_V1_BINDING_FIELDS.PaginatedLayoutV1, issues)
  if (!root) return { ok: false, issues }
  if (root.protocol !== DOCX_PAGINATED_LAYOUT_PROTOCOL) add(issues, 'UNSUPPORTED_PROTOCOL', '/protocol', `must equal ${DOCX_PAGINATED_LAYOUT_PROTOCOL}`)
  if (root.version !== DOCX_PAGINATED_LAYOUT_VERSION) add(issues, 'UNSUPPORTED_VERSION', '/version', `must equal ${DOCX_PAGINATED_LAYOUT_VERSION}`)
  const status = enumValue(root.status, '/status', ['paginated', 'refused'], issues)
  let settingsUnsupported = false
  let settingsTabTwips: number | undefined
  let shapedTabInterval: number | undefined
  const provenance = object(root.provenance, '/provenance', DOCX_PAGINATED_LAYOUT_V1_BINDING_FIELDS.ProvenanceV1, issues)
  if (provenance) {
    const documentID = stringValue(provenance.document_id, '/provenance/document_id', issues)
    const revision = stringValue(provenance.revision, '/provenance/revision', issues)
    const packageSHA = stringValue(provenance.package_sha256, '/provenance/package_sha256', issues, SHA256, 71)
    const mainPart = partName(provenance.main_part, '/provenance/main_part', issues)
    stringValue(provenance.body_story_id, '/provenance/body_story_id', issues)
    if (provenance.numbering_source !== undefined) validateNumberingSource(provenance.numbering_source, '/provenance/numbering_source', issues)
    const settings = decodeNativeDocxPaginationSettings(provenance.pagination_settings)
    if (!settings.ok) issues.push(...settings.issues.map((entry) => ({ ...entry, path: `/provenance/pagination_settings${entry.path}` })))
    else {
      settingsUnsupported = settings.value.profile !== 'word-modern-default'
      settingsTabTwips = settings.value.default_tab_stop_twips
      if (settings.value.document_id !== documentID || settings.value.revision !== revision || settings.value.package_sha256 !== packageSHA || !mainPart || canonicalPartKey(settings.value.main_part) !== canonicalPartKey(mainPart)) add(issues, 'BROKEN_REFERENCE', '/provenance/pagination_settings', 'pagination settings identity, package fingerprint, and main part must match output provenance')
    }
    const shaped = object(provenance.shaped_lines, '/provenance/shaped_lines', DOCX_PAGINATED_LAYOUT_V1_BINDING_FIELDS.ShapedSourceV1, issues)
    if (shaped) {
      if (shaped.protocol !== DOCX_SHAPED_LINES_PROTOCOL) add(issues, 'UNSUPPORTED_PROTOCOL', '/provenance/shaped_lines/protocol', `must equal ${DOCX_SHAPED_LINES_PROTOCOL}`)
      if (shaped.version !== DOCX_SHAPED_LINES_VERSION) add(issues, 'UNSUPPORTED_VERSION', '/provenance/shaped_lines/version', `must equal ${DOCX_SHAPED_LINES_VERSION}`)
      integer(shaped.available_width_millipoints, '/provenance/shaped_lines/available_width_millipoints', 1, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
      shapedTabInterval = integer(shaped.tab_interval_millipoints, '/provenance/shaped_lines/tab_interval_millipoints', 1, DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints, issues)
    }
    const manifest = object(provenance.font_manifest, '/provenance/font_manifest', DOCX_PAGINATED_LAYOUT_V1_BINDING_FIELDS.FontManifestV1, issues)
    if (manifest) for (const key of DOCX_PAGINATED_LAYOUT_V1_BINDING_FIELDS.FontManifestV1) stringValue(manifest[key], `/provenance/font_manifest/${key}`, issues)
    const providers = object(provenance.providers, '/provenance/providers', DOCX_PAGINATED_LAYOUT_V1_BINDING_FIELDS.ProvidersV1, issues)
    if (providers) for (const key of DOCX_PAGINATED_LAYOUT_V1_BINDING_FIELDS.ProvidersV1) stringValue(providers[key], `/provenance/providers/${key}`, issues)
    const settingsTabMilliPoints = settingsTabTwips !== undefined
      && Number.isSafeInteger(settingsTabTwips)
      && !Object.is(settingsTabTwips, -0)
      && settingsTabTwips >= 0
      && settingsTabTwips <= DOCX_MAX_TWIPS_FOR_MILLIPOINTS
      ? settingsTabTwips * 50
      : undefined
    if (settingsTabTwips !== undefined && shapedTabInterval !== undefined && settingsTabMilliPoints !== shapedTabInterval) add(issues, 'BROKEN_REFERENCE', '/provenance/shaped_lines/tab_interval_millipoints', 'shaped tab interval must equal the attested Word default tab stop')
  }
  const diagnosticList = array(root.diagnostics, '/diagnostics', DOCX_PAGINATION_LIMITS.maxDiagnostics, issues)
  let hasUnsupported = false
  diagnosticList.forEach((diagnosticValue, index) => {
    const path = `/diagnostics/${index}`
    const diagnostic = object(diagnosticValue, path, DOCX_PAGINATED_LAYOUT_V1_BINDING_FIELDS.DiagnosticV1, issues)
    if (!diagnostic) return
    enumValue(diagnostic.code, `${path}/code`, DOCX_PAGINATION_DIAGNOSTIC_CODES, issues)
    const severity = enumValue(diagnostic.severity, `${path}/severity`, ['unsupported', 'deferred'], issues)
    if (severity === 'unsupported') hasUnsupported = true
    stringValue(diagnostic.scope_id, `${path}/scope_id`, issues)
    if (diagnostic.source_code !== undefined) stringValue(diagnostic.source_code, `${path}/source_code`, issues)
    optionalMessage(diagnostic.source_message, `${path}/source_message`, issues)
    stringValue(diagnostic.message, `${path}/message`, issues, /[\s\S]+/, MAX_MESSAGE)
  })
  const sectionList = array(root.sections, '/sections', DOCX_NATIVE_LIMITS.maxCollectionItems, issues)
  const pageList = array(root.pages, '/pages', DOCX_PAGINATION_LIMITS.maxPages, issues)
  if (status === 'refused') {
    if (sectionList.length !== 0 || pageList.length !== 0) add(issues, 'INVALID_UNION', '', 'refused layout must expose empty sections and pages')
    if (!hasUnsupported) add(issues, 'REQUIRED', '/diagnostics', 'refused layout requires at least one unsupported diagnostic')
  }
  if (status === 'paginated' && hasUnsupported) add(issues, 'INVALID_UNION', '/diagnostics', 'paginated layout cannot carry unsupported diagnostics')
  if (status === 'paginated' && settingsUnsupported) add(issues, 'INVALID_UNION', '/provenance/pagination_settings/profile', 'paginated output cannot carry an unsupported settings attestation')
  const pageIDs = new Set<string>()
  const placedLineIDs = new Set<string>()
  const noteSourceLineIDs = new Set<string>()
  const sliceIDs = new Set<string>()
  const referencedLineIDs = new Set<string>()
  const paragraphSlices = new Map<string, SliceValidation[]>()
  const pageResults = pageList.map((page, index) => validatePage(page, `/pages/${index}`, index, issues, pageIDs, placedLineIDs, noteSourceLineIDs, sliceIDs, referencedLineIDs, paragraphSlices))
  validateTableRowFragments(pageList, issues)
  const sectionIDs = new Set<string>()
  const sectionOrder: string[] = []
  const sectionBreaks = new Map<string, string>()
  const sectionColumns = new Map<string, string[]>()
  sectionList.forEach((sectionValue, index) => {
    const path = `/sections/${index}`
    const section = object(sectionValue, path, DOCX_PAGINATED_LAYOUT_V1_BINDING_FIELDS.SectionV1, issues)
    if (!section) return
    const sectionID = stringValue(section.section_id, `${path}/section_id`, issues)
    if (sectionID && sectionIDs.has(sectionID)) add(issues, 'DUPLICATE_ID', `${path}/section_id`, 'paginated section id is duplicated')
    if (sectionID) sectionIDs.add(sectionID)
    if (sectionID) sectionOrder.push(sectionID)
    stringValue(section.starts_at_block_id, `${path}/starts_at_block_id`, issues)
    const breakType = enumValue(section.break_type, `${path}/break_type`, ['continuous', 'next-page', 'even-page', 'odd-page', 'next-column'], issues)
    if (sectionID && breakType) sectionBreaks.set(sectionID, breakType)
    const columnIDs = array(section.column_ids, `${path}/column_ids`, 45, issues).map((columnID, columnIndex) => stringValue(columnID, `${path}/column_ids/${columnIndex}`, issues, ID, 1024)).filter((value): value is string => value !== undefined)
    if (new Set(columnIDs).size !== columnIDs.length) add(issues, 'DUPLICATE_ID', `${path}/column_ids`, 'section column identities must be unique')
    if (sectionID) sectionColumns.set(sectionID, columnIDs)
    const pageRefs = array(section.page_ids, `${path}/page_ids`, DOCX_PAGINATION_LIMITS.maxPages, issues).map((pageID, pageIndex) => stringValue(pageID, `${path}/page_ids/${pageIndex}`, issues, ID, 1024))
    const actual = pageResults.filter((page) => page.sectionIDs.includes(sectionID ?? '')).map((page) => page.id)
    if (pageRefs.length !== actual.length || pageRefs.some((pageID, pageIndex) => pageID !== actual[pageIndex])) add(issues, 'BROKEN_REFERENCE', `${path}/page_ids`, 'section page ids must exactly match source-ordered physical pages containing the section')
  })
  pageResults.forEach((page, index) => {
    if (page.sectionID && !sectionIDs.has(page.sectionID)) add(issues, 'BROKEN_REFERENCE', `/pages/${index}/section_id`, 'page must reference a paginated section')
    for (const sectionID of page.sectionIDs) if (!sectionIDs.has(sectionID)) add(issues, 'BROKEN_REFERENCE', `/pages/${index}/section_ids`, 'page section identity must reference a paginated section')
    for (const sectionID of page.sectionIDs) {
      const expected = sectionColumns.get(sectionID) ?? []
      const actual = page.columns.filter((column) => column.sectionID === sectionID).map((column) => column.id)
      if (actual.length !== expected.length || actual.some((columnID, columnIndex) => columnID !== expected[columnIndex])) add(issues, 'BROKEN_REFERENCE', `/pages/${index}/columns`, 'page columns must exactly match each occurring section column identity and order')
    }
    if (page.parityBeforeSectionID) {
      const target = page.parityBeforeSectionID
      if (!sectionIDs.has(target) || sectionOrder.indexOf(target) !== sectionOrder.indexOf(page.sectionID ?? '') + 1) add(issues, 'BROKEN_REFERENCE', `/pages/${index}/parity_before_section_id`, 'parity filler must target the immediately following section')
      const requiredBreak = page.parityReason === 'odd-page-section' ? 'odd-page' : page.parityReason === 'even-page-section' ? 'even-page' : undefined
      if (requiredBreak && sectionBreaks.get(target) !== requiredBreak) add(issues, 'BROKEN_REFERENCE', `/pages/${index}/parity_reason`, 'parity reason must match the immediately following section break type')
      const following = pageResults[index + 1]
      if (!following || following.kind !== 'content' || following.sectionID !== target) add(issues, 'BROKEN_REFERENCE', `/pages/${index}/parity_before_section_id`, 'parity filler must immediately precede the target section content page')
      const targetOneBasedPage = index + 2
      if ((requiredBreak === 'odd-page' && targetOneBasedPage % 2 !== 1) || (requiredBreak === 'even-page' && targetOneBasedPage % 2 !== 0)) add(issues, 'INVALID_VALUE', `/pages/${index}/parity_reason`, 'parity filler does not place the target section on its requested one-based page parity')
    }
  })
  sectionOrder.forEach((sectionID, sectionIndex) => {
    if (sectionIndex === 0) return
    const breakType = sectionBreaks.get(sectionID)
    if (breakType !== 'odd-page' && breakType !== 'even-page') return
    const firstContentIndex = pageResults.findIndex((page) => page.sectionID === sectionID && page.kind === 'content')
    if (firstContentIndex < 0) {
      add(issues, 'BROKEN_REFERENCE', `/sections/${sectionIndex}/page_ids`, 'odd/even section requires a first content page')
      return
    }
    const oneBasedPage = firstContentIndex + 1
    if ((breakType === 'odd-page' && oneBasedPage % 2 !== 1) || (breakType === 'even-page' && oneBasedPage % 2 !== 0)) add(issues, 'INVALID_VALUE', `/sections/${sectionIndex}/break_type`, 'odd/even section must begin on its requested one-based physical page parity')
  })
  paragraphSlices.forEach((records) => records.forEach((record, index) => {
    const previous = records[index - 1]
    const next = records[index + 1]
    if (record.sliceOrdinal !== index) add(issues, 'INVALID_VALUE', `${record.path}/slice_ordinal`, 'paragraph slice ordinals must be globally contiguous from zero')
    if (record.repeated) {
      const original = records[0]
      if (index === 0 || original?.repeated || record.first !== 0 || record.last !== original?.last || record.pageOrdinal <= (previous?.pageOrdinal ?? record.pageOrdinal) || record.continued || record.continues || record.continuedColumn || record.continuesColumn) add(issues, 'INVALID_VALUE', record.path, 'repeated headers must repeat one complete original paragraph on a later page without continuation flags')
      return
    }
    if (index === 0) {
      if (record.first !== 0) add(issues, 'INVALID_VALUE', `${record.path}/first_line_ordinal`, 'first paragraph slice must begin at shaped line zero')
      if (record.continued !== false) add(issues, 'INVALID_VALUE', `${record.path}/continued_from_previous_page`, 'first paragraph slice cannot continue from a previous page')
      if (record.continuedColumn !== false) add(issues, 'INVALID_VALUE', `${record.path}/continued_from_previous_column`, 'first paragraph slice cannot continue from a previous column')
    } else {
      if (previous?.last !== undefined && record.first !== previous.last + 1) add(issues, 'INVALID_VALUE', `${record.path}/first_line_ordinal`, 'paragraph slices must cover globally contiguous shaped-line ranges')
      const pageTransition = previous !== undefined && record.pageOrdinal > previous.pageOrdinal
      const columnTransition = previous !== undefined && record.pageOrdinal === previous.pageOrdinal && record.columnOrdinal !== previous.columnOrdinal
      if (record.continued !== pageTransition || record.continuedColumn !== columnTransition) add(issues, 'INVALID_VALUE', record.path, 'later paragraph slice must identify its exact preceding page or column transition')
    }
    const nextPage = next !== undefined && !next.repeated && next.pageOrdinal > record.pageOrdinal
    const nextColumn = next !== undefined && !next.repeated && next.pageOrdinal === record.pageOrdinal && next.columnOrdinal !== record.columnOrdinal
    if (record.continues !== nextPage || record.continuesColumn !== nextColumn) add(issues, 'INVALID_VALUE', record.path, 'paragraph continuation flags must identify the exact next page or column transition')
  }))
  if (status === 'paginated' && (sectionList.length === 0 || pageList.length === 0)) add(issues, 'REQUIRED', '', 'paginated layout requires at least one section and page')
  issues.sort(compareNativeValidationIssues)
  return issues.length > 0 ? { ok: false, issues: issues.slice(0, DOCX_NATIVE_LIMITS.maxIssues) } : { ok: true, value: value as NativeDocxPaginatedLayoutV1 }
}

function validateTableRowFragments(pages: unknown[], issues: NativeDocxValidationIssue[]): void {
  const rows = new Map<string,{ ordinal:number; end:number; height:number; page:number; rowOrdinal:number; sectionID:string; path:string }>()
  let count = 0
  for (const [pageIndex,pageValue] of pages.entries()) {
    if(!isObject(pageValue)||pageValue.table_rows===undefined)continue
    const path=`/pages/${pageIndex}/table_rows`
    const fragments=array(pageValue.table_rows,path,DOCX_PAGINATION_LIMITS.maxLinePlacements,issues)
    if(pageValue.kind==='parity-blank'&&fragments.length)add(issues,'INVALID_UNION',path,'parity blank pages cannot contain table row fragments')
    for(const [index,value] of fragments.entries()) {
      count+=1
      if(count>DOCX_PAGINATION_LIMITS.maxLinePlacements){add(issues,'LIMIT_EXCEEDED',path,'table row fragments exceed the global placement budget');return}
      const rowPath=`${path}/${index}`,row=object(value,rowPath,DOCX_PAGINATED_LAYOUT_V1_BINDING_FIELDS.TableRowFragmentV1,issues)
      if(!row)continue
      const id=stringValue(row.id,`${rowPath}/id`,issues,ID,1024),tableID=stringValue(row.table_id,`${rowPath}/table_id`,issues,ID,1024),rowID=stringValue(row.row_id,`${rowPath}/row_id`,issues,ID,1024),sectionID=stringValue(row.section_id,`${rowPath}/section_id`,issues,ID,1024),columnID=stringValue(row.column_id,`${rowPath}/column_id`,issues,ID,1024)
      const rowOrdinal=integer(row.row_ordinal,`${rowPath}/row_ordinal`,0,DOCX_NATIVE_LIMITS.maxCollectionItems,issues),ordinal=integer(row.fragment_ordinal,`${rowPath}/fragment_ordinal`,0,DOCX_PAGINATION_LIMITS.maxLinePlacements,issues),columnOrdinal=integer(row.column_ordinal,`${rowPath}/column_ordinal`,0,44,issues)
      const x=integer(row.x_millipoints,`${rowPath}/x_millipoints`,0,DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints,issues),y=integer(row.y_millipoints,`${rowPath}/y_millipoints`,0,DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints,issues),width=integer(row.width_millipoints,`${rowPath}/width_millipoints`,1,DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints,issues),height=integer(row.height_millipoints,`${rowPath}/height_millipoints`,1,DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints,issues),sourceY=integer(row.source_y_millipoints,`${rowPath}/source_y_millipoints`,0,DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints,issues),sourceHeight=integer(row.source_height_millipoints,`${rowPath}/source_height_millipoints`,1,DOCX_PAGINATION_LIMITS.maxCoordinateMilliPoints,issues)
      const column=(Array.isArray(pageValue.columns)?pageValue.columns.slice(0,45):[]).find(value=>isObject(value)&&value.id===columnID)
      if(!isObject(column)||column.section_id!==sectionID||column.ordinal!==columnOrdinal)add(issues,'BROKEN_REFERENCE',rowPath,'row fragment must reference its exact page section column')
      else if(typeof x==='number'&&typeof y==='number'&&typeof width==='number'&&typeof height==='number'&&(x<(column.x_millipoints as number)||x+width>(column.x_millipoints as number)+(column.width_millipoints as number)||y<(column.y_millipoints as number)||y+height>(column.y_millipoints as number)+(column.height_millipoints as number)))add(issues,'OUT_OF_RANGE',rowPath,'row fragment exceeds its exact page column')
      if(tableID&&rowID&&sectionID&&ordinal!==undefined&&rowOrdinal!==undefined&&sourceY!==undefined&&sourceHeight!==undefined&&height!==undefined) {
        const key=JSON.stringify([tableID,rowID]),previous=rows.get(key)
        if(id!==`table-row:${tableID}:${rowID}:${ordinal}`)add(issues,'INVALID_VALUE',`${rowPath}/id`,'row fragment ID must derive from table, row and ordinal')
        if(sourceY+height>sourceHeight||(!previous&&(ordinal!==0||sourceY!==0))||(previous&&(ordinal!==previous.ordinal+1||sourceY!==previous.end||sourceHeight!==previous.height||rowOrdinal!==previous.rowOrdinal||sectionID!==previous.sectionID||pageIndex<=previous.page)))add(issues,'BROKEN_REFERENCE',rowPath,'row fragments must exactly and contiguously cover one source row across increasing pages')
        rows.set(key,{ordinal,end:sourceY+height,height:sourceHeight,page:pageIndex,rowOrdinal,sectionID,path:rowPath})
      }
    }
  }
  for(const row of rows.values())if(row.end!==row.height)add(issues,'BROKEN_REFERENCE',row.path,'final row fragment must cover the source row through its end')
}

/** Strict output decoding plus exact source-completeness validation. */
export function decodeNativeDocxPaginatedLayoutForRequest(value: unknown, requestValue: unknown): DecodeNativeDocxPaginatedLayoutResult {
  const request = decodeNativeDocxPaginationRequestV1(requestValue)
  const output = decodeNativeDocxPaginatedLayout(value)
  if (!request.ok || !output.ok) {
    const issues: NativeDocxValidationIssue[] = []
    if (!request.ok) issues.push(...request.issues.map((entry) => ({ ...entry, path: `/request${entry.path}` })))
    if (!output.ok) issues.push(...output.issues.map((entry) => ({ ...entry, path: `/output${entry.path}` })))
    return { ok: false, issues: issues.slice(0, DOCX_NATIVE_LIMITS.maxIssues) }
  }
  const issues = validateNativeDocxPaginatedLayoutSourceV1(output.value, request.value)
  return issues.length > 0 ? { ok: false, issues } : output
}
