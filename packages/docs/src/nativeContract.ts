/**
 * Native DOCX parse/layout boundary, version 1.
 *
 * The original OPC package remains authoritative. This model carries stable
 * identities and source anchors into browser layout/editing code; it is not a
 * second persistence representation and must never be serialized back into a
 * new DOCX wholesale.
 */

import { asciiLowerNative as asciiLower, compareNativeValidationIssues } from './nativeDeterminism.js'

export const DOCX_NATIVE_PROTOCOL = 'injoffice.docx.native'
export const DOCX_NATIVE_VERSION = 1 as const
export const DOCX_NATIVE_LIMITS = {
  maxJsonBytes: 8 * 1024 * 1024,
  maxDepth: 64,
  maxNodes: 100_000,
  maxCollectionItems: 10_000,
  maxTextLength: 1_048_576,
  maxIssues: 100,
} as const

/** Largest integer twip value whose exact x50 milli-point projection is safe. */
export const DOCX_MAX_TWIPS_FOR_MILLIPOINTS = 180_143_985_094_819

export type NativeDocxStoryKind = 'body' | 'header' | 'footer' | 'footnote' | 'endnote' | 'comment'
export type NativeDocxEditMode = 'read-write' | 'read-only'
export type NativeDocxEditOperation =
  | 'text.replace'
  | 'properties.patch'
  | 'block.insert_after'
  | 'block.delete'
  | 'drawing.replace'

export interface NativeDocxSourcePackageV1 {
  package_sha256: string
  main_part: string
}

export interface NativeDocxSourceAnchorV1 {
  part_name: string
  /** Namespace-qualified, parser-owned path; never a user-facing identity. */
  path: string
  start_byte: number
  end_byte: number
  xml_sha256: string
}

export interface NativeDocxRefusalV1 {
  code: string
  message: string
  preservation: 'preserve-verbatim' | 'refuse-mutation'
}

export interface NativeDocxEditPolicyV1 {
  mode: NativeDocxEditMode
  allowed_operations: NativeDocxEditOperation[]
  refusal?: NativeDocxRefusalV1
}

export interface NativeDocxCapabilityV1 {
  name: string
  level: 'read-write' | 'read-only' | 'passthrough' | 'unsupported'
  detail?: string
}

export interface NativeDocxPassthroughPartV1 {
  part_name: string
  content_type: string
  byte_length: number
  sha256: string
  policy: 'preserve-verbatim'
}

export interface NativeDocxRunPropertiesV1 {
  character_style_id?: string
  font_family?: string
  font_size_half_points?: number
  bold?: boolean
  italic?: boolean
  underline?: 'none' | 'single' | 'double' | 'words'
  vertical_alignment?: 'baseline' | 'subscript' | 'superscript'
  color?: string
  highlight?: string
  language?: string
  rtl?: boolean
  hidden?: boolean
}

export interface NativeDocxDrawingV1 {
  id: string
  anchor: NativeDocxSourceAnchorV1
  relationship_id?: string
  media_part?: string
  content_type?: string
  name?: string
  alt_text?: string
  placement: 'inline' | 'floating'
  width_emu: number
  height_emu: number
  rotation_degrees?: 0 | 90 | 180 | 270
  flip_horizontal?: boolean
  flip_vertical?: boolean
  source_crop?: { left: number; top: number; right: number; bottom: number }
  x_emu?: number
  y_emu?: number
  horizontal_relative_from?: string
  vertical_relative_from?: string
  wrap?: 'none' | 'square' | 'tight' | 'through' | 'top-and-bottom'
  edit_policy: NativeDocxEditPolicyV1
}

export interface NativeDocxReferenceV1 {
  kind: 'footnote' | 'endnote' | 'comment' | 'comment-range-start' | 'comment-range-end'
  target_id: string
  /** `label` is the exact w:footnoteRef/w:endnoteRef leaf in its owning note. */
  role?: 'anchor' | 'label'
}

export interface NativeDocxRunV1 {
  kind: 'text' | 'control' | 'reference' | 'drawing'
  id: string
  anchor: NativeDocxSourceAnchorV1
  properties?: NativeDocxRunPropertiesV1
  text?: string
  /** Source PAGE/NUMPAGES instruction; cached text is not a rendering authority. */
  page_field?: 'PAGE' | 'NUMPAGES'
  control?: 'tab' | 'line-break' | 'page-break' | 'column-break' | 'soft-hyphen'
  reference?: NativeDocxReferenceV1
  drawing?: NativeDocxDrawingV1
}

export interface NativeDocxNumberingReferenceV1 {
  num_id: string
  level: number
  abstract_num_id?: string
}

export interface NativeDocxParagraphPropertiesV1 {
  paragraph_style_id?: string
  numbering?: NativeDocxNumberingReferenceV1
  alignment?: 'left' | 'center' | 'right' | 'both' | 'distribute'
  keep_next?: boolean
  keep_lines?: boolean
  page_break_before?: boolean
  widow_control?: boolean
}

export interface NativeDocxParagraphV1 {
  id: string
  anchor: NativeDocxSourceAnchorV1
  edit_policy: NativeDocxEditPolicyV1
  properties: NativeDocxParagraphPropertiesV1
  runs: NativeDocxRunV1[]
}

export interface NativeDocxTableCellV1 {
  id: string
  anchor: NativeDocxSourceAnchorV1
  width_twips?: number
  grid_span: number
  vertical_merge: 'none' | 'restart' | 'continue'
  borders?: NativeDocxTableBordersV1
  shading_rgb?: string
  /** V1 is deliberately conservative: nested tables remain passthrough. */
  paragraphs: NativeDocxParagraphV1[]
}

export interface NativeDocxTableRowV1 {
  id: string
  anchor: NativeDocxSourceAnchorV1
  height_twips?: number
  height_rule?: 'atLeast' | 'exact'
  repeat_header: boolean
  cant_split?: boolean
  cells: NativeDocxTableCellV1[]
}

export interface NativeDocxTableV1 {
  id: string
  anchor: NativeDocxSourceAnchorV1
  edit_policy: NativeDocxEditPolicyV1
  table_style_id?: string
  width_twips?: number
  layout?: 'fixed'
  alignment?: 'left'
  indent_twips?: number
  grid_widths_twips?: number[]
  cell_margins?: NativeDocxTableCellMarginsV1
  borders?: NativeDocxTableBordersV1
  rows: NativeDocxTableRowV1[]
}

export interface NativeDocxTableBorderV1 {
  style: 'none' | 'single'
  size_eighth_points: number
  color_rgb?: string
}

export interface NativeDocxTableBordersV1 {
  top?: NativeDocxTableBorderV1
  right?: NativeDocxTableBorderV1
  bottom?: NativeDocxTableBorderV1
  left?: NativeDocxTableBorderV1
  inside_horizontal?: NativeDocxTableBorderV1
  inside_vertical?: NativeDocxTableBorderV1
}

export interface NativeDocxTableCellMarginsV1 {
  top_twips: number
  right_twips: number
  bottom_twips: number
  left_twips: number
}

export interface NativeDocxBlockV1 {
  kind: 'paragraph' | 'table'
  id: string
  paragraph?: NativeDocxParagraphV1
  table?: NativeDocxTableV1
}

export interface NativeDocxStoryV1 {
  id: string
  kind: NativeDocxStoryKind
  part_name: string
  /** Native numeric note/comment identity where the owning part uses one. */
  native_story_id?: string
  /** Main-document relationship that resolved the shared note part. */
  relationship_id?: string
  note_role?: 'content' | 'separator' | 'continuation-separator'
  anchor: NativeDocxSourceAnchorV1
  blocks: NativeDocxBlockV1[]
}

export interface NativeDocxHeaderFooterReferenceV1 {
  kind: 'default' | 'first' | 'even'
  story_id: string
  relationship_id: string
}

export interface NativeDocxPageMarginsV1 {
  top_twips: number
  right_twips: number
  bottom_twips: number
  left_twips: number
  header_twips: number
  footer_twips: number
  gutter_twips: number
}

export interface NativeDocxColumnV1 {
  /** Stable section-qualified identity; ordinal is the physical left-to-right order. */
  id: string
  ordinal: number
  /** Present only for explicit WordprocessingML column definitions. */
  width_twips?: number
  /** Present only for explicit definitions; the final column must use zero. */
  space_after_twips?: number
}

export interface NativeDocxPageGeometryV1 {
  width_twips: number
  height_twips: number
  orientation: 'portrait' | 'landscape'
  margins: NativeDocxPageMarginsV1
  columns: number
  column_spacing_twips: number
  column_layout: 'equal-width' | 'explicit'
  column_definitions: NativeDocxColumnV1[]
}

export interface NativeDocxSectionV1 {
  id: string
  anchor: NativeDocxSourceAnchorV1
  /** First body block governed by this section. */
  starts_at_block_id: string
  break_type: 'continuous' | 'next-page' | 'even-page' | 'odd-page' | 'next-column'
  /** Exact w:titlePg policy. False means the element is absent or explicitly off. */
  title_page: boolean
  page: NativeDocxPageGeometryV1
  header_refs: NativeDocxHeaderFooterReferenceV1[]
  footer_refs: NativeDocxHeaderFooterReferenceV1[]
}

export interface NativeDocxCommentV1 {
  id: string
  native_comment_id: string
  author: string
  initials?: string
  created_at?: string
  anchor: NativeDocxSourceAnchorV1
  body_story_id: string
}

export interface NativeDocxUnsupportedCapabilityV1 {
  id: string
  code: string
  capability: string
  scope_id: string
  anchor?: NativeDocxSourceAnchorV1
  preservation: 'preserve-verbatim' | 'refuse-mutation'
  message: string
}

export interface NativeDocxDocumentV1 {
  protocol: typeof DOCX_NATIVE_PROTOCOL
  version: typeof DOCX_NATIVE_VERSION
  document_id: string
  /** Opaque gateway revision used for compare-and-swap saves. */
  revision: string
  source: NativeDocxSourcePackageV1
  body: NativeDocxStoryV1
  sections: NativeDocxSectionV1[]
  headers: NativeDocxStoryV1[]
  footers: NativeDocxStoryV1[]
  notes: NativeDocxStoryV1[]
  comment_stories: NativeDocxStoryV1[]
  comments: NativeDocxCommentV1[]
  capabilities: NativeDocxCapabilityV1[]
  passthrough_parts: NativeDocxPassthroughPartV1[]
  unsupported: NativeDocxUnsupportedCapabilityV1[]
}

export type NativeDocxIssueCode =
  | 'UNKNOWN_FIELD'
  | 'REQUIRED'
  | 'INVALID_TYPE'
  | 'INVALID_VALUE'
  | 'OUT_OF_RANGE'
  | 'UNSUPPORTED_PROTOCOL'
  | 'UNSUPPORTED_VERSION'
  | 'DUPLICATE_ID'
  | 'BROKEN_REFERENCE'
  | 'INVALID_UNION'
  | 'LIMIT_EXCEEDED'

export interface NativeDocxValidationIssue {
  code: NativeDocxIssueCode
  /** RFC 6901 JSON Pointer into the submitted body. */
  path: string
  message: string
}

export type DecodeNativeDocxResult =
  | { ok: true; value: NativeDocxDocumentV1 }
  | { ok: false; issues: NativeDocxValidationIssue[] }

export class NativeDocxValidationError extends Error {
  readonly issues: NativeDocxValidationIssue[]

  constructor(issues: NativeDocxValidationIssue[]) {
    super(issues.map((entry) => `${entry.path || '/'}: ${entry.message}`).join('; '))
    this.name = 'NativeDocxValidationError'
    this.issues = issues
  }
}

type JsonObject = Record<string, unknown>

/** Kept in lockstep with schema/native-docx-v1.schema.json by tests. */
export const DOCX_NATIVE_V1_BINDING_FIELDS = {
  SourcePackageV1: ['package_sha256', 'main_part'],
  SourceAnchorV1: ['part_name', 'path', 'start_byte', 'end_byte', 'xml_sha256'],
  RefusalV1: ['code', 'message', 'preservation'],
  EditPolicyV1: ['mode', 'allowed_operations', 'refusal'],
  CapabilityV1: ['name', 'level', 'detail'],
  PassthroughPartV1: ['part_name', 'content_type', 'byte_length', 'sha256', 'policy'],
  RunPropertiesV1: ['character_style_id', 'font_family', 'font_size_half_points', 'bold', 'italic', 'underline', 'vertical_alignment', 'color', 'highlight', 'language', 'rtl', 'hidden'],
  DrawingV1: ['id', 'anchor', 'relationship_id', 'media_part', 'content_type', 'name', 'alt_text', 'placement', 'width_emu', 'height_emu', 'x_emu', 'y_emu', 'horizontal_relative_from', 'vertical_relative_from', 'wrap', 'edit_policy', 'rotation_degrees', 'flip_horizontal', 'flip_vertical', 'source_crop'],
  DrawingCropV1: ['left', 'top', 'right', 'bottom'],
  ReferenceV1: ['kind', 'target_id', 'role'],
  RunV1: ['kind', 'id', 'anchor', 'properties', 'text', 'page_field', 'control', 'reference', 'drawing'],
  NumberingReferenceV1: ['num_id', 'level', 'abstract_num_id'],
  ParagraphPropertiesV1: ['paragraph_style_id', 'numbering', 'alignment', 'keep_next', 'keep_lines', 'page_break_before', 'widow_control'],
  ParagraphV1: ['id', 'anchor', 'edit_policy', 'properties', 'runs'],
  TableBorderV1: ['style', 'size_eighth_points', 'color_rgb'],
  TableBordersV1: ['top', 'right', 'bottom', 'left', 'inside_horizontal', 'inside_vertical'],
  TableCellMarginsV1: ['top_twips', 'right_twips', 'bottom_twips', 'left_twips'],
  TableCellV1: ['id', 'anchor', 'width_twips', 'grid_span', 'vertical_merge', 'borders', 'shading_rgb', 'paragraphs'],
  TableRowV1: ['id', 'anchor', 'height_twips', 'height_rule', 'repeat_header', 'cant_split', 'cells'],
  TableV1: ['id', 'anchor', 'edit_policy', 'table_style_id', 'width_twips', 'layout', 'alignment', 'indent_twips', 'grid_widths_twips', 'cell_margins', 'borders', 'rows'],
  BlockV1: ['kind', 'id', 'paragraph', 'table'],
  StoryV1: ['id', 'kind', 'part_name', 'native_story_id', 'relationship_id', 'note_role', 'anchor', 'blocks'],
  HeaderFooterReferenceV1: ['kind', 'story_id', 'relationship_id'],
  PageMarginsV1: ['top_twips', 'right_twips', 'bottom_twips', 'left_twips', 'header_twips', 'footer_twips', 'gutter_twips'],
  ColumnV1: ['id', 'ordinal', 'width_twips', 'space_after_twips'],
  PageGeometryV1: ['width_twips', 'height_twips', 'orientation', 'margins', 'columns', 'column_spacing_twips', 'column_layout', 'column_definitions'],
  SectionV1: ['id', 'anchor', 'starts_at_block_id', 'break_type', 'title_page', 'page', 'header_refs', 'footer_refs'],
  CommentV1: ['id', 'native_comment_id', 'author', 'initials', 'created_at', 'anchor', 'body_story_id'],
  UnsupportedCapabilityV1: ['id', 'code', 'capability', 'scope_id', 'anchor', 'preservation', 'message'],
  DocumentV1: ['protocol', 'version', 'document_id', 'revision', 'source', 'body', 'sections', 'headers', 'footers', 'notes', 'comment_stories', 'comments', 'capabilities', 'passthrough_parts', 'unsupported'],
} as const

const ID = /^[A-Za-z0-9][A-Za-z0-9._:-]{0,255}$/
const NOTE_CONTENT_ID = /^[1-9][0-9]{0,18}$/
const PART_SEGMENT = /^(?:[A-Za-z0-9._~!$&'()*+,;=@-]|%[0-9A-F]{2})+$/
const SHA256 = /^sha256:[0-9a-f]{64}$/
const COLOR = /^(?:auto|[0-9A-F]{6})$/
const operations = ['text.replace', 'properties.patch', 'block.insert_after', 'block.delete', 'drawing.replace'] as const
const paragraphOperations: readonly NativeDocxEditOperation[] = ['text.replace', 'properties.patch', 'block.insert_after', 'block.delete']
const tableOperations: readonly NativeDocxEditOperation[] = ['properties.patch', 'block.insert_after', 'block.delete']
const drawingOperations: readonly NativeDocxEditOperation[] = ['drawing.replace']

function isObject(value: unknown): value is JsonObject {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

function add(issues: NativeDocxValidationIssue[], code: NativeDocxIssueCode, path: string, message: string): void {
  if (issues.length >= DOCX_NATIVE_LIMITS.maxIssues) return
  if (!issues.some((issue) => issue.code === code && issue.path === path && issue.message === message)) issues.push({ code, path, message })
}

function object(value: unknown, path: string, fields: readonly string[], issues: NativeDocxValidationIssue[]): JsonObject | null {
  if (!isObject(value)) {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, value === undefined ? 'field is required' : 'must be an object')
    return null
  }
  const allowed = new Set(fields)
  for (const key of Object.keys(value).sort()) {
    if (!allowed.has(key)) add(issues, 'UNKNOWN_FIELD', `${path}/${escapePointer(key)}`, `unknown field ${JSON.stringify(key)}`)
  }
  return value
}

function escapePointer(value: string): string {
  return value.replace(/~/g, '~0').replace(/\//g, '~1')
}

function codePointLength(value: string): number {
  return [...value].length
}

function stringValue(value: unknown, path: string, issues: NativeDocxValidationIssue[], pattern?: RegExp, maxLength = 4096): string | null {
  if (value === undefined) {
    add(issues, 'REQUIRED', path, 'field is required')
    return null
  }
  if (typeof value !== 'string') {
    add(issues, 'INVALID_TYPE', path, 'must be a string')
    return null
  }
  if (codePointLength(value) === 0 || codePointLength(value) > maxLength || (pattern && !pattern.test(value))) {
    add(issues, 'INVALID_VALUE', path, 'contains an invalid string value')
    return null
  }
  return value
}

function optionalString(value: unknown, path: string, issues: NativeDocxValidationIssue[], pattern?: RegExp, maxLength = 4096): string | undefined {
  if (value === undefined) return undefined
  return stringValue(value, path, issues, pattern, maxLength) ?? undefined
}

function enumValue<T extends string>(value: unknown, path: string, allowed: readonly T[], issues: NativeDocxValidationIssue[]): T | null {
  if (typeof value !== 'string' || !allowed.includes(value as T)) {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_VALUE', path, `must be one of ${allowed.join(', ')}`)
    return null
  }
  return value as T
}

function booleanValue(value: unknown, path: string, issues: NativeDocxValidationIssue[], required = true): boolean | undefined {
  if (value === undefined && !required) return undefined
  if (typeof value !== 'boolean') {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, 'must be a boolean')
    return undefined
  }
  return value
}

function integer(value: unknown, path: string, issues: NativeDocxValidationIssue[], min = 0, required = true): number | undefined {
  if (value === undefined && !required) return undefined
  if (typeof value !== 'number' || !Number.isInteger(value)) {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, 'must be a safe integer')
    return undefined
  }
  if (!Number.isSafeInteger(value)) {
    add(issues, 'OUT_OF_RANGE', path, 'must be a safe integer')
    return undefined
  }
  if (Object.is(value, -0)) {
    add(issues, 'INVALID_VALUE', path, 'negative zero is not permitted')
    return undefined
  }
  if (value < min) {
    add(issues, 'OUT_OF_RANGE', path, `must be at least ${min}`)
    return undefined
  }
  return value
}

function twipsInteger(value: unknown, path: string, issues: NativeDocxValidationIssue[], min = 0, required = true): number | undefined {
  const parsed = integer(value, path, issues, min, required)
  if (parsed !== undefined && parsed > DOCX_MAX_TWIPS_FOR_MILLIPOINTS) {
    add(issues, 'OUT_OF_RANGE', path, `must not exceed ${DOCX_MAX_TWIPS_FOR_MILLIPOINTS} before milli-point conversion`)
    return undefined
  }
  return parsed
}

function array(value: unknown, path: string, issues: NativeDocxValidationIssue[], maxItems: number = DOCX_NATIVE_LIMITS.maxCollectionItems): unknown[] {
  if (!Array.isArray(value)) {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, 'must be an array')
    return []
  }
  if (value.length > maxItems) add(issues, 'LIMIT_EXCEEDED', path, `must contain at most ${maxItems} items`)
  return value.slice(0, maxItems)
}

interface AnchorBounds {
  partName: string
  start: number
  end: number
}

function isCanonicalPartName(value: string): boolean {
  if (codePointLength(value) > 4096 || value.startsWith('/') || value.includes('\\') || value.includes('//')) return false
  const segments = value.split('/')
  return segments.every((segment) => {
    if (segment === '.' || segment === '..' || !PART_SEGMENT.test(segment)) return false
    try {
      const decoded = decodeURIComponent(segment)
      if (decoded === '.' || decoded === '..' || decoded.endsWith('.') || /[\\/?#%]/.test(decoded)) return false
      return ![...decoded].some((character) => {
        const code = character.codePointAt(0) ?? 0
        return code < 0x20 || code === 0x7f
      })
    } catch {
      return false
    }
  })
}

function partNameValue(value: unknown, path: string, issues: NativeDocxValidationIssue[]): string | null {
  const part = stringValue(value, path, issues)
  if (part !== null && !isCanonicalPartName(part)) {
    add(issues, 'INVALID_VALUE', path, 'contains a non-canonical OPC part name')
    return null
  }
  return part
}

function validateAnchor(value: unknown, path: string, issues: NativeDocxValidationIssue[], expectedPart?: string | null, parent?: AnchorBounds | null): AnchorBounds | null {
  const entry = object(value, path, DOCX_NATIVE_V1_BINDING_FIELDS.SourceAnchorV1, issues)
  if (!entry) return null
  const partName = partNameValue(entry.part_name, `${path}/part_name`, issues)
  stringValue(entry.path, `${path}/path`, issues)
  const start = integer(entry.start_byte, `${path}/start_byte`, issues)
  const end = integer(entry.end_byte, `${path}/end_byte`, issues, 1)
  if (start !== undefined && end !== undefined && end <= start) add(issues, 'OUT_OF_RANGE', `${path}/end_byte`, 'must be greater than start_byte')
  if (partName && expectedPart && partName !== expectedPart) add(issues, 'INVALID_VALUE', `${path}/part_name`, `must equal owning part ${JSON.stringify(expectedPart)}`)
  if (partName && parent && (partName !== parent.partName || (start !== undefined && start < parent.start) || (end !== undefined && end > parent.end))) {
    add(issues, 'OUT_OF_RANGE', path, 'anchor must be contained by its parent anchor in the same part')
  }
  stringValue(entry.xml_sha256, `${path}/xml_sha256`, issues, SHA256)
  return partName && start !== undefined && end !== undefined && end > start ? { partName, start, end } : null
}

function validateRefusal(value: unknown, path: string, issues: NativeDocxValidationIssue[]): void {
  const entry = object(value, path, DOCX_NATIVE_V1_BINDING_FIELDS.RefusalV1, issues)
  if (!entry) return
  stringValue(entry.code, `${path}/code`, issues, ID)
  stringValue(entry.message, `${path}/message`, issues)
  enumValue(entry.preservation, `${path}/preservation`, ['preserve-verbatim', 'refuse-mutation'], issues)
}

function validateEditPolicy(value: unknown, path: string, issues: NativeDocxValidationIssue[], supportedOperations: readonly NativeDocxEditOperation[]): void {
  const entry = object(value, path, DOCX_NATIVE_V1_BINDING_FIELDS.EditPolicyV1, issues)
  if (!entry) return
  const mode = enumValue(entry.mode, `${path}/mode`, ['read-write', 'read-only'], issues)
  const allowed = array(entry.allowed_operations, `${path}/allowed_operations`, issues, operations.length)
  const seen = new Set<string>()
  allowed.forEach((operation, index) => {
    const parsed = enumValue(operation, `${path}/allowed_operations/${index}`, operations, issues)
    if (parsed && seen.has(parsed)) add(issues, 'INVALID_VALUE', `${path}/allowed_operations/${index}`, 'operation is duplicated')
    if (parsed && !supportedOperations.includes(parsed)) add(issues, 'INVALID_VALUE', `${path}/allowed_operations/${index}`, 'operation is not valid for this object type')
    if (parsed) seen.add(parsed)
  })
  if (entry.refusal !== undefined) validateRefusal(entry.refusal, `${path}/refusal`, issues)
  if (mode === 'read-only' && entry.refusal === undefined) add(issues, 'REQUIRED', `${path}/refusal`, 'read-only content requires an explicit refusal')
  if (mode === 'read-only' && allowed.length > 0) add(issues, 'INVALID_VALUE', `${path}/allowed_operations`, 'read-only content cannot advertise write operations')
  if (mode === 'read-write' && entry.refusal !== undefined) add(issues, 'INVALID_VALUE', `${path}/refusal`, 'read-write content cannot carry a refusal')
}

function validateRunProperties(value: unknown, path: string, issues: NativeDocxValidationIssue[]): void {
  const entry = object(value, path, DOCX_NATIVE_V1_BINDING_FIELDS.RunPropertiesV1, issues)
  if (!entry) return
  optionalString(entry.character_style_id, `${path}/character_style_id`, issues, ID)
  optionalString(entry.font_family, `${path}/font_family`, issues)
  integer(entry.font_size_half_points, `${path}/font_size_half_points`, issues, 1, false)
  booleanValue(entry.bold, `${path}/bold`, issues, false)
  booleanValue(entry.italic, `${path}/italic`, issues, false)
  if (entry.underline !== undefined) enumValue(entry.underline, `${path}/underline`, ['none', 'single', 'double', 'words'], issues)
  if (entry.vertical_alignment !== undefined) enumValue(entry.vertical_alignment, `${path}/vertical_alignment`, ['baseline', 'subscript', 'superscript'], issues)
  optionalString(entry.color, `${path}/color`, issues, COLOR)
  optionalString(entry.highlight, `${path}/highlight`, issues)
  optionalString(entry.language, `${path}/language`, issues)
  booleanValue(entry.rtl, `${path}/rtl`, issues, false)
  booleanValue(entry.hidden, `${path}/hidden`, issues, false)
}

type ReferenceTarget = 'modeled-id' | 'body-block' | 'header-story' | 'footer-story' | 'footnote-story' | 'endnote-story' | 'comment' | 'comment-story' | 'media-part'

interface PendingReference {
  id: string
  path: string
  target: ReferenceTarget
}

function reference(refs: PendingReference[], id: string, path: string, target: ReferenceTarget): void {
  refs.push({ id, path, target })
}

function validateDrawing(value: unknown, path: string, issues: NativeDocxValidationIssue[], ids: Set<string>, refs: PendingReference[], ownerPart: string | null, parentAnchor: AnchorBounds | null): void {
  const entry = object(value, path, DOCX_NATIVE_V1_BINDING_FIELDS.DrawingV1, issues)
  if (!entry) return
  trackId(entry.id, `${path}/id`, issues, ids)
  validateAnchor(entry.anchor, `${path}/anchor`, issues, ownerPart, parentAnchor)
  optionalString(entry.relationship_id, `${path}/relationship_id`, issues, ID)
  const mediaPart = entry.media_part === undefined ? undefined : partNameValue(entry.media_part, `${path}/media_part`, issues) ?? undefined
  optionalString(entry.content_type, `${path}/content_type`, issues)
  optionalString(entry.name, `${path}/name`, issues)
  optionalString(entry.alt_text, `${path}/alt_text`, issues)
  const placement = enumValue(entry.placement, `${path}/placement`, ['inline', 'floating'], issues)
  integer(entry.width_emu, `${path}/width_emu`, issues, 1)
  integer(entry.height_emu, `${path}/height_emu`, issues, 1)
  if (entry.rotation_degrees !== undefined && ![0, 90, 180, 270].includes(entry.rotation_degrees as number)) add(issues, 'INVALID_VALUE', `${path}/rotation_degrees`, 'must equal 0, 90, 180 or 270')
  booleanValue(entry.flip_horizontal, `${path}/flip_horizontal`, issues, false)
  booleanValue(entry.flip_vertical, `${path}/flip_vertical`, issues, false)
  if (entry.source_crop !== undefined) {
    const crop = object(entry.source_crop, `${path}/source_crop`, DOCX_NATIVE_V1_BINDING_FIELDS.DrawingCropV1, issues)
    if (crop) {
      for (const key of ['left','top','right','bottom'] as const) {
        const value = integer(crop[key], `${path}/source_crop/${key}`, issues, 0)
        if (value !== undefined && value !== null && value > 99000) add(issues, 'OUT_OF_RANGE', `${path}/source_crop/${key}`, 'crop must retain at least one percent per axis')
      }
      if (typeof crop.left === 'number' && typeof crop.right === 'number' && crop.left + crop.right > 99000 || typeof crop.top === 'number' && typeof crop.bottom === 'number' && crop.top + crop.bottom > 99000) add(issues, 'OUT_OF_RANGE', `${path}/source_crop`, 'crop must retain at least one percent per axis')
    }
  }
  integer(entry.x_emu, `${path}/x_emu`, issues, Number.MIN_SAFE_INTEGER, false)
  integer(entry.y_emu, `${path}/y_emu`, issues, Number.MIN_SAFE_INTEGER, false)
  optionalString(entry.horizontal_relative_from, `${path}/horizontal_relative_from`, issues)
  optionalString(entry.vertical_relative_from, `${path}/vertical_relative_from`, issues)
  if (entry.wrap !== undefined) enumValue(entry.wrap, `${path}/wrap`, ['none', 'square', 'tight', 'through', 'top-and-bottom'], issues)
  if (placement === 'inline' && (entry.x_emu !== undefined || entry.y_emu !== undefined)) add(issues, 'INVALID_VALUE', path, 'inline drawings cannot carry floating offsets')
  validateEditPolicy(entry.edit_policy, `${path}/edit_policy`, issues, drawingOperations)
  if (mediaPart) reference(refs, mediaPart, `${path}/media_part`, 'media-part')
}

function validateRun(value: unknown, path: string, issues: NativeDocxValidationIssue[], ids: Set<string>, refs: PendingReference[], ownerPart: string | null, parentAnchor: AnchorBounds | null): void {
  const entry = object(value, path, DOCX_NATIVE_V1_BINDING_FIELDS.RunV1, issues)
  if (!entry) return
  const kind = enumValue(entry.kind, `${path}/kind`, ['text', 'control', 'reference', 'drawing'], issues)
  trackId(entry.id, `${path}/id`, issues, ids)
  const runAnchor = validateAnchor(entry.anchor, `${path}/anchor`, issues, ownerPart, parentAnchor)
  if (entry.properties !== undefined) validateRunProperties(entry.properties, `${path}/properties`, issues)
  if (entry.page_field !== undefined) {
    enumValue(entry.page_field, `${path}/page_field`, ['PAGE', 'NUMPAGES'], issues)
    if (kind !== 'text') add(issues, 'INVALID_UNION', `${path}/page_field`, 'page field requires a text run')
    if (entry.text !== '') add(issues, 'INVALID_VALUE', `${path}/text`, 'page-field source text must be empty; cached text is not authoritative')
  }
  const payloads = ['text', 'control', 'reference', 'drawing'].filter((key) => entry[key] !== undefined)
  if (payloads.length !== 1 || payloads[0] !== kind) add(issues, 'INVALID_UNION', path, 'run kind must match exactly one payload')
  if (kind === 'text') {
    if (typeof entry.text !== 'string') add(issues, entry.text === undefined ? 'REQUIRED' : 'INVALID_TYPE', `${path}/text`, 'must be a string')
    else if (codePointLength(entry.text) > DOCX_NATIVE_LIMITS.maxTextLength) add(issues, 'LIMIT_EXCEEDED', `${path}/text`, `must contain at most ${DOCX_NATIVE_LIMITS.maxTextLength} characters`)
  }
  if (kind === 'control') enumValue(entry.control, `${path}/control`, ['tab', 'line-break', 'page-break', 'column-break', 'soft-hyphen'], issues)
  if (kind === 'reference') {
    const refEntry = object(entry.reference, `${path}/reference`, DOCX_NATIVE_V1_BINDING_FIELDS.ReferenceV1, issues)
    if (refEntry) {
      const referenceKind = enumValue(refEntry.kind, `${path}/reference/kind`, ['footnote', 'endnote', 'comment', 'comment-range-start', 'comment-range-end'], issues)
      const target = stringValue(refEntry.target_id, `${path}/reference/target_id`, issues, ID)
      if (refEntry.role !== undefined) enumValue(refEntry.role, `${path}/reference/role`, ['anchor', 'label'], issues)
      if (target && referenceKind) {
        const targetKind: ReferenceTarget = referenceKind === 'footnote' ? 'footnote-story' : referenceKind === 'endnote' ? 'endnote-story' : 'comment'
        reference(refs, target, `${path}/reference/target_id`, targetKind)
      }
    }
  }
  if (kind === 'drawing') validateDrawing(entry.drawing, `${path}/drawing`, issues, ids, refs, ownerPart, runAnchor)
}

function validateParagraphProperties(value: unknown, path: string, issues: NativeDocxValidationIssue[]): void {
  const entry = object(value, path, DOCX_NATIVE_V1_BINDING_FIELDS.ParagraphPropertiesV1, issues)
  if (!entry) return
  optionalString(entry.paragraph_style_id, `${path}/paragraph_style_id`, issues, ID)
  if (entry.numbering !== undefined) {
    const numbering = object(entry.numbering, `${path}/numbering`, DOCX_NATIVE_V1_BINDING_FIELDS.NumberingReferenceV1, issues)
    if (numbering) {
      stringValue(numbering.num_id, `${path}/numbering/num_id`, issues, ID)
      integer(numbering.level, `${path}/numbering/level`, issues, 0)
      optionalString(numbering.abstract_num_id, `${path}/numbering/abstract_num_id`, issues, ID)
    }
  }
  if (entry.alignment !== undefined) enumValue(entry.alignment, `${path}/alignment`, ['left', 'center', 'right', 'both', 'distribute'], issues)
  for (const key of ['keep_next', 'keep_lines', 'page_break_before', 'widow_control']) booleanValue(entry[key], `${path}/${key}`, issues, false)
}

function validateParagraph(value: unknown, path: string, issues: NativeDocxValidationIssue[], ids: Set<string>, refs: PendingReference[], ownerPart: string | null, parentAnchor: AnchorBounds | null, trackIdentity = true): void {
  const entry = object(value, path, DOCX_NATIVE_V1_BINDING_FIELDS.ParagraphV1, issues)
  if (!entry) return
  if (trackIdentity) trackId(entry.id, `${path}/id`, issues, ids)
  else stringValue(entry.id, `${path}/id`, issues, ID)
  const paragraphAnchor = validateAnchor(entry.anchor, `${path}/anchor`, issues, ownerPart, parentAnchor)
  validateEditPolicy(entry.edit_policy, `${path}/edit_policy`, issues, paragraphOperations)
  validateParagraphProperties(entry.properties, `${path}/properties`, issues)
  array(entry.runs, `${path}/runs`, issues).forEach((run, index) => validateRun(run, `${path}/runs/${index}`, issues, ids, refs, ownerPart, paragraphAnchor))
}

function validateTableBorder(value: unknown, path: string, issues: NativeDocxValidationIssue[]): void {
  const entry = object(value, path, DOCX_NATIVE_V1_BINDING_FIELDS.TableBorderV1, issues)
  if (!entry) return
  const style = enumValue(entry.style, `${path}/style`, ['none', 'single'], issues)
  integer(entry.size_eighth_points, `${path}/size_eighth_points`, issues, 0, false)
  optionalString(entry.color_rgb, `${path}/color_rgb`, issues, /^[0-9A-F]{6}$/)
  if (style === 'none' && (entry.size_eighth_points !== 0 || entry.color_rgb !== undefined)) add(issues, 'INVALID_VALUE', path, 'none border must have zero size and no color')
  if (style === 'single' && (!Number.isSafeInteger(entry.size_eighth_points) || (entry.size_eighth_points as number) <= 0 || entry.color_rgb === undefined)) add(issues, 'INVALID_VALUE', path, 'single border requires positive size and explicit RGB color')
}

function validateTableBorders(value: unknown, path: string, issues: NativeDocxValidationIssue[]): void {
  const entry = object(value, path, DOCX_NATIVE_V1_BINDING_FIELDS.TableBordersV1, issues)
  if (!entry) return
  for (const key of DOCX_NATIVE_V1_BINDING_FIELDS.TableBordersV1) if (entry[key] !== undefined) validateTableBorder(entry[key], `${path}/${key}`, issues)
}

function validateCellMargins(value: unknown, path: string, issues: NativeDocxValidationIssue[]): void {
  const entry = object(value, path, DOCX_NATIVE_V1_BINDING_FIELDS.TableCellMarginsV1, issues)
  if (!entry) return
  for (const key of DOCX_NATIVE_V1_BINDING_FIELDS.TableCellMarginsV1) twipsInteger(entry[key], `${path}/${key}`, issues, 0)
}

function validateTable(value: unknown, path: string, issues: NativeDocxValidationIssue[], ids: Set<string>, refs: PendingReference[], ownerPart: string | null, parentAnchor: AnchorBounds | null, trackIdentity = true): void {
  const entry = object(value, path, DOCX_NATIVE_V1_BINDING_FIELDS.TableV1, issues)
  if (!entry) return
  if (trackIdentity) trackId(entry.id, `${path}/id`, issues, ids)
  else stringValue(entry.id, `${path}/id`, issues, ID)
  const tableAnchor = validateAnchor(entry.anchor, `${path}/anchor`, issues, ownerPart, parentAnchor)
  validateEditPolicy(entry.edit_policy, `${path}/edit_policy`, issues, tableOperations)
  optionalString(entry.table_style_id, `${path}/table_style_id`, issues, ID)
  twipsInteger(entry.width_twips, `${path}/width_twips`, issues, 1, false)
  if (entry.layout !== undefined) enumValue(entry.layout, `${path}/layout`, ['fixed'], issues)
  if (entry.alignment !== undefined) enumValue(entry.alignment, `${path}/alignment`, ['left'], issues)
  twipsInteger(entry.indent_twips, `${path}/indent_twips`, issues, 0, false)
  if (entry.grid_widths_twips !== undefined) array(entry.grid_widths_twips, `${path}/grid_widths_twips`, issues).forEach((width, index) => twipsInteger(width, `${path}/grid_widths_twips/${index}`, issues, 1))
  if (entry.cell_margins !== undefined) validateCellMargins(entry.cell_margins, `${path}/cell_margins`, issues)
  if (entry.borders !== undefined) validateTableBorders(entry.borders, `${path}/borders`, issues)
  array(entry.rows, `${path}/rows`, issues).forEach((rowValue, rowIndex) => {
    const rowPath = `${path}/rows/${rowIndex}`
    const row = object(rowValue, rowPath, DOCX_NATIVE_V1_BINDING_FIELDS.TableRowV1, issues)
    if (!row) return
    trackId(row.id, `${rowPath}/id`, issues, ids)
    const rowAnchor = validateAnchor(row.anchor, `${rowPath}/anchor`, issues, ownerPart, tableAnchor)
    twipsInteger(row.height_twips, `${rowPath}/height_twips`, issues, 0, false)
    if (row.height_rule !== undefined) enumValue(row.height_rule, `${rowPath}/height_rule`, ['atLeast', 'exact'], issues)
    if ((row.height_twips === undefined) !== (row.height_rule === undefined)) add(issues, 'INVALID_VALUE', `${rowPath}/height_rule`, 'row height requires both height_twips and height_rule')
    booleanValue(row.repeat_header, `${rowPath}/repeat_header`, issues)
    booleanValue(row.cant_split, `${rowPath}/cant_split`, issues, false)
    array(row.cells, `${rowPath}/cells`, issues).forEach((cellValue, cellIndex) => {
      const cellPath = `${rowPath}/cells/${cellIndex}`
      const cell = object(cellValue, cellPath, DOCX_NATIVE_V1_BINDING_FIELDS.TableCellV1, issues)
      if (!cell) return
      trackId(cell.id, `${cellPath}/id`, issues, ids)
      const cellAnchor = validateAnchor(cell.anchor, `${cellPath}/anchor`, issues, ownerPart, rowAnchor)
      twipsInteger(cell.width_twips, `${cellPath}/width_twips`, issues, 0, false)
      integer(cell.grid_span, `${cellPath}/grid_span`, issues, 1)
      enumValue(cell.vertical_merge, `${cellPath}/vertical_merge`, ['none', 'restart', 'continue'], issues)
      if (cell.borders !== undefined) validateTableBorders(cell.borders, `${cellPath}/borders`, issues)
      optionalString(cell.shading_rgb, `${cellPath}/shading_rgb`, issues, /^[0-9A-F]{6}$/)
      array(cell.paragraphs, `${cellPath}/paragraphs`, issues).forEach((paragraph, paragraphIndex) => validateParagraph(paragraph, `${cellPath}/paragraphs/${paragraphIndex}`, issues, ids, refs, ownerPart, cellAnchor))
    })
  })
}

function validateBlock(value: unknown, path: string, issues: NativeDocxValidationIssue[], ids: Set<string>, refs: PendingReference[], bodyBlockIds: Set<string>, ownerPart: string | null, parentAnchor: AnchorBounds | null): void {
  const entry = object(value, path, DOCX_NATIVE_V1_BINDING_FIELDS.BlockV1, issues)
  if (!entry) return
  const kind = enumValue(entry.kind, `${path}/kind`, ['paragraph', 'table'], issues)
  const id = trackId(entry.id, `${path}/id`, issues, ids)
  if (id) bodyBlockIds.add(id)
  const payloads = ['paragraph', 'table'].filter((key) => entry[key] !== undefined)
  if (payloads.length !== 1 || payloads[0] !== kind) add(issues, 'INVALID_UNION', path, 'block kind must match exactly one payload')
  if (kind === 'paragraph') {
    validateParagraph(entry.paragraph, `${path}/paragraph`, issues, ids, refs, ownerPart, parentAnchor, false)
    if (isObject(entry.paragraph) && id && entry.paragraph.id !== id) add(issues, 'INVALID_VALUE', `${path}/paragraph/id`, 'payload id must equal block id')
  }
  if (kind === 'table') {
    validateTable(entry.table, `${path}/table`, issues, ids, refs, ownerPart, parentAnchor, false)
    if (isObject(entry.table) && id && entry.table.id !== id) add(issues, 'INVALID_VALUE', `${path}/table/id`, 'payload id must equal block id')
  }
}

interface StoryValidationResult {
  id: string | null
  kind: NativeDocxStoryKind | null
  partName: string | null
  anchor: AnchorBounds | null
  nativeStoryId: string | undefined
  noteRole: NativeDocxStoryV1['note_role']
  relationshipId: string | undefined
}

function validateStory(value: unknown, path: string, expectedKinds: readonly NativeDocxStoryKind[], issues: NativeDocxValidationIssue[], ids: Set<string>, refs: PendingReference[], bodyBlockIds: Set<string>): StoryValidationResult {
  const entry = object(value, path, DOCX_NATIVE_V1_BINDING_FIELDS.StoryV1, issues)
  if (!entry) return { id: null, kind: null, partName: null, anchor: null, nativeStoryId: undefined, noteRole: undefined, relationshipId: undefined }
  const id = trackId(entry.id, `${path}/id`, issues, ids)
  const kind = enumValue(entry.kind, `${path}/kind`, expectedKinds, issues)
  const partName = partNameValue(entry.part_name, `${path}/part_name`, issues)
  const noteRole = entry.note_role === undefined ? undefined : enumValue(entry.note_role, `${path}/note_role`, ['content', 'separator', 'continuation-separator'], issues) as NativeDocxStoryV1['note_role']
  let nativeStoryId: string | undefined
  let relationshipId: string | undefined
  if (kind === 'footnote' || kind === 'endnote') {
    if (noteRole === undefined) add(issues, 'REQUIRED', `${path}/note_role`, 'note stories require an explicit content or separator role')
    const expected = noteRole === 'separator' ? '-1' : noteRole === 'continuation-separator' ? '0' : undefined
    if (typeof entry.native_story_id !== 'string' || (expected === undefined ? !NOTE_CONTENT_ID.test(entry.native_story_id) : entry.native_story_id !== expected)) add(issues, 'INVALID_VALUE', `${path}/native_story_id`, 'native note id must exactly match its content or separator role')
    else nativeStoryId = entry.native_story_id
    if (entry.relationship_id === undefined) add(issues, 'REQUIRED', `${path}/relationship_id`, 'note stories require their resolved main-document relationship')
    else relationshipId = optionalString(entry.relationship_id, `${path}/relationship_id`, issues, ID)
  } else {
    nativeStoryId = optionalString(entry.native_story_id, `${path}/native_story_id`, issues, ID)
    if (kind === 'comment' && nativeStoryId === undefined) add(issues, 'REQUIRED', `${path}/native_story_id`, 'comment stories require their native OOXML id')
    if (entry.relationship_id !== undefined || noteRole !== undefined) add(issues, 'INVALID_UNION', path, 'relationship_id and note_role are reserved for note stories')
  }
  const storyAnchor = validateAnchor(entry.anchor, `${path}/anchor`, issues, partName)
  array(entry.blocks, `${path}/blocks`, issues).forEach((block, index) => validateBlock(block, `${path}/blocks/${index}`, issues, ids, refs, kind === 'body' ? bodyBlockIds : new Set(), partName, storyAnchor))
  return { id, kind, partName, anchor: storyAnchor, nativeStoryId, noteRole, relationshipId }
}

function validateSection(value: unknown, path: string, issues: NativeDocxValidationIssue[], ids: Set<string>, refs: PendingReference[], mainPart: string | null, bodyAnchor: AnchorBounds | null): void {
  const entry = object(value, path, DOCX_NATIVE_V1_BINDING_FIELDS.SectionV1, issues)
  if (!entry) return
  trackId(entry.id, `${path}/id`, issues, ids)
  validateAnchor(entry.anchor, `${path}/anchor`, issues, mainPart, bodyAnchor)
  const start = stringValue(entry.starts_at_block_id, `${path}/starts_at_block_id`, issues, ID)
  if (start) reference(refs, start, `${path}/starts_at_block_id`, 'body-block')
  enumValue(entry.break_type, `${path}/break_type`, ['continuous', 'next-page', 'even-page', 'odd-page', 'next-column'], issues)
  booleanValue(entry.title_page, `${path}/title_page`, issues)
  const page = object(entry.page, `${path}/page`, DOCX_NATIVE_V1_BINDING_FIELDS.PageGeometryV1, issues)
  if (page) {
    twipsInteger(page.width_twips, `${path}/page/width_twips`, issues, 1)
    twipsInteger(page.height_twips, `${path}/page/height_twips`, issues, 1)
    enumValue(page.orientation, `${path}/page/orientation`, ['portrait', 'landscape'], issues)
    const margins = object(page.margins, `${path}/page/margins`, DOCX_NATIVE_V1_BINDING_FIELDS.PageMarginsV1, issues)
    if (margins) for (const key of DOCX_NATIVE_V1_BINDING_FIELDS.PageMarginsV1) twipsInteger(margins[key], `${path}/page/margins/${key}`, issues, 0)
    const columnCount = integer(page.columns, `${path}/page/columns`, issues, 1)
    if (columnCount !== undefined && columnCount > 45) add(issues, 'OUT_OF_RANGE', `${path}/page/columns`, 'must be from 1 through 45')
    twipsInteger(page.column_spacing_twips, `${path}/page/column_spacing_twips`, issues, 0)
    const columnLayout = enumValue(page.column_layout, `${path}/page/column_layout`, ['equal-width', 'explicit'], issues)
    const columns = array(page.column_definitions, `${path}/page/column_definitions`, issues, 45)
    if (columnCount !== undefined && columns.length !== columnCount) add(issues, 'INVALID_VALUE', `${path}/page/column_definitions`, 'must contain exactly one identity per declared column')
    columns.forEach((value, index) => {
      const columnPath = `${path}/page/column_definitions/${index}`
      const column = object(value, columnPath, DOCX_NATIVE_V1_BINDING_FIELDS.ColumnV1, issues)
      if (!column) return
      trackId(column.id, `${columnPath}/id`, issues, ids)
      const ordinal = integer(column.ordinal, `${columnPath}/ordinal`, issues, 0)
      if (ordinal !== undefined && ordinal > 44) add(issues, 'OUT_OF_RANGE', `${columnPath}/ordinal`, 'must be from 0 through 44')
      if (ordinal !== undefined && ordinal !== index) add(issues, 'INVALID_VALUE', `${columnPath}/ordinal`, 'must equal the source-order column index')
      if (columnLayout === 'equal-width') {
        if (column.width_twips !== undefined) add(issues, 'INVALID_VALUE', `${columnPath}/width_twips`, 'equal-width columns derive width from section geometry')
        if (column.space_after_twips !== undefined) add(issues, 'INVALID_VALUE', `${columnPath}/space_after_twips`, 'equal-width columns use column_spacing_twips')
      } else if (columnLayout === 'explicit') {
        twipsInteger(column.width_twips, `${columnPath}/width_twips`, issues, 1)
        twipsInteger(column.space_after_twips, `${columnPath}/space_after_twips`, issues, 0)
        if (index === columns.length - 1 && column.space_after_twips !== 0) add(issues, 'INVALID_VALUE', `${columnPath}/space_after_twips`, 'the final explicit column cannot have trailing inter-column space')
      }
    })
  }
  for (const key of ['header_refs', 'footer_refs'] as const) {
    array(entry[key], `${path}/${key}`, issues).forEach((refValue, index) => {
      const refPath = `${path}/${key}/${index}`
      const ref = object(refValue, refPath, DOCX_NATIVE_V1_BINDING_FIELDS.HeaderFooterReferenceV1, issues)
      if (!ref) return
      enumValue(ref.kind, `${refPath}/kind`, ['default', 'first', 'even'], issues)
      const storyId = stringValue(ref.story_id, `${refPath}/story_id`, issues, ID)
      if (storyId) reference(refs, storyId, `${refPath}/story_id`, key === 'header_refs' ? 'header-story' : 'footer-story')
      stringValue(ref.relationship_id, `${refPath}/relationship_id`, issues, ID)
    })
  }
}

function trackId(value: unknown, path: string, issues: NativeDocxValidationIssue[], ids: Set<string>): string | null {
  const id = stringValue(value, path, issues, ID)
  if (!id) return null
  if (ids.has(id)) add(issues, 'DUPLICATE_ID', path, `duplicate native id ${JSON.stringify(id)}`)
  ids.add(id)
  return id
}

interface PreflightState {
  nodes: number
  bounded: boolean
}

function preflightValue(value: unknown, path: string, depth: number, state: PreflightState, issues: NativeDocxValidationIssue[]): void {
  if (!state.bounded) return
  state.nodes += 1
  if (state.nodes > DOCX_NATIVE_LIMITS.maxNodes) {
    state.bounded = false
    add(issues, 'LIMIT_EXCEEDED', path, `document traversal exceeds ${DOCX_NATIVE_LIMITS.maxNodes} values`)
    return
  }
  if (depth > DOCX_NATIVE_LIMITS.maxDepth) {
    state.bounded = false
    add(issues, 'LIMIT_EXCEEDED', path, `document nesting exceeds ${DOCX_NATIVE_LIMITS.maxDepth} levels`)
    return
  }
  if (value === null) {
    add(issues, 'INVALID_VALUE', path, 'JSON null is not permitted anywhere in the native contract')
    return
  }
  if (typeof value === 'number' && Object.is(value, -0)) {
    add(issues, 'INVALID_VALUE', path, 'negative zero is not permitted')
    return
  }
  if (Array.isArray(value)) {
    if (value.length > DOCX_NATIVE_LIMITS.maxCollectionItems) add(issues, 'LIMIT_EXCEEDED', path, `must contain at most ${DOCX_NATIVE_LIMITS.maxCollectionItems} items`)
    for (let index = 0; index < Math.min(value.length, DOCX_NATIVE_LIMITS.maxCollectionItems); index += 1) {
      preflightValue(value[index], `${path}/${index}`, depth + 1, state, issues)
      if (!state.bounded) return
    }
    return
  }
  if (isObject(value)) {
    for (const key of Object.keys(value).sort()) {
      preflightValue(value[key], `${path}/${escapePointer(key)}`, depth + 1, state, issues)
      if (!state.bounded) return
    }
  }
}

/** Strictly decode a v1 model. Unknown fields and dangling identities fail. */
export function decodeNativeDocxDocument(value: unknown): DecodeNativeDocxResult {
  const issues: NativeDocxValidationIssue[] = []
  const preflight = { nodes: 0, bounded: true }
  preflightValue(value, '', 0, preflight, issues)
  if (!preflight.bounded) return { ok: false, issues }
  const root = object(value, '', DOCX_NATIVE_V1_BINDING_FIELDS.DocumentV1, issues)
  if (!root) return { ok: false, issues }
  if (root.protocol !== DOCX_NATIVE_PROTOCOL) add(issues, 'UNSUPPORTED_PROTOCOL', '/protocol', `must equal ${DOCX_NATIVE_PROTOCOL}`)
  if (root.version !== DOCX_NATIVE_VERSION) add(issues, 'UNSUPPORTED_VERSION', '/version', `must equal ${DOCX_NATIVE_VERSION}`)
  const ids = new Set<string>()
  const refs: PendingReference[] = []
  const bodyBlockIds = new Set<string>()
  const headerStoryIds = new Set<string>()
  const footerStoryIds = new Set<string>()
  const footnoteStoryIds = new Set<string>()
  const endnoteStoryIds = new Set<string>()
  const commentIds = new Set<string>()
  const commentStoryIds = new Set<string>()
  const commentStories = new Map<string, StoryValidationResult>()
  const noteNativeIds = new Set<string>()
  const noteRelationships = new Map<string, string>()
  const noteRelationshipByKind = new Map<string, string>()
  const notePartByKind = new Map<string, string>()
  const modeledParts = new Set<string>()
  trackId(root.document_id, '/document_id', issues, ids)
  stringValue(root.revision, '/revision', issues, ID)
  const source = object(root.source, '/source', DOCX_NATIVE_V1_BINDING_FIELDS.SourcePackageV1, issues)
  let mainPart: string | null = null
  if (source) {
    stringValue(source.package_sha256, '/source/package_sha256', issues, SHA256)
    mainPart = partNameValue(source.main_part, '/source/main_part', issues)
    if (mainPart) modeledParts.add(mainPart)
  }
  const body = validateStory(root.body, '/body', ['body'], issues, ids, refs, bodyBlockIds)
  if (body.partName) modeledParts.add(body.partName)
  if (body.partName && mainPart && body.partName !== mainPart) add(issues, 'INVALID_VALUE', '/body/part_name', 'body story part must equal source.main_part')
  array(root.sections, '/sections', issues).forEach((entry, index) => validateSection(entry, `/sections/${index}`, issues, ids, refs, mainPart, body.anchor))
  array(root.headers, '/headers', issues).forEach((entry, index) => {
    const story = validateStory(entry, `/headers/${index}`, ['header'], issues, ids, refs, bodyBlockIds)
    if (story.id && story.kind === 'header') headerStoryIds.add(story.id)
    if (story.partName) modeledParts.add(story.partName)
  })
  array(root.footers, '/footers', issues).forEach((entry, index) => {
    const story = validateStory(entry, `/footers/${index}`, ['footer'], issues, ids, refs, bodyBlockIds)
    if (story.id && story.kind === 'footer') footerStoryIds.add(story.id)
    if (story.partName) modeledParts.add(story.partName)
  })
  array(root.notes, '/notes', issues).forEach((entry, index) => {
    const story = validateStory(entry, `/notes/${index}`, ['footnote', 'endnote'], issues, ids, refs, bodyBlockIds)
    if (story.kind && story.nativeStoryId) {
      const key = `${story.kind}:${story.nativeStoryId}`
      if (noteNativeIds.has(key)) add(issues, 'DUPLICATE_ID', `/notes/${index}/native_story_id`, 'native note id is duplicated within its note kind')
      noteNativeIds.add(key)
    }
    if (story.kind && story.relationshipId) {
      const knownRelationship = noteRelationshipByKind.get(story.kind)
      if (knownRelationship !== undefined && knownRelationship !== story.relationshipId) add(issues, 'INVALID_VALUE', `/notes/${index}/relationship_id`, 'one note kind must resolve through exactly one relationship identity')
      noteRelationshipByKind.set(story.kind, story.relationshipId)
      const known = noteRelationships.get(story.relationshipId)
      if (known !== undefined && known !== story.kind) add(issues, 'DUPLICATE_ID', `/notes/${index}/relationship_id`, 'one relationship identity cannot resolve both footnotes and endnotes')
      noteRelationships.set(story.relationshipId, story.kind)
    }
    if (story.kind && story.partName) {
      const knownPart = notePartByKind.get(story.kind)
      if (knownPart !== undefined && knownPart !== story.partName) add(issues, 'INVALID_VALUE', `/notes/${index}/part_name`, 'one note kind must resolve to exactly one package part')
      notePartByKind.set(story.kind, story.partName)
    }
    if (story.id && story.kind === 'footnote') footnoteStoryIds.add(story.id)
    if (story.id && story.kind === 'endnote') endnoteStoryIds.add(story.id)
    if (story.partName) modeledParts.add(story.partName)
  })
  array(root.comment_stories, '/comment_stories', issues).forEach((entry, index) => {
    const story = validateStory(entry, `/comment_stories/${index}`, ['comment'], issues, ids, refs, bodyBlockIds)
    if (story.id && story.kind === 'comment') {
      commentStoryIds.add(story.id)
      commentStories.set(story.id, story)
    }
    if (story.partName) modeledParts.add(story.partName)
  })
  array(root.comments, '/comments', issues).forEach((entryValue, index) => {
    const path = `/comments/${index}`
    const entry = object(entryValue, path, DOCX_NATIVE_V1_BINDING_FIELDS.CommentV1, issues)
    if (!entry) return
    const id = trackId(entry.id, `${path}/id`, issues, ids)
    if (id) commentIds.add(id)
    stringValue(entry.native_comment_id, `${path}/native_comment_id`, issues, ID)
    stringValue(entry.author, `${path}/author`, issues)
    optionalString(entry.initials, `${path}/initials`, issues)
    optionalString(entry.created_at, `${path}/created_at`, issues)
    const commentAnchor = validateAnchor(entry.anchor, `${path}/anchor`, issues)
    const story = stringValue(entry.body_story_id, `${path}/body_story_id`, issues, ID)
    if (story) {
      reference(refs, story, `${path}/body_story_id`, 'comment-story')
      const bodyStory = commentStories.get(story)
      if (commentAnchor && bodyStory?.anchor && (bodyStory.anchor.partName !== commentAnchor.partName || bodyStory.anchor.start < commentAnchor.start || bodyStory.anchor.end > commentAnchor.end)) {
        add(issues, 'OUT_OF_RANGE', `${path}/body_story_id`, 'comment body story anchor must be contained by the owning comment anchor')
      }
      if (bodyStory?.nativeStoryId && typeof entry.native_comment_id === 'string' && bodyStory.nativeStoryId !== entry.native_comment_id) {
        add(issues, 'INVALID_VALUE', `${path}/body_story_id`, 'comment body story native_story_id must equal native_comment_id')
      }
    }
  })
  const capabilityNames = new Set<string>()
  array(root.capabilities, '/capabilities', issues).forEach((entryValue, index) => {
    const path = `/capabilities/${index}`
    const entry = object(entryValue, path, DOCX_NATIVE_V1_BINDING_FIELDS.CapabilityV1, issues)
    if (!entry) return
    const name = stringValue(entry.name, `${path}/name`, issues, ID)
    if (name && capabilityNames.has(name)) add(issues, 'INVALID_VALUE', `${path}/name`, 'capability name is duplicated')
    if (name) capabilityNames.add(name)
    enumValue(entry.level, `${path}/level`, ['read-write', 'read-only', 'passthrough', 'unsupported'], issues)
    optionalString(entry.detail, `${path}/detail`, issues)
  })
  const passthroughParts = new Set<string>()
  const mediaParts = new Set<string>()
  array(root.passthrough_parts, '/passthrough_parts', issues).forEach((entryValue, index) => {
    const path = `/passthrough_parts/${index}`
    const entry = object(entryValue, path, DOCX_NATIVE_V1_BINDING_FIELDS.PassthroughPartV1, issues)
    if (!entry) return
    const name = partNameValue(entry.part_name, `${path}/part_name`, issues)
    if (name && passthroughParts.has(name)) add(issues, 'INVALID_VALUE', `${path}/part_name`, 'passthrough part is duplicated')
    if (name && modeledParts.has(name)) add(issues, 'INVALID_VALUE', `${path}/part_name`, 'modeled story parts cannot also be passthrough parts')
    if (name) passthroughParts.add(name)
    const contentType = stringValue(entry.content_type, `${path}/content_type`, issues)
    if (name && contentType && /^(?:image|audio|video)\//.test(asciiLower(contentType))) mediaParts.add(name)
    integer(entry.byte_length, `${path}/byte_length`, issues, 0)
    stringValue(entry.sha256, `${path}/sha256`, issues, SHA256)
    enumValue(entry.policy, `${path}/policy`, ['preserve-verbatim'], issues)
  })
  array(root.unsupported, '/unsupported', issues).forEach((entryValue, index) => {
    const path = `/unsupported/${index}`
    const entry = object(entryValue, path, DOCX_NATIVE_V1_BINDING_FIELDS.UnsupportedCapabilityV1, issues)
    if (!entry) return
    trackId(entry.id, `${path}/id`, issues, ids)
    stringValue(entry.code, `${path}/code`, issues, ID)
    stringValue(entry.capability, `${path}/capability`, issues, ID)
    const scope = stringValue(entry.scope_id, `${path}/scope_id`, issues, ID)
    if (scope) reference(refs, scope, `${path}/scope_id`, 'modeled-id')
    if (entry.anchor !== undefined) validateAnchor(entry.anchor, `${path}/anchor`, issues)
    enumValue(entry.preservation, `${path}/preservation`, ['preserve-verbatim', 'refuse-mutation'], issues)
    stringValue(entry.message, `${path}/message`, issues)
  })
  const targets: Record<ReferenceTarget, Set<string>> = {
    'modeled-id': ids,
    'body-block': bodyBlockIds,
    'header-story': headerStoryIds,
    'footer-story': footerStoryIds,
    'footnote-story': footnoteStoryIds,
    'endnote-story': endnoteStoryIds,
    comment: commentIds,
    'comment-story': commentStoryIds,
    'media-part': mediaParts,
  }
  for (const pending of refs) {
    if (!targets[pending.target].has(pending.id)) add(issues, 'BROKEN_REFERENCE', pending.path, `reference ${JSON.stringify(pending.id)} does not resolve to a ${pending.target}`)
  }
  if (Array.isArray(root.sections) && root.sections.length === 0) add(issues, 'REQUIRED', '/sections', 'at least one section is required')
  issues.sort(compareNativeValidationIssues)
  return issues.length > 0 ? { ok: false, issues } : { ok: true, value: value as NativeDocxDocumentV1 }
}

/** Decode a bounded raw JSON payload before allocating the typed model. */
export function decodeNativeDocxJson(json: string): DecodeNativeDocxResult {
  if (json.length > DOCX_NATIVE_LIMITS.maxJsonBytes || new TextEncoder().encode(json).byteLength > DOCX_NATIVE_LIMITS.maxJsonBytes) {
    return { ok: false, issues: [{ code: 'LIMIT_EXCEEDED', path: '', message: `JSON payload exceeds ${DOCX_NATIVE_LIMITS.maxJsonBytes} bytes` }] }
  }
  try {
    return decodeNativeDocxDocument(JSON.parse(json) as unknown)
  } catch {
    return { ok: false, issues: [{ code: 'INVALID_VALUE', path: '', message: 'payload must be valid JSON' }] }
  }
}

function canonicalize(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(canonicalize)
  if (!isObject(value)) return value
  const ordered: JsonObject = {}
  for (const key of Object.keys(value).sort()) ordered[key] = canonicalize(value[key])
  return ordered
}

/** Validate and serialize with lexicographically ordered object keys. */
export function encodeNativeDocxDocument(value: NativeDocxDocumentV1): string {
  const decoded = decodeNativeDocxDocument(value)
  if (!decoded.ok) throw new NativeDocxValidationError(decoded.issues)
  const encoded = JSON.stringify(canonicalize(decoded.value))
  if (new TextEncoder().encode(encoded).byteLength > DOCX_NATIVE_LIMITS.maxJsonBytes) {
    throw new NativeDocxValidationError([{ code: 'LIMIT_EXCEEDED', path: '', message: `JSON payload exceeds ${DOCX_NATIVE_LIMITS.maxJsonBytes} bytes` }])
  }
  return encoded
}
