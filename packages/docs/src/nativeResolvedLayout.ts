/**
 * Strict TypeScript wire projection of Go docxpatch.NativeResolvedLayoutInputV1.
 *
 * The Go projection remains authoritative. This decoder rejects unknown fields,
 * unbounded collections, JSON null, negative zero, and dangling paragraph/run
 * references before shaping code observes the value.
 */

import { DOCX_MAX_TWIPS_FOR_MILLIPOINTS, DOCX_NATIVE_LIMITS, type NativeDocxIssueCode, type NativeDocxTableBordersV1, type NativeDocxValidationIssue } from './nativeContract.js'
import { compareNativeValidationIssues } from './nativeDeterminism.js'
import { asciiLower, hasAsciiEdgeWhitespace, normalizeFontFamilyName } from '@injoffice/font-metrics/layout'
import { sha256 } from '@noble/hashes/sha2.js'
import { bytesToHex } from '@noble/hashes/utils.js'

export const DOCX_RESOLVED_LAYOUT_PROTOCOL = 'injoffice.docx.resolved-layout'
export const DOCX_RESOLVED_LAYOUT_VERSION = 1 as const
export const DOCX_RESOLVED_LAYOUT_MAX_DIAGNOSTICS = 1_000

export interface NativeDocxResolvedSourcePartsV1 {
  main_part: string
  styles_part?: string
  numbering_part?: string
  theme_part?: string
  font_table_part?: string
}

export interface NativeDocxResolvedNumberingSourceV1 {
  relationships_part: string
  relationships_sha256: string
  relationship_id: string
  relationship_type: string
  relationship_target: string
  part_name: string
  content_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml'
  part_sha256: string
  model_sha256: string
}

export interface NativeDocxResolvedParagraphPropertiesV1 {
  alignment?: 'left' | 'right' | 'center' | 'both' | 'distribute' | 'start' | 'end'
  spacing_before_twips?: number
  spacing_after_twips?: number
  line?: number
  line_rule?: 'auto' | 'exact' | 'atLeast'
  indent_left_twips?: number
  indent_right_twips?: number
  indent_start_twips?: number
  indent_end_twips?: number
  first_line_twips?: number
  hanging_twips?: number
  bidi?: boolean
  keep_next?: boolean
  keep_lines?: boolean
  page_break_before?: boolean
  widow_control?: boolean
}

export interface NativeDocxResolvedRunPropertiesV1 {
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

export interface NativeDocxResolvedNumberingV1 {
  marker_id: string
  definition_sha256: string
  num_id: string
  abstract_num_id: string
  level: number
  level_style_id?: string
  start: number
  format: 'decimal' | 'lowerLetter' | 'upperLetter' | 'lowerRoman' | 'upperRoman' | 'bullet'
  text: string
  suffix: 'tab' | 'space' | 'nothing'
  alignment: 'left' | 'right' | 'center' | 'start' | 'end'
  restart_after_level?: number
  never_restart: boolean
  counter_value: number
  counter_values: Array<{ level: number; value: number; format: 'decimal' | 'lowerLetter' | 'upperLetter' | 'lowerRoman' | 'upperRoman' }>
  resolved_text: string
  label_start_twips: number
  label_end_twips: number
  text_start_twips: number
  numbering_tab_twips?: number
  marker_properties: NativeDocxResolvedRunPropertiesV1
}

export interface NativeDocxResolvedParagraphV1 {
  paragraph_id: string
  style_id?: string
  applied_styles: string[]
  properties: NativeDocxResolvedParagraphPropertiesV1
  paragraph_mark_properties: NativeDocxResolvedRunPropertiesV1
  numbering?: NativeDocxResolvedNumberingV1
}

export interface NativeDocxResolvedRunV1 {
  run_id: string
  paragraph_id: string
  character_style_id?: string
  applied_paragraph_styles: string[]
  applied_character_styles: string[]
  properties: NativeDocxResolvedRunPropertiesV1
}

export interface NativeDocxResolvedTableGeometryV1 {
  layout: 'fixed' | 'autofit'
  alignment: 'left'
  indent_twips: number
  width_type: 'auto' | 'dxa' | 'pct'
  width_value: number
  cell_margins: { top_twips: number; right_twips: number; bottom_twips: number; left_twips: number }
}

export interface NativeDocxResolvedTableV1 {
  table_id: string
  style_id?: string
  borders?: NativeDocxTableBordersV1
  cell_shading_rgb?: string
  geometry?: NativeDocxResolvedTableGeometryV1
}

export interface NativeDocxResolvedFontV1 {
  name: string
  alt_name?: string
}

export interface NativeDocxResolutionDiagnosticV1 {
  code: string
  severity: 'unsupported'
  scope_id: string
  part_name?: string
  path?: string
  preservation: 'preserve-verbatim'
  message: string
}

export interface NativeDocxResolvedLayoutInputV1 {
  protocol: typeof DOCX_RESOLVED_LAYOUT_PROTOCOL
  version: typeof DOCX_RESOLVED_LAYOUT_VERSION
  document_id: string
  revision: string
  source_parts: NativeDocxResolvedSourcePartsV1
  numbering_source?: NativeDocxResolvedNumberingSourceV1
  paragraphs: NativeDocxResolvedParagraphV1[]
  runs: NativeDocxResolvedRunV1[]
  tables: NativeDocxResolvedTableV1[]
  fonts: NativeDocxResolvedFontV1[]
  diagnostics: NativeDocxResolutionDiagnosticV1[]
}

export type DecodeNativeDocxResolvedLayoutResult =
  | { ok: true; value: NativeDocxResolvedLayoutInputV1 }
  | { ok: false; issues: NativeDocxValidationIssue[] }

/** Kept in field-for-field parity with the exported Go JSON structs. */
export const DOCX_RESOLVED_LAYOUT_V1_BINDING_FIELDS = {
  SourcePartsV1: ['main_part', 'styles_part', 'numbering_part', 'theme_part', 'font_table_part'],
  NumberingSourceV1: ['relationships_part', 'relationships_sha256', 'relationship_id', 'relationship_type', 'relationship_target', 'part_name', 'content_type', 'part_sha256', 'model_sha256'],
  ParagraphPropertiesV1: ['alignment', 'spacing_before_twips', 'spacing_after_twips', 'line', 'line_rule', 'indent_left_twips', 'indent_right_twips', 'indent_start_twips', 'indent_end_twips', 'first_line_twips', 'hanging_twips', 'bidi', 'keep_next', 'keep_lines', 'page_break_before', 'widow_control'],
  RunPropertiesV1: ['font_family', 'font_size_half_points', 'bold', 'italic', 'underline', 'vertical_alignment', 'color', 'highlight', 'language', 'rtl', 'hidden'],
  CounterValueV1: ['level', 'value', 'format'],
  NumberingV1: ['marker_id', 'definition_sha256', 'num_id', 'abstract_num_id', 'level', 'level_style_id', 'start', 'format', 'text', 'suffix', 'alignment', 'restart_after_level', 'never_restart', 'counter_value', 'counter_values', 'resolved_text', 'label_start_twips', 'label_end_twips', 'text_start_twips', 'numbering_tab_twips', 'marker_properties'],
  ParagraphV1: ['paragraph_id', 'style_id', 'applied_styles', 'properties', 'paragraph_mark_properties', 'numbering'],
  RunV1: ['run_id', 'paragraph_id', 'character_style_id', 'applied_paragraph_styles', 'applied_character_styles', 'properties'],
  TableBorderV1: ['style', 'size_eighth_points', 'color_rgb'],
  TableBordersV1: ['top', 'right', 'bottom', 'left', 'inside_horizontal', 'inside_vertical'],
  TableV1: ['table_id', 'style_id', 'borders', 'cell_shading_rgb', 'geometry'],
  FontV1: ['name', 'alt_name'],
  DiagnosticV1: ['code', 'severity', 'scope_id', 'part_name', 'path', 'preservation', 'message'],
  LayoutInputV1: ['protocol', 'version', 'document_id', 'revision', 'source_parts', 'numbering_source', 'paragraphs', 'runs', 'tables', 'fonts', 'diagnostics'],
} as const

type JsonObject = Record<string, unknown>
const ID = /^[A-Za-z0-9][A-Za-z0-9._:-]{0,255}$/
const DECIMAL_ID = /^(?:0|[1-9][0-9]{0,18})$/
const COLOR = /^[0-9A-F]{6}$/
const SHA256 = /^sha256:[0-9a-f]{64}$/
const ABSOLUTE_URI = /^[A-Za-z][A-Za-z0-9+.-]*:[^\u0000\r\n]*$/
const PART_SEGMENT = /^(?:[A-Za-z0-9._~!$&'()*+,;=@-]|%[0-9A-F]{2})+$/

function add(issues: NativeDocxValidationIssue[], code: NativeDocxIssueCode, path: string, message: string): void {
  if (issues.length >= DOCX_NATIVE_LIMITS.maxIssues) return
  if (!issues.some((issue) => issue.code === code && issue.path === path && issue.message === message)) issues.push({ code, path, message })
}

function isObject(value: unknown): value is JsonObject {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

function escapePointer(value: string): string {
  return value.replace(/~/g, '~0').replace(/\//g, '~1')
}

function object(value: unknown, path: string, fields: readonly string[], issues: NativeDocxValidationIssue[]): JsonObject | null {
  if (!isObject(value)) {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, value === undefined ? 'field is required' : 'must be an object')
    return null
  }
  const allowed = new Set(fields)
  for (const key of Object.keys(value).sort()) if (!allowed.has(key)) add(issues, 'UNKNOWN_FIELD', `${path}/${escapePointer(key)}`, `unknown field ${JSON.stringify(key)}`)
  return value
}

function array(value: unknown, path: string, issues: NativeDocxValidationIssue[], max: number = DOCX_NATIVE_LIMITS.maxCollectionItems): unknown[] {
  if (!Array.isArray(value)) {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, 'must be an array')
    return []
  }
  if (value.length > max) add(issues, 'LIMIT_EXCEEDED', path, `must contain at most ${max} items`)
  return value.slice(0, max)
}

function stringValue(value: unknown, path: string, issues: NativeDocxValidationIssue[], pattern: RegExp = ID, max = 4096): string | null {
  if (value === undefined) {
    add(issues, 'REQUIRED', path, 'field is required')
    return null
  }
  if (typeof value !== 'string') {
    add(issues, 'INVALID_TYPE', path, 'must be a string')
    return null
  }
  if (value.length === 0 || value.length > max || !pattern.test(value) || hasAsciiEdgeWhitespace(value) || /[\u0000\r\n]/.test(value)) {
    add(issues, 'INVALID_VALUE', path, 'contains an invalid string value')
    return null
  }
  return value
}

function optionalString(value: unknown, path: string, issues: NativeDocxValidationIssue[], pattern: RegExp = ID, max = 4096): string | undefined {
  if (value === undefined) return undefined
  return stringValue(value, path, issues, pattern, max) ?? undefined
}

function integer(value: unknown, path: string, issues: NativeDocxValidationIssue[], min = Number.MIN_SAFE_INTEGER, max = Number.MAX_SAFE_INTEGER, optional = true): number | undefined {
  if (value === undefined && optional) return undefined
  if (!Number.isSafeInteger(value) || Object.is(value, -0) || (value as number) < min || (value as number) > max) {
    add(issues, value === undefined ? 'REQUIRED' : 'OUT_OF_RANGE', path, `must be a safe integer from ${min} through ${max}`)
    return undefined
  }
  return value as number
}

function booleanValue(value: unknown, path: string, issues: NativeDocxValidationIssue[]): void {
  if (value !== undefined && typeof value !== 'boolean') add(issues, 'INVALID_TYPE', path, 'must be a boolean')
}

function enumValue(value: unknown, path: string, allowed: readonly string[], issues: NativeDocxValidationIssue[], optional = false): string | undefined {
  if (value === undefined && optional) return undefined
  if (typeof value !== 'string' || !allowed.includes(value)) {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_VALUE', path, `must be one of ${allowed.join(', ')}`)
    return undefined
  }
  return value
}

function isCanonicalPartName(value: string): boolean {
  if (value.length > 4096 || value.startsWith('/') || value.includes('\\') || value.includes('//')) return false
  return value.split('/').every((segment) => {
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

function canonicalPartKey(value: string): string {
  return asciiLower(value.split('/').map((segment) => decodeURIComponent(segment)).join('/'))
}

function relationshipsPartFor(owner: string): string {
  const segments = owner.split('/')
  const base = segments.pop()!
  return [...segments, '_rels', `${base}.rels`].join('/')
}

function resolveInternalRelationshipTarget(owner: string, target: string): string | undefined {
  if (target.length === 0 || target.startsWith('/') || hasAsciiEdgeWhitespace(target) || /[\\?#\u0000\r\n\t]/.test(target)) return undefined
  const segments = owner.split('/')
  segments.pop()
  for (const segment of target.split('/')) {
    if (segment === '') return undefined
    if (segment === '.') continue
    if (segment === '..') {
      if (segments.length === 0) return undefined
      segments.pop()
      continue
    }
    if (!PART_SEGMENT.test(segment)) return undefined
    segments.push(segment)
  }
  const resolved = segments.join('/')
  return isCanonicalPartName(resolved) ? resolved : undefined
}

function partName(value: unknown, path: string, issues: NativeDocxValidationIssue[]): string | null {
  const name = stringValue(value, path, issues, /[\s\S]+/)
  if (name && !isCanonicalPartName(name)) {
    add(issues, 'INVALID_VALUE', path, 'contains a non-canonical OPC part name')
    return null
  }
  return name
}

function styleChain(value: unknown, path: string, issues: NativeDocxValidationIssue[]): void {
  const entries = array(value, path, issues, DOCX_NATIVE_LIMITS.maxDepth)
  const seen = new Set<string>()
  entries.forEach((entry, index) => {
    const id = stringValue(entry, `${path}/${index}`, issues)
    if (id && seen.has(id)) add(issues, 'INVALID_VALUE', `${path}/${index}`, 'applied style is duplicated')
    if (id) seen.add(id)
  })
}

function validateRunProperties(value: unknown, path: string, issues: NativeDocxValidationIssue[]): void {
  const entry = object(value, path, DOCX_RESOLVED_LAYOUT_V1_BINDING_FIELDS.RunPropertiesV1, issues)
  if (!entry) return
  optionalString(entry.font_family, `${path}/font_family`, issues, /[\s\S]+/, 256)
  integer(entry.font_size_half_points, `${path}/font_size_half_points`, issues, 1, 3276)
  for (const key of ['bold', 'italic', 'rtl', 'hidden']) booleanValue(entry[key], `${path}/${key}`, issues)
  if (entry.underline !== undefined) enumValue(entry.underline, `${path}/underline`, ['none', 'single', 'double', 'words'], issues)
  if (entry.vertical_alignment !== undefined) enumValue(entry.vertical_alignment, `${path}/vertical_alignment`, ['baseline', 'subscript', 'superscript'], issues)
  optionalString(entry.color, `${path}/color`, issues, COLOR, 6)
  if (entry.highlight !== undefined) enumValue(entry.highlight, `${path}/highlight`, ['none', 'black', 'blue', 'cyan', 'green', 'magenta', 'red', 'yellow', 'white', 'darkBlue', 'darkCyan', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'darkGray', 'lightGray'], issues)
  optionalString(entry.language, `${path}/language`, issues, /[\s\S]+/, 256)
}

function validateParagraphProperties(value: unknown, path: string, issues: NativeDocxValidationIssue[]): void {
  const entry = object(value, path, DOCX_RESOLVED_LAYOUT_V1_BINDING_FIELDS.ParagraphPropertiesV1, issues)
  if (!entry) return
  if (entry.alignment !== undefined) enumValue(entry.alignment, `${path}/alignment`, ['left', 'right', 'center', 'both', 'distribute', 'start', 'end'], issues)
  for (const key of ['spacing_before_twips', 'spacing_after_twips', 'line', 'first_line_twips', 'hanging_twips']) integer(entry[key], `${path}/${key}`, issues, 0, DOCX_MAX_TWIPS_FOR_MILLIPOINTS)
  for (const key of ['indent_left_twips', 'indent_right_twips', 'indent_start_twips', 'indent_end_twips']) integer(entry[key], `${path}/${key}`, issues, -DOCX_MAX_TWIPS_FOR_MILLIPOINTS, DOCX_MAX_TWIPS_FOR_MILLIPOINTS)
  for (const key of ['bidi', 'keep_next', 'keep_lines', 'page_break_before', 'widow_control']) booleanValue(entry[key], `${path}/${key}`, issues)
  if (entry.line_rule !== undefined) enumValue(entry.line_rule, `${path}/line_rule`, ['auto', 'exact', 'atLeast'], issues)
  if ((entry.line === undefined) !== (entry.line_rule === undefined)) add(issues, 'INVALID_VALUE', path, 'line and line_rule must be supplied together')
  if (entry.first_line_twips !== undefined && entry.hanging_twips !== undefined) add(issues, 'INVALID_VALUE', path, 'first-line and hanging indents are mutually exclusive')
}

function validateNumbering(value: unknown, path: string, issues: NativeDocxValidationIssue[]): void {
  const entry = object(value, path, DOCX_RESOLVED_LAYOUT_V1_BINDING_FIELDS.NumberingV1, issues)
  if (!entry) return
  stringValue(entry.marker_id, `${path}/marker_id`, issues)
  stringValue(entry.definition_sha256, `${path}/definition_sha256`, issues, SHA256, 71)
  const num = stringValue(entry.num_id, `${path}/num_id`, issues, DECIMAL_ID, 19)
  if (num === '0') add(issues, 'INVALID_VALUE', `${path}/num_id`, 'num_id must be positive')
  if (num && BigInt(num) > 2_147_483_647n) add(issues, 'OUT_OF_RANGE', `${path}/num_id`, 'num_id must fit the Go v1 31-bit identifier bound')
  const abstract = stringValue(entry.abstract_num_id, `${path}/abstract_num_id`, issues, DECIMAL_ID, 19)
  if (abstract && BigInt(abstract) > 2_147_483_647n) add(issues, 'OUT_OF_RANGE', `${path}/abstract_num_id`, 'abstract_num_id must fit the Go v1 31-bit identifier bound')
  integer(entry.level, `${path}/level`, issues, 0, 8, false)
  optionalString(entry.level_style_id, `${path}/level_style_id`, issues)
  integer(entry.start, `${path}/start`, issues, 0, 2_147_483_647)
  enumValue(entry.format, `${path}/format`, ['decimal', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman', 'bullet'], issues)
  numberingText(entry.text, `${path}/text`, issues)
  enumValue(entry.suffix, `${path}/suffix`, ['tab', 'space', 'nothing'], issues)
  enumValue(entry.alignment, `${path}/alignment`, ['left', 'right', 'center', 'start', 'end'], issues)
  const level = Number.isSafeInteger(entry.level) ? entry.level as number : undefined
  if (entry.restart_after_level !== undefined) integer(entry.restart_after_level, `${path}/restart_after_level`, issues, 0, Math.max(0, (level ?? 0) - 1))
  if (typeof entry.never_restart !== 'boolean') add(issues, entry.never_restart === undefined ? 'REQUIRED' : 'INVALID_TYPE', `${path}/never_restart`, 'must be a boolean')
  if ((entry.restart_after_level !== undefined) === (entry.never_restart === true)) add(issues, 'INVALID_VALUE', path, 'restart_after_level and never_restart must encode exactly one restart policy')
  integer(entry.counter_value, `${path}/counter_value`, issues, 0, 2_147_483_647)
  const counterLevels = new Set<number>()
  array(entry.counter_values, `${path}/counter_values`, issues, 9).forEach((value, index) => {
    const counterPath = `${path}/counter_values/${index}`
    const counter = object(value, counterPath, DOCX_RESOLVED_LAYOUT_V1_BINDING_FIELDS.CounterValueV1, issues)
    if (!counter) return
    const counterLevel = integer(counter.level, `${counterPath}/level`, issues, 0, level ?? 8)
    if (counterLevel !== undefined && counterLevels.has(counterLevel)) add(issues, 'DUPLICATE_ID', `${counterPath}/level`, 'counter level is duplicated')
    if (counterLevel !== undefined) counterLevels.add(counterLevel)
    integer(counter.value, `${counterPath}/value`, issues, 0, 2_147_483_647)
    enumValue(counter.format, `${counterPath}/format`, ['decimal', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman'], issues)
  })
  const resolvedText = numberingText(entry.resolved_text, `${path}/resolved_text`, issues)
  const resolvedScalars = resolvedText === null ? undefined : unicodeScalarLength(resolvedText)
  if (resolvedScalars !== undefined && resolvedScalars > 31) add(issues, 'LIMIT_EXCEEDED', `${path}/resolved_text`, 'must contain at most 31 Unicode scalar values')
  const labelStart = integer(entry.label_start_twips, `${path}/label_start_twips`, issues, 0, DOCX_MAX_TWIPS_FOR_MILLIPOINTS)
  const labelEnd = integer(entry.label_end_twips, `${path}/label_end_twips`, issues, 1, DOCX_MAX_TWIPS_FOR_MILLIPOINTS)
  const textStart = integer(entry.text_start_twips, `${path}/text_start_twips`, issues, 1, DOCX_MAX_TWIPS_FOR_MILLIPOINTS)
  integer(entry.numbering_tab_twips, `${path}/numbering_tab_twips`, issues, 0, DOCX_MAX_TWIPS_FOR_MILLIPOINTS)
  if (labelStart !== undefined && labelEnd !== undefined && labelEnd <= labelStart) add(issues, 'INVALID_VALUE', `${path}/label_end_twips`, 'must follow label_start_twips')
  if (labelEnd !== undefined && textStart !== undefined && textStart !== labelEnd) add(issues, 'INVALID_VALUE', `${path}/text_start_twips`, 'must equal the hanging-indent label end')
  validateRunProperties(entry.marker_properties, `${path}/marker_properties`, issues)
}

function numberingText(value: unknown, path: string, issues: NativeDocxValidationIssue[]): string | null {
  if (typeof value !== 'string') {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, 'must be a string')
    return null
  }
  if (value.length === 0 || value.length > 1024 || unicodeScalarLength(value) === undefined || new TextEncoder().encode(value).length > 1024 || /[\u0000\r\n]/.test(value)) {
    add(issues, 'INVALID_VALUE', path, 'contains invalid or unbounded numbering text')
    return null
  }
  return value
}

function unicodeScalarLength(value: string): number | undefined {
  let length = 0
  for (let index = 0; index < value.length; index += 1) {
    const unit = value.charCodeAt(index)
    if (unit >= 0xd800 && unit <= 0xdbff) {
      const next = value.charCodeAt(index + 1)
      if (index + 1 >= value.length || next < 0xdc00 || next > 0xdfff) return undefined
      index += 1
    } else if (unit >= 0xdc00 && unit <= 0xdfff) return undefined
    length += 1
  }
  return length
}

function goOrderedRunProperties(properties: NativeDocxResolvedRunPropertiesV1): Record<string, unknown> {
  return Object.fromEntries((['font_family', 'font_size_half_points', 'bold', 'italic', 'underline', 'vertical_alignment', 'color', 'highlight', 'language', 'rtl', 'hidden'] as const)
    .flatMap((key) => properties[key] === undefined ? [] : [[key, properties[key]]]))
}

const NUMBERING_CANONICAL_SERIALIZER_V1 = 'injoffice.canonical-json/utf8-v1'

function canonicalNumberingValue(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(canonicalNumberingValue)
  if (isObject(value)) return Object.fromEntries(Object.keys(value).sort().map((key) => [key, canonicalNumberingValue(value[key])]))
  return value
}

function nativeDocxNumberingCanonicalSha256V1(value: unknown): string {
  const json = JSON.stringify(canonicalNumberingValue(value)).replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029')
  const encoded = new TextEncoder().encode(`${NUMBERING_CANONICAL_SERIALIZER_V1}\u0000${json}`)
  return `sha256:${bytesToHex(sha256(encoded))}`
}

/** Recomputes the canonical Go/TypeScript digest over source-ordered markers. */
export function nativeDocxResolvedNumberingModelSha256V1(paragraphs: readonly NativeDocxResolvedParagraphV1[], source?: Omit<NativeDocxResolvedNumberingSourceV1, 'model_sha256'> | NativeDocxResolvedNumberingSourceV1): string {
  const markers = paragraphs.flatMap((paragraph) => paragraph.numbering ? [paragraph.numbering] : []).map((marker) => ({
    marker_id: marker.marker_id,
    definition_sha256: marker.definition_sha256,
    num_id: marker.num_id,
    abstract_num_id: marker.abstract_num_id,
    level: marker.level,
    ...(marker.level_style_id !== undefined ? { level_style_id: marker.level_style_id } : {}),
    start: marker.start,
    format: marker.format,
    text: marker.text,
    suffix: marker.suffix,
    alignment: marker.alignment,
    ...(marker.restart_after_level !== undefined ? { restart_after_level: marker.restart_after_level } : {}),
    never_restart: marker.never_restart,
    counter_value: marker.counter_value,
    counter_values: marker.counter_values.map((counter) => ({ level: counter.level, value: counter.value, format: counter.format })),
    resolved_text: marker.resolved_text,
    label_start_twips: marker.label_start_twips,
    label_end_twips: marker.label_end_twips,
    text_start_twips: marker.text_start_twips,
    ...(marker.numbering_tab_twips !== undefined ? { numbering_tab_twips: marker.numbering_tab_twips } : {}),
    marker_properties: goOrderedRunProperties(marker.marker_properties),
  }))
  return nativeDocxNumberingCanonicalSha256V1({
    protocol: 'injoffice.docx.numbering-model/v1',
    source: {
      relationships_part: source?.relationships_part ?? '', relationships_sha256: source?.relationships_sha256 ?? '',
      relationship_id: source?.relationship_id ?? '', relationship_type: source?.relationship_type ?? '', relationship_target: source?.relationship_target ?? '',
      part_name: source?.part_name ?? '', content_type: source?.content_type ?? '', part_sha256: source?.part_sha256 ?? '',
    },
    markers,
  })
}

/** Recomputes the canonical definition digest that binds exported marker semantics to the exact numbering part. */
export function nativeDocxResolvedNumberingDefinitionSha256V1(marker: NativeDocxResolvedNumberingV1, partSha256: string): string {
  return nativeDocxNumberingCanonicalSha256V1({
    protocol: 'injoffice.docx.numbering-definition/v1',
    part_sha256: partSha256,
    num_id: marker.num_id,
    abstract_num_id: marker.abstract_num_id,
    level: marker.level,
    ...(marker.level_style_id !== undefined ? { level_style_id: marker.level_style_id } : {}),
    start: marker.start,
    format: marker.format,
    text: marker.text,
    suffix: marker.suffix,
    alignment: marker.alignment,
    ...(marker.restart_after_level !== undefined ? { restart_after_level: marker.restart_after_level } : {}),
    never_restart: marker.never_restart,
    ...(marker.numbering_tab_twips !== undefined ? { numbering_tab_twips: marker.numbering_tab_twips } : {}),
    label_start_twips: marker.label_start_twips,
    label_end_twips: marker.label_end_twips,
    text_start_twips: marker.text_start_twips,
  })
}

interface PreflightState { nodes: number; bounded: boolean }

function preflight(value: unknown, path: string, depth: number, state: PreflightState, issues: NativeDocxValidationIssue[]): void {
  if (!state.bounded) return
  state.nodes += 1
  if (state.nodes > DOCX_NATIVE_LIMITS.maxNodes) {
    state.bounded = false
    add(issues, 'LIMIT_EXCEEDED', path, `resolved layout traversal exceeds ${DOCX_NATIVE_LIMITS.maxNodes} values`)
    return
  }
  if (depth > DOCX_NATIVE_LIMITS.maxDepth) {
    state.bounded = false
    add(issues, 'LIMIT_EXCEEDED', path, `resolved layout nesting exceeds ${DOCX_NATIVE_LIMITS.maxDepth} levels`)
    return
  }
  if (value === null) {
    add(issues, 'INVALID_VALUE', path, 'JSON null is not permitted')
    return
  }
  if (typeof value === 'number' && Object.is(value, -0)) {
    add(issues, 'INVALID_VALUE', path, 'negative zero is not permitted')
    return
  }
  if (Array.isArray(value)) {
    if (value.length > DOCX_NATIVE_LIMITS.maxCollectionItems) add(issues, 'LIMIT_EXCEEDED', path, `must contain at most ${DOCX_NATIVE_LIMITS.maxCollectionItems} items`)
    value.slice(0, DOCX_NATIVE_LIMITS.maxCollectionItems).forEach((entry, index) => preflight(entry, `${path}/${index}`, depth + 1, state, issues))
  } else if (isObject(value)) {
    for (const key of Object.keys(value).sort()) preflight(value[key], `${path}/${escapePointer(key)}`, depth + 1, state, issues)
  }
}

export function decodeNativeDocxResolvedLayout(value: unknown): DecodeNativeDocxResolvedLayoutResult {
  const issues: NativeDocxValidationIssue[] = []
  const state = { nodes: 0, bounded: true }
  preflight(value, '', 0, state, issues)
  if (!state.bounded) return { ok: false, issues }
  const root = object(value, '', DOCX_RESOLVED_LAYOUT_V1_BINDING_FIELDS.LayoutInputV1, issues)
  if (!root) return { ok: false, issues }
  if (root.protocol !== DOCX_RESOLVED_LAYOUT_PROTOCOL) add(issues, 'UNSUPPORTED_PROTOCOL', '/protocol', `must equal ${DOCX_RESOLVED_LAYOUT_PROTOCOL}`)
  if (root.version !== DOCX_RESOLVED_LAYOUT_VERSION) add(issues, 'UNSUPPORTED_VERSION', '/version', `must equal ${DOCX_RESOLVED_LAYOUT_VERSION}`)
  const documentID = stringValue(root.document_id, '/document_id', issues)
  stringValue(root.revision, '/revision', issues)
  const parts = object(root.source_parts, '/source_parts', DOCX_RESOLVED_LAYOUT_V1_BINDING_FIELDS.SourcePartsV1, issues)
  if (parts) {
    const seen = new Set<string>()
    for (const key of DOCX_RESOLVED_LAYOUT_V1_BINDING_FIELDS.SourcePartsV1) {
      if (key !== 'main_part' && parts[key] === undefined) continue
      const name = partName(parts[key], `/source_parts/${key}`, issues)
      const folded = name ? canonicalPartKey(name) : undefined
      if (folded && seen.has(folded)) add(issues, 'INVALID_VALUE', `/source_parts/${key}`, 'source part is duplicated case-insensitively')
      if (folded) seen.add(folded)
    }
  }
  const numberingSource = root.numbering_source === undefined ? undefined : object(root.numbering_source, '/numbering_source', DOCX_RESOLVED_LAYOUT_V1_BINDING_FIELDS.NumberingSourceV1, issues)
  if (numberingSource) {
    const relationshipsPart = partName(numberingSource.relationships_part, '/numbering_source/relationships_part', issues)
    stringValue(numberingSource.relationships_sha256, '/numbering_source/relationships_sha256', issues, SHA256, 71)
    stringValue(numberingSource.relationship_id, '/numbering_source/relationship_id', issues)
    const relationshipType = stringValue(numberingSource.relationship_type, '/numbering_source/relationship_type', issues, ABSOLUTE_URI, 4096)
    if (relationshipType !== undefined && relationshipType !== 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering' && relationshipType !== 'http://purl.oclc.org/ooxml/officeDocument/relationships/numbering') add(issues, 'INVALID_VALUE', '/numbering_source/relationship_type', 'must be the Strict or Transitional numbering relationship type')
    stringValue(numberingSource.relationship_target, '/numbering_source/relationship_target', issues, /[^\u0000\r\n]+/, 4096)
    const numberingPart = partName(numberingSource.part_name, '/numbering_source/part_name', issues)
    if (numberingSource.content_type !== 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml') add(issues, 'INVALID_VALUE', '/numbering_source/content_type', 'must be the WordprocessingML numbering content type')
    stringValue(numberingSource.part_sha256, '/numbering_source/part_sha256', issues, SHA256, 71)
    stringValue(numberingSource.model_sha256, '/numbering_source/model_sha256', issues, SHA256, 71)
    const sourceNumberingPart = parts && typeof parts.numbering_part === 'string' ? parts.numbering_part : undefined
    if (!sourceNumberingPart || !numberingPart || canonicalPartKey(sourceNumberingPart) !== canonicalPartKey(numberingPart)) add(issues, 'BROKEN_REFERENCE', '/numbering_source/part_name', 'must match source_parts.numbering_part')
    if (relationshipsPart && numberingPart && canonicalPartKey(relationshipsPart) === canonicalPartKey(numberingPart)) add(issues, 'INVALID_VALUE', '/numbering_source/relationships_part', 'relationship and numbering parts must be distinct')
    const mainPart = parts && typeof parts.main_part === 'string' && isCanonicalPartName(parts.main_part) ? parts.main_part : undefined
    if (mainPart && relationshipsPart && canonicalPartKey(relationshipsPart) !== canonicalPartKey(relationshipsPartFor(mainPart))) add(issues, 'BROKEN_REFERENCE', '/numbering_source/relationships_part', 'must be the canonical relationship part owned by source_parts.main_part')
    const target = typeof numberingSource.relationship_target === 'string' && mainPart ? resolveInternalRelationshipTarget(mainPart, numberingSource.relationship_target) : undefined
    if (!target || !numberingPart || canonicalPartKey(target) !== canonicalPartKey(numberingPart)) add(issues, 'BROKEN_REFERENCE', '/numbering_source/relationship_target', 'must be an internal target resolving from source_parts.main_part to the numbering part')
  } else if (parts && parts.numbering_part !== undefined) {
    add(issues, 'REQUIRED', '/numbering_source', 'a numbering part requires exact relationship and digest provenance')
  }

  const paragraphIDs = new Set<string>()
  const runIDs = new Set<string>()
  const tableIDs = new Set<string>()
  array(root.paragraphs, '/paragraphs', issues).forEach((value, index) => {
    const path = `/paragraphs/${index}`
    const entry = object(value, path, DOCX_RESOLVED_LAYOUT_V1_BINDING_FIELDS.ParagraphV1, issues)
    if (!entry) return
    const id = stringValue(entry.paragraph_id, `${path}/paragraph_id`, issues)
    if (id && paragraphIDs.has(id)) add(issues, 'DUPLICATE_ID', `${path}/paragraph_id`, 'paragraph id is duplicated')
    if (id) paragraphIDs.add(id)
    optionalString(entry.style_id, `${path}/style_id`, issues)
    styleChain(entry.applied_styles, `${path}/applied_styles`, issues)
    validateParagraphProperties(entry.properties, `${path}/properties`, issues)
    validateRunProperties(entry.paragraph_mark_properties, `${path}/paragraph_mark_properties`, issues)
    if (entry.numbering !== undefined) {
      validateNumbering(entry.numbering, `${path}/numbering`, issues)
      if (isObject(entry.numbering) && entry.numbering.level_style_id !== undefined && (!Array.isArray(entry.applied_styles) || !entry.applied_styles.includes(entry.numbering.level_style_id))) add(issues, 'BROKEN_REFERENCE', `${path}/numbering/level_style_id`, 'must attest the paragraph-style cascade layer that selected the abstract numbering level')
    }
  })
  array(root.runs, '/runs', issues).forEach((value, index) => {
    const path = `/runs/${index}`
    const entry = object(value, path, DOCX_RESOLVED_LAYOUT_V1_BINDING_FIELDS.RunV1, issues)
    if (!entry) return
    const id = stringValue(entry.run_id, `${path}/run_id`, issues)
    if (id && runIDs.has(id)) add(issues, 'DUPLICATE_ID', `${path}/run_id`, 'run id is duplicated')
    if (id) runIDs.add(id)
    const paragraphID = stringValue(entry.paragraph_id, `${path}/paragraph_id`, issues)
    if (paragraphID && !paragraphIDs.has(paragraphID)) add(issues, 'BROKEN_REFERENCE', `${path}/paragraph_id`, 'must reference a resolved paragraph')
    optionalString(entry.character_style_id, `${path}/character_style_id`, issues)
    styleChain(entry.applied_paragraph_styles, `${path}/applied_paragraph_styles`, issues)
    styleChain(entry.applied_character_styles, `${path}/applied_character_styles`, issues)
    validateRunProperties(entry.properties, `${path}/properties`, issues)
  })
  array(root.tables, '/tables', issues).forEach((value, index) => {
    const path = `/tables/${index}`
    const entry = object(value, path, DOCX_RESOLVED_LAYOUT_V1_BINDING_FIELDS.TableV1, issues)
    if (!entry) return
    const id = stringValue(entry.table_id, `${path}/table_id`, issues)
    if (id && tableIDs.has(id)) add(issues, 'DUPLICATE_ID', `${path}/table_id`, 'table id is duplicated')
    if (id) tableIDs.add(id)
    optionalString(entry.style_id, `${path}/style_id`, issues)
    optionalString(entry.cell_shading_rgb, `${path}/cell_shading_rgb`, issues, COLOR, 6)
    if (entry.geometry !== undefined) {
      const geometryPath = `${path}/geometry`
      const geometry = object(entry.geometry, geometryPath, ['layout', 'alignment', 'indent_twips', 'width_type', 'width_value', 'cell_margins'], issues)
      if (geometry) {
        enumValue(geometry.layout, `${geometryPath}/layout`, ['fixed', 'autofit'], issues)
        enumValue(geometry.alignment, `${geometryPath}/alignment`, ['left'], issues)
        enumValue(geometry.width_type, `${geometryPath}/width_type`, ['auto', 'dxa', 'pct'], issues)
        integer(geometry.indent_twips, `${geometryPath}/indent_twips`, issues, 0, DOCX_MAX_TWIPS_FOR_MILLIPOINTS, false)
        integer(geometry.width_value, `${geometryPath}/width_value`, issues, geometry.width_type === 'auto' ? 0 : 1, geometry.width_type === 'auto' ? 0 : geometry.width_type === 'pct' ? 5000 : DOCX_MAX_TWIPS_FOR_MILLIPOINTS, false)
        const margins = object(geometry.cell_margins, `${geometryPath}/cell_margins`, ['top_twips', 'right_twips', 'bottom_twips', 'left_twips'], issues)
        if (margins) for (const side of ['top_twips', 'right_twips', 'bottom_twips', 'left_twips']) integer(margins[side], `${geometryPath}/cell_margins/${side}`, issues, 0, DOCX_MAX_TWIPS_FOR_MILLIPOINTS, false)
      }
    }
    if (entry.borders !== undefined) {
      const borders = object(entry.borders, `${path}/borders`, DOCX_RESOLVED_LAYOUT_V1_BINDING_FIELDS.TableBordersV1, issues)
      if (borders) for (const key of DOCX_RESOLVED_LAYOUT_V1_BINDING_FIELDS.TableBordersV1) if (borders[key] !== undefined) {
        const borderPath = `${path}/borders/${key}`
        const border = object(borders[key], borderPath, DOCX_RESOLVED_LAYOUT_V1_BINDING_FIELDS.TableBorderV1, issues)
        if (!border) continue
        enumValue(border.style, `${borderPath}/style`, ['none', 'single'], issues)
        integer(border.size_eighth_points, `${borderPath}/size_eighth_points`, issues, 0, 768, false)
        optionalString(border.color_rgb, `${borderPath}/color_rgb`, issues, COLOR, 6)
      }
    }
  })
  const fontNames = new Set<string>()
  array(root.fonts, '/fonts', issues).forEach((value, index) => {
    const path = `/fonts/${index}`
    const entry = object(value, path, DOCX_RESOLVED_LAYOUT_V1_BINDING_FIELDS.FontV1, issues)
    if (!entry) return
    const name = stringValue(entry.name, `${path}/name`, issues, /[\s\S]+/, 256)
    const folded = name === null ? undefined : normalizeFontFamilyName(name)
    if (folded && fontNames.has(folded)) add(issues, 'DUPLICATE_ID', `${path}/name`, 'font name is duplicated case-insensitively')
    if (folded) fontNames.add(folded)
    optionalString(entry.alt_name, `${path}/alt_name`, issues, /[\s\S]+/, 256)
  })
  array(root.diagnostics, '/diagnostics', issues, DOCX_RESOLVED_LAYOUT_MAX_DIAGNOSTICS).forEach((value, index) => {
    const path = `/diagnostics/${index}`
    const entry = object(value, path, DOCX_RESOLVED_LAYOUT_V1_BINDING_FIELDS.DiagnosticV1, issues)
    if (!entry) return
    stringValue(entry.code, `${path}/code`, issues)
    enumValue(entry.severity, `${path}/severity`, ['unsupported'], issues)
    const scope = stringValue(entry.scope_id, `${path}/scope_id`, issues)
    if (scope && scope !== documentID && !paragraphIDs.has(scope) && !runIDs.has(scope) && !tableIDs.has(scope)) add(issues, 'BROKEN_REFERENCE', `${path}/scope_id`, 'diagnostic scope is not modeled')
    if (entry.part_name !== undefined) partName(entry.part_name, `${path}/part_name`, issues)
    if (entry.path !== undefined) optionalString(entry.path, `${path}/path`, issues, /[\s\S]+/, 8192)
    if (entry.path !== undefined && entry.part_name === undefined) add(issues, 'REQUIRED', `${path}/part_name`, 'path requires part_name')
    enumValue(entry.preservation, `${path}/preservation`, ['preserve-verbatim'], issues)
    stringValue(entry.message, `${path}/message`, issues, /[\s\S]+/, DOCX_NATIVE_LIMITS.maxTextLength)
  })
  if (numberingSource && issues.length === 0) {
    for (const [index, paragraph] of (root.paragraphs as NativeDocxResolvedParagraphV1[]).entries()) {
      if (paragraph.numbering && paragraph.numbering.definition_sha256 !== nativeDocxResolvedNumberingDefinitionSha256V1(paragraph.numbering, numberingSource.part_sha256 as string)) add(issues, 'BROKEN_REFERENCE', `/paragraphs/${index}/numbering/definition_sha256`, 'must attest the exported numbering definition and exact numbering part hash')
    }
    const expected = nativeDocxResolvedNumberingModelSha256V1(root.paragraphs as NativeDocxResolvedParagraphV1[], numberingSource as unknown as NativeDocxResolvedNumberingSourceV1)
    if (numberingSource.model_sha256 !== expected) add(issues, 'BROKEN_REFERENCE', '/numbering_source/model_sha256', 'must equal the canonical source-ordered resolved marker model digest')
  }
  issues.sort(compareNativeValidationIssues)
  return issues.length > 0 ? { ok: false, issues } : { ok: true, value: value as NativeDocxResolvedLayoutInputV1 }
}

export function decodeNativeDocxResolvedLayoutJson(json: string): DecodeNativeDocxResolvedLayoutResult {
  if (new TextEncoder().encode(json).byteLength > DOCX_NATIVE_LIMITS.maxJsonBytes) return { ok: false, issues: [{ code: 'LIMIT_EXCEEDED', path: '', message: `JSON payload exceeds ${DOCX_NATIVE_LIMITS.maxJsonBytes} bytes` }] }
  try {
    return decodeNativeDocxResolvedLayout(JSON.parse(json) as unknown)
  } catch {
    return { ok: false, issues: [{ code: 'INVALID_VALUE', path: '', message: 'payload must be valid JSON' }] }
  }
}
