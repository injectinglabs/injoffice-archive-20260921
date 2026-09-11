/** Strict decoder for the renderer-neutral DOCX shaped-lines v1 wire output. */

import { DOCX_NATIVE_LIMITS, type NativeDocxIssueCode, type NativeDocxValidationIssue } from './nativeContract.js'
import { BIDI_UNICODE_VERSION, NATIVE_BIDI_PROVIDER_ID, NATIVE_BIDI_PROVIDER_REVISION, reorderNativeBidiLineV1 } from '@injoffice/font-metrics/bidi'
import { UNICODE_13_CLASSIFIER_REVISION } from '@injoffice/font-metrics/unicode13'
import {
  DOCX_SHAPED_LINES_LIMITS,
  DOCX_SHAPED_LINES_PROTOCOL,
  DOCX_SHAPED_LINES_VERSION,
  type NativeDocxLineFragmentV1,
  type NativeDocxPositionedGlyphV1,
  type NativeDocxShapedLineV1,
  type NativeDocxShapedLinesV1,
  type NativeDocxShapedParagraphV1,
} from './nativeShapingLines.js'
import { compareNativeValidationIssues } from './nativeDeterminism.js'
import { validateNativeDocxScriptTransformV1 } from './nativeScriptLayoutV1.js'

export type DecodeNativeDocxShapedLinesResult =
  | { ok: true; value: NativeDocxShapedLinesV1 }
  | { ok: false; issues: NativeDocxValidationIssue[] }

export const DOCX_SHAPED_LINES_V1_BINDING_FIELDS = {
  FontManifestV1: ['manifest_id', 'revision'],
  ProvidersV1: ['resolver_id', 'resolver_revision', 'shaper_id', 'shaper_revision', 'bidi_id', 'bidi_revision', 'bidi_unicode_version', 'unicode13_revision'],
  GlyphV1: ['glyph_id', 'advance_x_millipoints', 'advance_y_millipoints', 'offset_x_millipoints', 'offset_y_millipoints'],
  FragmentV1: ['id', 'source_kind', 'source_id', 'start_utf16', 'end_utf16', 'text', 'direction', 'bidi_level', 'logical_order', 'script', 'language', 'face_id', 'whitespace', 'advance_inline_millipoints', 'justification_expansion_millipoints', 'ascent_millipoints', 'descent_millipoints', 'line_gap_millipoints', 'underline_position_millipoints', 'underline_thickness_millipoints', 'glyphs', 'script_transform'],
  HardBreakV1: ['source_run_id', 'control'],
  LineV1: ['id', 'ordinal', 'available_width_millipoints', 'inline_offset_millipoints', 'exclusion_start_millipoints', 'advance_inline_millipoints', 'ascent_millipoints', 'descent_millipoints', 'line_gap_millipoints', 'line_height_millipoints', 'justified', 'logical_to_visual', 'fragments', 'hard_break_after'],
  NumberingSourceV1: ['relationships_part', 'relationships_sha256', 'relationship_id', 'relationship_type', 'relationship_target', 'part_name', 'content_type', 'part_sha256', 'model_sha256'],
  ListMarkerV1: ['marker_id', 'definition_sha256', 'numbering_part_sha256', 'model_sha256', 'num_id', 'abstract_num_id', 'level', 'counter_value', 'text', 'suffix', 'alignment', 'label_start_millipoints', 'label_end_millipoints', 'marker_start_millipoints', 'marker_advance_millipoints', 'text_start_millipoints'],
  ParagraphV1: ['paragraph_id', 'story_id', 'story_kind', 'direction', 'alignment', 'spacing_before_millipoints', 'spacing_after_millipoints', 'indent_start_millipoints', 'indent_end_millipoints', 'first_line_delta_millipoints', 'list_marker', 'block_advance_millipoints', 'lines'],
  DiagnosticV1: ['code', 'severity', 'scope_id', 'source_id', 'source_diagnostic_code', 'source_diagnostic_message', 'message'],
  ShapedLinesV1: ['protocol', 'version', 'document_id', 'revision', 'numbering_source', 'available_width_millipoints', 'tab_interval_millipoints', 'font_manifest', 'providers', 'paragraphs', 'diagnostics'],
} as const

type JsonObject = Record<string, unknown>
const ID = /^[A-Za-z0-9][A-Za-z0-9._:/-]{0,1023}$/
const SHORT_ID = /^[A-Za-z0-9][A-Za-z0-9._:/-]{0,255}$/
const TAG = /^[\x20-\x7e]{4}$/
const SHA256 = /^sha256:[0-9a-f]{64}$/
const ABSOLUTE_URI = /^[A-Za-z][A-Za-z0-9+.-]*:[^\u0000\r\n]{1,4090}$/
const PART_NAME = /^(?!\/)(?!.*(?:^|\/)\.\.?\/)(?!.*\\)[^\u0000\r\n?#]+$/
const MAX_METRIC = 1_000_000_000
const MAX_BLOCK = 1_000_000_000_000
const MAX_MESSAGE = 4_096
const MAX_FRAGMENT_TEXT_UTF16 = 262_144
const BIDI_TRAILING_RE = /^[\u0009-\u000d\u001c-\u001e\u0020\u0085\u2028\u2029]+$/u

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
  for (const key of Object.keys(value).sort()) if (!fields.includes(key)) add(issues, 'UNKNOWN_FIELD', `${path}/${escapePointer(key)}`, `unknown field ${JSON.stringify(key)}`)
  return value
}

function array(value: unknown, path: string, issues: NativeDocxValidationIssue[], max: number): unknown[] {
  if (!Array.isArray(value)) {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, 'must be an array')
    return []
  }
  if (value.length > max) add(issues, 'LIMIT_EXCEEDED', path, `must contain at most ${max} items`)
  return value.slice(0, max)
}

function stringValue(value: unknown, path: string, issues: NativeDocxValidationIssue[], options: { pattern?: RegExp; max?: number; allowEmpty?: boolean } = {}): string | null {
  if (typeof value !== 'string') {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, 'must be a string')
    return null
  }
  const max = options.max ?? 1_024
  if ((!options.allowEmpty && value.length === 0) || value.length > max || (options.pattern && !options.pattern.test(value)) || /[\u0000\r\n]/.test(value)) {
    add(issues, 'INVALID_VALUE', path, 'contains an invalid or unbounded string value')
    return null
  }
  return value
}

function optionalString(value: unknown, path: string, issues: NativeDocxValidationIssue[], max = 1_024): string | undefined {
  if (value === undefined) return undefined
  return stringValue(value, path, issues, { max }) ?? undefined
}

function integer(value: unknown, path: string, issues: NativeDocxValidationIssue[], min: number, max: number, optional = false): number | undefined {
  if (value === undefined && optional) return undefined
  if (!Number.isSafeInteger(value) || Object.is(value, -0) || (value as number) < min || (value as number) > max) {
    add(issues, value === undefined ? 'REQUIRED' : 'OUT_OF_RANGE', path, `must be a safe integer from ${min} through ${max}`)
    return undefined
  }
  return value as number
}

function enumValue(value: unknown, path: string, allowed: readonly string[], issues: NativeDocxValidationIssue[]): string | undefined {
  if (typeof value !== 'string' || !allowed.includes(value)) {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_VALUE', path, `must be one of ${allowed.join(', ')}`)
    return undefined
  }
  return value
}

function validateNumberingSource(value: unknown, path: string, issues: NativeDocxValidationIssue[]): void {
  const entry = object(value, path, DOCX_SHAPED_LINES_V1_BINDING_FIELDS.NumberingSourceV1, issues)
  if (!entry) return
  stringValue(entry.relationships_part, `${path}/relationships_part`, issues, { pattern: PART_NAME, max: 4096 })
  stringValue(entry.relationships_sha256, `${path}/relationships_sha256`, issues, { pattern: SHA256, max: 71 })
  stringValue(entry.relationship_id, `${path}/relationship_id`, issues, { pattern: SHORT_ID, max: 256 })
  const relationshipType = stringValue(entry.relationship_type, `${path}/relationship_type`, issues, { pattern: ABSOLUTE_URI, max: 4096 })
  if (relationshipType && relationshipType !== 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering' && relationshipType !== 'http://purl.oclc.org/ooxml/officeDocument/relationships/numbering') add(issues, 'INVALID_VALUE', `${path}/relationship_type`, 'must be the Strict or Transitional numbering relationship type')
  stringValue(entry.relationship_target, `${path}/relationship_target`, issues, { max: 4096 })
  stringValue(entry.part_name, `${path}/part_name`, issues, { pattern: PART_NAME, max: 4096 })
  if (entry.content_type !== 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml') add(issues, 'INVALID_VALUE', `${path}/content_type`, 'must be the WordprocessingML numbering content type')
  stringValue(entry.part_sha256, `${path}/part_sha256`, issues, { pattern: SHA256, max: 71 })
  stringValue(entry.model_sha256, `${path}/model_sha256`, issues, { pattern: SHA256, max: 71 })
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

function validateListMarker(value: unknown, path: string, issues: NativeDocxValidationIssue[]): void {
  const entry = object(value, path, DOCX_SHAPED_LINES_V1_BINDING_FIELDS.ListMarkerV1, issues)
  if (!entry) return
  stringValue(entry.marker_id, `${path}/marker_id`, issues, { pattern: ID, max: 1024 })
  for (const key of ['definition_sha256', 'numbering_part_sha256', 'model_sha256']) stringValue(entry[key], `${path}/${key}`, issues, { pattern: SHA256, max: 71 })
  stringValue(entry.num_id, `${path}/num_id`, issues, { pattern: SHORT_ID, max: 256 })
  stringValue(entry.abstract_num_id, `${path}/abstract_num_id`, issues, { pattern: SHORT_ID, max: 256 })
  integer(entry.level, `${path}/level`, issues, 0, 8)
  integer(entry.counter_value, `${path}/counter_value`, issues, 0, 2_147_483_647)
  const text = stringValue(entry.text, `${path}/text`, issues, { allowEmpty: false, max: 1024 })
  const scalarLength = text === null ? undefined : unicodeScalarLength(text)
  if (text !== null && (scalarLength === undefined || scalarLength > 31)) add(issues, 'LIMIT_EXCEEDED', `${path}/text`, 'must contain at most 31 valid Unicode scalar values')
  enumValue(entry.suffix, `${path}/suffix`, ['tab', 'space', 'nothing'], issues)
  enumValue(entry.alignment, `${path}/alignment`, ['left', 'right', 'center', 'start', 'end'], issues)
  const labelStart = integer(entry.label_start_millipoints, `${path}/label_start_millipoints`, issues, 0, MAX_METRIC)
  const labelEnd = integer(entry.label_end_millipoints, `${path}/label_end_millipoints`, issues, 1, MAX_METRIC)
  const markerStart = integer(entry.marker_start_millipoints, `${path}/marker_start_millipoints`, issues, 0, MAX_METRIC)
  const markerAdvance = integer(entry.marker_advance_millipoints, `${path}/marker_advance_millipoints`, issues, 1, MAX_METRIC)
  const textStart = integer(entry.text_start_millipoints, `${path}/text_start_millipoints`, issues, 1, MAX_METRIC)
  if (labelStart !== undefined && labelEnd !== undefined && labelEnd <= labelStart) add(issues, 'INVALID_VALUE', `${path}/label_end_millipoints`, 'must follow the label start')
  if (markerStart !== undefined && markerAdvance !== undefined && labelEnd !== undefined && markerStart + markerAdvance > labelEnd) add(issues, 'OUT_OF_RANGE', `${path}/marker_advance_millipoints`, 'shaped marker must fit the attested label region')
  if (markerStart !== undefined && markerAdvance !== undefined && textStart !== undefined && textStart < markerStart + markerAdvance) add(issues, 'OUT_OF_RANGE', `${path}/text_start_millipoints`, 'body text must not overlap the shaped marker')
}

interface PreflightState { nodes: number; bounded: boolean }

function preflight(value: unknown, path: string, depth: number, state: PreflightState, issues: NativeDocxValidationIssue[]): void {
  if (!state.bounded) return
  state.nodes += 1
  if (state.nodes > DOCX_NATIVE_LIMITS.maxNodes) {
    state.bounded = false
    add(issues, 'LIMIT_EXCEEDED', path, `shaped-lines traversal exceeds ${DOCX_NATIVE_LIMITS.maxNodes} values`)
    return
  }
  if (depth > DOCX_NATIVE_LIMITS.maxDepth) {
    state.bounded = false
    add(issues, 'LIMIT_EXCEEDED', path, `shaped-lines nesting exceeds ${DOCX_NATIVE_LIMITS.maxDepth} levels`)
    return
  }
  if (value === null) {
    add(issues, 'INVALID_VALUE', path, 'JSON null is not permitted')
    return
  }
  if (typeof value === 'number' && Object.is(value, -0)) add(issues, 'INVALID_VALUE', path, 'negative zero is not permitted')
  if (Array.isArray(value)) {
    if (value.length > DOCX_SHAPED_LINES_LIMITS.maxFragments) add(issues, 'LIMIT_EXCEEDED', path, `array exceeds ${DOCX_SHAPED_LINES_LIMITS.maxFragments} items`)
    value.slice(0, DOCX_SHAPED_LINES_LIMITS.maxFragments).forEach((entry, index) => preflight(entry, `${path}/${index}`, depth + 1, state, issues))
  } else if (isObject(value)) {
    for (const key of Object.keys(value).sort()) preflight(value[key], `${path}/${escapePointer(key)}`, depth + 1, state, issues)
  } else if (!['string', 'number', 'boolean'].includes(typeof value)) {
    add(issues, 'INVALID_TYPE', path, 'must contain only JSON wire values')
  }
}

function validateGlyph(value: unknown, path: string, issues: NativeDocxValidationIssue[]): void {
  const entry = object(value, path, DOCX_SHAPED_LINES_V1_BINDING_FIELDS.GlyphV1, issues)
  if (!entry) return
  integer(entry.glyph_id, `${path}/glyph_id`, issues, 0, 0xffffffff)
  for (const key of ['advance_x_millipoints', 'advance_y_millipoints', 'offset_x_millipoints', 'offset_y_millipoints']) integer(entry[key], `${path}/${key}`, issues, -MAX_METRIC, MAX_METRIC)
}

interface ShapedDecodeState {
  paragraphs: number
  lines: number
  fragments: number
  glyphs: number
  fragmentIDs: Set<string>
}

function validateFragment(value: unknown, path: string, issues: NativeDocxValidationIssue[], state: ShapedDecodeState, expectedID: string): void {
  const entry = object(value, path, DOCX_SHAPED_LINES_V1_BINDING_FIELDS.FragmentV1, issues)
  if (!entry) return
  state.fragments += 1
  if (state.fragments > DOCX_SHAPED_LINES_LIMITS.maxFragments) add(issues, 'LIMIT_EXCEEDED', path, `fragments exceed ${DOCX_SHAPED_LINES_LIMITS.maxFragments}`)
  const id = stringValue(entry.id, `${path}/id`, issues, { pattern: ID, max: 1024 })
  if (id && id !== expectedID) add(issues, 'INVALID_VALUE', `${path}/id`, 'must equal the deterministic paragraph/line/fragment id')
  if (id && state.fragmentIDs.has(id)) add(issues, 'DUPLICATE_ID', `${path}/id`, 'fragment id is duplicated')
  if (id) state.fragmentIDs.add(id)
  const sourceKind = enumValue(entry.source_kind, `${path}/source_kind`, ['run', 'list-marker', 'tab', 'image'], issues)
  if (entry.script_transform !== undefined && (sourceKind !== 'run' || !validateNativeDocxScriptTransformV1(entry.script_transform))) add(issues, 'INVALID_VALUE', `${path}/script_transform`, 'requires bounded font-authored subscript/superscript metrics on a text run')
  stringValue(entry.source_id, `${path}/source_id`, issues, { pattern: SHORT_ID, max: 256 })
  const start = integer(entry.start_utf16, `${path}/start_utf16`, issues, 0, MAX_FRAGMENT_TEXT_UTF16)
  const end = integer(entry.end_utf16, `${path}/end_utf16`, issues, 0, MAX_FRAGMENT_TEXT_UTF16)
  if (start !== undefined && end !== undefined && end < start) add(issues, 'INVALID_VALUE', `${path}/end_utf16`, 'must not precede start_utf16')
  stringValue(entry.text, `${path}/text`, issues, { max: MAX_FRAGMENT_TEXT_UTF16, allowEmpty: true })
  enumValue(entry.direction, `${path}/direction`, ['ltr', 'rtl'], issues)
  const bidiLevel = integer(entry.bidi_level, `${path}/bidi_level`, issues, 0, 125)
  const logicalOrder = integer(entry.logical_order, `${path}/logical_order`, issues, 0, MAX_FRAGMENT_TEXT_UTF16)
  if (bidiLevel !== undefined && entry.direction !== (bidiLevel % 2 === 0 ? 'ltr' : 'rtl')) add(issues, 'INVALID_VALUE', `${path}/direction`, 'must equal the resolved bidi-level parity')
  stringValue(entry.script, `${path}/script`, issues, { pattern: TAG, max: 4 })
  stringValue(entry.language, `${path}/language`, issues, { max: 256 })
  if (entry.face_id !== undefined) stringValue(entry.face_id, `${path}/face_id`, issues, { pattern: SHORT_ID, max: 256 })
  if (typeof entry.whitespace !== 'boolean') add(issues, entry.whitespace === undefined ? 'REQUIRED' : 'INVALID_TYPE', `${path}/whitespace`, 'must be a boolean')
  const advance = integer(entry.advance_inline_millipoints, `${path}/advance_inline_millipoints`, issues, 0, MAX_METRIC)
  const expansion = integer(entry.justification_expansion_millipoints, `${path}/justification_expansion_millipoints`, issues, 0, MAX_METRIC)
  if (advance !== undefined && expansion !== undefined && expansion > advance) add(issues, 'OUT_OF_RANGE', `${path}/justification_expansion_millipoints`, 'must not exceed the fragment advance')
  integer(entry.ascent_millipoints, `${path}/ascent_millipoints`, issues, 0, MAX_METRIC)
  integer(entry.descent_millipoints, `${path}/descent_millipoints`, issues, -MAX_METRIC, 0)
  integer(entry.line_gap_millipoints, `${path}/line_gap_millipoints`, issues, 0, MAX_METRIC)
  if (entry.underline_position_millipoints !== undefined || entry.underline_thickness_millipoints !== undefined) {
    integer(entry.underline_position_millipoints, `${path}/underline_position_millipoints`, issues, -MAX_METRIC, MAX_METRIC)
    integer(entry.underline_thickness_millipoints, `${path}/underline_thickness_millipoints`, issues, 0, MAX_METRIC)
    if (entry.face_id === undefined || sourceKind === 'image') add(issues, 'BROKEN_REFERENCE', path, 'underline metrics require a resolved font face')
  }
  const glyphs = array(entry.glyphs, `${path}/glyphs`, issues, MAX_FRAGMENT_TEXT_UTF16)
  if (sourceKind === 'image' && (entry.text !== '' || entry.start_utf16 !== 0 || entry.end_utf16 !== 0 || entry.face_id !== undefined || entry.whitespace !== false || glyphs.length !== 0)) add(issues, 'INVALID_VALUE', path, 'image fragment must be glyphless, face-less, non-whitespace, and carry an empty UTF-16 range')
  state.glyphs += glyphs.length
  if (state.glyphs > DOCX_SHAPED_LINES_LIMITS.maxFragments) add(issues, 'LIMIT_EXCEEDED', `${path}/glyphs`, `glyphs exceed ${DOCX_SHAPED_LINES_LIMITS.maxFragments}`)
  glyphs.forEach((glyph, index) => validateGlyph(glyph, `${path}/glyphs/${index}`, issues))
  if (expansion !== undefined && expansion > 0 && (entry.text !== ' ' || entry.whitespace !== true || (entry.source_kind !== 'run' && entry.source_kind !== 'list-marker') || start === undefined || end !== start + 1 || glyphs.length === 0)) add(issues, 'INVALID_VALUE', `${path}/justification_expansion_millipoints`, 'positive expansion requires one paintable authored U+0020 cluster')
  let glyphAdvance = 0
  let completeGlyphAdvance = glyphs.length > 0
  for (const glyph of glyphs) {
    if (!isObject(glyph) || !Number.isSafeInteger(glyph.advance_x_millipoints)) completeGlyphAdvance = false
    else glyphAdvance += glyph.advance_x_millipoints as number
  }
  if (completeGlyphAdvance && advance !== undefined && glyphAdvance !== advance) add(issues, 'INVALID_VALUE', `${path}/advance_inline_millipoints`, 'must equal the sum of glyph inline advances')
  void logicalOrder
}

function canonicalAlignmentOffset(alignment: string, direction: string, available: number, advance: number): number {
  const remaining = Math.max(0, available - advance)
  if (alignment === 'center') return Math.round(remaining / 2)
  if (alignment === 'left' || alignment === 'both' || alignment === 'distribute') return 0
  if (alignment === 'right') return remaining
  if (alignment === 'start') return direction === 'ltr' ? 0 : remaining
  return direction === 'ltr' ? remaining : 0
}

function validateLine(value: unknown, path: string, issues: NativeDocxValidationIssue[], paragraphID: string | null, paragraphDirection: string | undefined, paragraphAlignment: string | undefined, paragraphIndentStart: number | undefined, paragraphIndentEnd: number | undefined, firstLineDelta: number | undefined, expectedOrdinal: number, state: ShapedDecodeState): number | null {
  const entry = object(value, path, DOCX_SHAPED_LINES_V1_BINDING_FIELDS.LineV1, issues)
  if (!entry) return null
  state.lines += 1
  if (state.lines > DOCX_SHAPED_LINES_LIMITS.maxLines) add(issues, 'LIMIT_EXCEEDED', path, `lines exceed ${DOCX_SHAPED_LINES_LIMITS.maxLines}`)
  const id = stringValue(entry.id, `${path}/id`, issues, { pattern: ID, max: 1024 })
  const ordinal = integer(entry.ordinal, `${path}/ordinal`, issues, 0, DOCX_SHAPED_LINES_LIMITS.maxLines)
  if (ordinal !== undefined && ordinal !== expectedOrdinal) add(issues, 'INVALID_VALUE', `${path}/ordinal`, `must equal source-order ordinal ${expectedOrdinal}`)
  if (id && paragraphID && id !== `line:${paragraphID}:${expectedOrdinal}`) add(issues, 'INVALID_VALUE', `${path}/id`, 'must equal the deterministic paragraph/ordinal line id')
  const available = integer(entry.available_width_millipoints, `${path}/available_width_millipoints`, issues, 0, DOCX_SHAPED_LINES_LIMITS.maxWidthMilliPoints)
  const inlineOffset = integer(entry.inline_offset_millipoints, `${path}/inline_offset_millipoints`, issues, -MAX_METRIC, MAX_METRIC)
  const lineAdvance = integer(entry.advance_inline_millipoints, `${path}/advance_inline_millipoints`, issues, 0, MAX_METRIC)
  integer(entry.ascent_millipoints, `${path}/ascent_millipoints`, issues, 0, MAX_METRIC)
  integer(entry.descent_millipoints, `${path}/descent_millipoints`, issues, -MAX_METRIC, 0)
  integer(entry.line_gap_millipoints, `${path}/line_gap_millipoints`, issues, 0, MAX_METRIC)
  const height = integer(entry.line_height_millipoints, `${path}/line_height_millipoints`, issues, 1, MAX_METRIC)
  if (typeof entry.justified !== 'boolean') add(issues, entry.justified === undefined ? 'REQUIRED' : 'INVALID_TYPE', `${path}/justified`, 'must be a boolean')
  const fragments = array(entry.fragments, `${path}/fragments`, issues, DOCX_SHAPED_LINES_LIMITS.maxFragments)
  fragments.forEach((fragment, index) => validateFragment(fragment, `${path}/fragments/${index}`, issues, state, `fragment:${paragraphID ?? 'invalid'}:${expectedOrdinal}:${index}`))
  const mapping = array(entry.logical_to_visual, `${path}/logical_to_visual`, issues, DOCX_SHAPED_LINES_LIMITS.maxFragments)
  if (mapping.length !== fragments.length) add(issues, 'INVALID_VALUE', `${path}/logical_to_visual`, 'must contain one visual index per fragment')
  const seenVisual = new Set<number>()
  mapping.forEach((value, logical) => {
    const visual = integer(value, `${path}/logical_to_visual/${logical}`, issues, 0, Math.max(0, fragments.length - 1))
    if (visual !== undefined) {
      if (seenVisual.has(visual)) add(issues, 'INVALID_VALUE', `${path}/logical_to_visual/${logical}`, 'visual index must be a permutation')
      seenVisual.add(visual)
      const fragment = fragments[visual]
      if (isObject(fragment) && fragment.logical_order !== logical) add(issues, 'INVALID_VALUE', `${path}/fragments/${visual}/logical_order`, 'must invert logical_to_visual')
    }
  })
  let logicalFragments: JsonObject[] | undefined
  let canonicalVisualToLogical: readonly number[] | undefined
  if (seenVisual.size === fragments.length && fragments.every(isObject)) {
    logicalFragments = mapping.map((visual) => fragments[visual as number] as JsonObject)
    const reordered = reorderNativeBidiLineV1(logicalFragments.map((fragment) => fragment.bidi_level as number), paragraphDirection === 'rtl' ? 1 : 0, logicalFragments.map((fragment) => typeof fragment.text === 'string' && BIDI_TRAILING_RE.test(fragment.text)))
    if (!reordered.ok || reordered.value.logicalToVisual.some((visual, logical) => mapping[logical] !== visual)) add(issues, 'INVALID_VALUE', `${path}/logical_to_visual`, 'must equal UAX #9 L1/L2 ordering for the supplied cluster levels')
    else canonicalVisualToLogical = reordered.value.visualToLogical
  }
  let fragmentAdvance = 0
  let expansion = 0
  for (const fragment of fragments) if (isObject(fragment) && Number.isSafeInteger(fragment.advance_inline_millipoints) && Number.isSafeInteger(fragment.justification_expansion_millipoints)) {
    fragmentAdvance += fragment.advance_inline_millipoints as number
    expansion += fragment.justification_expansion_millipoints as number
  }
  if (fragments.length > 0 && lineAdvance !== undefined && fragmentAdvance !== lineAdvance) add(issues, 'INVALID_VALUE', `${path}/advance_inline_millipoints`, 'must equal the sum of visual fragment advances')
  if (entry.justified === false && expansion !== 0) add(issues, 'INVALID_VALUE', `${path}/justified`, 'must be true when a fragment has justification expansion')
  if (entry.justified === true && available !== undefined && lineAdvance !== available) add(issues, 'INVALID_VALUE', `${path}/advance_inline_millipoints`, 'a justified line must exactly fill its available width')
  if (entry.justified === true && paragraphAlignment !== 'both') add(issues, 'INVALID_VALUE', `${path}/justified`, 'only paragraph alignment both may emit a justified line')
  if (entry.justified === true && expansion > 0 && logicalFragments && canonicalVisualToLogical) {
    const firstContent = logicalFragments.findIndex((fragment) => fragment.whitespace !== true)
    let lastContent = -1
    for (let index = logicalFragments.length - 1; index >= 0; index--) if (logicalFragments[index]!.whitespace !== true) { lastContent = index; break }
    const opportunities = canonicalVisualToLogical.filter((logical) => {
      const fragment = logicalFragments![logical]!
      return logical > firstContent && logical < lastContent && fragment.text === ' ' && fragment.source_kind !== 'tab' && Array.isArray(fragment.glyphs) && fragment.glyphs.length > 0
    })
    if (opportunities.length === 0) add(issues, 'INVALID_VALUE', `${path}/justified`, 'positive justification requires an interior paintable U+0020 opportunity')
    else {
      const expected = new Array<number>(logicalFragments.length).fill(0)
      const quotient = Math.floor(expansion / opportunities.length)
      let remainder = expansion % opportunities.length
      for (const logical of opportunities) expected[logical] = quotient + (remainder-- > 0 ? 1 : 0)
      for (let logical = 0; logical < logicalFragments.length; logical++) if (logicalFragments[logical]!.justification_expansion_millipoints !== expected[logical]) add(issues, 'INVALID_VALUE', `${path}/fragments/${mapping[logical]}/justification_expansion_millipoints`, 'must equal the canonical visual-order U+0020 quotient/remainder assignment')
    }
  }
  if (inlineOffset !== undefined && available !== undefined && lineAdvance !== undefined && paragraphDirection && paragraphAlignment && paragraphIndentStart !== undefined && paragraphIndentEnd !== undefined && firstLineDelta !== undefined && typeof entry.justified === 'boolean') {
    let base = paragraphDirection === 'ltr' ? paragraphIndentStart + (expectedOrdinal === 0 ? firstLineDelta : 0) : paragraphIndentEnd
    if (entry.exclusion_start_millipoints !== undefined) {
      const exclusion = integer(entry.exclusion_start_millipoints, `${path}/exclusion_start_millipoints`, issues, 0, MAX_METRIC)
      if (exclusion !== undefined) {
        if (exclusion < base || paragraphDirection !== 'ltr' || !['left', 'start'].includes(paragraphAlignment)) add(issues, 'INVALID_VALUE', `${path}/exclusion_start_millipoints`, 'must be a nonnegative left-aligned LTR exclusion after paragraph indentation')
        base = exclusion
      }
    }
    const alignment = paragraphAlignment === 'both' && entry.justified === false ? 'start' : paragraphAlignment
    const expectedOffset = base + canonicalAlignmentOffset(alignment, paragraphDirection, available, lineAdvance)
    if (inlineOffset !== expectedOffset) add(issues, 'INVALID_VALUE', `${path}/inline_offset_millipoints`, 'must equal the canonical physical/logical alignment offset for paragraph direction and indents')
  }
  if (entry.hard_break_after !== undefined) {
    const hardBreak = object(entry.hard_break_after, `${path}/hard_break_after`, DOCX_SHAPED_LINES_V1_BINDING_FIELDS.HardBreakV1, issues)
    if (hardBreak) {
      stringValue(hardBreak.source_run_id, `${path}/hard_break_after/source_run_id`, issues, { pattern: SHORT_ID, max: 256 })
      enumValue(hardBreak.control, `${path}/hard_break_after/control`, ['line-break'], issues)
    }
  }
  return height ?? null
}

function validateParagraph(value: unknown, path: string, issues: NativeDocxValidationIssue[], state: ShapedDecodeState, paragraphIDs: Set<string>, lineIDs: Set<string>): void {
  const entry = object(value, path, DOCX_SHAPED_LINES_V1_BINDING_FIELDS.ParagraphV1, issues)
  if (!entry) return
  state.paragraphs += 1
  const paragraphID = stringValue(entry.paragraph_id, `${path}/paragraph_id`, issues, { pattern: SHORT_ID, max: 256 })
  if (paragraphID && paragraphIDs.has(paragraphID)) add(issues, 'DUPLICATE_ID', `${path}/paragraph_id`, 'shaped paragraph id is duplicated')
  if (paragraphID) paragraphIDs.add(paragraphID)
  stringValue(entry.story_id, `${path}/story_id`, issues, { pattern: SHORT_ID, max: 256 })
  enumValue(entry.story_kind, `${path}/story_kind`, ['body', 'header', 'footer', 'footnote', 'endnote', 'comment'], issues)
  const direction = enumValue(entry.direction, `${path}/direction`, ['ltr', 'rtl'], issues)
  const alignment = enumValue(entry.alignment, `${path}/alignment`, ['left', 'right', 'center', 'both', 'distribute', 'start', 'end'], issues)
  const before = integer(entry.spacing_before_millipoints, `${path}/spacing_before_millipoints`, issues, 0, MAX_METRIC)
  const after = integer(entry.spacing_after_millipoints, `${path}/spacing_after_millipoints`, issues, 0, MAX_METRIC)
  const indentStart = integer(entry.indent_start_millipoints, `${path}/indent_start_millipoints`, issues, -MAX_METRIC, MAX_METRIC)
  const indentEnd = integer(entry.indent_end_millipoints, `${path}/indent_end_millipoints`, issues, -MAX_METRIC, MAX_METRIC)
  const firstDelta = integer(entry.first_line_delta_millipoints, `${path}/first_line_delta_millipoints`, issues, -MAX_METRIC, MAX_METRIC)
  if (entry.list_marker !== undefined) validateListMarker(entry.list_marker, `${path}/list_marker`, issues)
  const block = integer(entry.block_advance_millipoints, `${path}/block_advance_millipoints`, issues, 1, MAX_BLOCK)
  const lines = array(entry.lines, `${path}/lines`, issues, DOCX_SHAPED_LINES_LIMITS.maxLines)
  if (lines.length === 0) add(issues, 'INVALID_VALUE', `${path}/lines`, 'a shaped paragraph must contain at least one line')
  let lineHeightSum = 0
  lines.forEach((line, index) => {
    const lineObject = isObject(line) ? line : null
    if (lineObject?.exclusion_start_millipoints !== undefined && entry.story_kind !== 'body') add(issues, 'INVALID_VALUE', `${path}/lines/${index}/exclusion_start_millipoints`, 'square-wrap exclusions are supported only in the body story')
    const lineID = lineObject ? stringValue(lineObject.id, `${path}/lines/${index}/id`, [], { pattern: ID, max: 1024 }) : null
    if (lineID && lineIDs.has(lineID)) add(issues, 'DUPLICATE_ID', `${path}/lines/${index}/id`, 'line id is duplicated')
    if (lineID) lineIDs.add(lineID)
    const height = validateLine(line, `${path}/lines/${index}`, issues, paragraphID, direction, alignment, indentStart, indentEnd, firstDelta, index, state)
    if (height !== null && Number.isSafeInteger(lineHeightSum + height)) lineHeightSum += height
    else if (height !== null) add(issues, 'OUT_OF_RANGE', `${path}/lines/${index}/line_height_millipoints`, 'line-height sum exceeds safe integer range')
  })
  const markerFragments = lines.flatMap((line, lineIndex) => isObject(line) && Array.isArray(line.fragments)
    ? line.fragments.map((fragment, fragmentIndex) => ({ fragment, lineIndex, fragmentIndex }))
    : []).filter(({ fragment }) => isObject(fragment) && fragment.source_kind === 'list-marker')
  if (entry.list_marker === undefined && markerFragments.length > 0) add(issues, 'BROKEN_REFERENCE', `${path}/lines`, 'list-marker fragments require paragraph marker provenance')
  if (entry.list_marker !== undefined) {
    if (markerFragments.length === 0) add(issues, 'BROKEN_REFERENCE', `${path}/list_marker`, 'marker provenance requires shaped list-marker fragments')
    for (const { fragment, lineIndex, fragmentIndex } of markerFragments) {
      if (!isObject(fragment)) continue
      const fragmentPath = `${path}/lines/${lineIndex}/fragments/${fragmentIndex}`
      if (lineIndex !== 0) add(issues, 'BROKEN_REFERENCE', `${fragmentPath}/source_kind`, 'list-marker fragments may appear only on the first shaped line')
      if (paragraphID && fragment.source_id !== paragraphID) add(issues, 'BROKEN_REFERENCE', `${fragmentPath}/source_id`, 'list-marker fragment must reference its owning paragraph')
    }
    if (isObject(entry.list_marker) && isObject(lines[0])) {
      const markerPrefix = Array.isArray(lines[0].fragments) ? lines[0].fragments.filter((fragment) => isObject(fragment) && fragment.source_kind === 'list-marker') : []
      if (Array.isArray(lines[0].fragments) && !lines[0].fragments.slice(0, markerPrefix.length).every((fragment) => isObject(fragment) && fragment.source_kind === 'list-marker')) add(issues, 'BROKEN_REFERENCE', `${path}/lines/0/fragments`, 'list-marker fragments must form one contiguous first-line prefix')
      const advance = markerPrefix.reduce((sum, fragment) => sum + (isObject(fragment) && Number.isSafeInteger(fragment.advance_inline_millipoints) ? fragment.advance_inline_millipoints as number : 0), 0)
      if (Number.isSafeInteger(lines[0].inline_offset_millipoints) && Number.isSafeInteger(entry.list_marker.text_start_millipoints) && advance !== (entry.list_marker.text_start_millipoints as number) - (lines[0].inline_offset_millipoints as number)) {
        add(issues, 'BROKEN_REFERENCE', `${path}/list_marker/text_start_millipoints`, 'must equal the first-line origin plus the exact shaped marker-prefix advance')
      }
      const markerText = typeof entry.list_marker.text === 'string' ? entry.list_marker.text : undefined
      const suffix = entry.list_marker.suffix
      if (markerText !== undefined) {
        let cursor = 0
        let visibleStarted = false
        let suffixSeen = false
        let prefixAdvance = 0
        for (const [fragmentIndex, fragment] of markerPrefix.entries()) {
          if (!isObject(fragment)) continue
          const fragmentPath = `${path}/lines/0/fragments/${fragmentIndex}`
          const glyphless = Array.isArray(fragment.glyphs) && fragment.glyphs.length === 0
          const virtualPrefix = fragment.text === '' && fragment.start_utf16 === 0 && fragment.end_utf16 === 0 && glyphless
          const tabSuffix = suffix === 'tab' && fragment.text === '\t' && fragment.start_utf16 === 0 && fragment.end_utf16 === 0 && glyphless
          const spaceSuffix = suffix === 'space' && fragment.text === ' ' && fragment.start_utf16 === markerText.length && fragment.end_utf16 === markerText.length + 1
          if (virtualPrefix && !visibleStarted && !suffixSeen) {
            prefixAdvance += Number.isSafeInteger(fragment.advance_inline_millipoints) ? fragment.advance_inline_millipoints as number : 0
            continue
          }
          if ((tabSuffix || spaceSuffix) && cursor === markerText.length && !suffixSeen) {
            suffixSeen = true
            continue
          }
          const authored = !suffixSeen
            && Number.isSafeInteger(fragment.start_utf16) && fragment.start_utf16 === cursor
            && Number.isSafeInteger(fragment.end_utf16) && (fragment.end_utf16 as number) > cursor && (fragment.end_utf16 as number) <= markerText.length
            && markerText.slice(cursor, fragment.end_utf16 as number) === fragment.text
          if (!authored) {
            add(issues, 'BROKEN_REFERENCE', fragmentPath, 'list-marker fragments must exactly and contiguously cover marker text before its declared suffix')
            continue
          }
          visibleStarted = true
          if (cursor === 0 && Number.isSafeInteger(lines[0].inline_offset_millipoints) && Number.isSafeInteger(entry.list_marker.marker_start_millipoints) && (lines[0].inline_offset_millipoints as number) + prefixAdvance !== entry.list_marker.marker_start_millipoints) add(issues, 'BROKEN_REFERENCE', `${path}/list_marker/marker_start_millipoints`, 'must equal the first-line origin plus the exact virtual marker prefix')
          cursor = fragment.end_utf16 as number
        }
        if (cursor !== markerText.length) add(issues, 'BROKEN_REFERENCE', `${path}/list_marker/text`, 'visible list-marker fragments must cover the complete marker text')
        if ((suffix === 'tab' || suffix === 'space') !== suffixSeen) add(issues, 'BROKEN_REFERENCE', `${path}/list_marker/suffix`, 'the declared list-marker suffix must appear exactly once after the complete marker text')
      }
    }
  }
  if (before !== undefined && after !== undefined && block !== undefined && block !== before + after + lineHeightSum) add(issues, 'INVALID_VALUE', `${path}/block_advance_millipoints`, 'must equal before spacing plus line heights plus after spacing')
}

function validateDiagnostic(value: unknown, path: string, issues: NativeDocxValidationIssue[]): void {
  const entry = object(value, path, DOCX_SHAPED_LINES_V1_BINDING_FIELDS.DiagnosticV1, issues)
  if (!entry) return
  stringValue(entry.code, `${path}/code`, issues, { pattern: SHORT_ID, max: 256 })
  enumValue(entry.severity, `${path}/severity`, ['unsupported', 'deferred'], issues)
  stringValue(entry.scope_id, `${path}/scope_id`, issues, { pattern: SHORT_ID, max: 256 })
  if (entry.source_id !== undefined) stringValue(entry.source_id, `${path}/source_id`, issues, { pattern: SHORT_ID, max: 256 })
  optionalString(entry.source_diagnostic_code, `${path}/source_diagnostic_code`, issues, 256)
  optionalString(entry.source_diagnostic_message, `${path}/source_diagnostic_message`, issues, MAX_MESSAGE)
  stringValue(entry.message, `${path}/message`, issues, { max: MAX_MESSAGE })
}

export function decodeNativeDocxShapedLines(value: unknown): DecodeNativeDocxShapedLinesResult {
  const issues: NativeDocxValidationIssue[] = []
  const preflightState = { nodes: 0, bounded: true }
  preflight(value, '', 0, preflightState, issues)
  if (!preflightState.bounded) return { ok: false, issues }
  const root = object(value, '', DOCX_SHAPED_LINES_V1_BINDING_FIELDS.ShapedLinesV1, issues)
  if (!root) return { ok: false, issues }
  if (root.protocol !== DOCX_SHAPED_LINES_PROTOCOL) add(issues, 'UNSUPPORTED_PROTOCOL', '/protocol', `must equal ${DOCX_SHAPED_LINES_PROTOCOL}`)
  if (root.version !== DOCX_SHAPED_LINES_VERSION) add(issues, 'UNSUPPORTED_VERSION', '/version', `must equal ${DOCX_SHAPED_LINES_VERSION}`)
  stringValue(root.document_id, '/document_id', issues, { pattern: SHORT_ID, max: 256 })
  stringValue(root.revision, '/revision', issues, { pattern: SHORT_ID, max: 256 })
  integer(root.available_width_millipoints, '/available_width_millipoints', issues, 1, DOCX_SHAPED_LINES_LIMITS.maxWidthMilliPoints)
  if (root.numbering_source !== undefined) validateNumberingSource(root.numbering_source, '/numbering_source', issues)
  integer(root.tab_interval_millipoints, '/tab_interval_millipoints', issues, 1, DOCX_SHAPED_LINES_LIMITS.maxWidthMilliPoints)
  const manifest = object(root.font_manifest, '/font_manifest', DOCX_SHAPED_LINES_V1_BINDING_FIELDS.FontManifestV1, issues)
  if (manifest) {
    stringValue(manifest.manifest_id, '/font_manifest/manifest_id', issues, { pattern: SHORT_ID, max: 256 })
    stringValue(manifest.revision, '/font_manifest/revision', issues, { pattern: SHORT_ID, max: 256 })
  }
  const providers = object(root.providers, '/providers', DOCX_SHAPED_LINES_V1_BINDING_FIELDS.ProvidersV1, issues)
  if (providers) {
    for (const key of DOCX_SHAPED_LINES_V1_BINDING_FIELDS.ProvidersV1) stringValue(providers[key], `/providers/${key}`, issues, { pattern: SHORT_ID, max: 256 })
    if (providers.bidi_id !== NATIVE_BIDI_PROVIDER_ID) add(issues, 'INVALID_VALUE', '/providers/bidi_id', `must equal ${NATIVE_BIDI_PROVIDER_ID}`)
    if (providers.bidi_revision !== NATIVE_BIDI_PROVIDER_REVISION) add(issues, 'INVALID_VALUE', '/providers/bidi_revision', `must equal the pinned ${NATIVE_BIDI_PROVIDER_ID} revision`)
    if (providers.bidi_unicode_version !== BIDI_UNICODE_VERSION) add(issues, 'INVALID_VALUE', '/providers/bidi_unicode_version', `must equal ${BIDI_UNICODE_VERSION}`)
    if (providers.unicode13_revision !== UNICODE_13_CLASSIFIER_REVISION) add(issues, 'INVALID_VALUE', '/providers/unicode13_revision', 'must equal the pinned generated Unicode 13 classifier revision')
  }
  const state: ShapedDecodeState = { paragraphs: 0, lines: 0, fragments: 0, glyphs: 0, fragmentIDs: new Set() }
  const paragraphIDs = new Set<string>()
  const lineIDs = new Set<string>()
  array(root.paragraphs, '/paragraphs', issues, DOCX_SHAPED_LINES_LIMITS.maxParagraphs).forEach((paragraph, index) => {
    validateParagraph(paragraph, `/paragraphs/${index}`, issues, state, paragraphIDs, lineIDs)
  })
  array(root.diagnostics, '/diagnostics', issues, DOCX_SHAPED_LINES_LIMITS.maxDiagnostics).forEach((diagnostic, index) => validateDiagnostic(diagnostic, `/diagnostics/${index}`, issues))
  issues.sort(compareNativeValidationIssues)
  return issues.length > 0 ? { ok: false, issues } : { ok: true, value: value as NativeDocxShapedLinesV1 }
}
