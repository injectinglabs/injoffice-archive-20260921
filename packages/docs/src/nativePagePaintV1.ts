/**
 * Canonical renderer-neutral DOCX page-paint projection, version 1.
 *
 * The compiler consumes the exact native shaping/pagination request and its
 * validated paginated output. It never shapes text and never asks a browser or
 * platform font API for metrics. An injected provider supplies content-addressed
 * glyph outlines in integer font design units; this module alone scales and
 * places those outlines into absolute integer milli-point page coordinates.
 */

import {
  scaleFontUnits,
  validateFontManifest,
  type MaybePromise,
  type NativeFontManifest,
} from '@injoffice/font-metrics/layout'
import { sha256 } from '@noble/hashes/sha2.js'
import { bytesToHex } from '@noble/hashes/utils.js'
import {
  DOCX_MAX_TWIPS_FOR_MILLIPOINTS,
  DOCX_NATIVE_LIMITS,
  type NativeDocxParagraphV1,
  type NativeDocxValidationIssue,
} from './nativeContract.js'
import {
  decodeNativeDocxPaginationRequestV1,
  type NativeDocxPaginationRequestV1,
  type NativeDocxPaginatedLayoutV1,
  type NativeDocxPaginatedPageV1,
  type NativeDocxPlacedLineV1,
  type NativeDocxPlacedNoteStoryV1,
} from './nativePaginationV1.js'
import { decodeNativeDocxPaginatedLayoutForRequest } from './nativePaginatedLayoutContract.js'
import { decodeNativeDocxPaginationSettings } from './nativePaginationSettings.js'
import type { NativeDocxShapedLinesV1 } from './nativeShapingLines.js'
import {
  DOCX_HEADER_FOOTER_LAYOUT_PROTOCOL,
  DOCX_HEADER_FOOTER_LAYOUT_VERSION,
  layoutNativeDocxHeadersFootersV1,
  type NativeDocxHeaderFooterLayoutV1,
  type NativeDocxHeaderFooterPageLayoutV1,
  type NativeDocxPlacedHeaderFooterLineV1,
} from './nativeHeaderFooterLayoutV1.js'
import { asciiLowerNative, compareNativeValidationIssues } from './nativeDeterminism.js'
import { layoutNativeDocxTableRowsV1, nativeDocxTableProjectionSha256V1, qualifyNativeDocxTablesV1, type NativeDocxQualifiedTableV1 } from './nativeTablePagePaintV1.js'
import {
  decodeNativeDocxPagePaintMediaAssetsV1,
  decodeNativeDocxPagePaintResourceListV1,
  qualifyNativeDocxInlineImageV1,
  type NativeDocxPagePaintMediaAssetV1,
} from './nativeImagePagePaintV1.js'
import type { NativeDocxResolvedNumberingSourceV1, NativeDocxResolvedRunPropertiesV1 } from './nativeResolvedLayout.js'

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

export interface NativeDocxPagePaintRequestV1 {
  protocol: typeof DOCX_PAGE_PAINT_REQUEST_PROTOCOL
  version: typeof DOCX_PAGE_PAINT_REQUEST_VERSION
  pagination_request: NativeDocxPaginationRequestV1
  paginated_layout: NativeDocxPaginatedLayoutV1
  /** The exact manifest named by shaped-lines provenance, including face digests. */
  font_manifest: NativeFontManifest
  media_assets: NativeDocxPagePaintMediaAssetV1[]
  integrity: { font_manifest_sha256: string; shaped_lines_sha256: string; paginated_layout_sha256: string; table_projection_sha256: string; media_assets_sha256: string }
  outline_provider: { provider_id: string; provider_revision: string }
}

export interface NativeDocxContentAddressedFaceV1 {
  face_id: string
  content_digest: string
  collection_index?: number
}

export interface NativeDocxGlyphOutlineRequestV1 {
  face: NativeDocxContentAddressedFaceV1
  glyph_id: number
}

export type NativeDocxGlyphDesignPathCommandV1 =
  | { kind: 'move_to'; x: number; y: number }
  | { kind: 'line_to'; x: number; y: number }
  | { kind: 'quadratic_to'; control_x: number; control_y: number; x: number; y: number }
  | { kind: 'cubic_to'; control_1_x: number; control_1_y: number; control_2_x: number; control_2_y: number; x: number; y: number }
  | { kind: 'close_path' }

export type NativeDocxGlyphOutlineResultV1 =
  | {
      status: 'outlined'
      face: NativeDocxContentAddressedFaceV1
      glyph_id: number
      units_per_em: number
      path: NativeDocxGlyphDesignPathCommandV1[]
    }
  | {
      /** Explicitly represents a valid glyph with no filled contour, such as space. */
      status: 'empty'
      face: NativeDocxContentAddressedFaceV1
      glyph_id: number
      units_per_em: number
    }
  | {
      status: 'refused'
      face: NativeDocxContentAddressedFaceV1
      glyph_id: number
      code: 'missing-font' | 'missing-glyph' | 'unsupported-font' | 'provider-refusal'
      message: string
    }

export interface NativeDocxGlyphOutlineProviderV1 {
  readonly providerId: string
  readonly providerRevision: string
  getGlyphOutline(request: Readonly<NativeDocxGlyphOutlineRequestV1>): MaybePromise<NativeDocxGlyphOutlineResultV1>
}

export type NativeDocxPaintPathCommandV1 =
  | { kind: 'move_to'; x_millipoints: number; y_millipoints: number }
  | { kind: 'line_to'; x_millipoints: number; y_millipoints: number }
  | { kind: 'quadratic_to'; control_x_millipoints: number; control_y_millipoints: number; x_millipoints: number; y_millipoints: number }
  | { kind: 'cubic_to'; control_1_x_millipoints: number; control_1_y_millipoints: number; control_2_x_millipoints: number; control_2_y_millipoints: number; x_millipoints: number; y_millipoints: number }
  | { kind: 'close_path' }

export interface NativeDocxFillGlyphPathCommandV1 {
  kind: 'fill_glyph_path'
  id: string
  line_id: string
  fragment_id: string
  source_id: string
  glyph_index: number
  face: NativeDocxContentAddressedFaceV1
  glyph_id: number
  font_size_millipoints: number
  fill_rgb: string
  fill_rule: 'nonzero'
  outline_kind: 'path' | 'empty'
  /** Absolute page coordinates. Empty only when outline_kind is `empty`. */
  path: NativeDocxPaintPathCommandV1[]
}

export interface NativeDocxFillTableCellCommandV1 {
  kind: 'fill_table_cell'
  id: string
  table_id: string
  row_id: string
  cell_id: string
  x_millipoints: number
  y_millipoints: number
  width_millipoints: number
  height_millipoints: number
  fill_rgb: string
}

export interface NativeDocxStrokeTableBorderCommandV1 {
  kind: 'stroke_table_border'
  id: string
  table_id: string
  row_id: string
  cell_id: string
  edge: 'top' | 'right' | 'bottom' | 'left'
  x1_millipoints: number
  y1_millipoints: number
  x2_millipoints: number
  y2_millipoints: number
  width_millipoints: number
  stroke_rgb: string
}

export interface NativeDocxStrokeNoteSeparatorCommandV1 {
  kind: 'stroke_note_separator'
  id: string
  line_id: string
  story_id: string
  x1_millipoints: number
  y1_millipoints: number
  x2_millipoints: number
  y2_millipoints: number
  width_millipoints: number
  stroke_rgb: '000000'
}

export interface NativeDocxPaintInlineImageCommandV1 {
  kind: 'paint_inline_image'
  id: string
  line_id: string
  fragment_id: string
  source_id: string
  drawing_id: string
  asset_id: string
  x_millipoints: number
  y_millipoints: number
  width_millipoints: number
  height_millipoints: number
  source_crop: { left: 0; top: 0; right: 0; bottom: 0; unit: 'one-hundred-thousandth' }
  transform: { rotation_degrees: 0; flip_horizontal: false; flip_vertical: false }
}

export type NativeDocxPagePaintCommandV1 = NativeDocxFillGlyphPathCommandV1 | NativeDocxFillTableCellCommandV1 | NativeDocxStrokeTableBorderCommandV1 | NativeDocxStrokeNoteSeparatorCommandV1 | NativeDocxPaintInlineImageCommandV1

export interface NativeDocxPaintLineV1 {
  placed_line_id: string
  line_id: string
  paragraph_id: string
  region: 'body' | 'header' | 'footer' | 'footnote' | 'endnote'
  section_id: string
  column_id?: string
  column_ordinal?: number
  source_line_ordinal: number
  x_millipoints: number
  y_millipoints: number
  width_millipoints: number
  height_millipoints: number
  baseline_y_millipoints: number
  command_ids: string[]
}

export interface NativeDocxPaintPageV1 {
  id: string
  ordinal: number
  section_id: string
  section_ids: string[]
  kind: 'content' | 'parity-blank'
  width_millipoints: number
  height_millipoints: number
  body_box: {
    x_millipoints: number
    y_millipoints: number
    width_millipoints: number
    height_millipoints: number
  }
  columns: Array<{ id: string; section_id: string; ordinal: number; x_millipoints: number; y_millipoints: number; width_millipoints: number; height_millipoints: number }>
  background_rgb: 'FFFFFF'
  clip_box: { x_millipoints: 0; y_millipoints: 0; width_millipoints: number; height_millipoints: number }
  lines: NativeDocxPaintLineV1[]
  commands: NativeDocxPagePaintCommandV1[]
}

export interface NativeDocxPagePaintProvenanceV1 {
  document_id: string
  revision: string
  package_sha256: string
  main_part: string
  body_story_id: string
  numbering_source?: NativeDocxResolvedNumberingSourceV1
  pagination_settings: NativeDocxPaginationRequestV1['pagination_settings']
  shaped_lines: {
    protocol: NativeDocxPaginationRequestV1['shaped_lines']['protocol']
    version: NativeDocxPaginationRequestV1['shaped_lines']['version']
    available_width_millipoints: number
    tab_interval_millipoints: number
    sha256: string
  }
  paginated_layout: { protocol: NativeDocxPaginatedLayoutV1['protocol']; version: NativeDocxPaginatedLayoutV1['version']; sha256: string }
  header_footer_layout: { protocol: typeof DOCX_HEADER_FOOTER_LAYOUT_PROTOCOL; version: typeof DOCX_HEADER_FOOTER_LAYOUT_VERSION; sha256: string }
  table_projection: { sha256: string }
  font_manifest: { manifest_id: string; revision: string; sha256: string }
  media_assets: { sha256: string }
  providers: NativeDocxPaginationRequestV1['shaped_lines']['providers'] & { outline_id: string; outline_revision: string }
}

export type NativeDocxPagePaintDiagnosticCode =
  | 'upstream-refused'
  | 'unsupported-diagnostic'
  | 'unsupported-source'
  | 'missing-font'
  | 'missing-glyph'
  | 'provider-refusal'
  | 'provider-mismatch'
  | 'provider-failure'
  | 'invalid-provider-output'
  | 'invalid-path'
  | 'resource-limit'
  | 'identity-mismatch'
  | 'incomplete-page'

export interface NativeDocxPagePaintDiagnosticV1 {
  code: NativeDocxPagePaintDiagnosticCode
  severity: 'unsupported'
  scope_id: string
  message: string
}

interface NativeDocxPagePaintBaseV1 {
  protocol: typeof DOCX_PAGE_PAINT_PROTOCOL
  version: typeof DOCX_PAGE_PAINT_VERSION
  provenance: NativeDocxPagePaintProvenanceV1
  diagnostics: NativeDocxPagePaintDiagnosticV1[]
}

export interface NativeDocxPagePaintSuccessV1 extends NativeDocxPagePaintBaseV1 {
  status: 'painted'
  resources: NativeDocxPagePaintMediaAssetV1[]
  pages: NativeDocxPaintPageV1[]
}

export interface NativeDocxPagePaintRefusedV1 extends NativeDocxPagePaintBaseV1 {
  status: 'refused'
  resources: []
  pages: []
}

export type NativeDocxPagePaintV1 = NativeDocxPagePaintSuccessV1 | NativeDocxPagePaintRefusedV1

export type DecodeNativeDocxPagePaintRequestV1Result =
  | { ok: true; value: NativeDocxPagePaintRequestV1 }
  | { ok: false; issues: NativeDocxValidationIssue[] }

export type DecodeNativeDocxPagePaintV1Result =
  | { ok: true; value: NativeDocxPagePaintV1 }
  | { ok: false; issues: NativeDocxValidationIssue[] }

export type CompileNativeDocxPagePaintV1Result =
  | { ok: true; value: NativeDocxPagePaintV1 }
  | { ok: false; issues: NativeDocxValidationIssue[] }

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

type JsonObject = Record<string, unknown>
const ID = /^[A-Za-z0-9][A-Za-z0-9._:/-]{0,1023}$/
const SHORT_ID = /^[A-Za-z0-9][A-Za-z0-9._:/-]{0,255}$/
const PROVIDER_ID = /^[A-Za-z0-9][A-Za-z0-9._:/-]{0,127}$/
const SHA256 = /^sha256:[0-9a-f]{64}$/
const RGB = /^[0-9A-F]{6}$/
const ABSOLUTE_URI = /^[A-Za-z][A-Za-z0-9+.-]*:[^\u0000\r\n]{1,4090}$/
const PART_SEGMENT = /^(?:[A-Za-z0-9._~!$&'()*+,;=@-]|%[0-9A-F]{2})+$/

function canonicalJsonValue(value: unknown, active = new WeakSet<object>(), depth = 0, state = { nodes: 0 }): unknown {
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

function canonicalWireSha256(value: unknown): string {
  return `sha256:${bytesToHex(sha256(new TextEncoder().encode(JSON.stringify(canonicalJsonValue(value)))))}`
}

/** Canonical content attestation for the exact validated manifest carried by page-paint v1. */
export function nativeDocxPagePaintFontManifestSha256V1(manifest: NativeFontManifest): string {
  const validated = validateFontManifest(manifest)
  if (!validated.ok) throw new TypeError('font manifest must satisfy the native text v1 contract before hashing')
  return canonicalWireSha256(validated.value)
}

/** Canonical content attestation for the complete shaped-lines projection carried by page-paint v1. */
export function nativeDocxPagePaintShapedLinesSha256V1(shapedLines: NativeDocxShapedLinesV1): string {
  return canonicalWireSha256(shapedLines)
}

export function nativeDocxPagePaintMediaAssetsSha256V1(assets: readonly NativeDocxPagePaintMediaAssetV1[]): string {
  return canonicalWireSha256(assets)
}

/** Canonical attestation for exact section, column, and line placement. */
export function nativeDocxPagePaintPaginatedLayoutSha256V1(layout: NativeDocxPaginatedLayoutV1): string {
  return canonicalWireSha256(layout)
}

/** Canonical hash of the complete strict request. The hash is carried by the compiler envelope, not recursively in the request. */
export function nativeDocxPagePaintRequestSha256V1(value: unknown): string {
  const decoded = decodeNativeDocxPagePaintRequestV1(value)
  if (!decoded.ok) throw new TypeError('page-paint request must validate before canonical hashing')
  return canonicalWireSha256(decoded.value)
}

/** Canonical hash of the complete strict output. The hash is carried by the compiler envelope, not recursively in the output. */
export function nativeDocxPagePaintOutputSha256V1(value: unknown): string {
  const decoded = decodeNativeDocxPagePaintV1(value)
  if (!decoded.ok) throw new TypeError('page-paint output must validate before canonical hashing')
  return canonicalWireSha256(decoded.value)
}

function issue(code: NativeDocxValidationIssue['code'], path: string, message: string): NativeDocxValidationIssue {
  return { code, path, message }
}

function add(issues: NativeDocxValidationIssue[], code: NativeDocxValidationIssue['code'], path: string, message: string): void {
  if (issues.length >= DOCX_NATIVE_LIMITS.maxIssues) return
  if (!issues.some((entry) => entry.code === code && entry.path === path && entry.message === message)) issues.push(issue(code, path, message))
}

function isObject(value: unknown): value is JsonObject {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

function pointer(value: string): string {
  return value.replace(/~/g, '~0').replace(/\//g, '~1')
}

function exactObject(value: unknown, path: string, fields: readonly string[], issues: NativeDocxValidationIssue[]): JsonObject | undefined {
  if (!isObject(value)) {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, 'must be an object')
    return undefined
  }
  const allowed = new Set(fields)
  for (const key of Object.keys(value).sort()) if (!allowed.has(key)) add(issues, 'UNKNOWN_FIELD', `${path}/${pointer(key)}`, `unknown field ${JSON.stringify(key)}`)
  for (const key of fields) if (!(key in value)) add(issues, 'REQUIRED', `${path}/${pointer(key)}`, 'field is required')
  return value
}

function exactUnionObject(value: unknown, path: string, fields: readonly string[], issues: NativeDocxValidationIssue[]): JsonObject | undefined {
  if (!isObject(value)) {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, 'must be an object')
    return undefined
  }
  const allowed = new Set(fields)
  for (const key of Object.keys(value).sort()) if (!allowed.has(key)) add(issues, 'UNKNOWN_FIELD', `${path}/${pointer(key)}`, `unknown field ${JSON.stringify(key)}`)
  return value
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

function integer(value: unknown, path: string, issues: NativeDocxValidationIssue[], min: number, max: number): number | undefined {
  if (!Number.isSafeInteger(value) || Object.is(value, -0) || (value as number) < min || (value as number) > max) {
    add(issues, value === undefined ? 'REQUIRED' : 'OUT_OF_RANGE', path, `must be a safe integer from ${min} through ${max}`)
    return undefined
  }
  return value as number
}

function safeClone<T>(value: T): T | undefined {
  try { return structuredClone(value) as T } catch { return undefined }
}

function preflightWire(value: unknown, label: string, maxNodes: number = DOCX_PAGE_PAINT_LIMITS.maxOutputNodes, maxArray: number = DOCX_PAGE_PAINT_LIMITS.maxGlyphs): NativeDocxValidationIssue[] {
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

function sameWire(left: unknown, right: unknown): boolean {
  if (left === right) return true
  if (Array.isArray(left) || Array.isArray(right)) return Array.isArray(left) && Array.isArray(right) && left.length === right.length && left.every((entry, index) => sameWire(entry, right[index]))
  if (!isObject(left) || !isObject(right)) return false
  const leftKeys = Object.keys(left).sort()
  const rightKeys = Object.keys(right).sort()
  return leftKeys.length === rightKeys.length && leftKeys.every((key, index) => key === rightKeys[index] && sameWire(left[key], right[key]))
}

function canonicalPart(value: string): string {
  try { return asciiLowerNative(value.split('/').map((segment) => decodeURIComponent(segment)).join('/')) } catch { return value }
}

function validPartName(value: string): boolean {
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

function validateNumberingSource(value: unknown, path: string, issues: NativeDocxValidationIssue[]): void {
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

function headerFooterLayout(request: NativeDocxPagePaintRequestV1): NativeDocxHeaderFooterLayoutV1 {
  return layoutNativeDocxHeadersFootersV1({
    document: request.pagination_request.document,
    resolved_layout: request.pagination_request.resolved_layout,
    shaped_lines: request.pagination_request.shaped_lines,
    pagination_settings: request.pagination_request.pagination_settings,
    paginated_layout: request.paginated_layout,
  })
}

function requestProvenance(request: NativeDocxPagePaintRequestV1, outlineID: string, outlineRevision: string, layout: NativeDocxHeaderFooterLayoutV1 = headerFooterLayout(request)): NativeDocxPagePaintProvenanceV1 {
  const pagination = request.pagination_request
  return {
    document_id: pagination.document.document_id,
    revision: pagination.document.revision,
    package_sha256: pagination.document.source.package_sha256,
    main_part: pagination.document.source.main_part,
    body_story_id: pagination.document.body.id,
    ...(pagination.shaped_lines.numbering_source ? { numbering_source: { ...pagination.shaped_lines.numbering_source } } : {}),
    pagination_settings: safeClone(pagination.pagination_settings)!,
    shaped_lines: {
      protocol: pagination.shaped_lines.protocol,
      version: pagination.shaped_lines.version,
      available_width_millipoints: pagination.shaped_lines.available_width_millipoints,
      tab_interval_millipoints: pagination.shaped_lines.tab_interval_millipoints,
      sha256: request.integrity.shaped_lines_sha256,
    },
    paginated_layout: { protocol: request.paginated_layout.protocol, version: request.paginated_layout.version, sha256: request.integrity.paginated_layout_sha256 },
    header_footer_layout: { protocol: layout.protocol, version: layout.version, sha256: layout.sha256 },
    table_projection: { sha256: request.integrity.table_projection_sha256 },
    font_manifest: { manifest_id: request.font_manifest.manifestId, revision: request.font_manifest.revision, sha256: request.integrity.font_manifest_sha256 },
    media_assets: { sha256: request.integrity.media_assets_sha256 },
    providers: { ...pagination.shaped_lines.providers, outline_id: outlineID, outline_revision: outlineRevision },
  }
}

/** Strictly decodes and exact-joins every engine-owned input projection. */
export function decodeNativeDocxPagePaintRequestV1(value: unknown): DecodeNativeDocxPagePaintRequestV1Result {
  const preflight = preflightWire(value, 'page-paint request')
  if (preflight.length > 0) return { ok: false, issues: preflight }
  const snapshot = safeClone(value)
  if (snapshot === undefined) return { ok: false, issues: [issue('INVALID_VALUE', '', 'page-paint request must be a cloneable JSON wire value')] }
  const issues: NativeDocxValidationIssue[] = []
  const root = exactObject(snapshot, '', DOCX_PAGE_PAINT_REQUEST_V1_BINDING_FIELDS.RequestV1, issues)
  if (!root) return { ok: false, issues }
  if (root.protocol !== DOCX_PAGE_PAINT_REQUEST_PROTOCOL) add(issues, 'UNSUPPORTED_PROTOCOL', '/protocol', `must equal ${DOCX_PAGE_PAINT_REQUEST_PROTOCOL}`)
  if (root.version !== DOCX_PAGE_PAINT_REQUEST_VERSION) add(issues, 'UNSUPPORTED_VERSION', '/version', `must equal ${DOCX_PAGE_PAINT_REQUEST_VERSION}`)
  const pagination = decodeNativeDocxPaginationRequestV1(root.pagination_request)
  if (!pagination.ok) issues.push(...pagination.issues.map((entry) => ({ ...entry, path: `/pagination_request${entry.path}` })))
  const paginated = decodeNativeDocxPaginatedLayoutForRequest(root.paginated_layout, root.pagination_request)
  if (!paginated.ok) issues.push(...paginated.issues.map((entry) => ({ ...entry, path: `/paginated_layout${entry.path}` })))
  const manifest = validateFontManifest(root.font_manifest)
  if (!manifest.ok) issues.push(...manifest.issues.slice(0, DOCX_NATIVE_LIMITS.maxIssues).map((entry) => issue(entry.code === 'unknown-field' ? 'UNKNOWN_FIELD' : entry.code === 'reference' ? 'BROKEN_REFERENCE' : 'INVALID_VALUE', `/font_manifest${entry.path === '$' ? '' : entry.path.slice(1).replace(/\./g, '/')}`, entry.message)))
  if (pagination.ok && manifest.ok) {
    if (manifest.value.manifestId !== pagination.value.shaped_lines.font_manifest.manifest_id || manifest.value.revision !== pagination.value.shaped_lines.font_manifest.revision) add(issues, 'BROKEN_REFERENCE', '/font_manifest', 'manifest id and revision must exactly match shaped-lines provenance')
    if (paginated.ok) {
      const provenance = paginated.value.provenance
      if (provenance.document_id !== pagination.value.document.document_id || provenance.revision !== pagination.value.document.revision || provenance.package_sha256 !== pagination.value.document.source.package_sha256 || canonicalPart(provenance.main_part) !== canonicalPart(pagination.value.document.source.main_part)) add(issues, 'BROKEN_REFERENCE', '/paginated_layout/provenance', 'paginated document, revision, package, and main-part identity must exactly match the joined request')
      if (provenance.font_manifest.manifest_id !== manifest.value.manifestId || provenance.font_manifest.revision !== manifest.value.revision || !sameWire(provenance.providers, pagination.value.shaped_lines.providers) || !sameWire(provenance.pagination_settings, pagination.value.pagination_settings)) add(issues, 'BROKEN_REFERENCE', '/paginated_layout/provenance', 'pagination settings and font/provider provenance must exactly match shaping and the supplied manifest')
      if (!sameWire(provenance.numbering_source, pagination.value.shaped_lines.numbering_source)) add(issues, 'BROKEN_REFERENCE', '/paginated_layout/provenance/numbering_source', 'numbering relationship, part, hashes, and canonical model digest must exactly match shaped-lines provenance')
    }
  }
  let mediaAssets: NativeDocxPagePaintMediaAssetV1[] | undefined
  if (pagination.ok) {
    try { mediaAssets = decodeNativeDocxPagePaintMediaAssetsV1(pagination.value.document, root.media_assets) } catch (error) {
      add(issues, 'INVALID_VALUE', '/media_assets', error instanceof Error ? error.message : 'media asset inventory is invalid')
    }
  }
  const integrity = exactObject(root.integrity, '/integrity', DOCX_PAGE_PAINT_REQUEST_V1_BINDING_FIELDS.IntegrityV1, issues)
  if (integrity) {
    stringValue(integrity.font_manifest_sha256, '/integrity/font_manifest_sha256', issues, SHA256, 71)
    stringValue(integrity.shaped_lines_sha256, '/integrity/shaped_lines_sha256', issues, SHA256, 71)
    stringValue(integrity.table_projection_sha256, '/integrity/table_projection_sha256', issues, SHA256, 71)
    stringValue(integrity.media_assets_sha256, '/integrity/media_assets_sha256', issues, SHA256, 71)
    stringValue(integrity.paginated_layout_sha256, '/integrity/paginated_layout_sha256', issues, SHA256, 71)
    if (manifest.ok && integrity.font_manifest_sha256 !== nativeDocxPagePaintFontManifestSha256V1(manifest.value)) add(issues, 'BROKEN_REFERENCE', '/integrity/font_manifest_sha256', 'must attest the complete validated font manifest carried by this request')
    if (pagination.ok && integrity.shaped_lines_sha256 !== nativeDocxPagePaintShapedLinesSha256V1(pagination.value.shaped_lines)) add(issues, 'BROKEN_REFERENCE', '/integrity/shaped_lines_sha256', 'must attest the complete strict shaped-lines projection carried by this request')
    if (pagination.ok) {
      const qualified = qualifyNativeDocxTablesV1(pagination.value.document, pagination.value.resolved_layout)
      const expected = qualified.status === 'qualified' ? qualified.sha256 : nativeDocxTableProjectionSha256V1([])
      if (integrity.table_projection_sha256 !== expected) add(issues, 'BROKEN_REFERENCE', '/integrity/table_projection_sha256', 'must attest the exact qualified table projection or the canonical empty projection for an upstream refusal')
    }
    if (mediaAssets && integrity.media_assets_sha256 !== nativeDocxPagePaintMediaAssetsSha256V1(mediaAssets)) add(issues, 'BROKEN_REFERENCE', '/integrity/media_assets_sha256', 'must attest the complete canonical digest-bound media inventory carried by this request')
    if (paginated.ok && integrity.paginated_layout_sha256 !== nativeDocxPagePaintPaginatedLayoutSha256V1(paginated.value)) add(issues, 'BROKEN_REFERENCE', '/integrity/paginated_layout_sha256', 'must attest the complete deterministic section, column, body, and note placement projection')
  }
  const outlineProvider = exactObject(root.outline_provider, '/outline_provider', DOCX_PAGE_PAINT_REQUEST_V1_BINDING_FIELDS.OutlineProviderV1, issues)
  if (outlineProvider) {
    stringValue(outlineProvider.provider_id, '/outline_provider/provider_id', issues, PROVIDER_ID, 128)
    stringValue(outlineProvider.provider_revision, '/outline_provider/provider_revision', issues, PROVIDER_ID, 128)
  }
  issues.sort(compareNativeValidationIssues)
  if (issues.length > 0 || !pagination.ok || !paginated.ok || !manifest.ok || !mediaAssets) return { ok: false, issues: issues.slice(0, DOCX_NATIVE_LIMITS.maxIssues) }
  return {
    ok: true,
    value: {
      protocol: DOCX_PAGE_PAINT_REQUEST_PROTOCOL,
      version: DOCX_PAGE_PAINT_REQUEST_VERSION,
      pagination_request: pagination.value,
      paginated_layout: paginated.value,
      font_manifest: manifest.value,
      media_assets: mediaAssets,
      integrity: root.integrity as NativeDocxPagePaintRequestV1['integrity'],
      outline_provider: root.outline_provider as NativeDocxPagePaintRequestV1['outline_provider'],
    },
  }
}

interface ProviderSnapshot {
  id: string
  revision: string
  live: NativeDocxGlyphOutlineProviderV1
  callable: NativeDocxGlyphOutlineProviderV1['getGlyphOutline']
  get: (request: Readonly<NativeDocxGlyphOutlineRequestV1>) => MaybePromise<NativeDocxGlyphOutlineResultV1>
}

function snapshotProvider(provider: NativeDocxGlyphOutlineProviderV1): { ok: true; value: ProviderSnapshot } | { ok: false; issues: NativeDocxValidationIssue[] } {
  try {
    if (!isObject(provider)) return { ok: false, issues: [issue('INVALID_TYPE', '/provider', 'outline provider must be an object')] }
    const id = provider.providerId
    const revision = provider.providerRevision
    const get = provider.getGlyphOutline
    if (typeof id !== 'string' || !PROVIDER_ID.test(id)) return { ok: false, issues: [issue('INVALID_VALUE', '/provider/provider_id', 'must be a bounded stable provider identifier')] }
    if (typeof revision !== 'string' || !PROVIDER_ID.test(revision)) return { ok: false, issues: [issue('INVALID_VALUE', '/provider/provider_revision', 'must be a bounded stable provider revision')] }
    if (typeof get !== 'function') return { ok: false, issues: [issue('INVALID_TYPE', '/provider/getGlyphOutline', 'must be a function')] }
    if (provider.providerId !== id || provider.providerRevision !== revision || provider.getGlyphOutline !== get) return { ok: false, issues: [issue('INVALID_VALUE', '/provider', 'provider identity and callable must remain stable while snapshotted')] }
    return { ok: true, value: { id, revision, live: provider, callable: get, get: get.bind(provider) } }
  } catch {
    return { ok: false, issues: [issue('INVALID_VALUE', '/provider', 'provider identity access must be deterministic and side-effect free')] }
  }
}

function providerStable(provider: ProviderSnapshot): boolean {
  try { return provider.live.providerId === provider.id && provider.live.providerRevision === provider.revision && provider.live.getGlyphOutline === provider.callable } catch { return false }
}

function faceSame(left: NativeDocxContentAddressedFaceV1, right: NativeDocxContentAddressedFaceV1): boolean {
  return left.face_id === right.face_id && left.content_digest === right.content_digest && left.collection_index === right.collection_index
}

function faceSnapshot(value: unknown): NativeDocxContentAddressedFaceV1 | undefined {
  if (!isObject(value)) return undefined
  const keys = Object.keys(value).sort()
  const expected = value.collection_index === undefined ? ['content_digest', 'face_id'] : ['collection_index', 'content_digest', 'face_id']
  if (keys.length !== expected.length || keys.some((key, index) => key !== expected[index])) return undefined
  if (typeof value.face_id !== 'string' || !SHORT_ID.test(value.face_id) || typeof value.content_digest !== 'string' || !SHA256.test(value.content_digest)) return undefined
  if (value.collection_index !== undefined && (!Number.isSafeInteger(value.collection_index) || Object.is(value.collection_index, -0) || (value.collection_index as number) < 0 || (value.collection_index as number) > 65_535)) return undefined
  return { face_id: value.face_id, content_digest: value.content_digest, ...(value.collection_index !== undefined ? { collection_index: value.collection_index as number } : {}) }
}

function designInteger(value: unknown): value is number {
  return Number.isSafeInteger(value) && !Object.is(value, -0) && Math.abs(value as number) <= DOCX_PAGE_PAINT_LIMITS.maxDesignCoordinate
}

function captureDesignPath(value: unknown): NativeDocxGlyphDesignPathCommandV1[] | undefined {
  if (!Array.isArray(value) || value.length === 0 || value.length > DOCX_PAGE_PAINT_LIMITS.maxPathCommandsPerGlyph) return undefined
  const result: NativeDocxGlyphDesignPathCommandV1[] = []
  let contourOpen = false
  let contourDrawn = false
  let minX = Infinity
  let maxX = -Infinity
  let minY = Infinity
  let maxY = -Infinity
  const bounds = (x: number, y: number): void => { minX = Math.min(minX, x); maxX = Math.max(maxX, x); minY = Math.min(minY, y); maxY = Math.max(maxY, y) }
  for (const entry of value) {
    if (!isObject(entry) || typeof entry.kind !== 'string') return undefined
    if (entry.kind === 'move_to') {
      if (contourOpen || Object.keys(entry).sort().join(',') !== 'kind,x,y' || !designInteger(entry.x) || !designInteger(entry.y)) return undefined
      contourOpen = true
      contourDrawn = false
      bounds(entry.x, entry.y)
      result.push({ kind: 'move_to', x: entry.x, y: entry.y })
    } else if (entry.kind === 'line_to') {
      if (!contourOpen || Object.keys(entry).sort().join(',') !== 'kind,x,y' || !designInteger(entry.x) || !designInteger(entry.y)) return undefined
      contourDrawn = true
      bounds(entry.x, entry.y)
      result.push({ kind: 'line_to', x: entry.x, y: entry.y })
    } else if (entry.kind === 'quadratic_to') {
      if (!contourOpen || Object.keys(entry).sort().join(',') !== 'control_x,control_y,kind,x,y' || !designInteger(entry.control_x) || !designInteger(entry.control_y) || !designInteger(entry.x) || !designInteger(entry.y)) return undefined
      contourDrawn = true
      bounds(entry.control_x, entry.control_y)
      bounds(entry.x, entry.y)
      result.push({ kind: 'quadratic_to', control_x: entry.control_x, control_y: entry.control_y, x: entry.x, y: entry.y })
    } else if (entry.kind === 'cubic_to') {
      if (!contourOpen || Object.keys(entry).sort().join(',') !== 'control_1_x,control_1_y,control_2_x,control_2_y,kind,x,y' || !designInteger(entry.control_1_x) || !designInteger(entry.control_1_y) || !designInteger(entry.control_2_x) || !designInteger(entry.control_2_y) || !designInteger(entry.x) || !designInteger(entry.y)) return undefined
      contourDrawn = true
      bounds(entry.control_1_x, entry.control_1_y)
      bounds(entry.control_2_x, entry.control_2_y)
      bounds(entry.x, entry.y)
      result.push({ kind: 'cubic_to', control_1_x: entry.control_1_x, control_1_y: entry.control_1_y, control_2_x: entry.control_2_x, control_2_y: entry.control_2_y, x: entry.x, y: entry.y })
    } else if (entry.kind === 'close_path') {
      if (!contourOpen || !contourDrawn || Object.keys(entry).length !== 1) return undefined
      contourOpen = false
      contourDrawn = false
      result.push({ kind: 'close_path' })
    } else return undefined
  }
  return contourOpen || !(maxX > minX && maxY > minY) ? undefined : result
}

function captureOutline(value: unknown, expectedFace: NativeDocxContentAddressedFaceV1, glyphID: number): NativeDocxGlyphOutlineResultV1 | undefined {
  if (!isObject(value) || typeof value.status !== 'string') return undefined
  const face = faceSnapshot(value.face)
  if (!face || !faceSame(face, expectedFace) || value.glyph_id !== glyphID) return undefined
  if (value.status === 'outlined') {
    if (Object.keys(value).sort().join(',') !== 'face,glyph_id,path,status,units_per_em' || !Number.isSafeInteger(value.units_per_em) || Object.is(value.units_per_em, -0) || (value.units_per_em as number) < 1 || (value.units_per_em as number) > DOCX_PAGE_PAINT_LIMITS.maxUnitsPerEm) return undefined
    const path = captureDesignPath(value.path)
    return path ? { status: 'outlined', face, glyph_id: glyphID, units_per_em: value.units_per_em as number, path } : undefined
  }
  if (value.status === 'empty') {
    if (Object.keys(value).sort().join(',') !== 'face,glyph_id,status,units_per_em' || !Number.isSafeInteger(value.units_per_em) || Object.is(value.units_per_em, -0) || (value.units_per_em as number) < 1 || (value.units_per_em as number) > DOCX_PAGE_PAINT_LIMITS.maxUnitsPerEm) return undefined
    return { status: 'empty', face, glyph_id: glyphID, units_per_em: value.units_per_em as number }
  }
  if (value.status === 'refused') {
    const codes = ['missing-font', 'missing-glyph', 'unsupported-font', 'provider-refusal']
    if (Object.keys(value).sort().join(',') !== 'code,face,glyph_id,message,status' || typeof value.code !== 'string' || !codes.includes(value.code) || typeof value.message !== 'string' || value.message.length === 0 || value.message.length > DOCX_PAGE_PAINT_LIMITS.maxProviderMessageLength || /[\u0000\r\n]/.test(value.message)) return undefined
    return { status: 'refused', face, glyph_id: glyphID, code: value.code as 'missing-font' | 'missing-glyph' | 'unsupported-font' | 'provider-refusal', message: value.message }
  }
  return undefined
}

function coordinate(origin: number, design: number, fontSize: number, unitsPerEm: number, invertY = false): number | undefined {
  let scaled: number
  try { scaled = scaleFontUnits(design, unitsPerEm, fontSize) } catch { return undefined }
  const result = origin + (invertY ? -scaled : scaled)
  return Number.isSafeInteger(result) && Math.abs(result) <= DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints ? result : undefined
}

function placePath(path: NativeDocxGlyphDesignPathCommandV1[], originX: number, originY: number, fontSize: number, unitsPerEm: number): NativeDocxPaintPathCommandV1[] | undefined {
  const output: NativeDocxPaintPathCommandV1[] = []
  for (const command of path) {
    if (command.kind === 'close_path') { output.push(command); continue }
    const x = coordinate(originX, command.x, fontSize, unitsPerEm)
    const y = coordinate(originY, command.y, fontSize, unitsPerEm, true)
    if (x === undefined || y === undefined) return undefined
    if (command.kind === 'move_to' || command.kind === 'line_to') output.push({ kind: command.kind, x_millipoints: x, y_millipoints: y })
    else if (command.kind === 'quadratic_to') {
      const controlX = coordinate(originX, command.control_x, fontSize, unitsPerEm)
      const controlY = coordinate(originY, command.control_y, fontSize, unitsPerEm, true)
      if (controlX === undefined || controlY === undefined) return undefined
      output.push({ kind: 'quadratic_to', control_x_millipoints: controlX, control_y_millipoints: controlY, x_millipoints: x, y_millipoints: y })
    } else {
      const control1X = coordinate(originX, command.control_1_x, fontSize, unitsPerEm)
      const control1Y = coordinate(originY, command.control_1_y, fontSize, unitsPerEm, true)
      const control2X = coordinate(originX, command.control_2_x, fontSize, unitsPerEm)
      const control2Y = coordinate(originY, command.control_2_y, fontSize, unitsPerEm, true)
      if (control1X === undefined || control1Y === undefined || control2X === undefined || control2Y === undefined) return undefined
      output.push({ kind: 'cubic_to', control_1_x_millipoints: control1X, control_1_y_millipoints: control1Y, control_2_x_millipoints: control2X, control_2_y_millipoints: control2Y, x_millipoints: x, y_millipoints: y })
    }
  }
  let minX = Infinity; let maxX = -Infinity; let minY = Infinity; let maxY = -Infinity
  const bounds = (x: number, y: number): void => { minX = Math.min(minX, x); maxX = Math.max(maxX, x); minY = Math.min(minY, y); maxY = Math.max(maxY, y) }
  for (const command of output) {
    if (command.kind === 'close_path') continue
    if (command.kind === 'quadratic_to') bounds(command.control_x_millipoints, command.control_y_millipoints)
    if (command.kind === 'cubic_to') {
      bounds(command.control_1_x_millipoints, command.control_1_y_millipoints)
      bounds(command.control_2_x_millipoints, command.control_2_y_millipoints)
    }
    bounds(command.x_millipoints, command.y_millipoints)
  }
  return maxX > minX && maxY > minY ? output : undefined
}

function refusal(provenance: NativeDocxPagePaintProvenanceV1, code: NativeDocxPagePaintDiagnosticCode, scopeID: string, message: string): NativeDocxPagePaintRefusedV1 {
  const safeMessage = message.replace(/[\u0000\r\n]/g, ' ').slice(0, DOCX_PAGE_PAINT_LIMITS.maxProviderMessageLength) || 'Native page paint refused'
  return {
    protocol: DOCX_PAGE_PAINT_PROTOCOL,
    version: DOCX_PAGE_PAINT_VERSION,
    status: 'refused',
    provenance,
    diagnostics: [{ code, severity: 'unsupported', scope_id: scopeID, message: safeMessage }],
    resources: [],
    pages: [],
  }
}

function boundedProviderError(error: unknown): string {
  try {
    const candidate = typeof error === 'string' ? error : isObject(error) ? Reflect.get(error, 'message') : undefined
    if (typeof candidate !== 'string') return 'unknown provider error'
    return candidate.replace(/[\u0000\r\n]/g, ' ').slice(0, 2_048) || 'unknown provider error'
  } catch { return 'unreadable provider error' }
}

function paintTwips(value: number): number | undefined {
  if (!Number.isSafeInteger(value) || Object.is(value, -0) || Math.abs(value) > DOCX_MAX_TWIPS_FOR_MILLIPOINTS || Math.abs(value) > Math.floor(DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints / 50)) return undefined
  return value * 50
}

function nativePaintParagraphs(request: NativeDocxPagePaintRequestV1): Map<string, NativeDocxParagraphV1> {
  const paragraphsByID = new Map<string, NativeDocxParagraphV1>()
  const document = request.pagination_request.document
  for (const story of [document.body, ...document.headers, ...document.footers, ...document.notes, ...document.comment_stories]) for (const block of story.blocks) {
    const paragraphs = block.paragraph ? [block.paragraph] : block.table ? block.table.rows.flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs)) : []
    for (const paragraph of paragraphs) paragraphsByID.set(paragraph.id, paragraph)
  }
  return paragraphsByID
}

interface NativeDocxTableCommandsV1 {
  fills: NativeDocxFillTableCellCommandV1[]
  borders: NativeDocxStrokeTableBorderCommandV1[]
}

/** Build table paint once for the complete page set; never rescan all tables for every page. */
function tableCommandsByPage(
  request: NativeDocxPagePaintRequestV1,
  pages: readonly NativeDocxPaginatedPageV1[],
  tables: readonly NativeDocxQualifiedTableV1[],
): Map<string, NativeDocxTableCommandsV1> | undefined {
  const shaped = new Map(request.pagination_request.shaped_lines.paragraphs.map((paragraph) => [paragraph.paragraph_id, paragraph]))
  const pageByID = new Map(pages.map((page) => [page.id, page]))
  const commands = new Map<string, NativeDocxTableCommandsV1>(pages.map((page) => [page.id, { fills: [], borders: [] }]))
  const placementsByParagraph = new Map<string, Array<{ pageID: string; line: NativeDocxPlacedLineV1 }>>()
  let work = 0
  for (const page of pages) {
    const seen = new Set<string>()
    for (const line of page.lines) {
      if (seen.has(line.paragraph_id)) continue
      seen.add(line.paragraph_id)
      const placements = placementsByParagraph.get(line.paragraph_id) ?? []
      placements.push({ pageID: page.id, line })
      placementsByParagraph.set(line.paragraph_id, placements)
      work += 1
      if (work > DOCX_PAGE_PAINT_LIMITS.maxOutputNodes) return undefined
    }
  }
  for (const table of tables) {
    const rows = layoutNativeDocxTableRowsV1(table, request.pagination_request.shaped_lines)
    if (!rows) return undefined
    const gridColumns = table.grid_widths_millipoints.length
    for (const [rowIndex, row] of rows.entries()) {
      const anchorCell = table.rows[rowIndex]!.cells.find((cell) => cell.vertical_merge !== 'continue') ?? table.rows[rowIndex]!.cells[0]
      const firstParagraph = anchorCell?.cell.paragraphs[0]
      const firstShaped = firstParagraph ? shaped.get(firstParagraph.id) : undefined
      const placements = firstParagraph ? placementsByParagraph.get(firstParagraph.id) : undefined
      if (!placements || !firstShaped) continue
      const topMargin = table.table.cell_margins!.top_twips * 50
      for (const placement of placements) {
        const page = pageByID.get(placement.pageID)
        const target = commands.get(placement.pageID)
        if (!page || !target) return undefined
        const rowY = placement.line.y_millipoints - topMargin - firstShaped.spacing_before_millipoints
        if (!Number.isSafeInteger(rowY)) return undefined
        for (const [cellIndex, cell] of row.cells.entries()) {
          work += 1
          if (work > DOCX_PAGE_PAINT_LIMITS.maxOutputNodes) return undefined
          if (cell.vertical_merge === 'continue') continue
          const x = page.body_box.x_millipoints + cell.x_millipoints
          const height = cell.height_millipoints
          if (cell.shading_rgb) target.fills.push({ kind: 'fill_table_cell', id: `paint:table:${table.table.id}:${rowIndex}:${cellIndex}:fill`, table_id: table.table.id, row_id: row.row_id, cell_id: cell.cell_id, x_millipoints: x, y_millipoints: rowY, width_millipoints: cell.width_millipoints, height_millipoints: height, fill_rgb: cell.shading_rgb })
          const source = table.table.borders
          const lastMergeRow = rowIndex + cell.row_span - 1
          const edges: Array<{ edge: NativeDocxStrokeTableBorderCommandV1['edge']; border?: import('./nativeContract.js').NativeDocxTableBorderV1; x1: number; y1: number; x2: number; y2: number }> = [
            { edge: 'top', border: rowIndex === 0 ? source?.top : undefined, x1: x, y1: rowY, x2: x + cell.width_millipoints, y2: rowY },
            { edge: 'left', border: cell.column_ordinal === 0 ? source?.left : undefined, x1: x, y1: rowY, x2: x, y2: rowY + height },
            { edge: 'right', border: cell.column_ordinal + cell.grid_span === gridColumns ? source?.right : source?.inside_vertical, x1: x + cell.width_millipoints, y1: rowY, x2: x + cell.width_millipoints, y2: rowY + height },
            { edge: 'bottom', border: lastMergeRow === rows.length - 1 ? source?.bottom : source?.inside_horizontal, x1: x, y1: rowY + height, x2: x + cell.width_millipoints, y2: rowY + height },
          ]
          for (const edge of edges) if (edge.border?.style === 'single' && edge.border.color_rgb) target.borders.push({
            kind: 'stroke_table_border', id: `paint:table:${table.table.id}:${rowIndex}:${cellIndex}:${edge.edge}`, table_id: table.table.id, row_id: row.row_id, cell_id: cell.cell_id, edge: edge.edge,
            x1_millipoints: edge.x1, y1_millipoints: edge.y1, x2_millipoints: edge.x2, y2_millipoints: edge.y2,
            width_millipoints: edge.border.size_eighth_points * 125, stroke_rgb: edge.border.color_rgb,
          })
        }
      }
    }
  }
  return commands
}

function paintCommandID(placedLineID: string, fragmentID: string, glyphIndex: number): string {
  return `paint:${placedLineID}:${fragmentID}:${glyphIndex}`
}

function paintImageCommandID(placedLineID: string, fragmentID: string): string {
  return `paint:${placedLineID}:${fragmentID}:image`
}

function paintNoteSeparatorCommandID(placedLineID: string): string {
  return `paint:${placedLineID}:separator`
}

function noteSeparatorCommand(page: NativeDocxPaginatedPageV1, placed: NativeDocxPlacedLineV1, storyID: string): NativeDocxStrokeNoteSeparatorCommandV1 | undefined {
  const column = page.columns[placed.column_ordinal]
  if (!column || column.id !== placed.column_id || column.section_id !== placed.section_id || column.ordinal !== placed.column_ordinal) return undefined
  const x2 = Math.min(column.x_millipoints + column.width_millipoints, column.x_millipoints + 144_000)
  const y = placed.y_millipoints + Math.min(6_000, Math.floor(placed.height_millipoints / 2))
  if (!Number.isSafeInteger(x2) || !Number.isSafeInteger(y) || x2 <= column.x_millipoints) return undefined
  return {
    kind: 'stroke_note_separator',
    id: paintNoteSeparatorCommandID(placed.id),
    line_id: placed.line_id,
    story_id: storyID,
    x1_millipoints: column.x_millipoints,
    y1_millipoints: y,
    x2_millipoints: x2,
    y2_millipoints: y,
    width_millipoints: 750,
    stroke_rgb: '000000',
  }
}

function placedLineIDMatches(placedLineID: string, shapedLineID: string): boolean {
  if (placedLineID === `placed:${shapedLineID}`) return true
  const suffix = `:${shapedLineID}`
  if (!placedLineID.endsWith(suffix)) return false
  const prefix = placedLineID.slice(0, -suffix.length)
  return /^placed-note:[0-9]+:[0-9]+$/.test(prefix) || prefix.startsWith('placed:header:') || prefix.startsWith('placed:footer:')
}

/**
 * Compiles a complete all-or-nothing page projection. Any provider/source
 * refusal discards every accumulated page before a result is returned.
 */
export async function compileNativeDocxPagePaintV1(value: unknown, outlineProvider: NativeDocxGlyphOutlineProviderV1): Promise<CompileNativeDocxPagePaintV1Result> {
  const decoded = decodeNativeDocxPagePaintRequestV1(value)
  if (!decoded.ok) return decoded
  const providerResult = snapshotProvider(outlineProvider)
  if (!providerResult.ok) return providerResult
  const request = decoded.value
  const provider = providerResult.value
  if (request.outline_provider.provider_id !== provider.id || request.outline_provider.provider_revision !== provider.revision) {
    const expectedProvenance = requestProvenance(request, request.outline_provider.provider_id, request.outline_provider.provider_revision)
    return { ok: true, value: refusal(expectedProvenance, 'provider-mismatch', request.pagination_request.document.document_id, 'Injected outline provider id and revision do not match the page-paint request') }
  }
  const headerFooter = headerFooterLayout(request)
  const provenance = requestProvenance(request, provider.id, provider.revision, headerFooter)
  const pagination = request.pagination_request
  const layout = request.paginated_layout
  const documentID = pagination.document.document_id
  if (layout.status !== 'paginated') return { ok: true, value: refusal(provenance, 'upstream-refused', documentID, 'Native pagination refused; no visual projection was emitted') }
  if (headerFooter.status === 'refused') return { ok: true, value: refusal(provenance, 'unsupported-source', headerFooter.diagnostics[0]?.scope_id ?? documentID, headerFooter.diagnostics[0]?.message ?? 'Native header/footer layout refused') }
  const blockingPaginationDiagnostics = layout.diagnostics.filter((entry) => entry.code !== 'header-footer-selection-deferred' && !(entry.code === 'source-diagnostic' && entry.severity === 'deferred' && entry.source_code === 'page-control-deferred'))
  const blockingShapingDiagnostics = pagination.shaped_lines.diagnostics.filter((entry) => entry.code !== 'page-control-deferred' || entry.source_id !== undefined)
  if (blockingShapingDiagnostics.length > 0 || blockingPaginationDiagnostics.length > 0 || pagination.resolved_layout.diagnostics.length > 0) {
    const first = blockingShapingDiagnostics[0] ?? blockingPaginationDiagnostics[0] ?? pagination.resolved_layout.diagnostics[0]
    return { ok: true, value: refusal(provenance, 'unsupported-diagnostic', documentID, `Page-paint v1 requires diagnostic-free shaping/resolution and permits only the exact header/footer selection handoff from pagination${first ? `: ${first.code}` : ''}`) }
  }
  const qualifiedTables = qualifyNativeDocxTablesV1(pagination.document, pagination.resolved_layout)
  if (qualifiedTables.status !== 'qualified') return { ok: true, value: refusal(provenance, 'unsupported-source', qualifiedTables.diagnostics[0]!.scope_id, qualifiedTables.diagnostics[0]!.message) }
  const tableCommandsIndex = tableCommandsByPage(request, layout.pages, qualifiedTables.tables)
  if (!tableCommandsIndex) return { ok: true, value: refusal(provenance, 'resource-limit', documentID, 'Table paint indexing exceeded its bounded work or geometry contract') }

  const faces = new Map(request.font_manifest.faces.map((face) => [face.faceId, face]))
  const paragraphs = new Map(pagination.shaped_lines.paragraphs.map((paragraph) => [paragraph.paragraph_id, paragraph]))
  const resolvedRuns = new Map(pagination.resolved_layout.runs.map((run) => [run.run_id, run]))
  const nativeParagraphs = nativePaintParagraphs(request)
  const nativeRuns = new Map([...nativeParagraphs.values()].flatMap((paragraph) => paragraph.runs.map((run) => [run.id, run] as const)))
  const noteNumbers = new Map<string, string>()
  for (const page of layout.pages) for (const placement of page.note_stories ?? []) if (placement.note_role === 'content' && placement.number !== undefined) {
    const value = String(placement.number)
    const existing = noteNumbers.get(placement.story_id)
    if (existing !== undefined && existing !== value) return { ok: true, value: refusal(provenance, 'identity-mismatch', placement.story_id, 'One note story has inconsistent painted numbering') }
    noteNumbers.set(placement.story_id, value)
  }
  const resolvedParagraphs = new Map(pagination.resolved_layout.paragraphs.map((paragraph) => [paragraph.paragraph_id, paragraph]))
  const outlineCache = new Map<string, NativeDocxGlyphOutlineResultV1>()
  const unitsByFace = new Map<string, number>()
  const sourceIntervals = new Map<string, Array<{ start: number; end: number }>>()
  const sourceControlCounts = new Map<string, number>()
  const sourceHardBreakCounts = new Map<string, number>()
  const sourceImageCounts = new Map<string, number>()
  const selectedParagraphIDs = new Set<string>()
  const coveredLineIDs = new Set<string>()
  const coveredFragmentIDs = new Set<string>()
  const pages: NativeDocxPaintPageV1[] = []
  const mediaAssets = new Map(request.media_assets.map((asset) => [asset.id, asset]))
  const headerFooterByPageID = new Map(headerFooter.pages.map((entry) => [entry.page_id, entry]))
  let glyphCount = 0
  let pathCommandCount = 0
  let providerCalls = 0

  for (const page of layout.pages) {
    const tableCommands = tableCommandsIndex.get(page.id)
    if (!tableCommands) return { ok: true, value: refusal(provenance, 'identity-mismatch', documentID, 'Table geometry could not exact-join paginated cell lines') }
    const contentCommands: Array<NativeDocxFillGlyphPathCommandV1 | NativeDocxPaintInlineImageCommandV1 | NativeDocxStrokeNoteSeparatorCommandV1> = []
    const paintLines: NativeDocxPaintLineV1[] = []
    const headerFooterPage = headerFooterByPageID.get(page.id)
    if (!headerFooterPage) return { ok: true, value: refusal(provenance, 'incomplete-page', page.id, 'Header/footer layout does not exactly cover the paginated page') }
    const headerLines = headerFooterPage.lines.filter((line) => line.region === 'header')
    const footerLines = headerFooterPage.lines.filter((line) => line.region === 'footer')
    const noteStories = page.note_stories ?? []
    const noteStoryByLineID = new Map(noteStories.flatMap((story) => story.lines.map((line) => [line.id, story] as const)))
    const placedLines: Array<NativeDocxPlacedHeaderFooterLineV1 | NativeDocxPaginatedPageV1['lines'][number]> = [...headerLines, ...page.lines, ...noteStories.flatMap((story) => story.lines), ...footerLines]
    for (const placed of placedLines) {
      const paragraph = paragraphs.get(placed.paragraph_id)
      const line = paragraph?.lines[placed.source_line_ordinal]
      if (!paragraph || !line || line.id !== placed.line_id) return { ok: true, value: refusal(provenance, 'identity-mismatch', placed.line_id, 'Placed line does not exact-join its shaped line') }
      const noteStory = noteStoryByLineID.get(placed.id)
      const expectedStoryKind = 'region' in placed ? placed.region : noteStory?.story_kind ?? 'body'
      const expectedStoryID = 'story_id' in placed ? placed.story_id : noteStory?.story_id
      if (paragraph.story_kind !== expectedStoryKind || expectedStoryID !== undefined && paragraph.story_id !== expectedStoryID) return { ok: true, value: refusal(provenance, 'unsupported-source', paragraph.paragraph_id, 'Painted line does not exact-join its selected body/header/footer/note story') }
      selectedParagraphIDs.add(paragraph.paragraph_id)
      if (paragraph.alignment === 'distribute') return { ok: true, value: refusal(provenance, 'unsupported-source', paragraph.paragraph_id, 'Distributed character expansion is outside page-paint v1') }
      if (line.line_height_millipoints !== line.ascent_millipoints - line.descent_millipoints + line.line_gap_millipoints) return { ok: true, value: refusal(provenance, 'unsupported-source', line.id, 'Page-paint v1 requires natural shaped line height for an exact baseline') }
      if (line.hard_break_after && !coveredLineIDs.has(line.id)) sourceHardBreakCounts.set(line.hard_break_after.source_run_id, (sourceHardBreakCounts.get(line.hard_break_after.source_run_id) ?? 0) + 1)
      coveredLineIDs.add(line.id)
      let fragmentX = placed.x_millipoints
      const baselineY = placed.y_millipoints + line.ascent_millipoints
      if (!Number.isSafeInteger(baselineY) || Math.abs(baselineY) > DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints) return { ok: true, value: refusal(provenance, 'resource-limit', line.id, 'Line baseline exceeds the bounded paint coordinate range') }
      const firstCommand = contentCommands.length
      if (noteStory?.note_role === 'separator' && noteStory.lines[0]?.id === placed.id) {
        if (noteStory.lines.length !== 1 || line.fragments.length !== 0) return { ok: true, value: refusal(provenance, 'unsupported-source', noteStory.story_id, 'Instruction-only note separator must paint exactly one derived rule and no text or glyph commands') }
        const separator = noteSeparatorCommand(page, placed as NativeDocxPlacedLineV1, noteStory.story_id)
        if (!separator) return { ok: true, value: refusal(provenance, 'identity-mismatch', noteStory.story_id, 'Ordinary note separator cannot exact-join its placed line and column geometry') }
        contentCommands.push(separator)
      }
      for (const fragment of line.fragments) {
        let properties: NativeDocxResolvedRunPropertiesV1
        if (fragment.source_kind === 'list-marker') {
          const numbering = resolvedParagraphs.get(paragraph.paragraph_id)?.numbering
          const marker = paragraph.list_marker
          if (!numbering || !marker || fragment.source_id !== paragraph.paragraph_id || marker.marker_id !== numbering.marker_id || marker.text !== numbering.resolved_text) return { ok: true, value: refusal(provenance, 'identity-mismatch', fragment.id, 'List-marker fragment does not exact-join its resolved and shaped marker provenance') }
          const authoredMarkerText = fragment.start_utf16 <= fragment.end_utf16 && fragment.end_utf16 <= marker.text.length && marker.text.slice(fragment.start_utf16, fragment.end_utf16) === fragment.text
          const virtualPrefix = fragment.text === '' && fragment.start_utf16 === 0 && fragment.end_utf16 === 0 && fragment.glyphs.length === 0
          const suffixTab = marker.suffix === 'tab' && fragment.text === '\t' && fragment.start_utf16 === 0 && fragment.end_utf16 === 0 && fragment.glyphs.length === 0
          const suffixSpace = marker.suffix === 'space' && fragment.text === ' ' && fragment.start_utf16 === marker.text.length && fragment.end_utf16 === marker.text.length + 1
          if (!authoredMarkerText && !virtualPrefix && !suffixTab && !suffixSpace) return { ok: true, value: refusal(provenance, 'identity-mismatch', fragment.id, 'List-marker text range must exactly cover resolved marker text or its declared suffix geometry') }
          if (authoredMarkerText && fragment.text.length > 0 && fragment.glyphs.length === 0) return { ok: true, value: refusal(provenance, 'missing-glyph', fragment.id, 'A non-empty visible list marker cannot paint without glyphs') }
          properties = numbering.marker_properties
        } else {
          const resolved = resolvedRuns.get(fragment.source_id)
          const nativeRun = nativeRuns.get(fragment.source_id)
          if (!resolved || !nativeRun || resolved.paragraph_id !== paragraph.paragraph_id) return { ok: true, value: refusal(provenance, 'identity-mismatch', fragment.id, 'Fragment source does not exact-join a native and resolved run') }
          properties = resolved.properties
          if (fragment.source_kind === 'run') {
            const noteMarker = nativeRun.kind === 'reference' && nativeRun.reference && (nativeRun.reference.kind === 'footnote' || nativeRun.reference.kind === 'endnote') ? noteNumbers.get(nativeRun.reference.target_id) : undefined
            const exactText = nativeRun.kind === 'text' && nativeRun.text !== undefined && nativeRun.text.slice(fragment.start_utf16, fragment.end_utf16) === fragment.text
            const exactMarker = noteMarker !== undefined && fragment.start_utf16 === 0 && fragment.end_utf16 === noteMarker.length && fragment.text === noteMarker
            if (!exactText && !exactMarker) return { ok: true, value: refusal(provenance, 'identity-mismatch', fragment.id, 'Run fragment text and UTF-16 range must exactly match native text or its placed note number') }
            if (fragment.text.length > 0 && fragment.glyphs.length === 0) return { ok: true, value: refusal(provenance, 'missing-glyph', fragment.id, 'A non-empty visible run fragment cannot paint without glyphs') }
            if (!coveredFragmentIDs.has(fragment.id)) {
              const intervals = sourceIntervals.get(nativeRun.id) ?? []
              intervals.push({ start: fragment.start_utf16, end: fragment.end_utf16 })
              sourceIntervals.set(nativeRun.id, intervals)
            }
          } else if (fragment.source_kind === 'tab') {
            if (nativeRun.kind !== 'control' || nativeRun.control !== 'tab' || fragment.text !== '\t' || fragment.start_utf16 !== 0 || fragment.end_utf16 !== 0 || fragment.glyphs.length !== 0) return { ok: true, value: refusal(provenance, 'identity-mismatch', fragment.id, 'Tab fragment must exactly match a glyphless native tab control') }
            if (!coveredFragmentIDs.has(fragment.id)) sourceControlCounts.set(nativeRun.id, (sourceControlCounts.get(nativeRun.id) ?? 0) + 1)
          } else if (fragment.source_kind === 'image') {
            if (nativeRun.kind !== 'drawing' || !nativeRun.drawing || fragment.text !== '' || fragment.start_utf16 !== 0 || fragment.end_utf16 !== 0 || fragment.glyphs.length !== 0 || fragment.face_id !== undefined || fragment.whitespace) return { ok: true, value: refusal(provenance, 'identity-mismatch', fragment.id, 'Image fragment must exactly match one glyphless native drawing run') }
            const qualified = qualifyNativeDocxInlineImageV1(pagination.document, nativeRun.id, nativeRun.drawing)
            if (!qualified.ok) return { ok: true, value: refusal(provenance, qualified.code === 'resource-limit' ? 'resource-limit' : 'unsupported-source', fragment.id, qualified.message) }
            const image = qualified.value
            const asset = mediaAssets.get(image.asset_id)
            if (!asset || asset.part_name !== image.part_name || asset.content_type !== image.content_type || asset.content_digest !== image.content_digest) return { ok: true, value: refusal(provenance, 'identity-mismatch', fragment.id, 'Image fragment does not exact-join one canonical content-addressed media asset') }
            if (fragment.advance_inline_millipoints !== image.width_millipoints || fragment.ascent_millipoints !== image.height_millipoints || fragment.descent_millipoints !== 0 || fragment.line_gap_millipoints !== 0) return { ok: true, value: refusal(provenance, 'identity-mismatch', fragment.id, 'Image fragment geometry changed after exact EMU projection') }
            const y = baselineY - image.height_millipoints
            if (!Number.isSafeInteger(y) || y < 0) return { ok: true, value: refusal(provenance, 'resource-limit', fragment.id, 'Image placement exceeds bounded non-negative page coordinates') }
            contentCommands.push({
              kind: 'paint_inline_image', id: paintImageCommandID(placed.id, fragment.id), line_id: line.id, fragment_id: fragment.id, source_id: fragment.source_id,
              drawing_id: image.drawing_id, asset_id: image.asset_id,
              x_millipoints: fragmentX, y_millipoints: y, width_millipoints: image.width_millipoints, height_millipoints: image.height_millipoints,
              source_crop: image.source_crop, transform: image.transform,
            })
            fragmentX += fragment.advance_inline_millipoints
            if (!Number.isSafeInteger(fragmentX)) return { ok: true, value: refusal(provenance, 'resource-limit', fragment.id, 'Image cursor exceeds safe integer coordinates') }
            if (!coveredFragmentIDs.has(fragment.id)) sourceImageCounts.set(nativeRun.id, (sourceImageCounts.get(nativeRun.id) ?? 0) + 1)
            coveredFragmentIDs.add(fragment.id)
            continue
          }
        }
        coveredFragmentIDs.add(fragment.id)
        if (properties.underline && properties.underline !== 'none' || properties.highlight && properties.highlight !== 'none') return { ok: true, value: refusal(provenance, 'unsupported-source', fragment.source_id, 'Underline and highlight paint are outside page-paint v1') }
        if (fragment.glyphs.length === 0) {
          fragmentX += fragment.advance_inline_millipoints
          if (!Number.isSafeInteger(fragmentX)) return { ok: true, value: refusal(provenance, 'resource-limit', fragment.id, 'Fragment cursor exceeds safe integer coordinates') }
          continue
        }
        const fontSize = properties.font_size_half_points === undefined ? undefined : properties.font_size_half_points * 500
        if (!fontSize || !Number.isSafeInteger(fontSize)) return { ok: true, value: refusal(provenance, 'missing-font', fragment.source_id, 'Resolved run has no bounded font size') }
        const fill = properties.color ?? '000000'
        if (!RGB.test(fill)) return { ok: true, value: refusal(provenance, 'unsupported-source', fragment.source_id, 'Resolved run color is not an explicit RGB value') }
        let glyphX = fragmentX
        let glyphAdvance = 0
        for (const [glyphIndex, glyph] of fragment.glyphs.entries()) {
          glyphCount += 1
          if (glyphCount > DOCX_PAGE_PAINT_LIMITS.maxGlyphs) return { ok: true, value: refusal(provenance, 'resource-limit', line.id, `Glyphs exceed ${DOCX_PAGE_PAINT_LIMITS.maxGlyphs}`) }
          if (glyph.advance_y_millipoints !== 0) return { ok: true, value: refusal(provenance, 'unsupported-source', fragment.id, 'Vertical glyph advances are outside page-paint v1') }
          if (glyph.glyph_id === 0) return { ok: true, value: refusal(provenance, 'missing-glyph', fragment.id, 'Glyph id zero is the missing-glyph sentinel and cannot be painted') }
          if (glyph.advance_x_millipoints < 0) return { ok: true, value: refusal(provenance, 'unsupported-source', fragment.id, 'Negative horizontal glyph advances are outside page-paint v1') }
          const manifestFace = fragment.face_id ? faces.get(fragment.face_id) : undefined
          if (!manifestFace || !manifestFace.source.contentDigest || !SHA256.test(manifestFace.source.contentDigest) || manifestFace.source.kind === 'system') return { ok: true, value: refusal(provenance, 'missing-font', fragment.id, 'Every painted glyph requires a non-system, manifest-backed content-addressed face') }
          const face: NativeDocxContentAddressedFaceV1 = {
            face_id: manifestFace.faceId,
            content_digest: manifestFace.source.contentDigest,
            ...(manifestFace.source.collectionIndex !== undefined ? { collection_index: manifestFace.source.collectionIndex } : {}),
          }
          const cacheKey = `${face.content_digest}\u0000${face.collection_index ?? ''}\u0000${glyph.glyph_id}`
          let outline = outlineCache.get(cacheKey)
          if (!outline) {
            if (outlineCache.size >= DOCX_PAGE_PAINT_LIMITS.maxUniqueGlyphOutlines || providerCalls >= DOCX_PAGE_PAINT_LIMITS.maxProviderCalls) return { ok: true, value: refusal(provenance, 'resource-limit', fragment.id, 'Unique glyph-outline/provider-call budget was exceeded') }
            providerCalls += 1
            let live: unknown
            try { live = await provider.get(Object.freeze({ face: Object.freeze({ ...face }), glyph_id: glyph.glyph_id })) } catch (error) {
              return { ok: true, value: refusal(provenance, 'provider-failure', fragment.id, `Glyph outline provider failed: ${boundedProviderError(error)}`) }
            }
            if (!providerStable(provider)) return { ok: true, value: refusal(provenance, 'provider-mismatch', fragment.id, 'Outline provider changed its snapshotted identity during compilation') }
            const providerPreflight = preflightWire(live, 'glyph outline provider result', DOCX_PAGE_PAINT_LIMITS.maxPathCommandsPerGlyph * 8 + 100, DOCX_PAGE_PAINT_LIMITS.maxPathCommandsPerGlyph)
            if (providerPreflight.length > 0) {
              const code: NativeDocxPagePaintDiagnosticCode = providerPreflight.some((entry) => entry.code === 'LIMIT_EXCEEDED') ? 'resource-limit' : 'invalid-provider-output'
              return { ok: true, value: refusal(provenance, code, fragment.id, 'Outline provider returned recursive, unreadable, or resource-unbounded output') }
            }
            const owned = safeClone(live)
            outline = owned === undefined ? undefined : captureOutline(owned, face, glyph.glyph_id)
            if (!outline) return { ok: true, value: refusal(provenance, 'invalid-provider-output', fragment.id, 'Outline provider returned malformed, mismatched, unbounded, or invalid path output') }
            outlineCache.set(cacheKey, outline)
          }
          if (outline.status === 'refused') {
            const code: NativeDocxPagePaintDiagnosticCode = outline.code === 'missing-font' ? 'missing-font' : outline.code === 'missing-glyph' ? 'missing-glyph' : 'provider-refusal'
            return { ok: true, value: refusal(provenance, code, fragment.id, outline.message) }
          }
          if (outline.status === 'empty' && !fragment.whitespace) return { ok: true, value: refusal(provenance, 'missing-glyph', fragment.id, 'A visible non-whitespace glyph cannot use an empty outline') }
          const faceUnitsKey = `${face.content_digest}\u0000${face.collection_index ?? ''}`
          const knownUnits = unitsByFace.get(faceUnitsKey)
          if (knownUnits !== undefined && knownUnits !== outline.units_per_em) return { ok: true, value: refusal(provenance, 'provider-mismatch', fragment.id, 'Provider changed units_per_em for one content-addressed face') }
          unitsByFace.set(faceUnitsKey, outline.units_per_em)
          const originX = glyphX + glyph.offset_x_millipoints
          const originY = baselineY - glyph.offset_y_millipoints
          if (!Number.isSafeInteger(originX) || !Number.isSafeInteger(originY)) return { ok: true, value: refusal(provenance, 'resource-limit', fragment.id, 'Glyph origin exceeds safe integer coordinates') }
          const path = outline.status === 'empty' ? [] : placePath(outline.path, originX, originY, fontSize, outline.units_per_em)
          if (!path) return { ok: true, value: refusal(provenance, 'invalid-path', fragment.id, 'Scaled glyph path exceeds bounded integer page coordinates') }
          pathCommandCount += path.length
          if (pathCommandCount > DOCX_PAGE_PAINT_LIMITS.maxPathCommands) return { ok: true, value: refusal(provenance, 'resource-limit', fragment.id, `Paint path commands exceed ${DOCX_PAGE_PAINT_LIMITS.maxPathCommands}`) }
          contentCommands.push({
            kind: 'fill_glyph_path',
            id: paintCommandID(placed.id, fragment.id, glyphIndex),
            line_id: line.id,
            fragment_id: fragment.id,
            source_id: fragment.source_id,
            glyph_index: glyphIndex,
            face,
            glyph_id: glyph.glyph_id,
            font_size_millipoints: fontSize,
            fill_rgb: fill,
            fill_rule: 'nonzero',
            outline_kind: outline.status === 'empty' ? 'empty' : 'path',
            path,
          })
          glyphX += glyph.advance_x_millipoints
          glyphAdvance += glyph.advance_x_millipoints
          if (!Number.isSafeInteger(glyphX)) return { ok: true, value: refusal(provenance, 'resource-limit', fragment.id, 'Glyph cursor exceeds safe integer coordinates') }
        }
        if (fragment.glyphs.length > 0 && glyphAdvance !== fragment.advance_inline_millipoints) return { ok: true, value: refusal(provenance, 'identity-mismatch', fragment.id, 'Glyph advances must exactly sum to the shaped fragment advance') }
        fragmentX += fragment.advance_inline_millipoints
        if (!Number.isSafeInteger(fragmentX)) return { ok: true, value: refusal(provenance, 'resource-limit', fragment.id, 'Fragment cursor exceeds safe integer coordinates') }
      }
      if (fragmentX !== placed.x_millipoints + line.advance_inline_millipoints) return { ok: true, value: refusal(provenance, 'identity-mismatch', line.id, 'Fragment advances must exactly sum to the shaped line advance') }
      paintLines.push({
        placed_line_id: placed.id,
        line_id: line.id,
        paragraph_id: paragraph.paragraph_id,
        region: expectedStoryKind,
        section_id: 'section_id' in placed ? placed.section_id : page.section_id,
        ...('column_id' in placed ? { column_id: placed.column_id, column_ordinal: placed.column_ordinal } : {}),
        source_line_ordinal: placed.source_line_ordinal,
        x_millipoints: placed.x_millipoints,
        y_millipoints: placed.y_millipoints,
        width_millipoints: placed.width_millipoints,
        height_millipoints: placed.height_millipoints,
        baseline_y_millipoints: baselineY,
        command_ids: contentCommands.slice(firstCommand).map((command) => command.id),
      })
    }
    pages.push({
      id: page.id,
      ordinal: page.ordinal,
      section_id: page.section_id,
      section_ids: [...page.section_ids],
      kind: page.kind,
      width_millipoints: page.width_millipoints,
      height_millipoints: page.height_millipoints,
      body_box: { ...page.body_box },
      columns: page.columns.map((column) => ({ ...column })),
      background_rgb: 'FFFFFF',
      clip_box: { x_millipoints: 0, y_millipoints: 0, width_millipoints: page.width_millipoints, height_millipoints: page.height_millipoints },
      lines: paintLines,
      commands: [...tableCommands.fills, ...contentCommands, ...tableCommands.borders],
    })
  }
  for (const paragraphID of selectedParagraphIDs) for (const nativeRun of nativeParagraphs.get(paragraphID)?.runs ?? []) {
    const resolved = resolvedRuns.get(nativeRun.id)
    if (!resolved || resolved.properties.hidden) continue
    if (nativeRun.kind === 'text' && nativeRun.text !== undefined) {
      const intervals = (sourceIntervals.get(nativeRun.id) ?? []).sort((left, right) => left.start - right.start || left.end - right.end)
      let cursor = 0
      for (const interval of intervals) {
        if (interval.start !== cursor || interval.end <= interval.start) return { ok: true, value: refusal(provenance, 'identity-mismatch', nativeRun.id, 'Painted visual clusters must exactly partition their native text run in logical order') }
        cursor = interval.end
      }
      if (cursor !== nativeRun.text.length) return { ok: true, value: refusal(provenance, 'identity-mismatch', nativeRun.id, 'Painted visual clusters must completely cover their native text run') }
    } else if (nativeRun.kind === 'control' && nativeRun.control === 'tab' && sourceControlCounts.get(nativeRun.id) !== 1) {
      return { ok: true, value: refusal(provenance, 'identity-mismatch', nativeRun.id, 'A painted native tab must have exactly one visual fragment') }
    } else if (nativeRun.kind === 'control' && nativeRun.control === 'line-break' && sourceHardBreakCounts.get(nativeRun.id) !== 1) {
      return { ok: true, value: refusal(provenance, 'identity-mismatch', nativeRun.id, 'A painted native line break must have exactly one hard-break attestation') }
    } else if (nativeRun.kind === 'drawing' && sourceImageCounts.get(nativeRun.id) !== 1) {
      return { ok: true, value: refusal(provenance, 'identity-mismatch', nativeRun.id, 'A painted native image must have exactly one visual fragment') }
    } else if (nativeRun.kind === 'control' && nativeRun.control === 'soft-hyphen') {
      return { ok: true, value: refusal(provenance, 'unsupported-source', nativeRun.id, 'Conditional soft-hyphen painting is outside page-paint v1') }
    }
  }
  if (pages.length !== layout.pages.length) return { ok: true, value: refusal(provenance, 'incomplete-page', documentID, 'Paint output did not cover every paginated page') }
  const success: NativeDocxPagePaintSuccessV1 = { protocol: DOCX_PAGE_PAINT_PROTOCOL, version: DOCX_PAGE_PAINT_VERSION, status: 'painted', provenance, diagnostics: [], resources: request.media_assets.map((asset) => ({ ...asset })), pages }
  const outputBound = preflightWire(success, 'generated page-paint output')
  if (outputBound.some((entry) => entry.code === 'LIMIT_EXCEEDED')) return { ok: true, value: refusal(provenance, 'resource-limit', documentID, `Generated page-paint output exceeds ${DOCX_PAGE_PAINT_LIMITS.maxOutputNodes} bounded values`) }
  if (outputBound.length > 0) return { ok: true, value: refusal(provenance, 'invalid-path', documentID, 'Generated page-paint output failed its JSON wire invariant') }
  const decodedSuccess = decodeNativeDocxPagePaintV1(success)
  if (!decodedSuccess.ok) {
    const first = decodedSuccess.issues[0]
    return { ok: true, value: refusal(provenance, 'invalid-path', documentID, `Generated page-paint output failed its canonical output contract${first ? `: ${first.path} ${first.message}` : ''}`) }
  }
  return decodedSuccess
}

function validateFace(value: unknown, path: string, issues: NativeDocxValidationIssue[]): NativeDocxContentAddressedFaceV1 | undefined {
  const entry = exactUnionObject(value, path, DOCX_PAGE_PAINT_V1_BINDING_FIELDS.FaceV1, issues)
  if (!entry) return undefined
  const faceID = stringValue(entry.face_id, `${path}/face_id`, issues)
  const digest = stringValue(entry.content_digest, `${path}/content_digest`, issues, SHA256, 71)
  let collection: number | undefined
  if (entry.collection_index !== undefined) collection = integer(entry.collection_index, `${path}/collection_index`, issues, 0, 65_535)
  return faceID && digest ? { face_id: faceID, content_digest: digest, ...(collection !== undefined ? { collection_index: collection } : {}) } : undefined
}

function validatePaintPath(value: unknown, path: string, outlineKind: string | undefined, issues: NativeDocxValidationIssue[], state: { pathCommands: number; nodes: number }): void {
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
    if (headerFooter) { if (headerFooter.protocol !== DOCX_HEADER_FOOTER_LAYOUT_PROTOCOL) add(issues, 'UNSUPPORTED_PROTOCOL', '/provenance/header_footer_layout/protocol', `must equal ${DOCX_HEADER_FOOTER_LAYOUT_PROTOCOL}`); integer(headerFooter.version, '/provenance/header_footer_layout/version', issues, 1, 1); stringValue(headerFooter.sha256, '/provenance/header_footer_layout/sha256', issues, SHA256, 71) }
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

/** Strict output decoding plus exact request/page/glyph identity joins. */
export function decodeNativeDocxPagePaintForRequestV1(value: unknown, requestValue: unknown, outlineProviderIdentity: { provider_id: string; provider_revision: string }): DecodeNativeDocxPagePaintV1Result {
  const request = decodeNativeDocxPagePaintRequestV1(requestValue)
  const output = decodeNativeDocxPagePaintV1(value)
  const issues: NativeDocxValidationIssue[] = []
  if (!request.ok) issues.push(...request.issues.map((entry) => ({ ...entry, path: `/request${entry.path}` })))
  if (!output.ok) issues.push(...output.issues.map((entry) => ({ ...entry, path: `/output${entry.path}` })))
  if (!request.ok || !output.ok) return { ok: false, issues: issues.slice(0, DOCX_NATIVE_LIMITS.maxIssues) }
  if (outlineProviderIdentity.provider_id !== request.value.outline_provider.provider_id || outlineProviderIdentity.provider_revision !== request.value.outline_provider.provider_revision) add(issues, 'BROKEN_REFERENCE', '/provider', 'outline provider identity must exactly match the page-paint request attestation')
  const expectedProvenance = requestProvenance(request.value, outlineProviderIdentity.provider_id, outlineProviderIdentity.provider_revision)
  if (!sameWire(output.value.provenance, expectedProvenance)) add(issues, 'BROKEN_REFERENCE', '/output/provenance', 'paint provenance must exactly match the joined page-paint request and outline provider')
  if (output.value.status === 'painted') {
    const headerFooter = headerFooterLayout(request.value)
    if (headerFooter.status !== 'placed') add(issues, 'BROKEN_REFERENCE', '/output/status', 'painted output cannot derive from refused header/footer placement')
    if (!sameWire(output.value.resources, request.value.media_assets)) add(issues, 'BROKEN_REFERENCE', '/output/resources', 'paint resources must exactly equal the canonical request media inventory')
    if (request.value.paginated_layout.status !== 'paginated') add(issues, 'BROKEN_REFERENCE', '/output/status', 'painted output cannot derive from refused pagination')
    else {
      if (output.value.pages.length !== request.value.paginated_layout.pages.length) add(issues, 'BROKEN_REFERENCE', '/output/pages', 'paint pages must exactly cover paginated pages')
      const shapedParagraphs = new Map(request.value.pagination_request.shaped_lines.paragraphs.map((paragraph) => [paragraph.paragraph_id, paragraph]))
      const resolvedRuns = new Map(request.value.pagination_request.resolved_layout.runs.map((run) => [run.run_id, run]))
      const nativeStories = [request.value.pagination_request.document.body, ...request.value.pagination_request.document.headers, ...request.value.pagination_request.document.footers, ...request.value.pagination_request.document.notes]
      const nativeRuns = new Map(nativeStories.flatMap((story) => story.blocks.flatMap((block) => block.paragraph?.runs ?? [])).map((run) => [run.id, run]))
      const resolvedParagraphs = new Map(request.value.pagination_request.resolved_layout.paragraphs.map((paragraph) => [paragraph.paragraph_id, paragraph]))
      const manifestFaces = new Map(request.value.font_manifest.faces.map((face) => [face.faceId, face]))
      const expectedGlyphs: Array<{ pageIndex: number; pageID: string; placedLineID: string; lineID: string; fragmentID: string; sourceID: string; glyphIndex: number; glyphID: number; face?: NativeDocxContentAddressedFaceV1; fontSize?: number; fill: string }> = []
      const expectedImages: Array<{ pageIndex: number; pageID: string; placedLineID: string; lineID: string; fragmentID: string; sourceID: string; drawingID: string; assetID: string; x: number; y: number; width: number; height: number }> = []
      const expectedSeparators: Array<{ pageIndex: number; command: NativeDocxStrokeNoteSeparatorCommandV1 }> = []
      const headerFooterByPageID = headerFooter.status === 'placed' ? new Map<string, NativeDocxHeaderFooterPageLayoutV1>(headerFooter.pages.map((entry) => [entry.page_id, entry])) : new Map<string, NativeDocxHeaderFooterPageLayoutV1>()
      const qualifiedTables = qualifyNativeDocxTablesV1(request.value.pagination_request.document, request.value.pagination_request.resolved_layout)
      const expectedTableByPageID = qualifiedTables.status === 'qualified'
        ? tableCommandsByPage(request.value, request.value.paginated_layout.pages, qualifiedTables.tables)
        : undefined
      if (!expectedTableByPageID) add(issues, 'BROKEN_REFERENCE', '/output/pages', 'table paint source could not be bounded and indexed for replay')
      const sourceLinesByPageID = new Map<string, Array<NativeDocxPlacedHeaderFooterLineV1 | NativeDocxPlacedLineV1>>()
      const noteStoryByPageLineID = new Map<string, Map<string, NativeDocxPlacedNoteStoryV1>>()
      request.value.paginated_layout.pages.forEach((page, pageIndex) => {
        const headerFooterPage = headerFooterByPageID.get(page.id)
        const placements = headerFooterPage ? [...headerFooterPage.lines.filter((line) => line.region === 'header'), ...page.lines, ...(page.note_stories ?? []).flatMap((story) => story.lines), ...headerFooterPage.lines.filter((line) => line.region === 'footer')] : [...page.lines, ...(page.note_stories ?? []).flatMap((story) => story.lines)]
        sourceLinesByPageID.set(page.id, placements)
        const noteStoryByLineID = new Map((page.note_stories ?? []).flatMap((story) => story.lines.map((line) => [line.id, story] as const)))
        noteStoryByPageLineID.set(page.id, noteStoryByLineID)
        placements.forEach((placed) => {
          const line = shapedParagraphs.get(placed.paragraph_id)?.lines[placed.source_line_ordinal]
          const indexedStory = noteStoryByLineID.get(placed.id)
          const separatorStory = indexedStory?.note_role === 'separator' ? indexedStory : undefined
          if (separatorStory) {
            const separator = noteSeparatorCommand(page, placed as NativeDocxPlacedLineV1, separatorStory.story_id)
            if (separator) expectedSeparators.push({ pageIndex, command: separator })
          }
          let fragmentX = placed.x_millipoints
          line?.fragments.forEach((fragment) => {
            const resolved = resolvedRuns.get(fragment.source_id)
            const properties = fragment.source_kind === 'list-marker' ? resolvedParagraphs.get(placed.paragraph_id)?.numbering?.marker_properties : resolved?.properties
            const manifestFace = fragment.face_id ? manifestFaces.get(fragment.face_id) : undefined
            const face = manifestFace?.source.contentDigest ? { face_id: manifestFace.faceId, content_digest: manifestFace.source.contentDigest, ...(manifestFace.source.collectionIndex !== undefined ? { collection_index: manifestFace.source.collectionIndex } : {}) } : undefined
            fragment.glyphs.forEach((glyph, glyphIndex) => expectedGlyphs.push({ pageIndex, pageID: page.id, placedLineID: placed.id, lineID: line.id, fragmentID: fragment.id, sourceID: fragment.source_id, glyphIndex, glyphID: glyph.glyph_id, face, fontSize: properties?.font_size_half_points === undefined ? undefined : properties.font_size_half_points * 500, fill: properties?.color ?? '000000' }))
            if (fragment.source_kind === 'image') {
              const run = nativeRuns.get(fragment.source_id)
              const qualified = run?.drawing ? qualifyNativeDocxInlineImageV1(request.value.pagination_request.document, run.id, run.drawing) : undefined
              if (qualified?.ok) expectedImages.push({ pageIndex, pageID: page.id, placedLineID: placed.id, lineID: line.id, fragmentID: fragment.id, sourceID: fragment.source_id, drawingID: qualified.value.drawing_id, assetID: qualified.value.asset_id, x: fragmentX, y: placed.y_millipoints + line.ascent_millipoints - qualified.value.height_millipoints, width: qualified.value.width_millipoints, height: qualified.value.height_millipoints })
            }
            fragmentX += fragment.advance_inline_millipoints
          })
        })
      })
      const actualGlyphs = output.value.pages.flatMap((page, pageIndex) => page.commands.flatMap((command, commandIndex) => command.kind === 'fill_glyph_path' ? [{ page, pageIndex, commandIndex, command }] : []))
      const actualImages = output.value.pages.flatMap((page, pageIndex) => page.commands.flatMap((command, commandIndex) => command.kind === 'paint_inline_image' ? [{ page, pageIndex, commandIndex, command }] : []))
      const actualSeparators = output.value.pages.flatMap((page, pageIndex) => page.commands.flatMap((command) => command.kind === 'stroke_note_separator' ? [{ pageIndex, command }] : []))
      if (actualGlyphs.length !== expectedGlyphs.length) add(issues, 'BROKEN_REFERENCE', '/output/pages', 'paint commands must exactly cover every shaped glyph once')
      if (actualImages.length !== expectedImages.length) add(issues, 'BROKEN_REFERENCE', '/output/pages', 'paint commands must exactly cover every qualified inline image once')
      if (!sameWire(actualSeparators, expectedSeparators)) add(issues, 'BROKEN_REFERENCE', '/output/pages', 'paint commands must exactly derive one deterministic rule from each placed ordinary note separator')
      output.value.pages.forEach((page, pageIndex) => {
        const source = request.value.paginated_layout.status === 'paginated' ? request.value.paginated_layout.pages[pageIndex] : undefined
        if (!source || page.id !== source.id || page.ordinal !== source.ordinal || page.section_id !== source.section_id || !sameWire(page.section_ids, source.section_ids) || !sameWire(page.columns, source.columns) || page.kind !== source.kind || page.width_millipoints !== source.width_millipoints || page.height_millipoints !== source.height_millipoints || !sameWire(page.body_box, source.body_box) || page.background_rgb !== 'FFFFFF' || !sameWire(page.clip_box, { x_millipoints: 0, y_millipoints: 0, width_millipoints: source.width_millipoints, height_millipoints: source.height_millipoints })) add(issues, 'BROKEN_REFERENCE', `/output/pages/${pageIndex}`, 'paint page section/column identity, background, clip, and geometry must exactly match pagination')
        const sourceLines = source ? sourceLinesByPageID.get(source.id) ?? [] : []
        const noteStoryByPlacedLineID = source ? noteStoryByPageLineID.get(source.id) ?? new Map() : new Map()
        if (source && page.lines.length !== sourceLines.length) add(issues, 'BROKEN_REFERENCE', `/output/pages/${pageIndex}/lines`, 'paint lines must exactly cover every body/header/footer/note placed line including glyphless lines')
        if (source) page.lines.forEach((line, lineIndex) => {
          const placed = sourceLines[lineIndex]
          const shaped = placed ? shapedParagraphs.get(placed.paragraph_id)?.lines[placed.source_line_ordinal] : undefined
          const indexedStory = placed ? noteStoryByPlacedLineID.get(placed.id) : undefined
          const separatorStory = indexedStory?.note_role === 'separator' ? indexedStory : undefined
          const expectedCommandIDs = [
            ...(separatorStory && placed ? [paintNoteSeparatorCommandID(placed.id)] : []),
            ...(shaped?.fragments.flatMap((fragment) => fragment.source_kind === 'image' ? [paintImageCommandID(placed!.id, fragment.id)] : fragment.glyphs.map((_, glyphIndex) => paintCommandID(placed!.id, fragment.id, glyphIndex))) ?? []),
          ]
          const expectedRegion = placed && 'region' in placed ? placed.region : placed ? noteStoryByPlacedLineID.get(placed.id)?.story_kind ?? 'body' : 'body'
          const expectedSectionID = placed && 'section_id' in placed ? placed.section_id : source.section_id
          const expectedColumnID = placed && 'column_id' in placed ? placed.column_id : undefined
          const expectedColumnOrdinal = placed && 'column_ordinal' in placed ? placed.column_ordinal : undefined
          if (!placed || !shaped || line.placed_line_id !== placed.id || line.line_id !== placed.line_id || line.paragraph_id !== placed.paragraph_id || line.region !== expectedRegion || line.section_id !== expectedSectionID || line.column_id !== expectedColumnID || line.column_ordinal !== expectedColumnOrdinal || line.source_line_ordinal !== placed.source_line_ordinal || line.x_millipoints !== placed.x_millipoints || line.y_millipoints !== placed.y_millipoints || line.width_millipoints !== placed.width_millipoints || line.height_millipoints !== placed.height_millipoints || line.baseline_y_millipoints !== placed.y_millipoints + shaped.ascent_millipoints || !sameWire(line.command_ids, expectedCommandIDs)) add(issues, 'BROKEN_REFERENCE', `/output/pages/${pageIndex}/lines/${lineIndex}`, 'paint line region/section/column identity, geometry, baseline, and glyph command coverage must exactly match its placed and shaped line')
        })
        if (source) {
          const expectedTable = expectedTableByPageID?.get(source.id)
          const actualTable = page.commands.filter((command) => command.kind === 'fill_table_cell' || command.kind === 'stroke_table_border')
          if (!expectedTable || !sameWire(actualTable, [...expectedTable.fills, ...expectedTable.borders])) add(issues, 'BROKEN_REFERENCE', `/output/pages/${pageIndex}/commands`, 'table fill and border commands must exactly derive from qualified source geometry and page placement')
        }
      })
      actualGlyphs.forEach((actual, index) => {
        const expected = expectedGlyphs[index]
        if (!expected || actual.pageIndex !== expected.pageIndex || actual.page.id !== expected.pageID || actual.command.line_id !== expected.lineID || actual.command.fragment_id !== expected.fragmentID || actual.command.source_id !== expected.sourceID || actual.command.glyph_index !== expected.glyphIndex || actual.command.glyph_id !== expected.glyphID || !sameWire(actual.command.face, expected.face) || actual.command.font_size_millipoints !== expected.fontSize || actual.command.fill_rgb !== expected.fill || actual.command.fill_rule !== 'nonzero' || actual.command.id !== paintCommandID(expected.placedLineID, expected.fragmentID, expected.glyphIndex)) add(issues, 'BROKEN_REFERENCE', `/output/pages/${actual.pageIndex}/commands/${actual.commandIndex}`, 'paint command must exact-join its page, line, fragment, face content address, glyph, size, and fill identity')
      })
      actualImages.forEach((actual, index) => {
        const expected = expectedImages[index]
        if (!expected || actual.pageIndex !== expected.pageIndex || actual.page.id !== expected.pageID || actual.command.line_id !== expected.lineID || actual.command.fragment_id !== expected.fragmentID || actual.command.source_id !== expected.sourceID || actual.command.drawing_id !== expected.drawingID || actual.command.asset_id !== expected.assetID || actual.command.x_millipoints !== expected.x || actual.command.y_millipoints !== expected.y || actual.command.width_millipoints !== expected.width || actual.command.height_millipoints !== expected.height || actual.command.id !== paintImageCommandID(expected.placedLineID, expected.fragmentID)) add(issues, 'BROKEN_REFERENCE', `/output/pages/${actual.pageIndex}/commands/${actual.commandIndex}`, 'image command must exact-join its page, line, fragment, drawing, media digest identity, and exact geometry')
      })
    }
  }
  return issues.length > 0 ? { ok: false, issues: issues.slice(0, DOCX_NATIVE_LIMITS.maxIssues) } : output
}

/** Recompiles with the exact provider and proves the entire canonical output, including paths and refusal. */
export async function validateNativeDocxPagePaintForRequestV1(value: unknown, requestValue: unknown, provider: NativeDocxGlyphOutlineProviderV1): Promise<DecodeNativeDocxPagePaintV1Result> {
  const decoded = decodeNativeDocxPagePaintV1(value)
  if (!decoded.ok) return decoded
  const compiled = await compileNativeDocxPagePaintV1(requestValue, provider)
  if (!compiled.ok) return compiled
  if (!sameWire(decoded.value, compiled.value)) return { ok: false, issues: [issue('BROKEN_REFERENCE', '', 'page-paint output must exactly equal canonical recompilation for the joined request and outline provider')] }
  return decoded
}
