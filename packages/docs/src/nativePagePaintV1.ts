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
  DOCX_PAGE_PAINT_REQUEST_PROTOCOL, DOCX_PAGE_PAINT_REQUEST_VERSION,
  DOCX_PAGE_PAINT_PROTOCOL, DOCX_PAGE_PAINT_VERSION, DOCX_PAGE_PAINT_LIMITS,
  DOCX_PAGE_PAINT_REQUEST_V1_BINDING_FIELDS, SHORT_ID, PROVIDER_ID, SHA256, RGB,
  canonicalWireSha256, issue, add, isObject, exactObject, stringValue, integer,
  safeClone, preflightWire, sameWire, canonicalPart, paintCommandID,
  paintImageCommandID, paintNoteSeparatorCommandID, decodeNativeDocxPagePaintV1,
} from './nativePagePaintWireV1.js'
import { validateNativeDocxBodyPageFieldSourceV1 } from './nativeBodyPageFieldsV1.js'
import { nativeDocxPageNumberV1 } from './nativePageNumbersV1.js'
import { isRenderNeutralLayoutDiagnostic } from './nativeRenderDiagnostics.js'
import {deriveNativeSquareWrapPlanV1, hasNativeSquareWrapV1} from './nativeSquareWrapV1.js'
import { hasNativeDocxPageFieldsV1 } from './nativePageFieldsV1.js'
import { approximatePagePreviewEnvelope, decodeNativeDocxApproximationEligibilityV1, type NativeDocxApproximatePagePreviewV1 } from './nativeApproximationV1.js'
import type {NativeDocxApproximationEligibilityV1} from './nativeApproximationV1.js'
import {qualifyApproximateLegacyTables} from './nativeLegacyTableOriginV1.js'
import {qualifyNativeDocxFontSubstitutionsV1,isQualifiedNativeDocxFontDiagnosticV1,nativeDocxFontSubstitutionDiagnosticV1,type NativeDocxFontSubstitutionV1} from './nativeFontSubstitutionEvidenceV1.js'
import {qualifyNativeDocxFontCompositionV1} from './nativeFontCompositionV1.js'
import {decodeNativeDocxFontSubstitutionPreviewV1,DOCX_FONT_SUBSTITUTION_PREVIEW_PROTOCOL,DOCX_FONT_SUBSTITUTION_WARNING,nativeDocxFontPolicySha256V1,type NativeDocxFontSubstitutionPreviewV1} from './nativeFontSubstitutionPreviewV1.js'
import {decodeExplicitFontPolicyV1,EXPLICIT_FONT_POLICY_V1} from '@injoffice/font-metrics/layout'
export {
  DOCX_PAGE_PAINT_REQUEST_PROTOCOL, DOCX_PAGE_PAINT_REQUEST_VERSION,
  DOCX_PAGE_PAINT_PROTOCOL, DOCX_PAGE_PAINT_VERSION, DOCX_PAGE_PAINT_LIMITS,
  DOCX_PAGE_PAINT_REQUEST_V1_BINDING_FIELDS, DOCX_PAGE_PAINT_V1_BINDING_FIELDS,
  decodeNativeDocxPagePaintV1,
} from './nativePagePaintWireV1.js'

import {
  scaleFontUnits,
  validateFontManifest,
  type MaybePromise,
  type NativeFontManifest,
} from '@injoffice/font-metrics/layout'
import {
  DOCX_MAX_TWIPS_FOR_MILLIPOINTS,
  DOCX_NATIVE_LIMITS,
  type NativeDocxParagraphV1,
  type NativeDocxValidationIssue,
} from './nativeContract.js'
import {
  decodeNativeDocxPaginationRequestV1,
  paginateNativeDocxApproximateLegacyV1,
  type NativeDocxPaginationRequestV1,
  type NativeDocxPaginatedLayoutV1,
  type NativeDocxPaginatedPageV1,
  type NativeDocxPlacedLineV1,
  type NativeDocxPlacedNoteStoryV1,
} from './nativePaginationV1.js'
import { decodeNativeDocxPaginatedLayoutForRequest, decodeNativeDocxApproximatePaginatedLayoutForRequest } from './nativePaginatedLayoutContract.js'
import type { NativeDocxShapedLinesV1 } from './nativeShapingLines.js'
import {
  DOCX_HEADER_FOOTER_LAYOUT_PROTOCOL,
  DOCX_HEADER_FOOTER_LAYOUT_VERSION,
  layoutNativeDocxHeadersFootersV1,
  layoutNativeDocxFontHeadersFootersV1,
  type NativeDocxHeaderFooterLayoutV1,
  type NativeDocxHeaderFooterPageLayoutV1,
  type NativeDocxPlacedHeaderFooterLineV1,
} from './nativeHeaderFooterLayoutV1.js'
import { asciiLowerNative, compareNativeValidationIssues } from './nativeDeterminism.js'
import { layoutNativeDocxTableRowsV1, nativeDocxTableProjectionSha256V1, qualifyNativeDocxTablesV1, type NativeDocxQualifiedTableV1 } from './nativeTablePagePaintV1.js'
import {
  decodeNativeDocxPagePaintMediaAssetsV1,
  qualifyNativeDocxInlineImageV1,
  type NativeDocxPagePaintMediaAssetV1,
} from './nativeImagePagePaintV1.js'
import type { NativeDocxResolvedNumberingSourceV1, NativeDocxResolvedRunPropertiesV1 } from './nativeResolvedLayout.js'
import { nativeTextHighlightCommandV1 } from './nativeTextHighlightV1.js'
import { nativeTextUnderlineCommandsV1 } from './nativeTextUnderlineV1.js'
import { nativeDocxScriptScaleV1, validateNativeDocxScriptTransformV1 } from './nativeScriptLayoutV1.js'
import { validateNativeDocxPageFieldVariantsV1, type NativeDocxPageFieldVariantV1 } from './nativePageFieldsV1.js'
import {validateNativeDocxFontPageFieldVariantsV1,type NativeDocxFontVariantPolicyV1,nativeDocxPageFieldDocumentV1} from './nativePageFieldsV1.js'

export interface NativeDocxPagePaintRequestV1 {
  protocol: typeof DOCX_PAGE_PAINT_REQUEST_PROTOCOL
  version: typeof DOCX_PAGE_PAINT_REQUEST_VERSION
  pagination_request: NativeDocxPaginationRequestV1
  paginated_layout: NativeDocxPaginatedLayoutV1
  /** The exact manifest named by shaped-lines provenance, including face digests. */
  font_manifest: NativeFontManifest
  media_assets: NativeDocxPagePaintMediaAssetV1[]
  page_field_variants?: NativeDocxPageFieldVariantV1[]
  body_field_source?: NativeDocxPaginationRequestV1['document']
  integrity: { font_manifest_sha256: string; shaped_lines_sha256: string; paginated_layout_sha256: string; table_projection_sha256: string; media_assets_sha256: string; body_field_source_sha256?: string }
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
  source_crop: { left: number; top: number; right: number; bottom: number; unit: 'one-hundred-thousandth' }
  transform: { rotation_degrees: 0 | 90 | 180 | 270; flip_horizontal: boolean; flip_vertical: boolean }
}

export interface NativeDocxPaintFloatingImageCommandV1 extends Omit<NativeDocxPaintInlineImageCommandV1, 'kind'> {
  kind: 'paint_floating_image'
  layer: 'behind' | 'front'
  stacking_order: number
}

export interface NativeDocxFillTextHighlightCommandV1 {
  kind: 'fill_text_highlight'
  id: string
  line_id: string
  fragment_id: string
  source_id: string
  x_millipoints: number
  y_millipoints: number
  width_millipoints: number
  height_millipoints: number
  fill_rgb: string
}

export interface NativeDocxStrokeTextUnderlineCommandV1 {
  kind: 'stroke_text_underline'
  id: string
  line_id: string
  fragment_id: string
  source_id: string
  stroke_index: 0 | 1
  x1_millipoints: number
  y1_millipoints: number
  x2_millipoints: number
  y2_millipoints: number
  width_millipoints: number
  stroke_rgb: string
}

export type NativeDocxPagePaintCommandV1 = NativeDocxFillGlyphPathCommandV1 | NativeDocxFillTextHighlightCommandV1 | NativeDocxStrokeTextUnderlineCommandV1 | NativeDocxFillTableCellCommandV1 | NativeDocxStrokeTableBorderCommandV1 | NativeDocxStrokeNoteSeparatorCommandV1 | NativeDocxPaintInlineImageCommandV1 | NativeDocxPaintFloatingImageCommandV1

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
  body_field_source_sha256?: string
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

/** Canonical content attestation for the exact validated manifest carried by page-paint v1. */
export function nativeDocxPagePaintFontManifestSha256V1(manifest: NativeFontManifest): string {
  const validated = validateFontManifest(manifest)
  if (!validated.ok) throw new TypeError('font manifest must satisfy the native text v1 contract before hashing')
  return canonicalWireSha256(validated.value)
}

/** Canonical content attestation for the complete shaped-lines projection carried by page-paint v1. */
export function nativeDocxPagePaintShapedLinesSha256V1(shapedLines: NativeDocxShapedLinesV1, pageFieldVariants?: NativeDocxPageFieldVariantV1[]): string {
  return canonicalWireSha256(pageFieldVariants ? { shaped_lines: shapedLines, page_field_variants: pageFieldVariants } : shapedLines)
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

function headerFooterLayout(request: NativeDocxPagePaintRequestV1,font?:NativeDocxFontVariantPolicyV1): NativeDocxHeaderFooterLayoutV1 {
  const input={
    document: request.pagination_request.document,
    resolved_layout: request.pagination_request.resolved_layout,
    shaped_lines: request.pagination_request.shaped_lines,
    pagination_settings: request.pagination_request.pagination_settings,
    paginated_layout: request.paginated_layout,
    page_field_variants: request.page_field_variants,
  }
  return font?layoutNativeDocxFontHeadersFootersV1(input,font):layoutNativeDocxHeadersFootersV1(input)
}

function requestProvenance(request: NativeDocxPagePaintRequestV1, outlineID: string, outlineRevision: string, layout: NativeDocxHeaderFooterLayoutV1 = headerFooterLayout(request)): NativeDocxPagePaintProvenanceV1 {
  const pagination = request.pagination_request
  return {
    ...(request.integrity.body_field_source_sha256 ? { body_field_source_sha256: request.integrity.body_field_source_sha256 } : {}),
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
  return decodePagePaintRequestForPolicy(value)
}

/** @internal Only the approximate renderer consumes this explicitly tagged result. */
export function decodeNativeDocxApproximateComputedPagePaintV1(value: unknown, eligibility: unknown): { fidelity: 'approximate'; request: NativeDocxPagePaintRequestV1 } {
  if (eligibility === undefined) throw new TypeError('Approximate computed layout requires explicit eligibility')
  const decoded = decodePagePaintRequestForPolicy(value, eligibility)
  if (!decoded.ok) throw new TypeError(`approximate computed request failed full source validation: ${decoded.issues.map(issue => issue.message).join('; ')}`)
  return { fidelity: 'approximate', request: decoded.value }
}

/** Internal tagged result for the read-only font renderer, never a strict prepared artifact. */
export function decodeNativeDocxFontComputedPagePaintV1(value:unknown,policy:unknown,eligibility?:unknown,descriptors?:NativeDocxFontVariantPolicyV1['descriptors']):{fidelity:'approximate';request:NativeDocxPagePaintRequestV1}{
  const decoded=decodePagePaintRequestForPolicy(value,eligibility,{policy,descriptors})
  if(!decoded.ok)throw new TypeError(`Font computed source validation failed: ${decoded.issues.map(i=>i.message).join('; ')}`)
  return {fidelity:'approximate',request:decoded.value}
}
function decodePagePaintRequestForPolicy(value: unknown, eligibility?: unknown,font?:Omit<NativeDocxFontVariantPolicyV1,'manifest'>): DecodeNativeDocxPagePaintRequestV1Result {
  const preflight = preflightWire(value, 'page-paint request')
  if (preflight.length > 0) return { ok: false, issues: preflight }
  const snapshot = safeClone(value)
  if (snapshot === undefined) return { ok: false, issues: [issue('INVALID_VALUE', '', 'page-paint request must be a cloneable JSON wire value')] }
  const issues: NativeDocxValidationIssue[] = []
  const root = exactObject(snapshot, '', [...DOCX_PAGE_PAINT_REQUEST_V1_BINDING_FIELDS.RequestV1, ...['page_field_variants', 'body_field_source'].filter(key => isObject(snapshot) && key in snapshot)], issues)
  if (!root) return { ok: false, issues }
  if (root.protocol !== DOCX_PAGE_PAINT_REQUEST_PROTOCOL) add(issues, 'UNSUPPORTED_PROTOCOL', '/protocol', `must equal ${DOCX_PAGE_PAINT_REQUEST_PROTOCOL}`)
  if (root.version !== DOCX_PAGE_PAINT_REQUEST_VERSION) add(issues, 'UNSUPPORTED_VERSION', '/version', `must equal ${DOCX_PAGE_PAINT_REQUEST_VERSION}`)
  const pagination = decodeNativeDocxPaginationRequestV1(root.pagination_request)
  if (!pagination.ok) issues.push(...pagination.issues.map((entry) => ({ ...entry, path: `/pagination_request${entry.path}` })))
  const paginated = eligibility === undefined ? decodeNativeDocxPaginatedLayoutForRequest(root.paginated_layout, root.pagination_request) : decodeNativeDocxApproximatePaginatedLayoutForRequest(root.paginated_layout, root.pagination_request, eligibility)
  if (!paginated.ok) issues.push(...paginated.issues.map((entry) => ({ ...entry, path: `/paginated_layout${entry.path}` })))
  if (pagination.ok && paginated.ok && paginated.value.status === 'paginated') {
    try { deriveNativeSquareWrapPlanV1(pagination.value.document, pagination.value.resolved_layout, pagination.value.shaped_lines, paginated.value, true) }
    catch (error) { add(issues, 'BROKEN_REFERENCE', '/paginated_layout', error instanceof Error ? error.message : 'Square-wrap source placement failed') }
  }
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
  const integrity = exactObject(root.integrity, '/integrity', [...DOCX_PAGE_PAINT_REQUEST_V1_BINDING_FIELDS.IntegrityV1, ...(root.body_field_source !== undefined ? ['body_field_source_sha256'] : [])], issues)
  if (integrity) {
    if (root.body_field_source !== undefined && integrity.body_field_source_sha256 !== canonicalWireSha256(root.body_field_source)) add(issues,'BROKEN_REFERENCE','/integrity/body_field_source_sha256','must bind the complete original body-field source')
    stringValue(integrity.font_manifest_sha256, '/integrity/font_manifest_sha256', issues, SHA256, 71)
    stringValue(integrity.shaped_lines_sha256, '/integrity/shaped_lines_sha256', issues, SHA256, 71)
    stringValue(integrity.table_projection_sha256, '/integrity/table_projection_sha256', issues, SHA256, 71)
    stringValue(integrity.media_assets_sha256, '/integrity/media_assets_sha256', issues, SHA256, 71)
    stringValue(integrity.paginated_layout_sha256, '/integrity/paginated_layout_sha256', issues, SHA256, 71)
    if (manifest.ok && integrity.font_manifest_sha256 !== nativeDocxPagePaintFontManifestSha256V1(manifest.value)) add(issues, 'BROKEN_REFERENCE', '/integrity/font_manifest_sha256', 'must attest the complete validated font manifest carried by this request')
    if (pagination.ok && integrity.shaped_lines_sha256 !== nativeDocxPagePaintShapedLinesSha256V1(pagination.value.shaped_lines, root.page_field_variants as NativeDocxPageFieldVariantV1[] | undefined)) add(issues, 'BROKEN_REFERENCE', '/integrity/shaped_lines_sha256', 'must attest the complete strict shaped-lines and page-field variants carried by this request')
    if (pagination.ok) {
      const qualified = qualifyApproximateLegacyTables(pagination.value.document, pagination.value.resolved_layout, pagination.value.shaped_lines,eligibility===undefined?undefined:decodeNativeDocxApproximationEligibilityV1(eligibility,pagination.value.pagination_settings))
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
  let pageFieldVariants: NativeDocxPageFieldVariantV1[] | undefined
  let bodyFieldSource: NativeDocxPaginationRequestV1['document'] | undefined
  if (pagination.ok && paginated.ok) {
    try { bodyFieldSource = validateNativeDocxBodyPageFieldSourceV1(root.body_field_source,pagination.value,paginated.value) }
    catch (error) { add(issues,'BROKEN_REFERENCE','/body_field_source',error instanceof Error ? error.message : 'Invalid body-field source') }
  }
  if (pagination.ok && paginated.ok) {
    try { pageFieldVariants = font&&manifest.ok?validateNativeDocxFontPageFieldVariantsV1(pagination.value,paginated.value,root.page_field_variants,{...font,manifest:manifest.value}):validateNativeDocxPageFieldVariantsV1(pagination.value, paginated.value, root.page_field_variants) }
    catch (error) { add(issues, 'BROKEN_REFERENCE', '/page_field_variants', error instanceof Error ? error.message : 'Invalid page-field variants') }
  }
  if (pagination.ok && manifest.ok) for (const shaped of [pagination.value.shaped_lines, ...(pageFieldVariants ?? []).map((variant) => variant.shaped_lines)]) for (const paragraph of shaped.paragraphs) for (const line of paragraph.lines) for (const fragment of line.fragments) {
    if (fragment.script_transform && fragment.script_transform.font_sha256 !== manifest.value.faces.find((face) => face.faceId === fragment.face_id)?.source.contentDigest) add(issues, 'BROKEN_REFERENCE', '/pagination_request/shaped_lines', 'Script metrics must bind the exact content-addressed shaped font face')
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
      ...(pageFieldVariants ? { page_field_variants: pageFieldVariants } : {}),
      ...(bodyFieldSource ? { body_field_source: bodyFieldSource } : {}),
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

function placePath(path: NativeDocxGlyphDesignPathCommandV1[], originX: number, originY: number, fontSize: number, unitsPerEm: number, fontSizeY = fontSize): NativeDocxPaintPathCommandV1[] | undefined {
  const output: NativeDocxPaintPathCommandV1[] = []
  for (const command of path) {
    if (command.kind === 'close_path') { output.push(command); continue }
    const x = coordinate(originX, command.x, fontSize, unitsPerEm)
    const y = coordinate(originY, command.y, fontSizeY, unitsPerEm, true)
    if (x === undefined || y === undefined) return undefined
    if (command.kind === 'move_to' || command.kind === 'line_to') output.push({ kind: command.kind, x_millipoints: x, y_millipoints: y })
    else if (command.kind === 'quadratic_to') {
      const controlX = coordinate(originX, command.control_x, fontSize, unitsPerEm)
      const controlY = coordinate(originY, command.control_y, fontSizeY, unitsPerEm, true)
      if (controlX === undefined || controlY === undefined) return undefined
      output.push({ kind: 'quadratic_to', control_x_millipoints: controlX, control_y_millipoints: controlY, x_millipoints: x, y_millipoints: y })
    } else {
      const control1X = coordinate(originX, command.control_1_x, fontSize, unitsPerEm)
      const control1Y = coordinate(originY, command.control_1_y, fontSizeY, unitsPerEm, true)
      const control2X = coordinate(originX, command.control_2_x, fontSize, unitsPerEm)
      const control2Y = coordinate(originY, command.control_2_y, fontSizeY, unitsPerEm, true)
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
  type Fragment = NonNullable<NativeDocxPaginatedPageV1['table_rows']>[number]
  const fragmentsByRow = new Map<string, Array<{ pageID: string; fragment: Fragment }>>()
  let work = 0
  for (const page of pages) {
    for (const fragment of page.table_rows ?? []) {
      const key = JSON.stringify([fragment.table_id, fragment.row_id]), entries = fragmentsByRow.get(key) ?? []
      entries.push({ pageID: page.id, fragment }); fragmentsByRow.set(key, entries)
      if (++work > DOCX_PAGE_PAINT_LIMITS.maxOutputNodes) return undefined
    }
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
      // Split-row geometry is source-replayed explicitly: a short/empty cell
      // has no line anchor on later pages but still needs its background/edges.
      const placements: Array<{ pageID: string; line?: NativeDocxPlacedLineV1; fragment?: Fragment }> | undefined = table.table.rows[rowIndex]!.cant_split !== true
        ? fragmentsByRow.get(JSON.stringify([table.table.id, row.row_id]))
        : firstParagraph ? placementsByParagraph.get(firstParagraph.id) : undefined
      if (!placements || !firstShaped) continue
      const topMargin = table.table.cell_margins!.top_twips * 50 + (table.border_reservation_policy?.above_content_millipoints ?? 0)
      for (const placement of placements) {
        const page = pageByID.get(placement.pageID)
        const target = commands.get(placement.pageID)
        if (!page || !target) return undefined
        const rowY = placement.fragment?.y_millipoints ?? (placement.line!.y_millipoints - topMargin - firstShaped.spacing_before_millipoints)
        if (!Number.isSafeInteger(rowY)) return undefined
        for (const [cellIndex, cell] of row.cells.entries()) {
          work += 1
          if (work > DOCX_PAGE_PAINT_LIMITS.maxOutputNodes) return undefined
          if (cell.vertical_merge === 'continue') continue
          const x = page.body_box.x_millipoints + cell.x_millipoints
          const height = placement.fragment?.height_millipoints ?? cell.height_millipoints
          const placementSuffix = placement.fragment ? `:fragment:${placement.fragment.fragment_ordinal}` : placement.line?.repeated_table_header ? `:repeat:${page.id}` : ''
          const firstFragment = !placement.fragment || placement.fragment.source_y_millipoints === 0
          const lastFragment = !placement.fragment || placement.fragment.source_y_millipoints + height === placement.fragment.source_height_millipoints
          if (cell.shading_rgb) target.fills.push({ kind: 'fill_table_cell', id: `paint:table:${table.table.id}:${rowIndex}:${cellIndex}:fill${placementSuffix}`, table_id: table.table.id, row_id: row.row_id, cell_id: cell.cell_id, x_millipoints: x, y_millipoints: rowY, width_millipoints: cell.width_millipoints, height_millipoints: height, fill_rgb: cell.shading_rgb })
          const source = table.table.borders
          const lastMergeRow = rowIndex + cell.row_span - 1
          const edges: Array<{ edge: NativeDocxStrokeTableBorderCommandV1['edge']; border?: import('./nativeContract.js').NativeDocxTableBorderV1; x1: number; y1: number; x2: number; y2: number }> = [
            { edge: 'top', border: firstFragment && rowIndex === 0 ? source?.top : undefined, x1: x, y1: rowY, x2: x + cell.width_millipoints, y2: rowY },
            { edge: 'left', border: cell.column_ordinal === 0 ? source?.left : undefined, x1: x, y1: rowY, x2: x, y2: rowY + height },
            { edge: 'right', border: cell.column_ordinal + cell.grid_span === gridColumns ? source?.right : source?.inside_vertical, x1: x + cell.width_millipoints, y1: rowY, x2: x + cell.width_millipoints, y2: rowY + height },
            { edge: 'bottom', border: lastFragment ? lastMergeRow === rows.length - 1 ? source?.bottom : source?.inside_horizontal : undefined, x1: x, y1: rowY + height, x2: x + cell.width_millipoints, y2: rowY + height },
          ]
          for (const edge of edges) if (edge.border?.style === 'single' && edge.border.color_rgb) target.borders.push({
            kind: 'stroke_table_border', id: `paint:table:${table.table.id}:${rowIndex}:${cellIndex}:${edge.edge}${placementSuffix}`, table_id: table.table.id, row_id: row.row_id, cell_id: cell.cell_id, edge: edge.edge,
            x1_millipoints: edge.x1, y1_millipoints: edge.y1, x2_millipoints: edge.x2, y2_millipoints: edge.y2,
            width_millipoints: edge.border.size_eighth_points * 125, stroke_rgb: edge.border.color_rgb,
          })
        }
      }
    }
  }
  return commands
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

/**
 * Compiles a complete all-or-nothing page projection. Any provider/source
 * refusal discards every accumulated page before a result is returned.
 */
export async function compileNativeDocxPagePaintV1(value: unknown, outlineProvider: NativeDocxGlyphOutlineProviderV1): Promise<CompileNativeDocxPagePaintV1Result> {
  const decoded = decodeNativeDocxPagePaintRequestV1(value)
  if (!decoded.ok) return decoded
  return compileDecodedPagePaint(decoded.value, outlineProvider)
}

/** Explicit read-only alternative. The original strict request is validated
 * before current-policy placement; no adjusted strict artifact is exposed. */
export async function compileNativeDocxApproximatePagePreviewV1(value: unknown, eligibilityValue: unknown, outlineProvider: NativeDocxGlyphOutlineProviderV1): Promise<NativeDocxApproximatePagePreviewV1> {
  const decoded = decodeNativeDocxPagePaintRequestV1(value)
  if (!decoded.ok) throw new TypeError('approximate preview requires a valid original strict request')
  const request = decoded.value
  const settings = request.pagination_request.pagination_settings
  const eligibility = decodeNativeDocxApproximationEligibilityV1(eligibilityValue, settings)
  const provenance = requestProvenance(request, request.outline_provider.provider_id, request.outline_provider.provider_revision)
  if (eligibility.status !== 'eligible' || request.body_field_source || hasNativeDocxPageFieldsV1(request.pagination_request.document) || hasNativeSquareWrapV1(request.pagination_request.document)) {
    return approximatePagePreviewEnvelope(settings, eligibility, refusal(provenance, 'unsupported-source', settings.document_id, 'Approximate legacy preview requires eligible settings and currently excludes page-field or square-wrap fixed-point layout'))
  }
  const approximate = paginateNativeDocxApproximateLegacyV1(request.pagination_request, eligibility)
  request.paginated_layout = approximate.layout
  request.integrity.paginated_layout_sha256 = nativeDocxPagePaintPaginatedLayoutSha256V1(approximate.layout)
  const originTables=qualifyApproximateLegacyTables(request.pagination_request.document,request.pagination_request.resolved_layout,request.pagination_request.shaped_lines,eligibility)
  request.integrity.table_projection_sha256=originTables.status==='qualified'?originTables.sha256:nativeDocxTableProjectionSha256V1([])
  const painted = await compileDecodedPagePaint(request, outlineProvider, eligibility)
  if (!painted.ok) throw new TypeError('approximate page painting failed bounded validation')
  return approximatePagePreviewEnvelope(settings, eligibility, painted.value)
}

/** @internal Paints only a source-validated current-policy fixed-point result.
 * The public strict compiler/decoder never accepts this as strict pagination. */
export async function compileNativeDocxApproximateComputedPagePreviewV1(value: unknown, eligibilityValue: unknown, outlineProvider: NativeDocxGlyphOutlineProviderV1): Promise<NativeDocxApproximatePagePreviewV1> {
  const { request } = decodeNativeDocxApproximateComputedPagePaintV1(value, eligibilityValue)
  const settings = request.pagination_request.pagination_settings
  const eligibility = decodeNativeDocxApproximationEligibilityV1(eligibilityValue, settings)
  const painted = await compileDecodedPagePaint(request, outlineProvider, eligibility)
  if (!painted.ok) throw new TypeError('Approximate computed page painting failed validation')
  return approximatePagePreviewEnvelope(settings, eligibility, painted.value)
}

/** Separate approximate envelope; never exports its prepared or strict paint. */
export async function compileNativeDocxFontSubstitutionPreviewV1(value:unknown,policyValue:unknown,outlineProvider:NativeDocxGlyphOutlineProviderV1,compositionValue?:unknown):Promise<NativeDocxFontSubstitutionPreviewV1>{
 const composition=compositionValue===undefined?undefined:qualifyNativeDocxFontCompositionV1(compositionValue)
 const decoded={ok:true as const,value:decodeNativeDocxFontComputedPagePaintV1(value,policyValue,composition?.legacy,composition?.descriptors).request}
 const request=decoded.value,policy=decodeExplicitFontPolicyV1(policyValue)
 if((!composition&&request.pagination_request.pagination_settings.profile!=='word-modern-default')||request.body_field_source||hasNativeSquareWrapV1(request.pagination_request.document))throw new TypeError('Font preview requires qualified composition and excludes body field/wrap policies')
 if(composition&&(canonicalWireSha256(request.pagination_request.document)!==canonicalWireSha256(composition.document)||canonicalWireSha256(request.pagination_request.resolved_layout)!==canonicalWireSha256(composition.resolved)||canonicalWireSha256(request.pagination_request.pagination_settings)!==canonicalWireSha256(composition.settings)))throw new TypeError('Font composition does not reproduce private layout projection')
 const records=qualifyNativeDocxFontSubstitutionsV1(request.pagination_request.shaped_lines,request.pagination_request.resolved_layout,request.font_manifest,policy,request.pagination_request.document,composition?.descriptors)
 const font={manifest:request.font_manifest,policy,descriptors:composition?.descriptors}
 validateNativeDocxFontPageFieldVariantsV1(request.pagination_request,request.paginated_layout,request.page_field_variants,font)
 const seen=new Set(records.map(r=>JSON.stringify([r.source_id,r.source_role])))
 for(const [index,variant]of (request.page_field_variants??[]).entries()){
  const page=request.paginated_layout.pages[index]!,document=nativeDocxPageFieldDocumentV1(request.pagination_request.document,page.ordinal,request.paginated_layout.pages.length,nativeDocxPageNumberV1(request.pagination_request.document,request.paginated_layout,page.ordinal))
  for(const r of qualifyNativeDocxFontSubstitutionsV1(variant.shaped_lines,request.pagination_request.resolved_layout,request.font_manifest,policy,document,composition?.descriptors)){const key=JSON.stringify([r.source_id,r.source_role]);if(!seen.has(key)){records.push(r);seen.add(key)}}
 }
 const painted=await compileDecodedPagePaint(request,outlineProvider,composition?.legacy??false,records,font)
 if(!painted.ok)throw new TypeError('Font preview painting failed validation')
 const paint=painted.value
 return decodeNativeDocxFontSubstitutionPreviewV1({protocol:DOCX_FONT_SUBSTITUTION_PREVIEW_PROTOCOL,version:1,fidelity:'approximate',read_only:true,policy:EXPLICIT_FONT_POLICY_V1,operator_policy:policy,policy_sha256:nativeDocxFontPolicySha256V1(policy),source:{document_id:paint.provenance.document_id,revision:paint.provenance.revision,package_sha256:paint.provenance.package_sha256},selected_font_manifest:request.font_manifest,substitutions:records,...(composition?{composition:composition.value,composition_sha256:composition.sha256}:{}),reasons:[DOCX_FONT_SUBSTITUTION_WARNING,...(composition?.reasons??[]),...records.map(r=>nativeDocxFontSubstitutionDiagnosticV1(r).message)],status:paint.status,pages:paint.pages,resources:paint.resources,diagnostics:paint.diagnostics,rendering_provenance:paint.provenance})
}

async function compileDecodedPagePaint(request: NativeDocxPagePaintRequestV1, outlineProvider: NativeDocxGlyphOutlineProviderV1, approximateLegacySettings:NativeDocxApproximationEligibilityV1|false = false,fontSubstitutions?:readonly NativeDocxFontSubstitutionV1[],font?:NativeDocxFontVariantPolicyV1): Promise<CompileNativeDocxPagePaintV1Result> {
  const providerResult = snapshotProvider(outlineProvider)
  if (!providerResult.ok) return providerResult
  const provider = providerResult.value
  if (request.outline_provider.provider_id !== provider.id || request.outline_provider.provider_revision !== provider.revision) {
    const expectedProvenance = requestProvenance(request, request.outline_provider.provider_id, request.outline_provider.provider_revision)
    return { ok: true, value: refusal(expectedProvenance, 'provider-mismatch', request.pagination_request.document.document_id, 'Injected outline provider id and revision do not match the page-paint request') }
  }
  const headerFooter = headerFooterLayout(request,font)
  const provenance = requestProvenance(request, provider.id, provider.revision, headerFooter)
  const pagination = request.pagination_request
  const layout = request.paginated_layout
  const documentID = pagination.document.document_id
  if(pagination.shaped_lines.font_substitutions?.length&&fontSubstitutions===undefined)return {ok:true,value:refusal(provenance,'unsupported-diagnostic',documentID,'Substituted fonts require a separate explicit approximate compiler; strict paint is unavailable')}
  if (layout.status !== 'paginated') {
    const result = refusal(provenance, 'upstream-refused', documentID, 'Native pagination refused; no visual projection was emitted')
    // Keep the v1 wire schema and atomic refusal, but expose actionable source
    // reasons. Bound the expansion independently of the upstream corpus size.
    const reasons = layout.diagnostics.filter((entry) => entry.severity !== 'deferred')
    result.diagnostics.push(...reasons.slice(0, 8).map((entry) => refusal(
      provenance, 'upstream-refused', entry.scope_id,
      `Pagination ${entry.code}${entry.source_code ? ` / ${entry.source_code}` : ''}: ${entry.source_message ?? entry.message}`,
    ).diagnostics[0]!))
    if (reasons.length > 8) result.diagnostics.push(refusal(provenance, 'upstream-refused', documentID, `${reasons.length - 8} additional pagination reasons omitted; inspect the prepared paginated_layout diagnostics`).diagnostics[0]!)
    return { ok: true, value: result }
  }
  if (headerFooter.status === 'refused') return { ok: true, value: refusal(provenance, 'unsupported-source', headerFooter.diagnostics[0]?.scope_id ?? documentID, headerFooter.diagnostics[0]?.message ?? 'Native header/footer layout refused') }
  const allowedFontPagination=(entry:typeof layout.diagnostics[number])=>fontSubstitutions!==undefined&&entry.code==='source-diagnostic'&&entry.severity==='deferred'&&fontSubstitutions.some(r=>{const d=nativeDocxFontSubstitutionDiagnosticV1(r);return entry.scope_id===d.scope_id&&entry.source_code===d.code&&entry.source_message===d.message&&entry.message===`Shaping diagnostic retained by pagination: ${d.code}: ${d.message}`})
  const blockingPaginationDiagnostics = layout.diagnostics.filter((entry) => !allowedFontPagination(entry)&&!(approximateLegacySettings && entry.code === 'settings-attestation-unsupported' && entry.severity === 'deferred') && entry.code !== 'header-footer-selection-deferred' && !(entry.code === 'source-diagnostic' && entry.severity === 'deferred' && entry.source_code === 'page-control-deferred'))
  const blockingShapingDiagnostics = pagination.shaped_lines.diagnostics.filter((entry) => !(fontSubstitutions!==undefined&&isQualifiedNativeDocxFontDiagnosticV1(entry,fontSubstitutions))&&(entry.code !== 'page-control-deferred' || entry.source_id !== undefined))
  const blockingResolutionDiagnostics = pagination.resolved_layout.diagnostics.filter((entry) => !isRenderNeutralLayoutDiagnostic(entry, pagination.resolved_layout))
  if (blockingShapingDiagnostics.length > 0 || blockingPaginationDiagnostics.length > 0 || blockingResolutionDiagnostics.length > 0) {
    const first = blockingShapingDiagnostics[0] ?? blockingPaginationDiagnostics[0] ?? blockingResolutionDiagnostics[0]
    return { ok: true, value: refusal(provenance, 'unsupported-diagnostic', documentID, `Page-paint v1 requires no blocking shaping/resolution diagnostics and permits only the exact header/footer selection handoff from pagination${first ? `: ${first.code}` : ''}`) }
  }
  const qualifiedTables = qualifyApproximateLegacyTables(pagination.document, pagination.resolved_layout, pagination.shaped_lines,approximateLegacySettings||undefined)
  if (qualifiedTables.status !== 'qualified') return { ok: true, value: refusal(provenance, 'unsupported-source', qualifiedTables.diagnostics[0]!.scope_id, qualifiedTables.diagnostics[0]!.message) }
  const tableCommandsIndex = tableCommandsByPage(request, layout.pages, qualifiedTables.tables)
  if (!tableCommandsIndex) return { ok: true, value: refusal(provenance, 'resource-limit', documentID, 'Table paint indexing exceeded its bounded work or geometry contract') }

  const faces = new Map(request.font_manifest.faces.map((face) => [face.faceId, face]))
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
  const selectedParagraphIDs = new Map<string, { paragraphID: string; prefix: string; ordinal: number }>()
  const coveredLineIDs = new Set<string>()
  const coveredFragmentIDs = new Set<string>()
  const pages: NativeDocxPaintPageV1[] = []
  const mediaAssets = new Map(request.media_assets.map((asset) => [asset.id, asset]))
  const headerFooterByPageID = new Map(headerFooter.pages.map((entry) => [entry.page_id, entry]))
  let glyphCount = 0
  let pathCommandCount = 0
  let providerCalls = 0

  for (const page of layout.pages) {
    const pageParagraphs = new Map((request.page_field_variants?.find((variant) => variant.page_id === page.id)?.shaped_lines ?? pagination.shaped_lines).paragraphs.map((paragraph) => [paragraph.paragraph_id, paragraph]))
    const tableCommands = tableCommandsIndex.get(page.id)
    if (!tableCommands) return { ok: true, value: refusal(provenance, 'identity-mismatch', documentID, 'Table geometry could not exact-join paginated cell lines') }
    const contentCommands: Array<NativeDocxFillGlyphPathCommandV1 | NativeDocxFillTextHighlightCommandV1 | NativeDocxStrokeTextUnderlineCommandV1 | NativeDocxPaintInlineImageCommandV1 | NativeDocxPaintFloatingImageCommandV1 | NativeDocxStrokeNoteSeparatorCommandV1> = []
    const paintLines: NativeDocxPaintLineV1[] = []
    const headerFooterPage = headerFooterByPageID.get(page.id)
    if (!headerFooterPage) return { ok: true, value: refusal(provenance, 'incomplete-page', page.id, 'Header/footer layout does not exactly cover the paginated page') }
    const headerLines = headerFooterPage.lines.filter((line) => line.region === 'header')
    const footerLines = headerFooterPage.lines.filter((line) => line.region === 'footer')
    const noteStories = page.note_stories ?? []
    const noteStoryByLineID = new Map(noteStories.flatMap((story) => story.lines.map((line) => [line.id, story] as const)))
    const placedLines: Array<NativeDocxPlacedHeaderFooterLineV1 | NativeDocxPaginatedPageV1['lines'][number]> = [...headerLines, ...page.lines, ...noteStories.flatMap((story) => story.lines), ...footerLines]
    for (const placed of placedLines) {
      const paragraph = pageParagraphs.get(placed.paragraph_id)
      const line = paragraph?.lines[placed.source_line_ordinal]
      if (!paragraph || !line || line.id !== placed.line_id) return { ok: true, value: refusal(provenance, 'identity-mismatch', placed.line_id, 'Placed line does not exact-join its shaped line') }
      const noteStory = noteStoryByLineID.get(placed.id)
      const expectedStoryKind = 'region' in placed ? placed.region : noteStory?.story_kind ?? 'body'
      const expectedStoryID = 'story_id' in placed ? placed.story_id : noteStory?.story_id
      if (paragraph.story_kind !== expectedStoryKind || expectedStoryID !== undefined && paragraph.story_id !== expectedStoryID) return { ok: true, value: refusal(provenance, 'unsupported-source', paragraph.paragraph_id, 'Painted line does not exact-join its selected body/header/footer/note story') }
      const coveragePrefix = request.page_field_variants && 'region' in placed ? `${page.id}:` : ''
      selectedParagraphIDs.set(coveragePrefix + paragraph.paragraph_id, { paragraphID: paragraph.paragraph_id, prefix: coveragePrefix, ordinal: page.ordinal })
      if (paragraph.alignment === 'distribute') return { ok: true, value: refusal(provenance, 'unsupported-source', paragraph.paragraph_id, 'Distributed character expansion is outside page-paint v1') }
      const naturalHeight = line.ascent_millipoints - line.descent_millipoints + line.line_gap_millipoints
      // Current-layout approximation deliberately anchors natural ascent at
      // the top of an expanded line box. It does not claim Word leading
      // distribution or permit clipping/compressed-line semantics.
      if (line.line_height_millipoints !== naturalHeight && !(approximateLegacySettings && line.line_height_millipoints >= naturalHeight)) return { ok: true, value: refusal(provenance, 'unsupported-source', line.id, 'Page-paint v1 requires natural shaped line height for an exact baseline; only explicit current-layout approximation supports expanded line boxes') }
      if (line.hard_break_after && !coveredLineIDs.has(coveragePrefix + line.id)) sourceHardBreakCounts.set(coveragePrefix + line.hard_break_after.source_run_id, (sourceHardBreakCounts.get(coveragePrefix + line.hard_break_after.source_run_id) ?? 0) + 1)
      coveredLineIDs.add(coveragePrefix + line.id)
      let fragmentX = placed.x_millipoints
      const baselineY = placed.y_millipoints + line.ascent_millipoints
      if (!Number.isSafeInteger(baselineY) || Math.abs(baselineY) > DOCX_PAGE_PAINT_LIMITS.maxPaintCoordinateMilliPoints) return { ok: true, value: refusal(provenance, 'resource-limit', line.id, 'Line baseline exceeds the bounded paint coordinate range') }
      const firstCommand = contentCommands.length
      const highlights: NativeDocxFillTextHighlightCommandV1[] = []
      const underlines: NativeDocxStrokeTextUnderlineCommandV1[] = []
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
            const sourceText = nativeRun.page_field ? String(nativeRun.page_field === 'PAGE' ? nativeDocxPageNumberV1(pagination.document,layout,page.ordinal) : layout.pages.length) : nativeRun.text
            const exactText = nativeRun.kind === 'text' && sourceText !== undefined && sourceText.slice(fragment.start_utf16, fragment.end_utf16) === fragment.text
            const exactMarker = noteMarker !== undefined && fragment.start_utf16 === 0 && fragment.end_utf16 === noteMarker.length && fragment.text === noteMarker
            if (!exactText && !exactMarker) return { ok: true, value: refusal(provenance, 'identity-mismatch', fragment.id, 'Run fragment text and UTF-16 range must exactly match native text or its placed note number') }
            if (fragment.text.length > 0 && fragment.glyphs.length === 0) return { ok: true, value: refusal(provenance, 'missing-glyph', fragment.id, 'A non-empty visible run fragment cannot paint without glyphs') }
            if (!coveredFragmentIDs.has(coveragePrefix + fragment.id)) {
              const intervals = sourceIntervals.get(coveragePrefix + nativeRun.id) ?? []
              intervals.push({ start: fragment.start_utf16, end: fragment.end_utf16 })
              sourceIntervals.set(coveragePrefix + nativeRun.id, intervals)
            }
          } else if (fragment.source_kind === 'tab') {
            if (nativeRun.kind !== 'control' || nativeRun.control !== 'tab' || fragment.text !== '\t' || fragment.start_utf16 !== 0 || fragment.end_utf16 !== 0 || fragment.glyphs.length !== 0) return { ok: true, value: refusal(provenance, 'identity-mismatch', fragment.id, 'Tab fragment must exactly match a glyphless native tab control') }
            if (!coveredFragmentIDs.has(coveragePrefix + fragment.id)) sourceControlCounts.set(coveragePrefix + nativeRun.id, (sourceControlCounts.get(coveragePrefix + nativeRun.id) ?? 0) + 1)
          } else if (fragment.source_kind === 'image') {
            if (nativeRun.kind !== 'drawing' || !nativeRun.drawing || fragment.text !== '' || fragment.start_utf16 !== 0 || fragment.end_utf16 !== 0 || fragment.glyphs.length !== 0 || fragment.face_id !== undefined || fragment.whitespace) return { ok: true, value: refusal(provenance, 'identity-mismatch', fragment.id, 'Image fragment must exactly match one glyphless native drawing run') }
            const qualified = qualifyNativeDocxInlineImageV1(pagination.document, nativeRun.id, nativeRun.drawing)
            if (!qualified.ok) return { ok: true, value: refusal(provenance, qualified.code === 'resource-limit' ? 'resource-limit' : 'unsupported-source', fragment.id, qualified.message) }
            const image = qualified.value
            const asset = mediaAssets.get(image.asset_id)
            if (!asset || asset.part_name !== image.part_name || asset.content_type !== image.content_type || asset.content_digest !== image.content_digest) return { ok: true, value: refusal(provenance, 'identity-mismatch', fragment.id, 'Image fragment does not exact-join one canonical content-addressed media asset') }
            if (fragment.advance_inline_millipoints !== (image.floating ? 0 : image.layout_width_millipoints) || fragment.ascent_millipoints !== (image.floating ? 0 : image.layout_ascent_millipoints) || fragment.descent_millipoints !== (image.floating ? 0 : image.layout_descent_millipoints) || fragment.line_gap_millipoints !== 0) return { ok: true, value: refusal(provenance, 'identity-mismatch', fragment.id, 'Image fragment geometry changed after exact EMU projection') }
            const y = image.floating?.y_millipoints ?? baselineY - image.height_millipoints
            const x = image.floating?.x_millipoints ?? fragmentX + image.content_offset_x_millipoints
            if (image.floating && (x + image.width_millipoints > page.width_millipoints || y + image.height_millipoints > page.height_millipoints)) return { ok: true, value: refusal(provenance, 'unsupported-source', fragment.id, 'Floating image must fit entirely inside its anchor paragraph page') }
            if (!Number.isSafeInteger(y) || y < 0) return { ok: true, value: refusal(provenance, 'resource-limit', fragment.id, 'Image placement exceeds bounded non-negative page coordinates') }
            contentCommands.push({
              ...(image.floating ? { kind: 'paint_floating_image' as const, layer: image.floating.layer, stacking_order: image.floating.stacking_order } : { kind: 'paint_inline_image' as const }), id: paintImageCommandID(placed.id, fragment.id), line_id: line.id, fragment_id: fragment.id, source_id: fragment.source_id,
              drawing_id: image.drawing_id, asset_id: image.asset_id,
              x_millipoints: x, y_millipoints: y, width_millipoints: image.width_millipoints, height_millipoints: image.height_millipoints,
              source_crop: image.source_crop, transform: image.transform,
            })
            fragmentX += fragment.advance_inline_millipoints
            if (!Number.isSafeInteger(fragmentX)) return { ok: true, value: refusal(provenance, 'resource-limit', fragment.id, 'Image cursor exceeds safe integer coordinates') }
            if (!coveredFragmentIDs.has(coveragePrefix + fragment.id)) sourceImageCounts.set(coveragePrefix + nativeRun.id, (sourceImageCounts.get(coveragePrefix + nativeRun.id) ?? 0) + 1)
            coveredFragmentIDs.add(coveragePrefix + fragment.id)
            continue
          }
        }
        coveredFragmentIDs.add(coveragePrefix + fragment.id)
        const scriptTransform = fragment.script_transform
        if (scriptTransform && (!validateNativeDocxScriptTransformV1(scriptTransform) || scriptTransform.kind !== properties.vertical_alignment || scriptTransform.font_sha256 !== (fragment.face_id ? faces.get(fragment.face_id)?.source.contentDigest : undefined))) return { ok: true, value: refusal(provenance, 'identity-mismatch', fragment.source_id, 'Script transform must exact-join source alignment and content-addressed font') }
        if (scriptTransform && (properties.underline && properties.underline !== 'none' || properties.highlight && properties.highlight !== 'none')) return { ok: true, value: refusal(provenance, 'unsupported-source', fragment.source_id, 'Combined script transforms and underline/highlight are not qualified') }
        const underline = nativeTextUnderlineCommandsV1(properties.underline, properties.color, fragment, placed.id, line.id, fragmentX, baselineY)
        if (!underline.ok) return { ok: true, value: refusal(provenance, 'unsupported-source', fragment.source_id, underline.message) }
        for (const command of underline.commands) underlines.push(command)
        const highlight = nativeTextHighlightCommandV1(properties.highlight, fragment, placed.id, line.id, fragmentX, baselineY)
        if (!highlight.ok) return { ok: true, value: refusal(provenance, 'unsupported-source', fragment.source_id, highlight.message) }
        if (highlight.command) highlights.push(highlight.command)
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
          const path = outline.status === 'empty' ? [] : placePath(outline.path, originX, originY, scriptTransform ? nativeDocxScriptScaleV1(fontSize,scriptTransform,'x') : fontSize, outline.units_per_em, scriptTransform ? nativeDocxScriptScaleV1(fontSize,scriptTransform,'y') : fontSize)
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
      // Backgrounds precede every glyph on this line so italic overhangs cannot
      // be covered by the next fragment's background. Avoid unbounded spread.
      if (highlights.length) {
        const foreground = contentCommands.splice(firstCommand)
        for (const command of highlights) contentCommands.push(command)
        for (const command of foreground) contentCommands.push(command)
      }
      for (const command of underlines) contentCommands.push(command)
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
      commands: [
        ...contentCommands.filter((command): command is NativeDocxPaintFloatingImageCommandV1 => command.kind === 'paint_floating_image' && command.layer === 'behind').sort((a,b) => a.stacking_order-b.stacking_order),
        ...tableCommands.fills, ...contentCommands.filter(command => command.kind !== 'paint_floating_image'), ...tableCommands.borders,
        ...contentCommands.filter((command): command is NativeDocxPaintFloatingImageCommandV1 => command.kind === 'paint_floating_image' && command.layer === 'front').sort((a,b) => a.stacking_order-b.stacking_order),
      ],
    })
  }
  for (const { paragraphID, prefix, ordinal } of selectedParagraphIDs.values()) for (const nativeRun of nativeParagraphs.get(paragraphID)?.runs ?? []) {
    const coverageID = prefix + nativeRun.id
    const resolved = resolvedRuns.get(nativeRun.id)
    if (!resolved || resolved.properties.hidden) continue
    if (nativeRun.kind === 'text' && nativeRun.text !== undefined) {
      const intervals = (sourceIntervals.get(coverageID) ?? []).sort((left, right) => left.start - right.start || left.end - right.end)
      let cursor = 0
      for (const interval of intervals) {
        if (interval.start !== cursor || interval.end <= interval.start) return { ok: true, value: refusal(provenance, 'identity-mismatch', nativeRun.id, 'Painted visual clusters must exactly partition their native text run in logical order') }
        cursor = interval.end
      }
      const expectedText = nativeRun.page_field ? String(nativeRun.page_field === 'PAGE' ? nativeDocxPageNumberV1(pagination.document,layout,ordinal) : layout.pages.length) : nativeRun.text
      if (cursor !== expectedText.length) return { ok: true, value: refusal(provenance, 'identity-mismatch', nativeRun.id, 'Painted visual clusters must completely cover their native text run') }
    } else if (nativeRun.kind === 'control' && nativeRun.control === 'tab' && sourceControlCounts.get(coverageID) !== 1) {
      return { ok: true, value: refusal(provenance, 'identity-mismatch', nativeRun.id, 'A painted native tab must have exactly one visual fragment') }
    } else if (nativeRun.kind === 'control' && nativeRun.control === 'line-break' && sourceHardBreakCounts.get(coverageID) !== 1) {
      return { ok: true, value: refusal(provenance, 'identity-mismatch', nativeRun.id, 'A painted native line break must have exactly one hard-break attestation') }
    } else if (nativeRun.kind === 'drawing' && sourceImageCounts.get(coverageID) !== 1) {
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
      const resolvedRuns = new Map(request.value.pagination_request.resolved_layout.runs.map((run) => [run.run_id, run]))
      const nativeStories = [request.value.pagination_request.document.body, ...request.value.pagination_request.document.headers, ...request.value.pagination_request.document.footers, ...request.value.pagination_request.document.notes]
      const nativeRuns = new Map(nativeStories.flatMap((story) => story.blocks.flatMap((block) => block.paragraph?.runs ?? [])).map((run) => [run.id, run]))
      const resolvedParagraphs = new Map(request.value.pagination_request.resolved_layout.paragraphs.map((paragraph) => [paragraph.paragraph_id, paragraph]))
      const manifestFaces = new Map(request.value.font_manifest.faces.map((face) => [face.faceId, face]))
      const expectedGlyphs: Array<{ pageIndex: number; pageID: string; placedLineID: string; lineID: string; fragmentID: string; sourceID: string; glyphIndex: number; glyphID: number; face?: NativeDocxContentAddressedFaceV1; fontSize?: number; fill: string }> = []
      const expectedImages: Array<{ floating?: { layer: 'behind' | 'front'; stacking_order: number }; pageIndex: number; pageID: string; placedLineID: string; lineID: string; fragmentID: string; sourceID: string; drawingID: string; assetID: string; x: number; y: number; width: number; height: number; transform: NativeDocxPaintInlineImageCommandV1['transform']; crop: NativeDocxPaintInlineImageCommandV1['source_crop'] }> = []
      const expectedSeparators: Array<{ pageIndex: number; command: NativeDocxStrokeNoteSeparatorCommandV1 }> = []
      const expectedHighlights: Array<{ pageIndex: number; command: NativeDocxFillTextHighlightCommandV1 }> = []
      const highlightIDsByPlacement = new Map<string, string[]>()
      const expectedUnderlines: Array<{ pageIndex: number; command: NativeDocxStrokeTextUnderlineCommandV1 }> = []
      const underlineIDsByPlacement = new Map<string, string[]>()
      const headerFooterByPageID = headerFooter.status === 'placed' ? new Map<string, NativeDocxHeaderFooterPageLayoutV1>(headerFooter.pages.map((entry) => [entry.page_id, entry])) : new Map<string, NativeDocxHeaderFooterPageLayoutV1>()
      const qualifiedTables = qualifyNativeDocxTablesV1(request.value.pagination_request.document, request.value.pagination_request.resolved_layout, request.value.pagination_request.shaped_lines)
      const expectedTableByPageID = qualifiedTables.status === 'qualified'
        ? tableCommandsByPage(request.value, request.value.paginated_layout.pages, qualifiedTables.tables)
        : undefined
      if (!expectedTableByPageID) add(issues, 'BROKEN_REFERENCE', '/output/pages', 'table paint source could not be bounded and indexed for replay')
      const sourceLinesByPageID = new Map<string, Array<NativeDocxPlacedHeaderFooterLineV1 | NativeDocxPlacedLineV1>>()
      const noteStoryByPageLineID = new Map<string, Map<string, NativeDocxPlacedNoteStoryV1>>()
      request.value.paginated_layout.pages.forEach((page, pageIndex) => {
        const shapedParagraphs = new Map((request.value.page_field_variants?.find((variant) => variant.page_id === page.id)?.shaped_lines ?? request.value.pagination_request.shaped_lines).paragraphs.map((paragraph) => [paragraph.paragraph_id, paragraph]))
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
          const highlightIDs: string[] = []
          highlightIDsByPlacement.set(`${pageIndex}\0${placed.id}`, highlightIDs)
          const underlineIDs: string[] = []
          underlineIDsByPlacement.set(`${pageIndex}\0${placed.id}`, underlineIDs)
          line?.fragments.forEach((fragment) => {
            const resolved = resolvedRuns.get(fragment.source_id)
            const properties = fragment.source_kind === 'list-marker' ? resolvedParagraphs.get(placed.paragraph_id)?.numbering?.marker_properties : resolved?.properties
            if (fragment.source_kind !== 'image') {
              const underline = nativeTextUnderlineCommandsV1(properties?.underline, properties?.color, fragment, placed.id, line.id, fragmentX, placed.y_millipoints + line.ascent_millipoints)
              if (!underline.ok) add(issues, 'BROKEN_REFERENCE', '/output/pages', underline.message)
              else for (const command of underline.commands) { expectedUnderlines.push({ pageIndex, command }); underlineIDs.push(command.id) }
              const highlight = nativeTextHighlightCommandV1(properties?.highlight, fragment, placed.id, line.id, fragmentX, placed.y_millipoints + line.ascent_millipoints)
              if (!highlight.ok) add(issues, 'BROKEN_REFERENCE', '/output/pages', highlight.message)
              else if (highlight.command) { expectedHighlights.push({ pageIndex, command: highlight.command }); highlightIDs.push(highlight.command.id) }
            }
            const manifestFace = fragment.face_id ? manifestFaces.get(fragment.face_id) : undefined
            const face = manifestFace?.source.contentDigest ? { face_id: manifestFace.faceId, content_digest: manifestFace.source.contentDigest, ...(manifestFace.source.collectionIndex !== undefined ? { collection_index: manifestFace.source.collectionIndex } : {}) } : undefined
            fragment.glyphs.forEach((glyph, glyphIndex) => expectedGlyphs.push({ pageIndex, pageID: page.id, placedLineID: placed.id, lineID: line.id, fragmentID: fragment.id, sourceID: fragment.source_id, glyphIndex, glyphID: glyph.glyph_id, face, fontSize: properties?.font_size_half_points === undefined ? undefined : properties.font_size_half_points * 500, fill: properties?.color ?? '000000' }))
            if (fragment.source_kind === 'image') {
              const run = nativeRuns.get(fragment.source_id)
              const qualified = run?.drawing ? qualifyNativeDocxInlineImageV1(request.value.pagination_request.document, run.id, run.drawing) : undefined
              if (qualified?.ok) expectedImages.push({ floating: qualified.value.floating, pageIndex, pageID: page.id, placedLineID: placed.id, lineID: line.id, fragmentID: fragment.id, sourceID: fragment.source_id, drawingID: qualified.value.drawing_id, assetID: qualified.value.asset_id, x: qualified.value.floating?.x_millipoints ?? fragmentX + qualified.value.content_offset_x_millipoints, y: qualified.value.floating?.y_millipoints ?? placed.y_millipoints + line.ascent_millipoints - qualified.value.height_millipoints, width: qualified.value.width_millipoints, height: qualified.value.height_millipoints, transform: qualified.value.transform, crop: qualified.value.source_crop })
            }
            fragmentX += fragment.advance_inline_millipoints
          })
        })
      })
      const actualGlyphs = output.value.pages.flatMap((page, pageIndex) => page.commands.flatMap((command, commandIndex) => command.kind === 'fill_glyph_path' ? [{ page, pageIndex, commandIndex, command }] : []))
      const actualImages = output.value.pages.flatMap((page, pageIndex) => page.commands.flatMap((command, commandIndex) => command.kind === 'paint_inline_image' || command.kind === 'paint_floating_image' ? [{ page, pageIndex, commandIndex, command }] : []))
      const actualSeparators = output.value.pages.flatMap((page, pageIndex) => page.commands.flatMap((command) => command.kind === 'stroke_note_separator' ? [{ pageIndex, command }] : []))
      const actualHighlights = output.value.pages.flatMap((page, pageIndex) => page.commands.flatMap((command) => command.kind === 'fill_text_highlight' ? [{ pageIndex, command }] : []))
      if (!sameWire(actualHighlights, expectedHighlights)) add(issues, 'BROKEN_REFERENCE', '/output/pages', 'highlights must exactly cover source run colors and shaped font-metric rectangles')
      const actualUnderlines = output.value.pages.flatMap((page, pageIndex) => page.commands.flatMap(command => command.kind === 'stroke_text_underline' ? [{ pageIndex, command }] : []))
      if (!sameWire(actualUnderlines, expectedUnderlines)) add(issues, 'BROKEN_REFERENCE', '/output/pages', 'underlines must exactly cover source style and resolved font metrics')
      if (actualGlyphs.length !== expectedGlyphs.length) add(issues, 'BROKEN_REFERENCE', '/output/pages', 'paint commands must exactly cover every shaped glyph once')
      if (actualImages.length !== expectedImages.length) add(issues, 'BROKEN_REFERENCE', '/output/pages', 'paint commands must exactly cover every qualified inline image once')
      if (!sameWire(actualSeparators, expectedSeparators)) add(issues, 'BROKEN_REFERENCE', '/output/pages', 'paint commands must exactly derive one deterministic rule from each placed ordinary note separator')
      output.value.pages.forEach((page, pageIndex) => {
        const shapedParagraphs = new Map((request.value.page_field_variants?.find((variant) => variant.page_id === page.id)?.shaped_lines ?? request.value.pagination_request.shaped_lines).paragraphs.map((paragraph) => [paragraph.paragraph_id, paragraph]))
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
            ...(placed ? highlightIDsByPlacement.get(`${pageIndex}\0${placed.id}`) ?? [] : []),
            ...(separatorStory && placed ? [paintNoteSeparatorCommandID(placed.id)] : []),
            ...(shaped?.fragments.flatMap((fragment) => fragment.source_kind === 'image' ? [paintImageCommandID(placed!.id, fragment.id)] : fragment.glyphs.map((_, glyphIndex) => paintCommandID(placed!.id, fragment.id, glyphIndex))) ?? []),
            ...(placed ? underlineIDsByPlacement.get(`${pageIndex}\0${placed.id}`) ?? [] : []),
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
      const expectedImageByID = new Map(expectedImages.map(expected => [paintImageCommandID(expected.placedLineID, expected.fragmentID), expected]))
      actualImages.forEach((actual) => {
        const expected = expectedImageByID.get(actual.command.id)
        if (expected && (actual.command.kind !== (expected.floating ? 'paint_floating_image' : 'paint_inline_image') || actual.command.kind === 'paint_floating_image' && (actual.command.layer !== expected.floating?.layer || actual.command.stacking_order !== expected.floating?.stacking_order))) add(issues, 'BROKEN_REFERENCE', `/output/pages/${actual.pageIndex}/commands/${actual.commandIndex}`, 'image layering must exactly match source anchor semantics')
        if (expected && !sameWire(actual.command.transform, expected.transform)) add(issues, 'BROKEN_REFERENCE', `/output/pages/${actual.pageIndex}/commands/${actual.commandIndex}/transform`, 'image orientation must exactly match the source drawing transform')
        if (expected && !sameWire(actual.command.source_crop, expected.crop)) add(issues, 'BROKEN_REFERENCE', `/output/pages/${actual.pageIndex}/commands/${actual.commandIndex}/source_crop`, 'image crop must exactly match the source drawing rectangle')
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
