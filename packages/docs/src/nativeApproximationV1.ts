import type { NativeDocxPaginationSettingsV1 } from './nativePaginationSettings.js'
import {validLegacyTableOrigins,DOCX_LEGACY_TABLE_ORIGIN_WARNING,DOCX_TABLE_BORDER_RESERVATION_WARNING,DOCX_TABLE_GRID_FIT_WARNING,type NativeDocxLegacyTableOriginV1} from './nativeLegacyTableOriginV1.js'
import type { NativeDocxPagePaintV1 } from './nativePagePaintV1.js'
import { preflightWire, decodeNativeDocxPagePaintV1, DOCX_PAGE_PAINT_PROTOCOL, DOCX_PAGE_PAINT_VERSION } from './nativePagePaintWireV1.js'
import type { NativeDocxValidationIssue } from './nativeContract.js'
import { nativeApproximationSettingReason, validNativeDocxApproximatedSettingV1, type NativeDocxApproximatedSettingV1 } from './nativeApproximationSettingsV1.js'
import { DOCX_LATIN_FONT_FALLBACK_WARNING, validNativeDocxApproximatedFontFacesV1, validNativeDocxLatinFontFallbacksV1, type NativeDocxLatinFontFallbackV1 } from './nativeLatinFontFallbackV1.js'
import { DOCX_ABSENT_FONT_SIZE_WARNING, validNativeDocxAbsentDefaultSizeShapeV1, validNativeDocxAbsentFontSizesV1, validNativeDocxApproximatedFontSizesV1, type NativeDocxAbsentDefaultSizeShapeV1, type NativeDocxAbsentFontSizeV1, type NativeDocxApproximatedFontSizeV1 } from './nativeAbsentFontSizeV1.js'
import { DOCX_ABSENT_FONT_FAMILY_WARNING, validNativeDocxAbsentFontFamiliesV1, validNativeDocxApproximatedFontFamiliesV1, type NativeDocxAbsentFontFamilyV1, type NativeDocxApproximatedFontFamilyV1 } from './nativeAbsentFontFamilyV1.js'
import { DOCX_APPROXIMATE_IMAGE_EXTENT_WARNING, DOCX_APPROXIMATE_MAX_IMAGE_EXTENT_FACTS, validNativeDocxApproximatedImageExtentsV1, type NativeDocxApproximatedImageExtentV1 } from './nativeApproximateImageExtentV1.js'
import { collectNativeDocxApproximateOmissionsV1, nativeDocxApproximateRefusalOmissionsV1, validNativeDocxApproximateOmissionsV1, DOCX_APPROXIMATE_OMITTED_CONTENT_WARNING, type NativeDocxApproximateOmissionsV1, type NativeDocxApproximateOmissionSourceV1 } from './nativeApproximateOmittedContentV1.js'

/** Bound on the typed not-applied disclosure vector, mirrored by the Go
 * extractor's nativeApproximationMaxFacts. It bounds wire size only: a legacy
 * w:compat block routinely carries dozens of option leaves that this tier
 * records rather than applies, so it matches the grouped-member bound. */
export const DOCX_APPROXIMATE_MAX_SETTING_FACTS = 64
export const DOCX_APPROXIMATE_PREVIEW_PROTOCOL = 'injoffice.docx.approximate-page-preview' as const
export const DOCX_APPROXIMATE_PREVIEW_POLICY = 'current-layout-approximate-v1' as const
/** Source/resolution codes approximate preview omits while painting remaining glyphs. */
export const DOCX_APPROXIMATE_OMITTED_SOURCE_UNSUPPORTED = new Set([
  'UNMODELED_RUN_CONTENT',
  'UNMODELED_PARAGRAPH_CONTENT',
  'UNMODELED_BODY_BLOCK',
  'UNMODELED_FONT_METADATA',
  'FONT_MATCHING_METADATA_PRESERVED',
  'UNMODELED_FONT_SELECTION',
  'THEME_COLOR_PRESERVED',
  'THEME_UNDERLINE_COLOR_PRESERVED',
  'UNDERLINE_COLOR_PRESERVED',
  'INVALID_COLOR',
  'INVALID_KERNING_THRESHOLD',
  'UNSUPPORTED_HIGHLIGHT',
  'UNSUPPORTED_UNDERLINE',
  'AUTO_PARAGRAPH_SPACING_PRESERVED',
  'UNMODELED_PARAGRAPH_PROPERTY',
  'UNMODELED_PARAGRAPH_MARK_PROPERTIES',
  'UNMODELED_SECTION_PROPERTY',
  'UNMODELED_TABLE_PROPERTY',
  'NESTED_TABLE_OR_CELL_MARKUP',
  'PARTIAL_PARAGRAPH_PROPERTIES',
  'PARTIAL_RUN_PROPERTIES',
  'PICTURE_GRAPHIC_REQUIRED',
  // A picture the extractor refuses at the transform gate is a drawing this
  // tier does not model, exactly like PICTURE_GRAPHIC_REQUIRED and
  // UNMODELED_DRAWING: the run carries no drawing, nothing in the document
  // model represents the picture, and there is no image to paint. Omitting it
  // and painting the paragraph's remaining runs is the same decision already
  // taken for every other unmodelled drawing; refusing the whole page instead
  // produced no preview at all.
  'PICTURE_TRANSFORM_PRESERVED',
  'UNMODELED_DRAWING',
  'UNRESOLVED_COMMENT_RANGE',
  'UNRESOLVED_COMMENT_REFERENCE',
  'FIELD_SEMANTICS',
  'WRAPPED_RUN_MARKUP',
  'NUMBERING_STYLE_PRESERVED',
  // Tracked-move markup that states no content and no formatting: a
  // content-free range endpoint between table rows, and the CT_TrackChange
  // annotation that names which revision a paragraph mark belongs to. The
  // approximate tier already omits the w:ins/w:del run content they accompany.
  'NON_VISUAL_RANGE_MARKER',
  'TRACKED_MARK_REVISION_PRESERVED',
  // w14:ligatures w14:val="standardContextual" names the standard and
  // contextual ligature sets, which the declared HarfBuzz shaping defaults
  // already apply to every run. Both the extractor and the resolver emit this
  // code only for that one value, so it states shaping this tier already
  // performs and cannot move an advance; every other ligature mode stays
  // foreign markup.
  'LIGATURE_MODE_MATCHES_SHAPER',
  // An enabled w14:cntxtAlts turns on the OpenType contextual alternates
  // feature, which the declared HarfBuzz shaping defaults already apply to
  // every horizontal run. Shaping with calt=1 and with the defaults returns
  // byte-identical glyphs and advances, so this states shaping this tier
  // already performs; a disabled w14:val stays foreign markup.
  'CONTEXTUAL_ALTERNATES_MATCH_SHAPER',
  // A w:bdr that states ST_Border none or nil asks for no border. It paints
  // no stroke and reserves no space around the run, so the run occupies the
  // same box as the same run without the element and cannot move a line or a
  // page. Both the extractor and the resolver emit this code only for those
  // two values; a border that paints stays foreign markup.
  'RUN_BORDER_ABSENT_PRESERVED',
  // w:webHidden hides a run in Word's Web Layout view only. Paginated layout
  // draws it like any other run, so the fact is recorded and the run is
  // painted; it cannot move a line or a page here.
  'WEB_LAYOUT_HIDDEN_RUN_PRESERVED',
  // w:cols w:sep="1" asks for a vertical rule drawn in the inter-column gap.
  // It is ink this tier cannot paint - the paint plan binds its only straight
  // rule, stroke_note_separator, to a placed note-separator line, and a column
  // separator has no line to bind to - but it moves no column box. Measured on
  // Word's own export of multi-column-separator-with-line.docx: the rule is a
  // filled 0.96 pt wide bar centred on x=306 pt, the exact centre of the gap
  // between the two columns, while the column text origins are 50.4 pt and
  // 324.0 pt - value-for-value the equal-width boxes this model derives from
  // w:pgSz, w:pgMar and w:cols with w:sep ignored. So the omission drops that
  // one bar and nothing else; it is disclosed as omitted content and the
  // strict tier keeps refusing it.
  'COLUMN_SEPARATOR_UNSUPPORTED',
  // Word 2010 run-typography extensions this tier records and does not apply:
  // w14:stylisticSets, a w14:ligatures mode outside the shaper defaults,
  // w14:numForm, w14:numSpacing and the 3D text effect w14:props3d. Unlike
  // LIGATURE_MODE_MATCHES_SHAPER and CONTEXTUAL_ALTERNATES_MATCH_SHAPER, these
  // do NOT state shaping this tier already performs - the page is painted with
  // the feature unapplied, at advances that are measurably not Word's. Shaped
  // with harfbuzzjs 1.6.0 over 224 face/sample pairs: ss02 changes 20 of them
  // (Arial +28,745 font units across the sample), ss04 6 (Calibri +112),
  // numForm="oldStyle" 68, numSpacing="proportional" 68 (a ten-digit Calibri
  // run 312 font units narrower at 2048 upem). They are admitted here only
  // because every one of them is also a member of
  // DOCX_APPROXIMATE_OMITTED_CONTENT_CODES, so each unapplied property is named
  // in omitted_content and the preview reports content_status 'partial'; the
  // strict tier keeps refusing all five. Any other foreign run property stays
  // FOREIGN_RUN_PROPERTY and still refuses on both tiers.
  'STYLISTIC_SET_UNAPPLIED',
  'LIGATURE_MODE_UNAPPLIED',
  'NUMBER_FORM_UNAPPLIED',
  'NUMBER_SPACING_UNAPPLIED',
  'TEXT_EFFECT_3D_UNAPPLIED',
  // w:tcPr/w:textDirection asks for the cell's text to be rotated 90 or 270
  // degrees, or stacked vertically. This tier lays the cell out and paints it
  // horizontally, so the cell's glyphs are present but the rotation is not:
  // it is a visual result the preview does not produce, disclosed as omitted
  // content rather than a refusal that produces no page at all.
  'CELL_TEXT_DIRECTION_UNSUPPORTED',
  // A negative w:line measurement. ECMA-376 17.3.1.33 types w:line as
  // ST_SignedTwipsMeasure and Word reads a negative one as an EXACT line height
  // of its absolute value, compressing the lines until they overlap and
  // ignoring the authored w:lineRule. This tier has no compressed line box -
  // DOCX_APPROXIMATE_LINE_BOX_WARNING already says so - so the resolver drops
  // the measurement and the paragraph keeps the line spacing it inherits. The
  // page is painted with the compression UNAPPLIED and its lines sit further
  // apart than Word's, by the difference between the inherited spacing and the
  // absolute authored value. Measured on Word's own export of
  // tdf125469_singleSpacing.docx: Word puts that document's paragraphs at
  // baselines 32.64, 46.56 and 29.28 pt apart where the inherited spacing this
  // tier keeps is a 43.95 pt natural line box at 36 pt. Admitted here only
  // because it is also a member of
  // DOCX_APPROXIMATE_OMITTED_CONTENT_CODES, so the unapplied compression is
  // named in omitted_content; the strict tier keeps refusing it, and a negative
  // w:before or w:after, or any unparseable measurement, stays
  // UNMODELED_PARAGRAPH_SPACING / INVALID_LINE_SPACING and refuses on both.
  'NEGATIVE_LINE_SPACING_UNAPPLIED',
])
export const DOCX_APPROXIMATE_PREVIEW_WARNING = 'Approximate read-only preview: current InjOffice layout, not Microsoft Word compatibility-mode fidelity.' as const
/** Declared whenever the approximate preview produced no page. The refusal
 * cause is otherwise reachable only by reading the paint diagnostics array, so
 * callers that render `reasons` saw an empty preview with no explanation. */
export const DOCX_APPROXIMATE_PAINT_REFUSED_WARNING = 'Approximate read-only preview produced no page. The reasons below beginning "Approximate paint refused" name the exact source facts that stopped it.' as const
/** Bound on how many painter diagnostics are restated as envelope reasons. */
export const DOCX_APPROXIMATE_PAINT_REFUSED_MAX_REASONS = 8
/** Byte-identical restatement of one painter diagnostic as an envelope reason. */
export function nativeDocxApproximatePaintRefusalReason(diagnostic: NativeDocxPagePaintV1['diagnostics'][number]): string {
  return `Approximate paint refused (${diagnostic.code} at ${diagnostic.scope_id}): ${diagnostic.message}`
}
export const DOCX_APPROXIMATE_LINE_BOX_WARNING = 'Current-layout policy places natural ascent at the top of an expanded automatic line box, leaving extra leading below the text; an expanded exact or at-least line box instead seats the descent on the box bottom, leaving the leading above the text. Compressed line boxes remain unsupported.' as const
export interface NativeDocxApproximationEligibilityV1 {
  protocol: 'injoffice.docx.approximation-eligibility'
  version: 1
  document_id: string
  revision: string
  package_sha256: string
  settings_sha256: string | null
  status: 'eligible' | 'ineligible'
  legacy_compatibility_mode: 12 | 14 | 15 | null
  reasons: string[]
  approximated_settings?: NativeDocxApproximatedSettingV1[]
  absent_font_sizes?: NativeDocxAbsentFontSizeV1[]
  /** Which source shape proved the absence. Microsoft Word resolves the two
   * shapes to different sizes, so the host default is selected from it. */
  absent_font_size_shape?: NativeDocxAbsentDefaultSizeShapeV1
  absent_font_families?: NativeDocxAbsentFontFamilyV1[]
  latin_font_fallbacks?: NativeDocxLatinFontFallbackV1[]
  legacy_table_origins?: NativeDocxLegacyTableOriginV1[]
}

export interface NativeDocxApproximatePagePreviewV1 {
  protocol: typeof DOCX_APPROXIMATE_PREVIEW_PROTOCOL
  version: 1
  fidelity: 'approximate'
  policy: typeof DOCX_APPROXIMATE_PREVIEW_POLICY
  read_only: true
  table_border_layout_policy?: 'collapsed-horizontal-border-reservation-v1'
  table_width_policy?: 'approximate-authored-grid-fitted-v1'
  status: 'painted' | 'refused'
  source: Pick<NativeDocxApproximationEligibilityV1, 'document_id' | 'revision' | 'package_sha256' | 'settings_sha256'>
  reasons: string[]
  source_settings_diagnostics: NativeDocxPaginationSettingsV1['diagnostics']
  approximated_settings?: NativeDocxApproximatedSettingV1[]
  source_absent_font_sizes?: NativeDocxAbsentFontSizeV1[]
  /** The source shape the host default was selected for, carried so a reader
   * can check the applied size without re-reading the package. */
  source_absent_font_size_shape?: NativeDocxAbsentDefaultSizeShapeV1
  approximated_font_sizes?: NativeDocxApproximatedFontSizeV1[]
  source_absent_font_families?: NativeDocxAbsentFontFamilyV1[]
  approximated_font_families?: NativeDocxApproximatedFontFamilyV1[]
  source_latin_font_fallbacks?: NativeDocxLatinFontFallbackV1[]
  approximated_font_faces?: NativeDocxLatinFontFallbackV1[]
  /** Picture extents this preview moved onto the milli-point lattice, with both
   * the authored and the painted EMU. Present only when one was moved. */
  approximated_image_extents?: NativeDocxApproximatedImageExtentV1[]
  legacy_table_origins?: NativeDocxLegacyTableOriginV1[]
  rendering_provenance: NativeDocxPagePaintV1['provenance']
  diagnostics: NativeDocxPagePaintV1['diagnostics']
  resources: NativeDocxPagePaintV1['resources']
  pages: NativeDocxPagePaintV1['pages']
  /** 'partial' whenever source content was dropped or a page painted nothing;
   * `status: 'painted'` alone never attests completeness. */
  content_status: NativeDocxApproximateOmissionsV1['content_status']
  omitted_content: NativeDocxApproximateOmissionsV1['omitted_content']
  omitted_content_total: NativeDocxApproximateOmissionsV1['omitted_content_total']
  unpainted_pages: NativeDocxApproximateOmissionsV1['unpainted_pages']
}

/** Eligibility is produced from original package bytes by the native extractor,
 * not inferred from a generic unsupported diagnostic code. */
export function decodeNativeDocxApproximationEligibilityV1(value: unknown, settings: NativeDocxPaginationSettingsV1): NativeDocxApproximationEligibilityV1 {
  const candidate = value as Partial<NativeDocxApproximationEligibilityV1> | null
  const issues = preflightWire(value, 'approximation eligibility', 20_000, 1000).filter(issue => !(issue.code === 'INVALID_VALUE' && ((issue.path === '/settings_sha256' && candidate?.settings_sha256 === null) || (issue.path === '/legacy_compatibility_mode' && candidate?.legacy_compatibility_mode === null))))
  if (issues.length) throw new TypeError('invalid approximation eligibility wire')
  const input = structuredClone(value) as NativeDocxApproximationEligibilityV1
  if (!input || typeof input !== 'object' || Object.keys(input).filter(key => key !== 'approximated_settings' && key !== 'absent_font_sizes' && key !== 'absent_font_size_shape' && key !== 'absent_font_families' && key !== 'latin_font_fallbacks' && key !== 'legacy_table_origins').sort().join(',') !== 'document_id,legacy_compatibility_mode,package_sha256,protocol,reasons,revision,settings_sha256,status,version'
    || input.protocol !== 'injoffice.docx.approximation-eligibility' || input.version !== 1
    || !['eligible', 'ineligible'].includes(input.status) || ![null, 12, 14, 15].includes(input.legacy_compatibility_mode)
    || !Array.isArray(input.reasons) || input.reasons.length > 256 || input.reasons.some(reason => typeof reason !== 'string' || reason.length > 8192)
    || input.document_id !== settings.document_id || input.revision !== settings.revision || input.package_sha256 !== settings.package_sha256 || input.settings_sha256 !== (settings.settings_sha256 ?? null)) throw new TypeError('approximation eligibility does not exact-join original settings')
  const facts = input.approximated_settings ?? []
  if(input.legacy_table_origins!==undefined&&(!validLegacyTableOrigins(input.legacy_table_origins,input.package_sha256)||input.legacy_compatibility_mode!==12||input.status!=='eligible'))throw new TypeError('Invalid legacy table origin evidence')
  if (input.absent_font_sizes !== undefined && !validNativeDocxAbsentFontSizesV1(input.absent_font_sizes, input.package_sha256)) throw new TypeError('Invalid source-absent font-size evidence')
  if ((input.absent_font_size_shape !== undefined) !== ((input.absent_font_sizes?.length ?? 0) > 0) || (input.absent_font_size_shape !== undefined && !validNativeDocxAbsentDefaultSizeShapeV1(input.absent_font_size_shape))) throw new TypeError('Invalid source-absent font-size shape')
  if (input.absent_font_families !== undefined && !validNativeDocxAbsentFontFamiliesV1(input.absent_font_families, input.package_sha256)) throw new TypeError('Invalid source-absent font-family evidence')
  if (input.latin_font_fallbacks !== undefined && !validNativeDocxLatinFontFallbacksV1(input.latin_font_fallbacks, input.package_sha256)) throw new TypeError('Invalid Latin font fallback evidence')
  if (!Array.isArray(facts) || facts.length > DOCX_APPROXIMATE_MAX_SETTING_FACTS || facts.some(fact => !validNativeDocxApproximatedSettingV1(fact)) || new Set(facts.map(fact => fact.kind)).size !== facts.length || new Set(facts.map(fact => fact.path)).size !== facts.length) throw new TypeError('invalid approximated settings source facts')
  if (input.status === 'eligible' && (input.legacy_compatibility_mode === null || settings.profile === 'word-modern-default' || (input.legacy_compatibility_mode === 15) !== (settings.compatibility_mode === 15) || !coveredSettingsDiagnostics(settings, facts) || facts.some(fact => !input.reasons.includes(nativeApproximationSettingReason(fact))) || input.reasons.length === 0)) throw new TypeError('approximation eligibility conflicts with strict settings facts')
  return input
}

export function approximatePagePreviewEnvelope(settings: NativeDocxPaginationSettingsV1, eligibility: NativeDocxApproximationEligibilityV1, paint: NativeDocxPagePaintV1, source?: NativeDocxApproximateOmissionSourceV1): NativeDocxApproximatePagePreviewV1 {
  if (paint.status === 'painted' && !source) throw new TypeError('Painted approximate output requires its source for omitted-content disclosure')
  const omissions = source ? collectNativeDocxApproximateOmissionsV1(source, paint) : nativeDocxApproximateRefusalOmissionsV1()
  // A refused paint carries its cause only in the paint diagnostics. Restate it
  // in the reasons vector so the disclosure surface every caller already reads
  // says why no page exists, within the decoder's bounded reason count.
  const refused = paint.status === 'refused'
    ? [DOCX_APPROXIMATE_PAINT_REFUSED_WARNING, ...paint.diagnostics.slice(0, DOCX_APPROXIMATE_PAINT_REFUSED_MAX_REASONS).map(nativeDocxApproximatePaintRefusalReason)]
    : []
  return {
    protocol: DOCX_APPROXIMATE_PREVIEW_PROTOCOL, version: 1, fidelity: 'approximate', policy: DOCX_APPROXIMATE_PREVIEW_POLICY, read_only: true,
    table_border_layout_policy:'collapsed-horizontal-border-reservation-v1',
    table_width_policy:'approximate-authored-grid-fitted-v1',
    status: paint.status,
    source: { document_id: settings.document_id, revision: settings.revision, package_sha256: settings.package_sha256, settings_sha256: settings.settings_sha256 ?? null },
    reasons: boundedApproximateReasons([...eligibility.reasons, DOCX_APPROXIMATE_PREVIEW_WARNING, DOCX_APPROXIMATE_LINE_BOX_WARNING,DOCX_TABLE_BORDER_RESERVATION_WARNING,DOCX_TABLE_GRID_FIT_WARNING,...(eligibility.legacy_table_origins?.length?[DOCX_LEGACY_TABLE_ORIGIN_WARNING]:[]),...(omissions.omitted_content.length||omissions.unpainted_pages.length?[DOCX_APPROXIMATE_OMITTED_CONTENT_WARNING]:[])], refused),
    ...(eligibility.legacy_table_origins?{legacy_table_origins:structuredClone(eligibility.legacy_table_origins)}:{}),
    source_settings_diagnostics: structuredClone(settings.diagnostics),
    ...(eligibility.approximated_settings ? { approximated_settings: structuredClone(eligibility.approximated_settings) } : {}),
    ...(eligibility.absent_font_sizes ? { source_absent_font_sizes: structuredClone(eligibility.absent_font_sizes) } : {}),
    ...(eligibility.absent_font_size_shape ? { source_absent_font_size_shape: eligibility.absent_font_size_shape } : {}),
    ...(eligibility.absent_font_families ? { source_absent_font_families: structuredClone(eligibility.absent_font_families) } : {}),
    ...(eligibility.latin_font_fallbacks ? { source_latin_font_fallbacks: structuredClone(eligibility.latin_font_fallbacks) } : {}),
    rendering_provenance: paint.provenance,
    diagnostics: paint.diagnostics, resources: paint.resources, pages: paint.pages,
    ...omissions,
  }
}

/** Browser-safe structural validation of the distinct approximate artifact.
 * This does not turn its pages into a strict source-qualified V1 projection. */
export function decodeNativeDocxApproximatePagePreviewV1(value: unknown): { ok: true; value: NativeDocxApproximatePagePreviewV1 } | { ok: false; issues: NativeDocxValidationIssue[] } {
  const invalid = (message: string): { ok: false; issues: NativeDocxValidationIssue[] } => ({ ok: false, issues: [{ code: 'INVALID_VALUE', path: '', message }] })
  try {
    const candidate = value as Partial<NativeDocxApproximatePagePreviewV1> | null
    const issues = preflightWire(value, 'approximate page preview').filter(issue => !(issue.code === 'INVALID_VALUE' && issue.path === '/source/settings_sha256' && candidate?.source?.settings_sha256 === null))
    if (issues.length) return { ok: false, issues }
    const input = structuredClone(value) as NativeDocxApproximatePagePreviewV1
    if (!input || typeof input !== 'object' || Object.keys(input).filter(key => !['approximated_settings', 'source_absent_font_sizes', 'source_absent_font_size_shape', 'approximated_font_sizes', 'source_absent_font_families', 'approximated_font_families', 'source_latin_font_fallbacks', 'approximated_font_faces','approximated_image_extents','legacy_table_origins','table_border_layout_policy','table_width_policy'].includes(key)).sort().join(',') !== 'content_status,diagnostics,fidelity,omitted_content,omitted_content_total,pages,policy,protocol,read_only,reasons,rendering_provenance,resources,source,source_settings_diagnostics,status,unpainted_pages,version'
      || input.protocol !== DOCX_APPROXIMATE_PREVIEW_PROTOCOL || input.version !== 1 || input.fidelity !== 'approximate' || input.policy !== DOCX_APPROXIMATE_PREVIEW_POLICY || input.read_only !== true
      || !Array.isArray(input.reasons) || input.reasons.length < 1 || input.reasons.length > 264 || !input.reasons.includes(DOCX_APPROXIMATE_PREVIEW_WARNING) || !input.reasons.includes(DOCX_APPROXIMATE_LINE_BOX_WARNING) || input.reasons.some(reason => typeof reason !== 'string' || reason.length > 8192)) return invalid('invalid approximate envelope or missing fidelity warning')
    // A refused approximate preview must say so in its reasons, and a painted one
    // must not claim a refusal it did not make.
    if ((input.status === 'refused') !== input.reasons.includes(DOCX_APPROXIMATE_PAINT_REFUSED_WARNING)) return invalid('Refused approximate preview must declare its refusal in reasons')
    const paint = decodeNativeDocxPagePaintV1({ protocol: DOCX_PAGE_PAINT_PROTOCOL, version: DOCX_PAGE_PAINT_VERSION, status: input.status, provenance: input.rendering_provenance, diagnostics: input.diagnostics, resources: input.resources, pages: input.pages })
    if(input.table_border_layout_policy!==undefined&&(input.table_border_layout_policy!=='collapsed-horizontal-border-reservation-v1'||!input.reasons.includes(DOCX_TABLE_BORDER_RESERVATION_WARNING)))return invalid('Invalid declared table border reservation policy or warning')
    if(input.table_width_policy!==undefined&&(input.table_width_policy!=='approximate-authored-grid-fitted-v1'||!input.reasons.includes(DOCX_TABLE_GRID_FIT_WARNING)))return invalid('Invalid declared table width policy or warning')
    if (!paint.ok) return paint
    if (!validNativeDocxApproximateOmissionsV1(input, paint.value) || (input.omitted_content.length > 0 || input.unpainted_pages.length > 0) !== input.reasons.includes(DOCX_APPROXIMATE_OMITTED_CONTENT_WARNING)) return invalid('Omitted-content disclosure does not agree with painted pages or its warning')
    const settings = paint.value.provenance.pagination_settings
    if(input.legacy_table_origins!==undefined&&(!validLegacyTableOrigins(input.legacy_table_origins,settings.package_sha256)||(input.legacy_table_origins.length>0&&!input.reasons.includes(DOCX_LEGACY_TABLE_ORIGIN_WARNING))))return invalid('Missing legacy origin evidence or warning')
    const absent = input.source_absent_font_sizes ?? []
    if (!validNativeDocxAbsentFontSizesV1(absent, settings.package_sha256)) return invalid('Invalid source font-size omissions')
    if ((input.source_absent_font_size_shape !== undefined) !== (absent.length > 0)) return invalid('Source font-size omissions do not declare their proven shape')
    if (input.approximated_font_sizes !== undefined) {
      if (!validNativeDocxApproximatedFontSizesV1(input.approximated_font_sizes, absent, settings.package_sha256, input.source_absent_font_size_shape) || !input.reasons.includes(DOCX_ABSENT_FONT_SIZE_WARNING)) return invalid('Missing explicit host size policy, source evidence or warning')
    } else if (input.status === 'painted' && absent.length > 0) return invalid('Painted source-absent sizes require declared host policy')
    const absentFamilies = input.source_absent_font_families ?? []
    if (!validNativeDocxAbsentFontFamiliesV1(absentFamilies, settings.package_sha256)) return invalid('Invalid source font-family omissions')
    if (input.approximated_font_families !== undefined && (!validNativeDocxApproximatedFontFamiliesV1(input.approximated_font_families, absentFamilies, settings.package_sha256) || !input.reasons.includes(DOCX_ABSENT_FONT_FAMILY_WARNING))) return invalid('Applied host default families require retained source evidence and the declared warning')
    if (input.approximated_font_families === undefined && input.reasons.includes(DOCX_ABSENT_FONT_FAMILY_WARNING)) return invalid('Declared host default family requires its applied facts')
    const fallbacks = input.source_latin_font_fallbacks ?? []
    if (!validNativeDocxLatinFontFallbacksV1(fallbacks, settings.package_sha256)) return invalid('Invalid Latin font fallback evidence')
    if (input.approximated_font_faces !== undefined && (!validNativeDocxApproximatedFontFacesV1(input.approximated_font_faces, fallbacks, settings.package_sha256) || !input.reasons.includes(DOCX_LATIN_FONT_FALLBACK_WARNING))) return invalid('Applied Latin font fallbacks require retained source evidence and the declared warning')
    // An approximation is recorded, never silent: a moved picture extent must
    // carry both its authored and painted EMU and declare the rounding warning.
    if (input.approximated_image_extents !== undefined && (!validNativeDocxApproximatedImageExtentsV1(input.approximated_image_extents) || input.approximated_image_extents.length === 0 || input.approximated_image_extents.length > DOCX_APPROXIMATE_MAX_IMAGE_EXTENT_FACTS || !input.reasons.includes(DOCX_APPROXIMATE_IMAGE_EXTENT_WARNING))) return invalid('Rounded picture extents require exact source/painted facts and the declared warning')
    if (input.approximated_image_extents === undefined && input.reasons.includes(DOCX_APPROXIMATE_IMAGE_EXTENT_WARNING)) return invalid('Declared picture extent rounding requires its retained source facts')
    const facts = input.approximated_settings ?? []
    if (!Array.isArray(facts) || facts.length > DOCX_APPROXIMATE_MAX_SETTING_FACTS || facts.some(fact => !validNativeDocxApproximatedSettingV1(fact)) || new Set(facts.map(fact => fact.kind)).size !== facts.length || new Set(facts.map(fact => fact.path)).size !== facts.length) return invalid('invalid retained approximate settings facts')
    if (facts.some(fact => !input.reasons.includes(nativeApproximationSettingReason(fact))) || (input.status === 'painted' && !coveredSettingsDiagnostics(settings, facts))) return invalid('missing approximate setting coverage or warning')
    if (paint.value.provenance.document_id !== settings.document_id || paint.value.provenance.revision !== settings.revision || paint.value.provenance.package_sha256 !== settings.package_sha256) return invalid('approximate rendering identity does not match retained source settings')
    if (!input.source || Object.keys(input.source).sort().join(',') !== 'document_id,package_sha256,revision,settings_sha256' || input.source.document_id !== settings.document_id || input.source.revision !== settings.revision || input.source.package_sha256 !== settings.package_sha256 || input.source.settings_sha256 !== (settings.settings_sha256 ?? null) || JSON.stringify(input.source_settings_diagnostics) !== JSON.stringify(settings.diagnostics)) return invalid('approximate source facts and retained settings diagnostics do not exact-join')
    return { ok: true, value: input }
  } catch { return invalid('approximate output could not be safely inspected') }
}

/** Keeps the declared refusal reasons inside the decoder's 264-reason bound by
 * dropping repeated eligibility detail first, never the refusal declaration. */
function boundedApproximateReasons(declared: string[], refused: string[]): string[] {
  const limit = 264
  if (declared.length + refused.length <= limit) return [...declared, ...refused]
  return [...declared.slice(0, Math.max(0, limit - refused.length)), ...refused.slice(0, limit)]
}

function coveredSettingsDiagnostics(settings: NativeDocxPaginationSettingsV1, facts: NativeDocxApproximatedSettingV1[]): boolean {
  return settings.diagnostics.every(reason => {
    if (reason.code === 'COMPATIBILITY_SETTING_UNSUPPORTED' && (['/w:settings[1]', '/w:settings[1]/w:compat[1]', '/w:settings[1]/w:compat[1]/w:compatSetting[1]', '/w:settings[1]/w:compat[1]/w:applyBreakingRules[1]'].includes(reason.path) || /^\/w:settings\[1\]\/w:compat\[1\]\/w:compatSetting\[[5-9]\]$/.test(reason.path))) return true
    if (reason.code === 'PAGINATION_SETTING_UNSUPPORTED' || reason.code === 'UNKNOWN_SETTINGS_ELEMENT' || reason.code === 'DUPLICATE_SETTINGS_PROPERTY') return true
    // Grouped repeated compatSetting attestations disclose each member by its own
    // settings.xml path, so the join is on the member key, not the anchor path.
    if (reason.code === 'COMPATIBILITY_SETTING_UNSUPPORTED' && facts.some(fact => fact.kind === 'repeatedCompatSettings' && reason.path in fact.values)) return true
    return facts.some(fact => fact.path === reason.path && reason.code === (fact.kind === 'mathPr' ? 'UNKNOWN_SETTINGS_ELEMENT' : fact.kind === 'characterSpacingControl' ? 'CHARACTER_SPACING_CONTROL_UNSUPPORTED' : fact.path.includes('/w:compat[1]/') ? 'COMPATIBILITY_SETTING_UNSUPPORTED' : 'PAGINATION_SETTING_UNSUPPORTED'))
  }) && facts.every(fact => settings.diagnostics.some(reason => reason.path === fact.path))
}
