import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { NativeFontManifest } from '@injoffice/font-metrics/layout'
import { BIDI_UNICODE_VERSION, NATIVE_BIDI_PROVIDER_ID, NATIVE_BIDI_PROVIDER_REVISION } from '@injoffice/font-metrics/bidi'
import { UNICODE_13_CLASSIFIER_REVISION } from '@injoffice/font-metrics/unicode13'
import {
  DOCX_NATIVE_LIMITS,
  DOCX_NATIVE_PROTOCOL,
  DOCX_NATIVE_VERSION,
  type NativeDocxDocumentV1,
} from './nativeContract.js'
import {
  DOCX_RESOLVED_LAYOUT_PROTOCOL,
  DOCX_RESOLVED_LAYOUT_VERSION,
  nativeDocxResolvedNumberingDefinitionSha256V1,
  nativeDocxResolvedNumberingModelSha256V1,
  type NativeDocxResolvedLayoutInputV1,
} from './nativeResolvedLayout.js'
import {
  DOCX_SHAPED_LINES_PROTOCOL,
  DOCX_SHAPED_LINES_VERSION,
  type NativeDocxShapedLinesV1,
} from './nativeShapingLines.js'
import { nativeDocxTableProjectionSha256V1 } from './nativeTablePagePaintV1.js'
import {
  DOCX_DEFAULT_TAB_STOP_TWIPS,
  DOCX_PAGINATION_SETTINGS_PROTOCOL,
  DOCX_PAGINATION_SETTINGS_VERSION,
  type NativeDocxPaginationSettingsV1,
} from './nativePaginationSettings.js'
import {
  DOCX_PAGINATION_REQUEST_PROTOCOL,
  DOCX_PAGINATION_REQUEST_VERSION,
  paginateNativeDocxV1,
  type NativeDocxPaginationRequestV1,
} from './nativePaginationV1.js'
import {
  DOCX_PAGE_PAINT_PROTOCOL,
  DOCX_PAGE_PAINT_REQUEST_PROTOCOL,
  DOCX_PAGE_PAINT_REQUEST_V1_BINDING_FIELDS,
  DOCX_PAGE_PAINT_REQUEST_VERSION,
  DOCX_PAGE_PAINT_V1_BINDING_FIELDS,
  compileNativeDocxPagePaintV1,
  compileNativeDocxApproximatePagePreviewV1,
  decodeNativeDocxPagePaintForRequestV1,
  decodeNativeDocxPagePaintRequestV1,
  decodeNativeDocxPagePaintV1,
  nativeDocxPagePaintFontManifestSha256V1,
  nativeDocxPagePaintMediaAssetsSha256V1,
  nativeDocxPagePaintPaginatedLayoutSha256V1,
  nativeDocxPagePaintShapedLinesSha256V1,
  validateNativeDocxPagePaintForRequestV1,
  type NativeDocxGlyphOutlineProviderV1,
  type NativeDocxGlyphOutlineRequestV1,
  type NativeDocxGlyphOutlineResultV1,
  type NativeDocxPagePaintRequestV1,
} from './nativePagePaintV1.js'
import { decodeNativeDocxApproximatePagePreviewV1, decodeNativeDocxApproximationEligibilityV1 } from './nativeApproximationV1.js'

const HASH = `sha256:${'a'.repeat(64)}` as `sha256:${string}`
const RELATIONSHIPS_HASH = `sha256:${'b'.repeat(64)}`
const SETTINGS_PART = 'word/settings.xml'
const RELATIONSHIPS_PART = 'word/_rels/document.xml.rels'
const NUMBERING_PART = 'word/numbering.xml'
const NUMBERING_HASH = `sha256:${'c'.repeat(64)}`

function anchor(path: string, start: number, end: number) {
  return { part_name: 'word/document.xml', path, start_byte: start, end_byte: end, xml_sha256: HASH }
}

function paginationRequest(): NativeDocxPaginationRequestV1 {
  const paragraph = {
    id: 'paragraph:1',
    anchor: anchor('/w:document[1]/w:body[1]/w:p[1]', 100, 190),
    edit_policy: { mode: 'read-only' as const, allowed_operations: [], refusal: { code: 'NATIVE_READ_ONLY', message: 'Fixture is immutable.', preservation: 'refuse-mutation' as const } },
    properties: {},
    runs: [{ kind: 'text' as const, id: 'run:1', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]', 110, 180), text: 'AA' }],
  }
  const document: NativeDocxDocumentV1 = {
    protocol: DOCX_NATIVE_PROTOCOL,
    version: DOCX_NATIVE_VERSION,
    document_id: 'document:test',
    revision: 'revision:1',
    source: { package_sha256: HASH, main_part: 'word/document.xml' },
    body: { id: 'story:body', kind: 'body', part_name: 'word/document.xml', anchor: anchor('/w:document[1]/w:body[1]', 1, 3_000), blocks: [{ kind: 'paragraph', id: paragraph.id, paragraph }] },
    sections: [{
      id: 'section:1',
      anchor: anchor('/w:document[1]/w:body[1]/w:sectPr[1]', 2_100, 2_190),
      starts_at_block_id: paragraph.id,
      break_type: 'next-page',
      title_page: false,
      page: {
        width_twips: 1_000,
        height_twips: 1_000,
        orientation: 'portrait',
        margins: { top_twips: 100, right_twips: 100, bottom_twips: 100, left_twips: 100, header_twips: 50, footer_twips: 50, gutter_twips: 0 },
        columns: 1,
        column_spacing_twips: 100,
        column_layout: 'equal-width',
        column_definitions: [{ id: 'column:section:1:0', ordinal: 0 }],
      },
      header_refs: [],
      footer_refs: [],
    }],
    headers: [], footers: [], notes: [], comment_stories: [], comments: [], capabilities: [],
    passthrough_parts: [
      { part_name: SETTINGS_PART, content_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml', byte_length: 1, sha256: HASH, policy: 'preserve-verbatim' },
      { part_name: RELATIONSHIPS_PART, content_type: 'application/vnd.openxmlformats-package.relationships+xml', byte_length: 1, sha256: RELATIONSHIPS_HASH, policy: 'preserve-verbatim' },
    ],
    unsupported: [],
  }
  const resolved: NativeDocxResolvedLayoutInputV1 = {
    protocol: DOCX_RESOLVED_LAYOUT_PROTOCOL,
    version: DOCX_RESOLVED_LAYOUT_VERSION,
    document_id: document.document_id,
    revision: document.revision,
    source_parts: { main_part: 'WORD/document.xml' },
    paragraphs: [{ paragraph_id: paragraph.id, applied_styles: [], properties: {}, paragraph_mark_properties: { font_family: 'Test', font_size_half_points: 20 } }],
    runs: [{ run_id: 'run:1', paragraph_id: paragraph.id, applied_paragraph_styles: [], applied_character_styles: [], properties: { font_family: 'Test', font_size_half_points: 20, color: '123456' } }],
    tables: [], fonts: [{ name: 'Test' }], diagnostics: [],
  }
  const shaped: NativeDocxShapedLinesV1 = {
    protocol: DOCX_SHAPED_LINES_PROTOCOL,
    version: DOCX_SHAPED_LINES_VERSION,
    document_id: document.document_id,
    revision: document.revision,
    available_width_millipoints: 40_000,
    tab_interval_millipoints: DOCX_DEFAULT_TAB_STOP_TWIPS * 50,
    font_manifest: { manifest_id: 'manifest:test', revision: 'manifest-revision:1' },
    providers: { resolver_id: 'resolver:test', resolver_revision: '1', shaper_id: 'shaper:test', shaper_revision: '1', bidi_id: NATIVE_BIDI_PROVIDER_ID, bidi_revision: NATIVE_BIDI_PROVIDER_REVISION, bidi_unicode_version: BIDI_UNICODE_VERSION, unicode13_revision: UNICODE_13_CLASSIFIER_REVISION },
    paragraphs: [{
      paragraph_id: paragraph.id,
      story_id: document.body.id,
      story_kind: 'body',
      direction: 'ltr',
      alignment: 'start',
      spacing_before_millipoints: 0,
      spacing_after_millipoints: 0,
      indent_start_millipoints: 0,
      indent_end_millipoints: 0,
      first_line_delta_millipoints: 0,
      block_advance_millipoints: 10_000,
      lines: [{
        id: `line:${paragraph.id}:0`,
        ordinal: 0,
        available_width_millipoints: 40_000,
        inline_offset_millipoints: 0,
        advance_inline_millipoints: 10_000,
        ascent_millipoints: 8_000,
        descent_millipoints: -2_000,
        line_gap_millipoints: 0,
        line_height_millipoints: 10_000,
        justified: false,
        logical_to_visual: [0],
        fragments: [{
          id: `fragment:${paragraph.id}:0:0`,
          source_kind: 'run',
          source_id: 'run:1',
          start_utf16: 0,
          end_utf16: 2,
          text: 'AA',
          direction: 'ltr',
          bidi_level: 0,
          logical_order: 0,
          script: 'Latn',
          language: 'en-US',
          face_id: 'face:test',
          whitespace: false,
          advance_inline_millipoints: 10_000,
          justification_expansion_millipoints: 0,
          ascent_millipoints: 8_000,
          descent_millipoints: -2_000,
          line_gap_millipoints: 0,
          glyphs: [
            { glyph_id: 7, advance_x_millipoints: 5_000, advance_y_millipoints: 0, offset_x_millipoints: 0, offset_y_millipoints: 0 },
            { glyph_id: 7, advance_x_millipoints: 5_000, advance_y_millipoints: 0, offset_x_millipoints: 0, offset_y_millipoints: 0 },
          ],
        }],
      }],
    }],
    diagnostics: [],
  }
  const settings: NativeDocxPaginationSettingsV1 = {
    protocol: DOCX_PAGINATION_SETTINGS_PROTOCOL,
    version: DOCX_PAGINATION_SETTINGS_VERSION,
    document_id: document.document_id,
    revision: document.revision,
    package_sha256: HASH,
    main_part: 'word/document.xml',
    relationships_part: RELATIONSHIPS_PART,
    relationships_sha256: RELATIONSHIPS_HASH,
    relationship_id: 'rIdSettings',
    settings_part: SETTINGS_PART,
    settings_sha256: HASH,
    profile: 'word-modern-default',
    default_tab_stop_twips: DOCX_DEFAULT_TAB_STOP_TWIPS,
    mirror_margins: false,
    gutter_at_top: false,
    even_and_odd_headers: false,
    compatibility_mode: 15,
    diagnostics: [],
  }
  return { protocol: DOCX_PAGINATION_REQUEST_PROTOCOL, version: DOCX_PAGINATION_REQUEST_VERSION, document, resolved_layout: resolved, shaped_lines: shaped, pagination_settings: settings }
}

function manifest(): NativeFontManifest {
  return {
    version: 1,
    manifestId: 'manifest:test',
    revision: 'manifest-revision:1',
    faces: [{ faceId: 'face:test', family: 'Test', weight: 400, style: 'normal', stretch: 100, source: { kind: 'bundled', resourceId: 'font:test', contentDigest: HASH } }],
    fallbackChains: [],
  }
}

function fixture(): NativeDocxPagePaintRequestV1 {
  const pagination = paginationRequest()
  const layout = paginateNativeDocxV1(pagination)
  if (!layout.ok || layout.value.status !== 'paginated') throw new Error(JSON.stringify(layout))
  const fontManifest = manifest()
  return {
    protocol: DOCX_PAGE_PAINT_REQUEST_PROTOCOL,
    version: DOCX_PAGE_PAINT_REQUEST_VERSION,
    pagination_request: pagination,
    paginated_layout: layout.value,
    font_manifest: fontManifest,
    media_assets: [],
    integrity: {
      font_manifest_sha256: nativeDocxPagePaintFontManifestSha256V1(fontManifest),
      shaped_lines_sha256: nativeDocxPagePaintShapedLinesSha256V1(pagination.shaped_lines),
      table_projection_sha256: nativeDocxTableProjectionSha256V1([]),
      media_assets_sha256: nativeDocxPagePaintMediaAssetsSha256V1([]),
      paginated_layout_sha256: nativeDocxPagePaintPaginatedLayoutSha256V1(layout.value),
    },
    outline_provider: { provider_id: 'outline:test', provider_revision: '1' },
  }
}

class FixtureProvider implements NativeDocxGlyphOutlineProviderV1 {
  providerId = 'outline:test'
  providerRevision = '1'
  calls = 0
  result?: (request: Readonly<NativeDocxGlyphOutlineRequestV1>) => NativeDocxGlyphOutlineResultV1

  getGlyphOutline(request: Readonly<NativeDocxGlyphOutlineRequestV1>): ReturnType<NativeDocxGlyphOutlineProviderV1['getGlyphOutline']> {
    this.calls += 1
    if (this.result) return this.result(request)
    return {
      status: 'outlined', face: { ...request.face }, glyph_id: request.glyph_id, units_per_em: 1_000,
      path: [
        { kind: 'move_to', x: 0, y: 0 },
        { kind: 'line_to', x: 500, y: 0 },
        { kind: 'line_to', x: 500, y: 700 },
        { kind: 'line_to', x: 0, y: 700 },
        { kind: 'close_path' },
      ],
    }
  }
}

async function painted(request = fixture(), provider = new FixtureProvider()) {
  const result = await compileNativeDocxPagePaintV1(request, provider)
  expect(result.ok, JSON.stringify(result)).toBe(true)
  if (!result.ok || result.value.status !== 'painted') throw new Error(JSON.stringify(result))
  return { value: result.value, provider }
}

describe('native DOCX page-paint v1', () => {
  it('keeps explicit legacy approximation separate from strict output and preserves original reasons', async () => {
    const request = fixture()
    const settings = request.pagination_request.pagination_settings
    settings.profile = 'unsupported'
    delete settings.compatibility_mode
    settings.diagnostics = [{ code: 'COMPATIBILITY_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: SETTINGS_PART, path: '/w:settings[1]/w:compat[1]', preservation: 'preserve-verbatim', message: 'Exact legacy mode 14 requires a different layout policy' }]
    const refused = paginateNativeDocxV1(request.pagination_request)
    expect(refused.ok).toBe(true)
    if (!refused.ok) return
    request.paginated_layout = refused.value
    request.integrity.paginated_layout_sha256 = nativeDocxPagePaintPaginatedLayoutSha256V1(refused.value)
    const original = structuredClone(request)
    const eligibility = { protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: settings.document_id, revision: settings.revision, package_sha256: settings.package_sha256, settings_sha256: settings.settings_sha256, status: 'eligible', legacy_compatibility_mode: 14, reasons: ['Legacy mode 14 uses current layout only in explicit approximate preview'] }
    const strict = await compileNativeDocxPagePaintV1(request, new FixtureProvider())
    expect(strict).toMatchObject({ ok: true, value: { status: 'refused', pages: [] } })
    const approximate = await compileNativeDocxApproximatePagePreviewV1(request, eligibility, new FixtureProvider())
    expect(approximate).toMatchObject({ protocol: 'injoffice.docx.approximate-page-preview', fidelity: 'approximate', read_only: true, status: 'painted' })
    expect(approximate.pages).toHaveLength(1)
    expect(approximate.source_settings_diagnostics).toEqual(settings.diagnostics)
    expect(approximate.rendering_provenance.pagination_settings).toEqual(settings)
    expect(request).toEqual(original)
    expect(decodeNativeDocxPagePaintV1(approximate).ok).toBe(false)
    expect(decodeNativeDocxApproximatePagePreviewV1(approximate).ok).toBe(true)
    expect(decodeNativeDocxApproximatePagePreviewV1({ ...approximate, source_settings_diagnostics: [] }).ok).toBe(false)
    expect(decodeNativeDocxApproximatePagePreviewV1({ ...approximate, reasons: [] }).ok).toBe(false)
    expect(decodeNativeDocxApproximatePagePreviewV1({ ...approximate, reasons: [''] }).ok).toBe(false)
    expect(decodeNativeDocxApproximatePagePreviewV1({ ...approximate, reasons: ['Everything is exact'] }).ok).toBe(false)
    expect(decodeNativeDocxApproximatePagePreviewV1({ ...approximate, fidelity: 'exact' }).ok).toBe(false)
    expect(decodeNativeDocxApproximatePagePreviewV1({ ...approximate, rendering_provenance: { ...approximate.rendering_provenance, package_sha256: `sha256:${'f'.repeat(64)}` } }).ok).toBe(false)
    await expect(compileNativeDocxApproximatePagePreviewV1(request, { ...eligibility, package_sha256: 'wrong' }, new FixtureProvider())).rejects.toThrow('exact-join')
    const ineligible = await compileNativeDocxApproximatePagePreviewV1(request, { ...eligibility, status: 'ineligible', legacy_compatibility_mode: null }, new FixtureProvider())
    expect(ineligible).toMatchObject({ fidelity: 'approximate', status: 'refused', pages: [] })
  })
  it('retains known approximate settings values and requires matching facts and warnings', async () => {
    const request = fixture()
    const settings = request.pagination_request.pagination_settings
    settings.profile = 'unsupported'
    delete settings.compatibility_mode
    const fact = { kind: 'decimalSymbol', path: '/w:settings[1]/w:decimalSymbol[1]', values: { val: ',' } }
    const warning = `Current-layout approximation disregards ${fact.kind} at ${fact.path}; source values are retained and Word layout may differ`
    settings.diagnostics = [{ code: 'PAGINATION_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: SETTINGS_PART, path: fact.path, preservation: 'preserve-verbatim', message: 'Original decimal setting is not strictly qualified' }]
    const refused = paginateNativeDocxV1(request.pagination_request)
    if (!refused.ok) throw new Error('invalid fixture')
    request.paginated_layout = refused.value
    request.integrity.paginated_layout_sha256 = nativeDocxPagePaintPaginatedLayoutSha256V1(refused.value)
    const original = structuredClone(request)
    const eligibility = { protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: settings.document_id, revision: settings.revision, package_sha256: settings.package_sha256, settings_sha256: settings.settings_sha256, status: 'eligible', legacy_compatibility_mode: 12, reasons: [warning], approximated_settings: [fact] }
    const approximate = await compileNativeDocxApproximatePagePreviewV1(request, eligibility, new FixtureProvider())
    expect(approximate).toMatchObject({ status: 'painted', approximated_settings: [fact] })
    expect(decodeNativeDocxApproximatePagePreviewV1(approximate).ok).toBe(true)
    expect(decodeNativeDocxApproximatePagePreviewV1({ ...approximate, approximated_settings: [] }).ok).toBe(false)
    expect(decodeNativeDocxApproximatePagePreviewV1({ ...approximate, reasons: approximate.reasons.filter(reason => reason !== warning) }).ok).toBe(false)
    await expect(compileNativeDocxApproximatePagePreviewV1(request, { ...eligibility, approximated_settings: [] }, new FixtureProvider())).rejects.toThrow('conflicts')
    await expect(compileNativeDocxApproximatePagePreviewV1(request, { ...eligibility, reasons: ['missing setting warning'] }, new FixtureProvider())).rejects.toThrow('conflicts')
    expect(request).toEqual(original)
    expect(await compileNativeDocxPagePaintV1(request, new FixtureProvider())).toMatchObject({ ok: true, value: { status: 'refused' } })
    const flag = { kind: 'enableOpenTypeFeatures', path: '/w:settings[1]/w:compat[1]/w:compatSetting[2]', values: { val: '1' } }
    const flagSettings = structuredClone(settings)
    flagSettings.diagnostics[0] = { ...flagSettings.diagnostics[0]!, code: 'COMPATIBILITY_SETTING_UNSUPPORTED', path: flag.path }
    const flagEligibility = { ...eligibility, approximated_settings: [flag], reasons: [`Current-layout approximation disregards ${flag.kind} at ${flag.path}; source values are retained and Word layout may differ`] }
    expect(decodeNativeDocxApproximationEligibilityV1(flagEligibility, flagSettings).status).toBe('eligible')
    expect(() => decodeNativeDocxApproximationEligibilityV1({ ...flagEligibility, approximated_settings: [], reasons: ['Legacy mode'] }, flagSettings)).toThrow('conflicts')
  })
  it('bounds upstream refusal reasons and keeps valid atomic output', async () => {
    const request = fixture()
    const settings = request.pagination_request.pagination_settings
    settings.profile = 'unsupported'
    settings.diagnostics = Array.from({ length: 12 }, (_, index) => ({
      code: 'PAGINATION_SETTING_UNSUPPORTED' as const, severity: 'unsupported' as const, preservation: 'preserve-verbatim' as const,
      part_name: settings.settings_part!, path: '/w:settings[1]', message: `Unsupported setting ${index}`,
    }))
    const layout = paginateNativeDocxV1(request.pagination_request)
    expect(layout.ok, JSON.stringify(layout)).toBe(true)
    if (!layout.ok) return
    request.paginated_layout = layout.value
    request.integrity.paginated_layout_sha256 = nativeDocxPagePaintPaginatedLayoutSha256V1(layout.value)
    const provider = new FixtureProvider()
    const result = await compileNativeDocxPagePaintV1(request, provider)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.value.status).toBe('refused')
    expect(result.value.pages).toEqual([])
    expect(result.value.resources).toEqual([])
    expect(result.value.diagnostics).toHaveLength(10)
    expect(result.value.diagnostics[1]!.message).toContain('PAGINATION_SETTING_UNSUPPORTED')
    expect(result.value.diagnostics.at(-1)!.message).toContain('4 additional pagination reasons omitted')
    expect(result.value.diagnostics.every((entry) => entry.message.length <= 4096)).toBe(true)
    expect(decodeNativeDocxPagePaintV1(result.value).ok).toBe(true)
    expect(provider.calls).toBe(0)
  })

  it('paints with qualified latent metadata without removing preservation evidence', async () => {
    const request = fixture()
    const resolved = request.pagination_request.resolved_layout
    resolved.source_parts.styles_part = 'word/styles.xml'
    resolved.diagnostics.push({ code: 'LATENT_STYLE_BEHAVIOR_PRESERVED', severity: 'unsupported', scope_id: resolved.document_id, part_name: 'word/styles.xml', path: '/w:styles[1]/w:latentStyles[1]', preservation: 'preserve-verbatim', message: 'Exact UI metadata retained' })
    const before = JSON.stringify(request.pagination_request.document)
    const result = await compileNativeDocxPagePaintV1(request, new FixtureProvider())
    expect(result.ok && result.value.status).toBe('painted')
    expect(resolved.diagnostics).toHaveLength(1)
    expect(JSON.stringify(request.pagination_request.document)).toBe(before)
    resolved.diagnostics[0]!.code = 'LATENT_STYLES_PRESERVED'
    const unknown = await compileNativeDocxPagePaintV1(request, new FixtureProvider())
    expect(unknown.ok && unknown.value.status).toBe('refused')
  })
  it('retains qualified font matching diagnostics while painting supplied glyphs', async () => {
    const request = fixture()
    const resolved = request.pagination_request.resolved_layout
    resolved.source_parts.font_table_part = 'word/fonts.xml'
    resolved.diagnostics.push({ code: 'FONT_MATCHING_METADATA_PRESERVED', severity: 'unsupported', scope_id: resolved.document_id, part_name: 'word/fonts.xml', path: '/w:fonts[1]/w:font[1]/w:panose1[1]', preservation: 'preserve-verbatim', message: 'Matching metadata retained' })
    const before = JSON.stringify(request.pagination_request.document)
    const result = await compileNativeDocxPagePaintV1(request, new FixtureProvider())
    expect(result.ok && result.value.status).toBe('painted')
    expect(resolved.diagnostics).toHaveLength(1)
    expect(JSON.stringify(request.pagination_request.document)).toBe(before)
    resolved.diagnostics[0]!.code = 'UNMODELED_FONT_METADATA'
    const unknown = await compileNativeDocxPagePaintV1(request, new FixtureProvider())
    expect(unknown.ok && unknown.value.status).toBe('refused')
  })

  it('retains empty default numbering style evidence without blocking paint', async () => {
    const request = fixture(), resolved = request.pagination_request.resolved_layout
    resolved.source_parts.styles_part = 'word/styles.xml'
    resolved.diagnostics.push({ code: 'EMPTY_NUMBERING_STYLE_PRESERVED', severity: 'unsupported', scope_id: resolved.document_id, part_name: 'word/styles.xml', path: '/w:styles[1]/w:style[4]', preservation: 'preserve-verbatim', message: 'Exact empty default numbering style retained' })
    const result = await compileNativeDocxPagePaintV1(request, new FixtureProvider())
    expect(result.ok && result.value.status).toBe('painted')
    expect(resolved.diagnostics).toHaveLength(1)
    resolved.diagnostics[0]!.code = 'NUMBERING_STYLE_PRESERVED'
    const unknown = await compileNativeDocxPagePaintV1(request, new FixtureProvider())
    expect(unknown.ok && unknown.value.status).toBe('refused')
  })

  it('bounds paginated-layout hashing and keeps object-key order irrelevant', () => {
    const layout = fixture().paginated_layout
    const reordered = Object.fromEntries(Object.entries(structuredClone(layout)).reverse()) as typeof layout
    expect(nativeDocxPagePaintPaginatedLayoutSha256V1(reordered)).toBe(nativeDocxPagePaintPaginatedLayoutSha256V1(layout))
    const cyclic: any = structuredClone(layout)
    cyclic.loop = cyclic
    expect(() => nativeDocxPagePaintPaginatedLayoutSha256V1(cyclic)).toThrow(/acyclic/)
    const deep: any = structuredClone(layout)
    let cursor = deep
    for (let index = 0; index <= DOCX_NATIVE_LIMITS.maxDepth; index += 1) cursor = cursor.extra = {}
    expect(() => nativeDocxPagePaintPaginatedLayoutSha256V1(deep)).toThrow(/bounded traversal/)
  })

  it('keeps strict binding-field manifests in parity with request and output objects', async () => {
    const request = fixture()
    const { value } = await painted(request)
    const page = value.pages[0]!
    const line = page.lines[0]!
    const command = page.commands[0]!
    if (command.kind !== 'fill_glyph_path') throw new Error('expected glyph command')
    const fields = (value: object) => Object.keys(value).sort()
    expect(fields(request)).toEqual([...DOCX_PAGE_PAINT_REQUEST_V1_BINDING_FIELDS.RequestV1].sort())
    expect(fields(request.integrity)).toEqual([...DOCX_PAGE_PAINT_REQUEST_V1_BINDING_FIELDS.IntegrityV1].sort())
    expect(fields(request.outline_provider)).toEqual([...DOCX_PAGE_PAINT_REQUEST_V1_BINDING_FIELDS.OutlineProviderV1].sort())
    expect(fields(value)).toEqual([...DOCX_PAGE_PAINT_V1_BINDING_FIELDS.OutputV1].sort())
    expect(fields(value.provenance)).toEqual(DOCX_PAGE_PAINT_V1_BINDING_FIELDS.ProvenanceV1.filter((key) => key !== 'numbering_source' && key !== 'body_field_source_sha256').sort())
    expect(fields(value.provenance.font_manifest)).toEqual([...DOCX_PAGE_PAINT_V1_BINDING_FIELDS.ManifestV1].sort())
    expect(fields(value.provenance.media_assets)).toEqual([...DOCX_PAGE_PAINT_V1_BINDING_FIELDS.MediaSourceV1].sort())
    expect(fields(value.provenance.providers)).toEqual([...DOCX_PAGE_PAINT_V1_BINDING_FIELDS.ProvidersV1].sort())
    expect(fields(page)).toEqual([...DOCX_PAGE_PAINT_V1_BINDING_FIELDS.PageV1].sort())
    expect(fields(line)).toEqual([...DOCX_PAGE_PAINT_V1_BINDING_FIELDS.BodyLineV1].sort())
    expect(fields(command)).toEqual([...DOCX_PAGE_PAINT_V1_BINDING_FIELDS.GlyphCommandV1].sort())
    expect(fields(command.face)).toEqual(['content_digest', 'face_id'])
    expect(fields(command.path[0]!)).toEqual([...DOCX_PAGE_PAINT_V1_BINDING_FIELDS.PathMoveV1].sort())
    expect(fields(command.path.at(-1)!)).toEqual([...DOCX_PAGE_PAINT_V1_BINDING_FIELDS.PathCloseV1].sort())
  })

  it('emits non-empty absolute integer glyph paths, explicit page geometry, and caches exact face/glyph keys', async () => {
    const request = fixture()
    const { value, provider } = await painted(request)
    expect(value.protocol).toBe(DOCX_PAGE_PAINT_PROTOCOL)
    expect(value.pages).toHaveLength(1)
    expect(value.pages[0]).toMatchObject({ id: 'page:section:1:0', ordinal: 0, width_millipoints: 50_000, height_millipoints: 50_000, body_box: { x_millipoints: 5_000, y_millipoints: 5_000, width_millipoints: 40_000, height_millipoints: 40_000 } })
    expect(value.pages[0]!.commands).toHaveLength(2)
    const command = value.pages[0]!.commands[0]!
    if (command.kind !== 'fill_glyph_path') throw new Error('expected glyph command')
    expect(command).toMatchObject({ glyph_id: 7, font_size_millipoints: 10_000, fill_rgb: '123456', outline_kind: 'path' })
    expect(command.path[0]).toEqual({ kind: 'move_to', x_millipoints: 5_000, y_millipoints: 13_000 })
    expect(command.path.some((entry) => entry.kind === 'line_to')).toBe(true)
    expect(provider.calls).toBe(1)
    expect(decodeNativeDocxPagePaintV1(value).ok).toBe(true)
    expect(decodeNativeDocxPagePaintForRequestV1(value, request, { provider_id: provider.providerId, provider_revision: provider.providerRevision }).ok).toBe(true)
    expect((await validateNativeDocxPagePaintForRequestV1(value, request, new FixtureProvider())).ok).toBe(true)
  })

  it('paints exact native list-marker glyphs with marker font/color and carries numbering hashes through provenance', async () => {
    const request = fixture()
    const pagination = request.pagination_request
    const nativeParagraph = pagination.document.body.blocks[0]!.paragraph!
    nativeParagraph.properties.numbering = { num_id: '7', abstract_num_id: '3', level: 0 }
    pagination.document.passthrough_parts.push({ part_name: NUMBERING_PART, content_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml', byte_length: 1, sha256: NUMBERING_HASH, policy: 'preserve-verbatim' })
    const resolvedParagraph = pagination.resolved_layout.paragraphs[0]!
    resolvedParagraph.properties = { indent_start_twips: 200, hanging_twips: 200 }
    resolvedParagraph.numbering = {
      marker_id: 'marker:paragraph:1', definition_sha256: HASH, num_id: '7', abstract_num_id: '3', level: 0, start: 1,
      format: 'decimal', text: '%1', suffix: 'nothing', alignment: 'start', never_restart: true,
      counter_value: 1, counter_values: [{ level: 0, value: 1, format: 'decimal' }], resolved_text: '1',
      label_start_twips: 0, label_end_twips: 200, text_start_twips: 200,
      marker_properties: { font_family: 'Test', font_size_half_points: 20, color: '654321', language: 'en-US' },
    }
    pagination.resolved_layout.source_parts.numbering_part = NUMBERING_PART
    const numberingSourceBase = {
      relationships_part: RELATIONSHIPS_PART, relationships_sha256: RELATIONSHIPS_HASH,
      relationship_id: 'rIdNumbering', relationship_type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering', relationship_target: 'numbering.xml',
      part_name: NUMBERING_PART, content_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml' as const, part_sha256: NUMBERING_HASH,
    }
    resolvedParagraph.numbering.definition_sha256 = nativeDocxResolvedNumberingDefinitionSha256V1(resolvedParagraph.numbering, NUMBERING_HASH)
    const numberingSource = { ...numberingSourceBase, model_sha256: nativeDocxResolvedNumberingModelSha256V1(pagination.resolved_layout.paragraphs, numberingSourceBase) }
    pagination.resolved_layout.numbering_source = numberingSource
    const shapedParagraph = pagination.shaped_lines.paragraphs[0]!
    shapedParagraph.indent_start_millipoints = 10_000
    shapedParagraph.first_line_delta_millipoints = -10_000
    shapedParagraph.list_marker = {
      marker_id: resolvedParagraph.numbering.marker_id, definition_sha256: resolvedParagraph.numbering.definition_sha256,
      numbering_part_sha256: NUMBERING_HASH, model_sha256: numberingSource.model_sha256,
      num_id: '7', abstract_num_id: '3', level: 0, counter_value: 1, text: '1', suffix: 'nothing', alignment: 'start',
      label_start_millipoints: 0, label_end_millipoints: 10_000, marker_start_millipoints: 0, marker_advance_millipoints: 5_000, text_start_millipoints: 5_000,
    }
    pagination.shaped_lines.numbering_source = numberingSource
    const line = shapedParagraph.lines[0]!
    line.advance_inline_millipoints = 15_000
    line.fragments.forEach((fragment) => { fragment.logical_order += 1 })
    line.logical_to_visual = [0, 1]
    line.fragments[0]!.id = 'fragment:paragraph:1:0:1'
    line.fragments.unshift({
      id: 'fragment:paragraph:1:0:0', source_kind: 'list-marker', source_id: 'paragraph:1', start_utf16: 0, end_utf16: 1, text: '1',
      direction: 'ltr', bidi_level: 0, logical_order: 0, script: 'Zyyy', language: 'en-US', face_id: 'face:test', whitespace: false,
      advance_inline_millipoints: 5_000, ascent_millipoints: 8_000, descent_millipoints: -2_000, line_gap_millipoints: 0,
      justification_expansion_millipoints: 0,
      glyphs: [{ glyph_id: 8, advance_x_millipoints: 5_000, advance_y_millipoints: 0, offset_x_millipoints: 0, offset_y_millipoints: 0 }],
    })
    const layout = paginateNativeDocxV1(pagination)
    expect(layout.ok && layout.value.status, JSON.stringify(layout)).toBe('paginated')
    if (!layout.ok || layout.value.status !== 'paginated') return

    const forgedGeometry = structuredClone(pagination)
    const forgedParagraph = forgedGeometry.shaped_lines.paragraphs[0]!
    const forgedMarker = forgedParagraph.list_marker!
    const forgedLine = forgedParagraph.lines[0]!
    forgedMarker.marker_start_millipoints = 1_000
    forgedMarker.text_start_millipoints = 6_000
    forgedLine.advance_inline_millipoints += 1_000
    forgedLine.fragments.forEach((fragment) => { fragment.logical_order += 1 })
    forgedLine.logical_to_visual = [0, 1, 2]
    forgedLine.fragments[0]!.id = 'fragment:paragraph:1:0:1'
    forgedLine.fragments[1]!.id = 'fragment:paragraph:1:0:2'
    forgedLine.fragments.unshift({
      id: 'fragment:paragraph:1:0:0', source_kind: 'list-marker', source_id: 'paragraph:1', start_utf16: 0, end_utf16: 0, text: '',
      direction: 'ltr', bidi_level: 0, logical_order: 0, script: 'Zyyy', language: 'en-US', whitespace: false,
      advance_inline_millipoints: 1_000, ascent_millipoints: 0, descent_millipoints: 0, line_gap_millipoints: 0, glyphs: [],
      justification_expansion_millipoints: 0,
    })
    expect(paginateNativeDocxV1(forgedGeometry).ok).toBe(false)

    request.paginated_layout = layout.value
    request.integrity.shaped_lines_sha256 = nativeDocxPagePaintShapedLinesSha256V1(pagination.shaped_lines)
    request.integrity.paginated_layout_sha256 = nativeDocxPagePaintPaginatedLayoutSha256V1(layout.value)
    const { value } = await painted(request)
    expect(value.pages[0]!.commands.flatMap((command) => command.kind === 'fill_glyph_path' ? [[command.source_id, command.fill_rgb]] : [])).toEqual([
      ['paragraph:1', '654321'], ['run:1', '123456'], ['run:1', '123456'],
    ])
    expect(value.provenance.numbering_source).toEqual(numberingSource)
    expect(decodeNativeDocxPagePaintForRequestV1(value, request, { provider_id: 'outline:test', provider_revision: '1' }).ok).toBe(true)
  })

  it('is deterministic and does not mutate caller-owned request data', async () => {
    const request = fixture()
    const before = structuredClone(request)
    const first = await painted(request, new FixtureProvider())
    const second = await painted(request, new FixtureProvider())
    expect(second.value).toEqual(first.value)
    expect(JSON.stringify(second.value)).toBe(JSON.stringify(first.value))
    expect(request).toEqual(before)
    const layoutHash = nativeDocxPagePaintPaginatedLayoutSha256V1(request.paginated_layout)
    expect(nativeDocxPagePaintPaginatedLayoutSha256V1(structuredClone(request.paginated_layout))).toBe(layoutHash)
    const changedColumn = structuredClone(request.paginated_layout)
    changedColumn.pages[0]!.columns[0]!.id = 'column:section:1:changed'
    expect(nativeDocxPagePaintPaginatedLayoutSha256V1(changedColumn)).not.toBe(layoutHash)
  })

  it('strictly decodes request/output exact keys and request-bound glyph identity', async () => {
    const request = fixture()
    expect(Object.keys(request).sort()).toEqual([...DOCX_PAGE_PAINT_REQUEST_V1_BINDING_FIELDS.RequestV1].sort())
    expect(decodeNativeDocxPagePaintRequestV1(request).ok).toBe(true)
    for (const mutate of [
      (value: any) => { value.extra = true },
      (value: any) => { value.version = 2 },
      (value: any) => { value.outline_provider.provider_id = null },
      (value: any) => { value.pagination_request.shaped_lines.paragraphs[0].lines[0].ordinal = -0 },
      (value: any) => { value.paginated_layout.pages[0].lines[0].x_millipoints += 1 },
      (value: any) => { value.font_manifest.revision = 'wrong' },
      (value: any) => { value.font_manifest.faces[0].source.contentDigest = `sha256:${'d'.repeat(64)}` },
    ]) {
      const invalid: any = structuredClone(request)
      mutate(invalid)
      expect(decodeNativeDocxPagePaintRequestV1(invalid).ok).toBe(false)
    }

    const { value } = await painted(request)
    expect(Object.keys(value).sort()).toEqual([...DOCX_PAGE_PAINT_V1_BINDING_FIELDS.OutputV1].sort())
    const tampered: any = structuredClone(value)
    tampered.pages[0].commands[0].glyph_id = 8
    expect(decodeNativeDocxPagePaintV1(tampered).ok).toBe(true)
    expect(decodeNativeDocxPagePaintForRequestV1(tampered, request, { provider_id: 'outline:test', provider_revision: '1' }).ok).toBe(false)

    for (const mutate of [
      (output: any) => { output.pages[0].commands[0].face.content_digest = `sha256:${'c'.repeat(64)}` },
      (output: any) => { output.pages[0].commands[0].face.collection_index = 1 },
      (output: any) => { output.pages[0].commands[0].font_size_millipoints += 1 },
      (output: any) => { output.pages[0].commands[0].fill_rgb = '654321' },
      (output: any) => { output.pages[0].lines[0].command_ids = [] },
    ]) {
      const invalid: any = structuredClone(value)
      mutate(invalid)
      expect(decodeNativeDocxPagePaintForRequestV1(invalid, request, { provider_id: 'outline:test', provider_revision: '1' }).ok).toBe(false)
    }

    for (const mutate of [
      (output: any) => { output.extra = true },
      (output: any) => { output.provenance.shaped_lines.protocol = 'forged' },
      (output: any) => { output.provenance.paginated_layout.protocol = 'forged' },
      (output: any) => { output.provenance.pagination_settings.extra = true },
      (output: any) => { output.pages[0].section_ids.push(output.pages[0].section_id) },
      (output: any) => { output.pages[0].columns[0].section_id = 'section:forged' },
      (output: any) => { output.pages[0].lines[0].column_ordinal = 1 },
      (output: any) => { output.pages[0].columns[0].x_millipoints = output.pages[0].width_millipoints },
      (output: any) => { output.pages[0].lines[0].x_millipoints = output.pages[0].columns[0].x_millipoints + output.pages[0].columns[0].width_millipoints },
      (output: any) => { output.pages[0].commands[0].fill_rule = 'evenodd' },
      (output: any) => { output.pages[0].commands[0].path[0].x_millipoints = -0 },
      (output: any) => { output.pages[0].commands[0].path.pop() },
    ]) {
      const invalid: any = structuredClone(value)
      mutate(invalid)
      expect(decodeNativeDocxPagePaintV1(invalid).ok).toBe(false)
    }

    const cyclicRequest: any = fixture()
    cyclicRequest.loop = cyclicRequest
    expect(decodeNativeDocxPagePaintRequestV1(cyclicRequest).ok).toBe(false)
    const cyclicOutput: any = structuredClone(value)
    cyclicOutput.pages[0].commands[0].path.push(cyclicOutput)
    expect(decodeNativeDocxPagePaintV1(cyclicOutput).ok).toBe(false)
  })

  it('refuses a structurally valid shaped projection that omits native source clusters', async () => {
    const request = fixture()
    request.pagination_request.document.body.blocks[0]!.paragraph!.runs[0]!.text = 'AAA'
    const provider = new FixtureProvider()
    const result = await compileNativeDocxPagePaintV1(request, provider)
    expect(result.ok).toBe(true)
    if (result.ok) expect(result.value).toMatchObject({ status: 'refused', diagnostics: [expect.objectContaining({ code: 'identity-mismatch', scope_id: 'run:1' })], pages: [] })
  })

  it('atomically refuses missing faces, provider mismatch/refusal/failure, and invalid paths', async () => {
    const cases: Array<{ request?: NativeDocxPagePaintRequestV1; provider: FixtureProvider; code: string }> = []
    const noDigest = fixture()
    delete (noDigest.font_manifest.faces[0]!.source as { contentDigest?: string }).contentDigest
    noDigest.integrity.font_manifest_sha256 = nativeDocxPagePaintFontManifestSha256V1(noDigest.font_manifest)
    cases.push({ request: noDigest, provider: new FixtureProvider(), code: 'missing-font' })

    const wrongIdentity = new FixtureProvider()
    wrongIdentity.providerRevision = '2'
    cases.push({ provider: wrongIdentity, code: 'provider-mismatch' })

    const missing = new FixtureProvider()
    missing.result = (request) => ({ status: 'refused', face: { ...request.face }, glyph_id: request.glyph_id, code: 'missing-glyph', message: 'glyph absent' })
    cases.push({ provider: missing, code: 'missing-glyph' })

    const invalid = new FixtureProvider()
    invalid.result = (request) => ({ status: 'outlined', face: { ...request.face }, glyph_id: request.glyph_id, units_per_em: 1_000, path: [{ kind: 'move_to', x: 0, y: 0 }, { kind: 'line_to', x: 1, y: 1 }] })
    cases.push({ provider: invalid, code: 'invalid-provider-output' })

    const collapsed = new FixtureProvider()
    collapsed.result = (request) => ({
      status: 'outlined', face: { ...request.face }, glyph_id: request.glyph_id, units_per_em: 1_000_000,
      path: [{ kind: 'move_to', x: 0, y: 0 }, { kind: 'line_to', x: 1, y: 0 }, { kind: 'line_to', x: 0, y: 1 }, { kind: 'close_path' }],
    })
    cases.push({ provider: collapsed, code: 'invalid-path' })

    const emptyVisible = new FixtureProvider()
    emptyVisible.result = (request) => ({ status: 'empty', face: { ...request.face }, glyph_id: request.glyph_id, units_per_em: 1_000 })
    cases.push({ provider: emptyVisible, code: 'missing-glyph' })

    const degenerate = new FixtureProvider()
    degenerate.result = (request) => ({ status: 'outlined', face: { ...request.face }, glyph_id: request.glyph_id, units_per_em: 1_000, path: [{ kind: 'move_to', x: 1, y: 1 }, { kind: 'line_to', x: 1, y: 1 }, { kind: 'close_path' }] })
    cases.push({ provider: degenerate, code: 'invalid-provider-output' })

    for (const entry of cases) {
      const result = await compileNativeDocxPagePaintV1(entry.request ?? fixture(), entry.provider)
      expect(result.ok).toBe(true)
      if (!result.ok) continue
      expect(result.value).toMatchObject({ status: 'refused', pages: [], diagnostics: [{ code: entry.code }] })
      expect(decodeNativeDocxPagePaintV1(result.value).ok).toBe(true)
    }

    const throwing = new FixtureProvider()
    throwing.getGlyphOutline = () => { throw new Error('boom') }
    const failure = await compileNativeDocxPagePaintV1(fixture(), throwing)
    expect(failure.ok && failure.value.status === 'refused' && failure.value.pages).toEqual([])
    if (failure.ok) expect(failure.value.diagnostics[0]?.code).toBe('provider-failure')
  })

  it('owns provider output and detects provider identity mutation during an async call', async () => {
    const provider = new FixtureProvider()
    const resultObject = provider.getGlyphOutline({ face: { face_id: 'face:test', content_digest: HASH }, glyph_id: 7 })
    provider.getGlyphOutline = (async (request) => {
      provider.providerRevision = 'mutated'
      return { ...(resultObject as any), face: { ...request.face }, glyph_id: request.glyph_id }
    }) as NativeDocxGlyphOutlineProviderV1['getGlyphOutline']
    const result = await compileNativeDocxPagePaintV1(fixture(), provider)
    expect(result.ok && result.value.status === 'refused' ? result.value.diagnostics[0]?.code : undefined).toBe('provider-mismatch')
    expect(result.ok && result.value.pages).toEqual([])
  })

  it('bounds hostile provider values and resource-heavy paths without leaking partial pages', async () => {
    const huge = new FixtureProvider()
    huge.result = (request) => ({
      status: 'outlined', face: { ...request.face }, glyph_id: request.glyph_id, units_per_em: 1_000,
      path: Array.from({ length: 65_537 }, () => ({ kind: 'close_path' as const })),
    })
    const overflow = await compileNativeDocxPagePaintV1(fixture(), huge)
    expect(overflow.ok && overflow.value.status === 'refused' ? overflow.value.diagnostics[0]?.code : undefined).toBe('resource-limit')
    expect(overflow.ok && overflow.value.pages).toEqual([])

    const unreadable = new FixtureProvider()
    unreadable.getGlyphOutline = (() => {
      const result: Record<string, unknown> = {}
      Object.defineProperty(result, 'status', { enumerable: true, get() { throw new Error('hostile getter') } })
      return result as unknown as NativeDocxGlyphOutlineResultV1
    }) as NativeDocxGlyphOutlineProviderV1['getGlyphOutline']
    const hostile = await compileNativeDocxPagePaintV1(fixture(), unreadable)
    expect(hostile.ok && hostile.value.status === 'refused' ? hostile.value.diagnostics[0]?.code : undefined).toBe('invalid-provider-output')
    expect(hostile.ok && hostile.value.pages).toEqual([])

    const errorGetter = new FixtureProvider()
    errorGetter.getGlyphOutline = (() => {
      const thrown: Record<string, unknown> = {}
      Object.defineProperty(thrown, 'message', { get() { throw new Error('nested') } })
      throw thrown
    }) as NativeDocxGlyphOutlineProviderV1['getGlyphOutline']
    const failure = await compileNativeDocxPagePaintV1(fixture(), errorGetter)
    expect(failure.ok && failure.value.status === 'refused' ? failure.value.diagnostics[0]?.message : '').toContain('unreadable provider error')
    expect(failure.ok ? decodeNativeDocxPagePaintV1(failure.value).ok : false).toBe(true)

    const lateRequest = fixture()
    lateRequest.pagination_request.shaped_lines.paragraphs[0]!.lines[0]!.fragments[0]!.glyphs[1]!.glyph_id = 8
    const lateLayout = paginateNativeDocxV1(lateRequest.pagination_request)
    if (!lateLayout.ok || lateLayout.value.status !== 'paginated') throw new Error(JSON.stringify(lateLayout))
    lateRequest.paginated_layout = lateLayout.value
    lateRequest.integrity.shaped_lines_sha256 = nativeDocxPagePaintShapedLinesSha256V1(lateRequest.pagination_request.shaped_lines)
    const late = new FixtureProvider()
    late.result = (request) => request.glyph_id === 8
      ? { status: 'refused', face: { ...request.face }, glyph_id: request.glyph_id, code: 'missing-glyph', message: 'second glyph missing' }
      : new FixtureProvider().getGlyphOutline(request) as NativeDocxGlyphOutlineResultV1
    const atomic = await compileNativeDocxPagePaintV1(lateRequest, late)
    expect(late.calls).toBe(2)
    expect(atomic.ok && atomic.value.status === 'refused' ? atomic.value.diagnostics[0]?.code : undefined).toBe('missing-glyph')
    expect(atomic.ok && atomic.value.pages).toEqual([])
  })

  it('refuses unsupported paint semantics before calling the outline provider', async () => {
    const request = fixture()
    request.pagination_request.resolved_layout.runs[0]!.properties.underline = 'single'
    const provider = new FixtureProvider()
    const result = await compileNativeDocxPagePaintV1(request, provider)
    expect(result.ok && result.value.status === 'refused' ? result.value.diagnostics[0]?.code : undefined).toBe('unsupported-source')
    expect(provider.calls).toBe(0)

    const tabRequest = fixture()
    tabRequest.pagination_request.resolved_layout.runs[0]!.properties.underline = 'single'
    const nativeRun: any = tabRequest.pagination_request.document.body.blocks[0]!.paragraph!.runs[0]!
    nativeRun.kind = 'control'; nativeRun.control = 'tab'; delete nativeRun.text
    const fragment: any = tabRequest.pagination_request.shaped_lines.paragraphs[0]!.lines[0]!.fragments[0]!
    fragment.source_kind = 'tab'; fragment.text = '\t'; fragment.start_utf16 = 0; fragment.end_utf16 = 0; fragment.whitespace = true; fragment.glyphs = []
    const tabLayout = paginateNativeDocxV1(tabRequest.pagination_request)
    if (!tabLayout.ok || tabLayout.value.status !== 'paginated') throw new Error(JSON.stringify(tabLayout))
    tabRequest.paginated_layout = tabLayout.value
    tabRequest.integrity.shaped_lines_sha256 = nativeDocxPagePaintShapedLinesSha256V1(tabRequest.pagination_request.shaped_lines)
    const tabProvider = new FixtureProvider()
    const tabResult = await compileNativeDocxPagePaintV1(tabRequest, tabProvider)
    expect(tabResult.ok, JSON.stringify(tabResult)).toBe(true)
    expect(tabResult.ok && tabResult.value.status === 'refused' ? tabResult.value.diagnostics[0]?.code : undefined).toBe('unsupported-source')
    expect(tabProvider.calls).toBe(0)

    const drawingRequest = fixture()
    const drawingRun: any = drawingRequest.pagination_request.document.body.blocks[0]!.paragraph!.runs[0]!
    delete drawingRun.text
    drawingRun.kind = 'drawing'
    drawingRun.drawing = {
      id: 'drawing:1', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:drawing[1]', 120, 170),
      placement: 'floating', width_emu: 914400, height_emu: 914400, x_emu: 0, y_emu: 0,
      horizontal_relative_from: 'page', vertical_relative_from: 'page', wrap: 'square',
      edit_policy: { mode: 'read-only', allowed_operations: [], refusal: { code: 'EXTRACT_ONLY', message: 'Drawing placement unavailable.', preservation: 'refuse-mutation' } },
    }
    const refusedLayout = paginateNativeDocxV1(drawingRequest.pagination_request)
    if (!refusedLayout.ok || refusedLayout.value.status !== 'refused') throw new Error(JSON.stringify(refusedLayout))
    drawingRequest.paginated_layout = refusedLayout.value
    drawingRequest.integrity.paginated_layout_sha256 = nativeDocxPagePaintPaginatedLayoutSha256V1(refusedLayout.value)
    const drawingProvider = new FixtureProvider()
    const drawingResult = await compileNativeDocxPagePaintV1(drawingRequest, drawingProvider)
    expect(drawingResult.ok && drawingResult.value.status === 'refused' ? drawingResult.value.diagnostics[0]?.code : undefined).toBe('upstream-refused')
    expect(drawingResult.ok && drawingResult.value.diagnostics.slice(1)).toEqual(expect.arrayContaining([expect.objectContaining({ message: expect.stringContaining(refusedLayout.value.diagnostics[0]!.code) })]))
    if (drawingResult.ok) expect(decodeNativeDocxPagePaintV1(drawingResult.value).ok).toBe(true)
    expect(drawingResult.ok && drawingResult.value.pages).toEqual([])
    expect(drawingProvider.calls).toBe(0)
  })

  it('contains no browser text/layout authority', () => {
    const source = ['./nativePagePaintV1.ts', './nativeTablePagePaintV1.ts', './nativePagePaintCompilerV1.ts'].map((file) => readFileSync(new URL(file, import.meta.url), 'utf8')).join('\n')
    for (const forbidden of ['CanvasRenderingContext2D', 'fill' + 'Text', 'measure' + 'Text', 'new Font' + 'Face', 'document.createElement', 'window.', 'getComputedStyle', 'system-font', 'pdfjs', 'pdf-lib', 'screen' + 'shot']) expect(source).not.toContain(forbidden)
  })
})
