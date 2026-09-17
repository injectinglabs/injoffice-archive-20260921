import { describe, expect, it } from 'vitest'
import { BIDI_UNICODE_VERSION, NATIVE_BIDI_PROVIDER_ID, NATIVE_BIDI_PROVIDER_REVISION } from '@injoffice/font-metrics/bidi'
import { UNICODE_13_CLASSIFIER_REVISION } from '@injoffice/font-metrics/unicode13'
import { DOCX_NATIVE_PROTOCOL, DOCX_NATIVE_VERSION, type NativeDocxDocumentV1, type NativeDocxParagraphV1 } from './nativeContract.js'
import { DOCX_RESOLVED_LAYOUT_PROTOCOL, DOCX_RESOLVED_LAYOUT_VERSION, type NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'
import { DOCX_SHAPED_LINES_PROTOCOL, DOCX_SHAPED_LINES_VERSION, type NativeDocxShapedLinesV1 } from './nativeShapingLines.js'
import { DOCX_DEFAULT_TAB_STOP_TWIPS, DOCX_PAGINATION_SETTINGS_PROTOCOL, DOCX_PAGINATION_SETTINGS_VERSION } from './nativePaginationSettings.js'
import {
  DOCX_PAGINATION_REQUEST_PROTOCOL, DOCX_PAGINATION_REQUEST_VERSION,
  paginateNativeDocxV1, paginateNativeDocxApproximateLegacyV1,
  type NativeDocxPaginationRequestV1,
} from './nativePaginationV1.js'
import type { NativeDocxApproximationEligibilityV1 } from './nativeApproximationV1.js'

/**
 * Word-seating corrections measured against a genuine Microsoft Word 16.112.4
 * render of NumberedList.docx at 96 DPI (office-hard-v1). Both corrections are
 * scoped to the explicitly approximate read-only preview lane, so every case
 * here also pins the strict pagination output byte-for-byte.
 */

const HASH = `sha256:${'a'.repeat(64)}`
const SETTINGS_PART = 'word/settings.xml'
const RELATIONSHIPS_PART = 'word/_rels/document.xml.rels'
const anchor = (path: string, start: number, end: number) => ({ part_name: 'word/document.xml', path, start_byte: start, end_byte: end, xml_sha256: HASH })
const policy = { mode: 'read-only' as const, allowed_operations: [] as const, refusal: { code: 'NATIVE_READ_ONLY', message: 'Source remains authoritative.', preservation: 'refuse-mutation' as const } }

/** Space-after of the paragraph that precedes the table, in milli-points. */
const BODY_SPACE_AFTER = 4_000
/** Space-before of the first paragraph inside the first table cell. */
const CELL_SPACE_BEFORE = 3_000
const LINE_HEIGHT = 6_000

function para(id: string, path: string, start: number, empty = false): NativeDocxParagraphV1 {
  return {
    id, anchor: anchor(path, start, start + 30), edit_policy: { ...policy, allowed_operations: [] }, properties: {},
    runs: empty ? [] : [{ kind: 'text', id: `run:${id}`, anchor: anchor(`${path}/w:r[1]/w:t[1]`, start + 5, start + 20), properties: {}, text: 'x' }],
  }
}

/** One body paragraph carrying space-after, then a single-cell table whose
 * first paragraph carries space-before — the NumberedList.docx shape. */
function fixture(): NativeDocxPaginationRequestV1 {
  // Word's own spacer paragraph before a table is an empty one (NumberedList's
  // "Table Spacing" style), which is exactly the shape the table qualifier admits.
  const lead = para('paragraph:lead', '/w:document[1]/w:body[1]/w:p[1]', 100, true)
  const cell = para('paragraph:cell', '/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]/w:p[1]', 300)
  const table = {
    id: 'table:1', anchor: anchor('/w:document[1]/w:body[1]/w:tbl[1]', 200, 400), edit_policy: { ...policy, allowed_operations: [] },
    width_twips: 400, layout: 'fixed' as const, alignment: 'left' as const, indent_twips: 0, grid_widths_twips: [400],
    cell_margins: { top_twips: 0, right_twips: 0, bottom_twips: 0, left_twips: 0 },
    borders: {},
    rows: [{
      id: 'row:1', anchor: anchor('/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]', 210, 390), repeat_header: false, cant_split: true,
      height_rule: 'atLeast' as const, height_twips: 400,
      cells: [{ id: 'cell:1', anchor: anchor('/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]', 220, 380), width_twips: 400, grid_span: 1, vertical_merge: 'none' as const, paragraphs: [cell] }],
    }],
  }
  const document: NativeDocxDocumentV1 = {
    protocol: DOCX_NATIVE_PROTOCOL, version: DOCX_NATIVE_VERSION, document_id: 'document:seating', revision: 'revision:1',
    source: { package_sha256: HASH, main_part: 'word/document.xml' },
    body: {
      id: 'story:body', kind: 'body', part_name: 'word/document.xml', anchor: anchor('/w:document[1]/w:body[1]', 1, 500),
      blocks: [{ kind: 'paragraph', id: lead.id, paragraph: lead }, { kind: 'table', id: table.id, table }],
    },
    sections: [{
      id: 'section:1', anchor: anchor('/w:document[1]/w:body[1]/w:sectPr[1]', 450, 490), starts_at_block_id: lead.id,
      break_type: 'next-page', title_page: false,
      page: {
        width_twips: 440, height_twips: 800, orientation: 'portrait',
        margins: { top_twips: 20, right_twips: 20, bottom_twips: 20, left_twips: 20, header_twips: 10, footer_twips: 10, gutter_twips: 0 },
        columns: 1, column_spacing_twips: 100, column_layout: 'equal-width', column_definitions: [{ id: 'section:1:column:0', ordinal: 0 }],
      },
      header_refs: [], footer_refs: [],
    }],
    headers: [], footers: [], notes: [], comment_stories: [], comments: [], capabilities: [],
    passthrough_parts: [
      { part_name: SETTINGS_PART, content_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml', byte_length: 1, sha256: HASH, policy: 'preserve-verbatim' },
      { part_name: RELATIONSHIPS_PART, content_type: 'application/vnd.openxmlformats-package.relationships+xml', byte_length: 1, sha256: HASH, policy: 'preserve-verbatim' },
    ],
    unsupported: [],
  }
  const resolved: NativeDocxResolvedLayoutInputV1 = {
    protocol: DOCX_RESOLVED_LAYOUT_PROTOCOL, version: DOCX_RESOLVED_LAYOUT_VERSION, document_id: document.document_id, revision: document.revision,
    source_parts: { main_part: 'word/document.xml' },
    paragraphs: [
      { paragraph_id: lead.id, applied_styles: [], properties: { spacing_after_twips: BODY_SPACE_AFTER / 50 }, paragraph_mark_properties: { font_family: 'Test', font_size_half_points: 20 } },
      { paragraph_id: cell.id, applied_styles: [], properties: { spacing_before_twips: CELL_SPACE_BEFORE / 50 }, paragraph_mark_properties: { font_family: 'Test', font_size_half_points: 20 } },
    ],
    runs: [cell].map(entry => ({ run_id: entry.runs[0]!.id, paragraph_id: entry.id, applied_paragraph_styles: [], applied_character_styles: [], properties: { font_family: 'Test', font_size_half_points: 20 } })),
    tables: [{ table_id: table.id }], fonts: [], diagnostics: [],
  }
  const line = (id: string, availableWidth = 20_000) => ({
    id, ordinal: 0, available_width_millipoints: availableWidth, inline_offset_millipoints: 0, advance_inline_millipoints: 0,
    ascent_millipoints: 5_000, descent_millipoints: -1_000, line_gap_millipoints: 0, line_height_millipoints: LINE_HEIGHT,
    justified: false, logical_to_visual: [], fragments: [],
  })
  const shaped: NativeDocxShapedLinesV1 = {
    protocol: DOCX_SHAPED_LINES_PROTOCOL, version: DOCX_SHAPED_LINES_VERSION, document_id: document.document_id, revision: document.revision,
    available_width_millipoints: 20_000, tab_interval_millipoints: DOCX_DEFAULT_TAB_STOP_TWIPS * 50,
    font_manifest: { manifest_id: 'manifest:test', revision: 'revision:1' },
    providers: { resolver_id: 'resolver:test', resolver_revision: '1', shaper_id: 'shaper:test', shaper_revision: '1', bidi_id: NATIVE_BIDI_PROVIDER_ID, bidi_revision: NATIVE_BIDI_PROVIDER_REVISION, bidi_unicode_version: BIDI_UNICODE_VERSION, unicode13_revision: UNICODE_13_CLASSIFIER_REVISION },
    diagnostics: [],
    paragraphs: [
      { paragraph_id: lead.id, story_id: document.body.id, story_kind: 'body', direction: 'ltr', alignment: 'start', spacing_before_millipoints: 0, spacing_after_millipoints: BODY_SPACE_AFTER, indent_start_millipoints: 0, indent_end_millipoints: 0, first_line_delta_millipoints: 0, block_advance_millipoints: LINE_HEIGHT + BODY_SPACE_AFTER, lines: [line(`line:${lead.id}:0`)] },
      { paragraph_id: cell.id, story_id: document.body.id, story_kind: 'body', direction: 'ltr', alignment: 'start', spacing_before_millipoints: CELL_SPACE_BEFORE, spacing_after_millipoints: 0, indent_start_millipoints: 0, indent_end_millipoints: 0, first_line_delta_millipoints: 0, block_advance_millipoints: LINE_HEIGHT + CELL_SPACE_BEFORE, lines: [line(`line:${cell.id}:0`)] },
    ],
  }
  return {
    protocol: DOCX_PAGINATION_REQUEST_PROTOCOL, version: DOCX_PAGINATION_REQUEST_VERSION, document, resolved_layout: resolved, shaped_lines: shaped,
    pagination_settings: {
      protocol: DOCX_PAGINATION_SETTINGS_PROTOCOL, version: DOCX_PAGINATION_SETTINGS_VERSION, document_id: document.document_id, revision: document.revision,
      package_sha256: HASH, main_part: 'word/document.xml', relationships_part: RELATIONSHIPS_PART, relationships_sha256: HASH, relationship_id: 'rIdSettings',
      settings_part: SETTINGS_PART, settings_sha256: HASH, profile: 'word-modern-default', default_tab_stop_twips: DOCX_DEFAULT_TAB_STOP_TWIPS,
      mirror_margins: false, gutter_at_top: false, even_and_odd_headers: false, compatibility_mode: 15, diagnostics: [],
    },
  }
}

/** The same shape with a real caption above the table, as TableWithAboveCaptions.docx
 * has it: the preceding paragraph carries text and states its own space-after. */
function captionFixture(): NativeDocxPaginationRequestV1 {
  const request = fixture()
  const lead = request.document.body.blocks[0]!.paragraph!
  lead.runs = [{ kind: 'text', id: `run:${lead.id}`, anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:t[1]', 105, 125), properties: {}, text: 'Table 1' }]
  request.resolved_layout.runs.push({ run_id: `run:${lead.id}`, paragraph_id: lead.id, applied_paragraph_styles: [], applied_character_styles: [], properties: { font_family: 'Test', font_size_half_points: 20 } })
  return request
}

/** Turn the request into one the approximate lane accepts, plus its eligibility. */
function approximate(request: NativeDocxPaginationRequestV1): NativeDocxApproximationEligibilityV1 {
  const settings = request.pagination_settings
  settings.profile = 'unsupported'
  delete settings.compatibility_mode
  settings.diagnostics = [{ code: 'COMPATIBILITY_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: SETTINGS_PART, path: '/w:settings[1]/w:compat[1]', preservation: 'preserve-verbatim', message: 'Legacy Word mode 14 requires different semantics' }]
  return {
    protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: settings.document_id, revision: settings.revision,
    package_sha256: settings.package_sha256, settings_sha256: settings.settings_sha256, status: 'eligible', legacy_compatibility_mode: 14,
    reasons: ['Legacy mode 14 uses current layout'],
  } as NativeDocxApproximationEligibilityV1
}

const cellLineY = (layout: { pages: Array<{ lines: Array<{ paragraph_id: string; y_millipoints: number }> }> }) =>
  layout.pages[0]!.lines.find(entry => entry.paragraph_id === 'paragraph:cell')!.y_millipoints

const leadLineY = (layout: { pages: Array<{ lines: Array<{ paragraph_id: string; y_millipoints: number }> }> }) =>
  layout.pages[0]!.lines.find(entry => entry.paragraph_id === 'paragraph:lead')!.y_millipoints

describe('approximate table seating below a spaced paragraph', () => {
  it('leaves the preceding space-after above the table and still applies the cell space-before', () => {
    const strictRequest = fixture()
    const strictResult = paginateNativeDocxV1(strictRequest)
    expect(strictResult.ok).toBe(true)
    if (!strictResult.ok) throw new Error('strict fixture must decode')
    if (strictResult.value.status !== 'paginated') throw new Error(`strict fixture must paginate: ${JSON.stringify(strictResult.value.diagnostics)}`)
    const strict = strictResult.value
    // Strict keeps the historical seating: the table starts at the bare cursor.
    expect(cellLineY(strict) - leadLineY(strict)).toBe(LINE_HEIGHT + CELL_SPACE_BEFORE)

    const request = fixture()
    const eligibility = approximate(request)
    const { layout } = paginateNativeDocxApproximateLegacyV1(request, eligibility)
    expect(layout.status).toBe('paginated')
    if (layout.status !== 'paginated') return
    // Word: a table has no space-before of its own, so nothing collapses the
    // preceding paragraph's space-after; the cell paragraph's own space-before
    // is then applied inside the cell.
    expect(cellLineY(layout) - leadLineY(layout)).toBe(LINE_HEIGHT + BODY_SPACE_AFTER + CELL_SPACE_BEFORE)
  })

  it('keeps the strict pagination output byte-identical whether or not the approximate lane ran', () => {
    const before = paginateNativeDocxV1(fixture())
    const input = fixture()
    const snapshot = structuredClone(input)
    const eligibility = approximate(structuredClone(input))
    // Run the approximate lane over its own copy, then re-run strict.
    const approximateRequest = fixture()
    paginateNativeDocxApproximateLegacyV1(approximateRequest, approximate(approximateRequest))
    const after = paginateNativeDocxV1(fixture())
    expect(JSON.stringify(after)).toBe(JSON.stringify(before))
    expect(input).toEqual(snapshot)
    expect(eligibility.status).toBe('eligible')
  })

  it('seats a table under a caption paragraph that states a real space-after', () => {
    // Strict pagination still starts a table at the bare cursor, so it keeps
    // refusing rather than dropping the caption's space-after silently.
    const strictResult = paginateNativeDocxV1(captionFixture())
    expect(strictResult.ok).toBe(true)
    if (!strictResult.ok) return
    expect(strictResult.value.status).toBe('refused')
    expect(strictResult.value.diagnostics.some(entry => entry.code === 'body-table-unsupported' && entry.message.includes('Paragraph spacing adjacent to a table must be explicit zero in v1'))).toBe(true)

    const request = captionFixture()
    const { layout } = paginateNativeDocxApproximateLegacyV1(request, approximate(request))
    expect(layout.status).toBe('paginated')
    if (layout.status !== 'paginated') return
    expect(cellLineY(layout) - leadLineY(layout)).toBe(LINE_HEIGHT + BODY_SPACE_AFTER + CELL_SPACE_BEFORE)
  })

  it('does not add the gap when the table starts an empty column', () => {
    const request = fixture()
    // Drop the leading paragraph: the table now opens the first column, where
    // Word has no preceding space-after to carry.
    request.document.body.blocks = request.document.body.blocks.filter(block => block.kind === 'table')
    request.document.sections[0]!.starts_at_block_id = 'table:1'
    request.resolved_layout.paragraphs = request.resolved_layout.paragraphs.filter(entry => entry.paragraph_id !== 'paragraph:lead')
    request.resolved_layout.runs = request.resolved_layout.runs.filter(entry => entry.paragraph_id !== 'paragraph:lead')
    request.shaped_lines.paragraphs = request.shaped_lines.paragraphs.filter(entry => entry.paragraph_id !== 'paragraph:lead')
    const { layout } = paginateNativeDocxApproximateLegacyV1(request, approximate(request))
    expect(layout.status).toBe('paginated')
    if (layout.status !== 'paginated') return
    const body = layout.pages[0]!.body_box.y_millipoints
    expect(cellLineY(layout) - body).toBe(CELL_SPACE_BEFORE)
  })
})
