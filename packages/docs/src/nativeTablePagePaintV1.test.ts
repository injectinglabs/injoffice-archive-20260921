import { describe, expect, it } from 'vitest'
import { BIDI_UNICODE_VERSION, NATIVE_BIDI_PROVIDER_ID, NATIVE_BIDI_PROVIDER_REVISION } from '@injoffice/font-metrics/bidi'
import { UNICODE_13_CLASSIFIER_REVISION } from '@injoffice/font-metrics/unicode13'
import { DOCX_NATIVE_PROTOCOL, DOCX_NATIVE_VERSION, type NativeDocxDocumentV1, type NativeDocxParagraphV1 } from './nativeContract.js'
import { DOCX_RESOLVED_LAYOUT_PROTOCOL, DOCX_RESOLVED_LAYOUT_VERSION, decodeNativeDocxResolvedLayout, type NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'
import { DOCX_SHAPED_LINES_PROTOCOL, DOCX_SHAPED_LINES_VERSION, type NativeDocxShapedLinesV1 } from './nativeShapingLines.js'
import { DOCX_DEFAULT_TAB_STOP_TWIPS, DOCX_PAGINATION_SETTINGS_PROTOCOL, DOCX_PAGINATION_SETTINGS_VERSION } from './nativePaginationSettings.js'
import { DOCX_PAGINATION_REQUEST_PROTOCOL, DOCX_PAGINATION_REQUEST_VERSION, paginateNativeDocxApproximateLegacyV1, paginateNativeDocxV1, type NativeDocxPaginationRequestV1 } from './nativePaginationV1.js'
import { layoutNativeDocxTableRowsV1, nativeDocxTableProjectionSha256V1, qualifyNativeDocxTablesV1 } from './nativeTablePagePaintV1.js'
import { qualifyApproximateLegacyTables } from './nativeLegacyTableOriginV1.js'
import { isRenderNeutralLayoutDiagnostic } from './nativeRenderDiagnostics.js'
import { decodeNativeDocxPaginatedLayoutForRequest } from './nativePaginatedLayoutContract.js'
import { DOCX_AUTO_BORDER_POLICY, type NativeDocxAutomaticBorderEvidenceV1 } from './nativeAutomaticBorderEvidenceV1.js'

const HASH = `sha256:${'a'.repeat(64)}`
const anchor = (path: string, start: number, end: number) => ({ part_name: 'word/document.xml', path, start_byte: start, end_byte: end, xml_sha256: HASH })
const policy = { mode: 'read-only' as const, allowed_operations: [] as const, refusal: { code: 'NATIVE_READ_ONLY', message: 'Source remains authoritative.', preservation: 'refuse-mutation' as const } }

function paragraph(id: string, ordinal: number): NativeDocxParagraphV1 {
  return { id, anchor: anchor(`/w:document[1]/w:body[1]/w:tbl[1]/w:tr[${ordinal + 1}]/w:tc[1]/w:p[1]`, 100 + ordinal * 40, 110 + ordinal * 40), edit_policy: { ...policy, allowed_operations: [] }, properties: {}, runs: [] }
}

/**
 * A balanced two-column section holding four body paragraphs and, between the
 * second and the third, a table that floats against the margin. This is the
 * shape of `floating-table-section-columns.docx` in hard-v2, whose Word
 * 16.112.4 export is the oracle for both halves of the behaviour: the float is
 * anchored to the top it would have had inline in the balance, and the second
 * line of BOTH columns is moved to the float's bottom edge.
 */
function balancedFloatFixture(tblpYTwips = -40): NativeDocxPaginationRequestV1 {
  const request = fixture(40_000)
  narrow(request, 360, 0)
  const table = request.document.body.blocks[0]!.table!
  table.rows = [table.rows[0]!]
  table.floating_position = { horizontal_anchor: 'margin', vertical_anchor: 'text', y_twips: tblpYTwips, left_from_text_twips: 180, right_from_text_twips: 180, top_from_text_twips: 0, bottom_from_text_twips: 0 }
  const body: NativeDocxParagraphV1[] = ['a1', 'a2', 'a3', 'a4'].map((name, index) => ({
    id: `paragraph:${name}`, anchor: anchor(`/w:document[1]/w:body[1]/w:p[${index + 1}]`, index < 2 ? 10 + index * 20 : 130 + (index - 2) * 20, index < 2 ? 20 + index * 20 : 140 + (index - 2) * 20),
    edit_policy: { ...policy, allowed_operations: [] }, properties: {}, runs: [],
  }))
  const cellShaped = request.shaped_lines.paragraphs[0]!
  request.document.body.blocks = [
    ...body.slice(0, 2).map((entry) => ({ kind: 'paragraph' as const, id: entry.id, paragraph: entry })),
    { kind: 'table' as const, id: table.id, table },
    ...body.slice(2).map((entry) => ({ kind: 'paragraph' as const, id: entry.id, paragraph: entry })),
  ]
  request.document.sections[0]!.starts_at_block_id = body[0]!.id
  request.document.sections[0]!.page.columns = 2
  request.document.sections[0]!.page.column_definitions = [{ id: 'section:1:column:0', ordinal: 0 }, { id: 'section:1:column:1', ordinal: 1 }]
  request.resolved_layout.paragraphs = [
    ...request.resolved_layout.paragraphs.slice(0, 1),
    ...body.map((entry) => ({ paragraph_id: entry.id, applied_styles: [], properties: {}, paragraph_mark_properties: {} })),
  ]
  request.shaped_lines.paragraphs = [
    cellShaped,
    ...body.map((entry) => ({ ...structuredClone(cellShaped), paragraph_id: entry.id, lines: [{ ...structuredClone(cellShaped.lines[0]!), id: `line:${entry.id}:0`, available_width_millipoints: 7_500 }] })),
  ]
  return request
}

/** Re-states a fixture table at a narrower authored width and indent, so the
 * text column and the margin box are distinguishable anchor boxes. */
function narrow(request: NativeDocxPaginationRequestV1, widthTwips: number, indentTwips: number): void {
  const table = request.document.body.blocks[0]!.table!
  table.width_twips = widthTwips
  table.grid_widths_twips = [widthTwips]
  table.indent_twips = indentTwips
  for (const row of table.rows) { row.cells[0]!.width_twips = widthTwips; row.cant_split = false }
  const content = (widthTwips - 2 * table.cell_margins!.left_twips) * 50
  for (const entry of request.shaped_lines.paragraphs) for (const line of entry.lines) line.available_width_millipoints = content
}

function fixture(bodyHeight = 10_000): NativeDocxPaginationRequestV1 {
  const paragraphs = [paragraph('paragraph:cell:1', 0), paragraph('paragraph:cell:2', 1)]
  const table = {
    id: 'table:1', anchor: anchor('/w:document[1]/w:body[1]/w:tbl[1]', 80, 180), edit_policy: { ...policy, allowed_operations: [] },
    width_twips: 400, layout: 'fixed' as const, alignment: 'left' as const, indent_twips: 0, grid_widths_twips: [400],
    cell_margins: { top_twips: 10, right_twips: 10, bottom_twips: 10, left_twips: 10 },
    borders: {
      top: { style: 'single' as const, size_eighth_points: 8, color_rgb: '112233' }, right: { style: 'single' as const, size_eighth_points: 8, color_rgb: '112233' },
      bottom: { style: 'single' as const, size_eighth_points: 8, color_rgb: '112233' }, left: { style: 'single' as const, size_eighth_points: 8, color_rgb: '112233' },
      inside_horizontal: { style: 'single' as const, size_eighth_points: 4, color_rgb: '445566' },
    },
    rows: paragraphs.map((entry, index) => ({
      id: `row:${index + 1}`, anchor: anchor(`/w:document[1]/w:body[1]/w:tbl[1]/w:tr[${index + 1}]`, 90 + index * 40, 125 + index * 40), repeat_header: false, cant_split: true,
      cells: [{ id: `cell:${index + 1}`, anchor: anchor(`/w:document[1]/w:body[1]/w:tbl[1]/w:tr[${index + 1}]/w:tc[1]`, 95 + index * 40, 120 + index * 40), width_twips: 400, grid_span: 1, vertical_merge: 'none' as const, ...(index === 0 ? { shading_rgb: 'DDEEFF' } : {}), paragraphs: [entry] }],
    })),
  }
  const document: NativeDocxDocumentV1 = {
    protocol: DOCX_NATIVE_PROTOCOL, version: DOCX_NATIVE_VERSION, document_id: 'document:table', revision: 'revision:1', source: { package_sha256: HASH, main_part: 'word/document.xml' },
    body: { id: 'story:body', kind: 'body', part_name: 'word/document.xml', anchor: anchor('/w:document[1]/w:body[1]', 1, 200), blocks: [{ kind: 'table', id: table.id, table }] },
    sections: [{ id: 'section:1', anchor: anchor('/w:document[1]/w:body[1]/w:sectPr[1]', 181, 199), starts_at_block_id: table.id, break_type: 'next-page', title_page: false, page: { width_twips: 440, height_twips: bodyHeight / 50 + 40, orientation: bodyHeight / 50 + 40 >= 440 ? 'portrait' : 'landscape', margins: { top_twips: 20, right_twips: 20, bottom_twips: 20, left_twips: 20, header_twips: 10, footer_twips: 10, gutter_twips: 0 }, columns: 1, column_spacing_twips: 100, column_layout: 'equal-width', column_definitions: [{ id: 'section:1:column:0', ordinal: 0 }] }, header_refs: [], footer_refs: [] }],
    headers: [], footers: [], notes: [], comment_stories: [], comments: [], capabilities: [], passthrough_parts: [
      { part_name: 'word/settings.xml', content_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml', byte_length: 1, sha256: HASH, policy: 'preserve-verbatim' },
      { part_name: 'word/_rels/document.xml.rels', content_type: 'application/vnd.openxmlformats-package.relationships+xml', byte_length: 1, sha256: HASH, policy: 'preserve-verbatim' },
    ], unsupported: [],
  }
  const resolved: NativeDocxResolvedLayoutInputV1 = {
    protocol: DOCX_RESOLVED_LAYOUT_PROTOCOL, version: DOCX_RESOLVED_LAYOUT_VERSION, document_id: document.document_id, revision: document.revision,
    source_parts: { main_part: 'word/document.xml' }, paragraphs: paragraphs.map((entry) => ({ paragraph_id: entry.id, applied_styles: [], properties: {}, paragraph_mark_properties: {} })), runs: [], tables: [{ table_id: table.id }], fonts: [], diagnostics: [],
  }
  const shaped: NativeDocxShapedLinesV1 = {
    protocol: DOCX_SHAPED_LINES_PROTOCOL, version: DOCX_SHAPED_LINES_VERSION, document_id: document.document_id, revision: document.revision,
    available_width_millipoints: 20_000, tab_interval_millipoints: DOCX_DEFAULT_TAB_STOP_TWIPS * 50,
    font_manifest: { manifest_id: 'manifest:test', revision: 'revision:1' }, providers: { resolver_id: 'resolver:test', resolver_revision: '1', shaper_id: 'shaper:test', shaper_revision: '1', bidi_id: NATIVE_BIDI_PROVIDER_ID, bidi_revision: NATIVE_BIDI_PROVIDER_REVISION, bidi_unicode_version: BIDI_UNICODE_VERSION, unicode13_revision: UNICODE_13_CLASSIFIER_REVISION }, diagnostics: [],
    paragraphs: paragraphs.map((entry) => ({ paragraph_id: entry.id, story_id: document.body.id, story_kind: 'body', direction: 'ltr', alignment: 'start', spacing_before_millipoints: 0, spacing_after_millipoints: 0, indent_start_millipoints: 0, indent_end_millipoints: 0, first_line_delta_millipoints: 0, block_advance_millipoints: 6_000, lines: [{ id: `line:${entry.id}:0`, ordinal: 0, available_width_millipoints: 19_000, inline_offset_millipoints: 0, advance_inline_millipoints: 0, ascent_millipoints: 5_000, descent_millipoints: -1_000, line_gap_millipoints: 0, line_height_millipoints: 6_000, justified: false, logical_to_visual: [], fragments: [] }] })),
  }
  return {
    protocol: DOCX_PAGINATION_REQUEST_PROTOCOL, version: DOCX_PAGINATION_REQUEST_VERSION, document, resolved_layout: resolved, shaped_lines: shaped,
    pagination_settings: { protocol: DOCX_PAGINATION_SETTINGS_PROTOCOL, version: DOCX_PAGINATION_SETTINGS_VERSION, document_id: document.document_id, revision: document.revision, package_sha256: HASH, main_part: 'word/document.xml', relationships_part: 'word/_rels/document.xml.rels', relationships_sha256: HASH, relationship_id: 'rIdSettings', settings_part: 'word/settings.xml', settings_sha256: HASH, profile: 'word-modern-default', default_tab_stop_twips: DOCX_DEFAULT_TAB_STOP_TWIPS, mirror_margins: false, gutter_at_top: false, even_and_odd_headers: false, compatibility_mode: 15, diagnostics: [] },
  }
}

/** The approximate lane only runs on a document whose settings are ineligible
 * for the strict profile; these two mirror the server's declared eligibility. */
function approximateRequest(request: NativeDocxPaginationRequestV1): NativeDocxPaginationRequestV1 {
  const copy = structuredClone(request)
  copy.pagination_settings.profile = 'unsupported'
  delete copy.pagination_settings.compatibility_mode
  copy.pagination_settings.diagnostics = [{ code: 'COMPATIBILITY_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: 'word/settings.xml', path: '/w:settings[1]/w:compat[1]', preservation: 'preserve-verbatim', message: 'Legacy Word mode 14 requires different semantics' }]
  return copy
}

function approximateEligibility(request: NativeDocxPaginationRequestV1): unknown {
  return {
    protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: request.document.document_id, revision: request.document.revision,
    package_sha256: HASH, settings_sha256: HASH, status: 'eligible', legacy_compatibility_mode: 14, reasons: ['Legacy mode 14 uses current layout'],
  }
}

function appendRow(request: NativeDocxPaginationRequestV1, ordinal: number): void {
  const table = request.document.body.blocks[0]!.table!
  const row = structuredClone(table.rows[1]!)
  row.id = `row:${ordinal}`; row.repeat_header = false; row.cells[0]!.id = `cell:${ordinal}`
  row.cells[0]!.paragraphs[0]!.id = `paragraph:cell:${ordinal}`
  table.rows.push(row)
  request.resolved_layout.paragraphs.push({ ...structuredClone(request.resolved_layout.paragraphs[1]!), paragraph_id: `paragraph:cell:${ordinal}` })
  const shaped = structuredClone(request.shaped_lines.paragraphs[1]!)
  shaped.paragraph_id = `paragraph:cell:${ordinal}`; shaped.lines[0]!.id = `line:paragraph:cell:${ordinal}:0`
  request.shaped_lines.paragraphs.push(shaped)
}

describe('bounded native DOCX table page-paint geometry', () => {
  /** Mirrors 2_table_doc.docx: TableGrid style, tblW auto, tblGrid 4428+4428 with
   * matching tcW, Letter page with 1800-twip side margins (8640-twip column). */
  function autoGridFixture(grid: number[]): NativeDocxPaginationRequestV1 {
    const request = fixture()
    const table = request.document.body.blocks[0]!.table!
    delete table.layout; delete table.alignment; delete table.indent_twips; delete table.width_twips; delete table.cell_margins
    table.table_style_id = 'TableGrid'
    table.grid_widths_twips = [...grid]
    for (const row of table.rows) {
      const first = row.cells[0]!
      row.cells = grid.map((width, index) => {
        if (index === 0) return { ...first, width_twips: width }
        const paragraphID = `${first.paragraphs[0]!.id}:${index}`
        request.resolved_layout.paragraphs.push({ ...structuredClone(request.resolved_layout.paragraphs[0]!), paragraph_id: paragraphID })
        const shaped = structuredClone(request.shaped_lines.paragraphs[0]!)
        shaped.paragraph_id = paragraphID; shaped.lines[0]!.id = `line:${paragraphID}:0`
        request.shaped_lines.paragraphs.push(shaped)
        return { ...structuredClone(first), id: `${first.id}:${index}`, width_twips: width, paragraphs: [{ ...structuredClone(first.paragraphs[0]!), id: paragraphID }] }
      })
    }
    const page = request.document.sections[0]!.page
    page.width_twips = 12_240; page.height_twips = 15_840; page.orientation = 'portrait'
    page.margins = { ...page.margins, left_twips: 1_800, right_twips: 1_800, top_twips: 1_440, bottom_twips: 1_440 }
    request.resolved_layout.tables = [{ table_id: table.id, style_id: 'TableGrid', geometry: { layout: 'autofit', alignment: 'left', indent_twips: 0, width_type: 'auto', width_value: 0, cell_margins: { top_twips: 0, right_twips: 108, bottom_twips: 0, left_twips: 108 } } }]
    return request
  }

  it('approximate preview fits an authored auto-table grid that exceeds the column to the body width', () => {
    const request = autoGridFixture([4428, 4428])
    const original = structuredClone(request.document)
    // Strict paint keeps content autofit: the preferences exceed the column, so
    // today's policy collapses to shaped content width. Only approximate changes.
    const strict = qualifyNativeDocxTablesV1(request.document, request.resolved_layout, request.shaped_lines)
    expect(strict.status === 'qualified' ? strict.tables[0]!.width_policy?.name : strict.status).toBe('shaped-content-minmax-v1')
    if (strict.status === 'qualified') expect(strict.tables[0]!.width_millipoints).toBeLessThan(8640 * 50)
    const approximate = qualifyApproximateLegacyTables(request.document, request.resolved_layout, request.shaped_lines, { legacy_compatibility_mode: 14 })
    expect(approximate.status).toBe('qualified')
    const [entry] = approximate.tables
    expect(entry!.width_policy).toEqual({ name: 'approximate-authored-grid-fitted-v1', section_id: 'section:1', container_width_twips: 8640, available_width_twips: 8640, source_grid_widths_twips: [4428, 4428], source_cell_widths_twips: [[4428, 4428], [4428, 4428]], fitted_grid_widths_twips: [4320, 4320] })
    expect(entry!.grid_widths_millipoints).toEqual([4320 * 50, 4320 * 50])
    expect(entry!.width_millipoints).toBe(8640 * 50)
    expect(entry!.rows[0]!.cells.reduce((sum, cell) => sum + cell.width_millipoints, 0)).toBe(8640 * 50)
    expect(entry!.rows[0]!.cells.map((cell) => cell.content_width_millipoints)).toEqual([(4320 - 216) * 50, (4320 - 216) * 50])
    expect(entry!.table.layout).toBe('fixed')
    expect(request.document).toEqual(original)
    expect(qualifyNativeDocxTablesV1(request.document, request.resolved_layout, request.shaped_lines, false)).toEqual(strict)
    // Without declared eligibility the wrapper reproduces the strict projection.
    expect(qualifyApproximateLegacyTables(request.document, request.resolved_layout, request.shaped_lines)).toEqual(strict)
  })

  /** Mirrors Table_cell_auto_width_fdo69656.docx: one cell states
   * `<w:tcW w:w="0" w:type="auto"/>`, which the extractor models as an absent
   * width_twips. Word draws that cell at its gridCol width like every other. */
  it('qualifies a cell that states no absolute preferred width at its grid slice', () => {
    const request = autoGridFixture([4428, 4428])
    for (const row of request.document.body.blocks[0]!.table!.rows) delete row.cells[1]!.width_twips
    const approximate = qualifyApproximateLegacyTables(request.document, request.resolved_layout, request.shaped_lines, { legacy_compatibility_mode: 14 })
    expect(approximate.status).toBe('qualified')
    expect(approximate.status === 'qualified' ? approximate.tables[0]!.width_policy : undefined)
      .toEqual({ name: 'approximate-authored-grid-fitted-v1', section_id: 'section:1', container_width_twips: 8640, available_width_twips: 8640, source_grid_widths_twips: [4428, 4428], source_cell_widths_twips: [[4428, null], [4428, null]], fitted_grid_widths_twips: [4320, 4320] })
    expect(approximate.status === 'qualified' ? approximate.tables[0]!.rows[0]!.cells.map((cell) => cell.width_millipoints) : undefined).toEqual([4320 * 50, 4320 * 50])
  })

  /** An authored width that disagrees with the grid is still a conflict. */
  it('still refuses a fixed-grid cell whose stated width disagrees with its grid slice', () => {
    const request = fixture()
    request.document.body.blocks[0]!.table!.rows[0]!.cells[0]!.width_twips = 401
    expect(qualifyNativeDocxTablesV1(request.document, request.resolved_layout, request.shaped_lines).status).toBe('refused')
  })

  /** A fixed-grid cell that states no preferred width takes its grid slice. */
  it('qualifies a fixed-grid cell that states no absolute preferred width', () => {
    const request = fixture()
    delete request.document.body.blocks[0]!.table!.rows[0]!.cells[0]!.width_twips
    const qualified = qualifyNativeDocxTablesV1(request.document, request.resolved_layout, request.shaped_lines)
    expect(qualified.status).toBe('qualified')
    expect(qualified.status === 'qualified' ? qualified.tables[0]!.rows[0]!.cells[0]!.width_millipoints : undefined).toBe(400 * 50)
  })

  it('approximate preview scales an uneven authored grid to the column with largest-remainder twips', () => {
    const request = autoGridFixture([6000, 3000])
    const approximate = qualifyApproximateLegacyTables(request.document, request.resolved_layout, request.shaped_lines, { legacy_compatibility_mode: 14 })
    expect(approximate.status).toBe('qualified')
    const [entry] = approximate.tables
    expect(entry!.width_policy).toMatchObject({ name: 'approximate-authored-grid-fitted-v1', available_width_twips: 8640, source_grid_widths_twips: [6000, 3000], fitted_grid_widths_twips: [5760, 2880] })
    expect(entry!.grid_widths_millipoints).toEqual([5760 * 50, 2880 * 50])
    expect(entry!.width_millipoints).toBe(8640 * 50)
    expect(entry!.rows[1]!.cells.map((cell) => cell.width_millipoints)).toEqual([5760 * 50, 2880 * 50])
    const uneven = autoGridFixture([5000, 5000, 5000])
    const scaled = qualifyApproximateLegacyTables(uneven.document, uneven.resolved_layout, uneven.shaped_lines, { legacy_compatibility_mode: 14 })
    expect(scaled.status === 'qualified' ? scaled.tables[0]!.width_policy : scaled.status).toMatchObject({ fitted_grid_widths_twips: [2880, 2880, 2880] })
    const remainder = autoGridFixture([5000, 5000, 4000])
    const split = qualifyApproximateLegacyTables(remainder.document, remainder.resolved_layout, remainder.shaped_lines, { legacy_compatibility_mode: 14 })
    expect(split.status === 'qualified' ? split.tables[0]!.width_policy : split.status).toMatchObject({ fitted_grid_widths_twips: [3086, 3086, 2468] })
  })

  it('approximate preview without a tblGrid or with a fitting grid keeps the existing behavior', () => {
    const missing = autoGridFixture([4428, 4428])
    delete missing.document.body.blocks[0]!.table!.grid_widths_twips
    const missingStrict = qualifyNativeDocxTablesV1(missing.document, missing.resolved_layout, missing.shaped_lines)
    expect(missingStrict.status).toBe('refused')
    expect(qualifyApproximateLegacyTables(missing.document, missing.resolved_layout, missing.shaped_lines, { legacy_compatibility_mode: 14 })).toEqual(missingStrict)
    const fitting = autoGridFixture([4000, 4000])
    const approximate = qualifyApproximateLegacyTables(fitting.document, fitting.resolved_layout, fitting.shaped_lines, { legacy_compatibility_mode: 14 })
    expect(approximate.status).toBe('qualified')
    expect(approximate.tables[0]!.width_policy?.name).toBe('source-preferred-nonconflicting-v1')
    expect(approximate.tables[0]!.width_millipoints).toBe(8000 * 50)
    const conflicting = autoGridFixture([4428, 4428])
    conflicting.document.body.blocks[0]!.table!.rows[0]!.cells[1]!.width_twips = 4000
    const fallback = qualifyApproximateLegacyTables(conflicting.document, conflicting.resolved_layout, conflicting.shaped_lines, { legacy_compatibility_mode: 14 })
    expect(fallback.status === 'qualified' ? fallback.tables[0]!.width_policy?.name : fallback.status).not.toBe('approximate-authored-grid-fitted-v1')
  })

  /** fdo80800b_tableStyle.docx and tdf118812_tableStyles-comprehensive.docx:
   * auto-width TableGrid tables whose cells content autofit cannot measure. The
   * approximate lane already sizes an auto-width table from its authored grid
   * whenever no resolved geometry is published; an unmeasurable one is not a
   * different document, so it takes that same declared policy instead of
   * discarding every other block on the page. Strict paint still refuses. */
  it('approximate preview falls back to the authored grid when content autofit cannot measure', () => {
    const unmeasurable = autoGridFixture([4_000, 4_000])
    unmeasurable.document.body.blocks[0]!.table!.rows[0]!.cells[0]!.vertical_merge = 'restart'
    const strict = qualifyNativeDocxTablesV1(unmeasurable.document, unmeasurable.resolved_layout, unmeasurable.shaped_lines)
    expect(strict).toMatchObject({ status: 'refused', diagnostics: [{ message: expect.stringContaining('Content autofit requires') }] })
    const approximate = qualifyApproximateLegacyTables(unmeasurable.document, unmeasurable.resolved_layout, unmeasurable.shaped_lines, { legacy_compatibility_mode: 14 })
    expect(approximate.status).toBe('qualified')
    expect(approximate.tables[0]!.width_policy).toBeUndefined()
    expect(approximate.tables[0]!.width_millipoints).toBe(8_000 * 50)
    expect(approximate.tables[0]!.grid_widths_millipoints).toEqual([4_000 * 50, 4_000 * 50])
    // Withheld measurements select the same policy; strict still needs them.
    const withheld = autoGridFixture([4_000, 4_000])
    expect(qualifyNativeDocxTablesV1(withheld.document, withheld.resolved_layout).status).toBe('refused')
    const unmeasured = qualifyApproximateLegacyTables(withheld.document, withheld.resolved_layout, undefined, { legacy_compatibility_mode: 14 })
    expect(unmeasured.status).toBe('qualified')
    expect(unmeasured.tables[0]!.width_millipoints).toBe(8_000 * 50)
    // A measurement that does not join this source is still an integrity
    // failure, not a licence to guess, on either lane.
    expect(qualifyApproximateLegacyTables(withheld.document, withheld.resolved_layout, { ...withheld.shaped_lines, revision: 'stale' }, { legacy_compatibility_mode: 14 }).status).toBe('refused')
  })

  /** table-rtl.docx and conditionalstyles-tbllook.docx: w:tblW auto with no
   * w:tblLayout element at all, so no resolved table geometry is published and
   * the authored grid is 216 twips (two default cell margins) wider than the
   * text column. The approximate fallback used to paint that grid sum as a
   * fixed width, and pagination then refused the whole document with
   * line-geometry-invalid because a fixed table wider than its column has no
   * lawful placement. */
  it('fits a cascade-default auto table whose authored grid exceeds its column', () => {
    const request = autoGridFixture([4428, 4428])
    request.resolved_layout.tables = [{ table_id: 'table:1' }]
    const table = request.document.body.blocks[0]!.table!
    expect(table.layout).toBeUndefined()
    expect(table.width_twips).toBeUndefined()
    const original = structuredClone(request.document)
    const approximate = qualifyApproximateLegacyTables(request.document, request.resolved_layout, request.shaped_lines, { legacy_compatibility_mode: 14 })
    expect(approximate.status).toBe('qualified')
    const [entry] = approximate.tables
    expect(entry!.width_policy).toMatchObject({ name: 'approximate-authored-grid-fitted-v1', available_width_twips: 8640, source_grid_widths_twips: [4428, 4428], fitted_grid_widths_twips: [4320, 4320] })
    expect(entry!.width_millipoints).toBe(8640 * 50)
    expect(entry!.x_millipoints + entry!.width_millipoints).toBeLessThanOrEqual(8640 * 50)
    expect(request.document).toEqual(original)
    // Strict qualification still refuses: it never derives a width from the grid.
    expect(qualifyNativeDocxTablesV1(request.document, request.resolved_layout, request.shaped_lines).status).toBe('refused')
    const forPagination = approximateRequest(request)
    forPagination.shaped_lines.available_width_millipoints = 8640 * 50
    const paginated = paginateNativeDocxApproximateLegacyV1(forPagination, approximateEligibility(request))
    expect(paginated.layout.status).toBe('paginated')
    expect(paginated.layout.pages.length).toBe(1)
  })

  /** tdf135943_shapeWithText_LayoutInCell0_compat15.docx: w:tblW 5000 pct on a
   * 6123-twip text column with an authored grid of 3006 + 3111 = 6117 and cells
   * whose w:tcW state 3007 and 3115. The exact policy refuses twice over -- the
   * proportional split is 3008.95 / 3114.05, and the cell preferences disagree
   * with their grid slices -- so the approximate lane paints the authored grid,
   * which is what Word paints: measured on Word's own PDF export, the rule
   * between the cells is at 207.00 pt = the 56.7 pt margin plus 3006 twips, and
   * the right edge at 362.52 pt against 362.55 pt for 6117 twips, where the
   * stretched grid would put them at 207.15 pt and 362.85 pt. */
  function percentFixture(percent: number, grid: number[], cellWidths: number[], page: { width: number; height: number; margin: number }): NativeDocxPaginationRequestV1 {
    const request = autoGridFixture(grid)
    const table = request.document.body.blocks[0]!.table!
    delete table.table_style_id
    delete table.layout
    table.width_percent_fiftieths = percent
    for (const row of table.rows) for (const [index, cell] of row.cells.entries()) cell.width_twips = cellWidths[index]!
    const geometry = request.document.sections[0]!.page
    geometry.width_twips = page.width; geometry.height_twips = page.height
    geometry.orientation = page.height >= page.width ? 'portrait' : 'landscape'
    geometry.margins = { ...geometry.margins, left_twips: page.margin, right_twips: page.margin }
    request.resolved_layout.tables = [{ table_id: table.id }]
    return request
  }

  it('approximate preview paints a percentage table at its authored grid, as Word does', () => {
    const request = percentFixture(5000, [3006, 3111], [3007, 3115], { width: 8391, height: 5953, margin: 1134 })
    const original = structuredClone(request.document)
    const strict = qualifyNativeDocxTablesV1(request.document, request.resolved_layout, request.shaped_lines)
    expect(strict).toMatchObject({ status: 'refused', diagnostics: [{ message: expect.stringContaining('Percentage table width requires') }] })
    const approximate = qualifyApproximateLegacyTables(request.document, request.resolved_layout, request.shaped_lines, { legacy_compatibility_mode: 14 })
    expect(approximate.status).toBe('qualified')
    const [entry] = approximate.tables
    expect(entry!.width_policy).toEqual({
      name: 'approximate-percent-authored-grid-v1', section_id: 'section:1', container_width_twips: 6123, percent_fiftieths: 5000,
      percent_width_twips: 6123, percent_width_exact: true, source_grid_widths_twips: [3006, 3111],
      source_cell_widths_twips: [[3007, 3115], [3007, 3115]], painted_grid_widths_twips: [3006, 3111], painted_width_twips: 6117,
    })
    // 56.7 pt margin + 3006 twips = 207.00 pt, Word's own interior rule.
    expect(entry!.grid_widths_millipoints).toEqual([3006 * 50, 3111 * 50])
    expect(entry!.width_millipoints).toBe(6117 * 50)
    expect(entry!.rows[0]!.cells.map((cell) => cell.x_millipoints)).toEqual([0, 3006 * 50])
    expect(entry!.table.rows[0]!.cells.map((cell) => cell.width_twips)).toEqual([3006, 3111])
    expect(request.document).toEqual(original)
  })

  /** lvlPicBulletId.docx: w:tblW 4850 pct on a 9360-twip text column, so the
   * percentage is 9079.2 twips and the exact policy refuses on the fraction.
   * Word truncated its own resolution when it authored the file -- the gridCol
   * is 9079 -- and paints the table 454.0 pt wide, so the authored grid is both
   * what Word states and what Word draws. */
  it('records but does not resolve a percentage that is not a whole twip', () => {
    const request = percentFixture(4850, [9079], [0], { width: 12240, height: 15840, margin: 1440 })
    for (const row of request.document.body.blocks[0]!.table!.rows) delete row.cells[0]!.width_twips
    expect(qualifyNativeDocxTablesV1(request.document, request.resolved_layout, request.shaped_lines).status).toBe('refused')
    const approximate = qualifyApproximateLegacyTables(request.document, request.resolved_layout, request.shaped_lines, { legacy_compatibility_mode: 14 })
    expect(approximate.status).toBe('qualified')
    expect(approximate.tables[0]!.width_policy).toMatchObject({ container_width_twips: 9360, percent_width_twips: 9079, percent_width_exact: false, painted_width_twips: 9079 })
    expect(approximate.tables[0]!.width_millipoints).toBe(9079 * 50)
    // A percentage table whose authored grid does not fit its column has no
    // lawful placement and keeps the existing refusal on both tiers.
    const oversized = percentFixture(5000, [9400], [9400], { width: 12240, height: 15840, margin: 1440 })
    expect(qualifyApproximateLegacyTables(oversized.document, oversized.resolved_layout, oversized.shaped_lines, { legacy_compatibility_mode: 14 }).status).toBe('refused')
  })

  /** The same file's cells each state w:tcBorders single/sz=2/000000 on all four
   * edges against a table that states none, so every shared edge is stated twice
   * with the same value and there is no conflict to resolve. Word draws exactly
   * that grid: 0.24 pt bars centred on x = 56.76, 207.00, 362.52 pt. */
  it('approximate preview projects one uniform cell border set onto a table that states none', () => {
    const request = percentFixture(5000, [3006, 3111], [3007, 3115], { width: 8391, height: 5953, margin: 1134 })
    const table = request.document.body.blocks[0]!.table!
    delete table.borders
    const border = { style: 'single' as const, size_eighth_points: 2, color_rgb: '000000' }
    for (const row of table.rows) for (const cell of row.cells) cell.borders = { top: { ...border }, right: { ...border }, bottom: { ...border }, left: { ...border } }
    const strict = qualifyNativeDocxTablesV1(request.document, request.resolved_layout, request.shaped_lines)
    expect(strict).toMatchObject({ status: 'refused' })
    const approximate = qualifyApproximateLegacyTables(request.document, request.resolved_layout, request.shaped_lines, { legacy_compatibility_mode: 14 })
    expect(approximate.status).toBe('qualified')
    const [entry] = approximate.tables
    expect(entry!.cell_border_policy).toEqual({ name: 'approximate-uniform-cell-borders-projected-v1', source_border: border, source_cell_count: 4 })
    expect(entry!.table.borders).toEqual({ top: border, right: border, bottom: border, left: border, inside_horizontal: border, inside_vertical: border })
    expect(entry!.table.rows.every((row) => row.cells.every((cell) => cell.borders === undefined))).toBe(true)
    // Two cells that disagree on their shared edge are a real conflict and keep
    // the existing refusal on both lanes, as does a table that states its own.
    const conflicting = structuredClone(request)
    conflicting.document.body.blocks[0]!.table!.rows[0]!.cells[1]!.borders!.left = { style: 'single', size_eighth_points: 8, color_rgb: '000000' }
    expect(qualifyApproximateLegacyTables(conflicting.document, conflicting.resolved_layout, conflicting.shaped_lines, { legacy_compatibility_mode: 14 }).status).toBe('refused')
    const partial = structuredClone(request)
    delete partial.document.body.blocks[0]!.table!.rows[0]!.cells[0]!.borders!.top
    expect(qualifyApproximateLegacyTables(partial.document, partial.resolved_layout, partial.shaped_lines, { legacy_compatibility_mode: 14 }).status).toBe('refused')
  })

  it('uses authored tblGrid as approximate fixed width when source layout is auto', () => {
    const request = fixture()
    const table = request.document.body.blocks[0]!.table!
    delete table.layout
    delete table.alignment
    delete table.indent_twips
    delete table.width_twips
    delete table.cell_margins
    table.grid_widths_twips = [400]
    const original = structuredClone(request.document)
    expect(qualifyNativeDocxTablesV1(request.document, request.resolved_layout).status).toBe('refused')
    const approximate = qualifyApproximateLegacyTables(request.document, request.resolved_layout, request.shaped_lines, { legacy_compatibility_mode: 14 })
    expect(approximate.status).toBe('qualified')
    expect(approximate.tables[0]!.table.layout).toBe('fixed')
    expect(approximate.tables[0]!.table.width_twips).toBe(400)
    expect(request.document).toEqual(original)
    // Strict pagination shares the qualifier and qualifies strictly without eligibility.
    const strictPagination = paginateNativeDocxV1(request)
    expect(strictPagination).toMatchObject({ ok: true, value: { status: 'refused', pages: [] } })
    if (strictPagination.ok) expect(strictPagination.value.diagnostics.some(diagnostic => diagnostic.code === 'body-table-unsupported')).toBe(true)
    expect(request.document).toEqual(original)
  })

  /** cell-sdt-redline.docx: w:tblPr states only w:tblLayout fixed, with no
   * w:tblW, no w:tblInd and no w:tblCellMar. A fixed-layout table sizes its
   * columns from w:tblGrid (ECMA-376 17.4.53), so the grid sum is the width the
   * source states; the approximate lane used to refuse the whole document for
   * lack of a w:tblW that a fixed table does not need. */
  it('uses the authored tblGrid when an explicitly fixed table states no tblW', () => {
    const request = fixture()
    const table = request.document.body.blocks[0]!.table!
    table.layout = 'fixed'
    delete table.alignment
    delete table.indent_twips
    delete table.width_twips
    delete table.cell_margins
    table.grid_widths_twips = [400]
    const original = structuredClone(request.document)
    expect(qualifyNativeDocxTablesV1(request.document, request.resolved_layout).status).toBe('refused')
    const approximate = qualifyApproximateLegacyTables(request.document, request.resolved_layout, request.shaped_lines, { legacy_compatibility_mode: 14 })
    expect(approximate.status).toBe('qualified')
    expect(approximate.tables[0]!.table.layout).toBe('fixed')
    expect(approximate.tables[0]!.table.width_twips).toBe(400)
    expect(approximate.tables[0]!.table.alignment).toBe('left')
    expect(approximate.tables[0]!.table.indent_twips).toBe(0)
    expect(approximate.tables[0]!.table.cell_margins).toEqual({ top_twips: 0, right_twips: 115, bottom_twips: 0, left_twips: 115 })
    expect(request.document).toEqual(original)
    // A fixed table that does state its width keeps that width, not the grid sum.
    const stated = fixture()
    stated.document.body.blocks[0]!.table!.width_twips = 400
    stated.document.body.blocks[0]!.table!.grid_widths_twips = [400]
    expect(qualifyApproximateLegacyTables(stated.document, stated.resolved_layout, stated.shaped_lines, { legacy_compatibility_mode: 14 }))
      .toEqual(qualifyNativeDocxTablesV1(stated.document, stated.resolved_layout, stated.shaped_lines))
  })

  /** floating-table-section-columns.docx: w:tblPr states w:tblW 10998 dxa and
   * w:tblLayout fixed, and states no w:jc, no w:tblInd and no w:tblCellMar --
   * the twin of the case above, where the width is stated and the placement is
   * left to the cascade. The three cascade defaults of ECMA-376 17.4 (left
   * alignment, zero indent, Word's default cell margin) are the same three the
   * auto-width and percentage policies already apply, so a stated width was the
   * one shape whose silence still discarded the whole document body. */
  it('applies the cascade placement defaults to a fixed dxa table that states none', () => {
    const request = fixture()
    const table = request.document.body.blocks[0]!.table!
    table.layout = 'fixed'
    table.width_twips = 400
    table.grid_widths_twips = [400]
    delete table.alignment
    delete table.indent_twips
    delete table.cell_margins
    const original = structuredClone(request.document)
    const strict = qualifyNativeDocxTablesV1(request.document, request.resolved_layout, request.shaped_lines)
    expect(strict).toMatchObject({ status: 'refused', diagnostics: [{ message: 'Table requires explicit fixed dxa width, left alignment, indent, and all four cell margins' }] })
    const approximate = qualifyApproximateLegacyTables(request.document, request.resolved_layout, request.shaped_lines, { legacy_compatibility_mode: 14 })
    expect(approximate.status).toBe('qualified')
    expect(approximate.tables[0]!.table.width_twips).toBe(400)
    expect(approximate.tables[0]!.table.alignment).toBe('left')
    expect(approximate.tables[0]!.table.indent_twips).toBe(0)
    expect(approximate.tables[0]!.table.cell_margins).toEqual({ top_twips: 0, right_twips: 115, bottom_twips: 0, left_twips: 115 })
    expect(approximate.tables[0]!.x_millipoints).toBe(0)
    expect(request.document).toEqual(original)
    // A table that states all three keeps exactly what it states, on both lanes.
    const authored = fixture()
    authored.document.body.blocks[0]!.table!.indent_twips = 7
    expect(qualifyApproximateLegacyTables(authored.document, authored.resolved_layout, authored.shaped_lines, { legacy_compatibility_mode: 14 }))
      .toEqual(qualifyNativeDocxTablesV1(authored.document, authored.resolved_layout, authored.shaped_lines))
  })

  it('uses bounded resolved geometry without changing source or overriding direct properties', () => {
    const request=fixture(), table=request.document.body.blocks[0]!.table!
    const expected=structuredClone(qualifyNativeDocxTablesV1(request.document,request.resolved_layout))
    request.resolved_layout.tables[0]!.geometry={layout:'fixed',alignment:'left',indent_twips:0,width_type:'dxa',width_value:400,cell_margins:{...table.cell_margins!}}
    delete table.layout;delete table.alignment;delete table.indent_twips;delete table.width_twips;delete table.cell_margins
    const original=structuredClone(request.document)
    expect(decodeNativeDocxResolvedLayout(request.resolved_layout).ok).toBe(true)
    expect(qualifyNativeDocxTablesV1(request.document,request.resolved_layout)).toEqual(expected)
    expect(request.document).toEqual(original)
    expect(qualifyNativeDocxTablesV1(request.document,{...request.resolved_layout,revision:'stale'}).status).toBe('refused')
    expect(qualifyNativeDocxTablesV1(request.document,{...request.resolved_layout,document_id:'document:other'}).status).toBe('refused')
    table.width_twips=400;request.resolved_layout.tables[0]!.geometry!.width_value=800
    expect(qualifyNativeDocxTablesV1(request.document,request.resolved_layout)).toEqual(expected)
    for(const change of [ {width_type:'auto',width_value:400}, {width_value:-1}, {layout:'future'}, {cell_margins:{left_twips:10}}, {extra:true} ]) {
      const invalid=structuredClone(request.resolved_layout)
      Object.assign(invalid.tables[0]!.geometry!,change)
      expect(decodeNativeDocxResolvedLayout(invalid).ok).toBe(false)
    }
  })
  /**
   * `w:tblpPr` with `vertAnchor="text"`: the float keeps the top it would have
   * had inline and is displaced by `w:tblpY`, and its left edge comes from the
   * anchor box `w:horzAnchor` names rather than from the text column.
   * Validated against Microsoft Word 16.112.4 PDF exports of the three hard-v2
   * packages carrying such a frame, to 0.12 pt or better.
   */
  it('lifts a w:tblpPr table out of the inline flow at its own anchor', () => {
    const inline = fixture(), inlineTable = inline.document.body.blocks[0]!.table!
    narrow(inline, 300, 40)
    const inlineRows = paginateNativeDocxV1(inline)
    expect(inlineRows).toMatchObject({ ok: true, value: { status: 'paginated' } })
    if (!inlineRows.ok) throw new Error('invalid request')
    const inlineFirst = inlineRows.value.pages[0]!.table_rows![0]!
    // The inline table sits at the column origin plus its own w:tblInd.
    expect([inlineFirst.x_millipoints, inlineFirst.y_millipoints]).toEqual([3_000, 1_000])

    const floated = structuredClone(inline)
    floated.document.body.blocks[0]!.table!.floating_position = {
      horizontal_anchor: 'margin', vertical_anchor: 'text', y_twips: 12,
      left_from_text_twips: 180, right_from_text_twips: 180, top_from_text_twips: 0, bottom_from_text_twips: 0,
    }
    const result = paginateNativeDocxV1(floated)
    expect(result).toMatchObject({ ok: true, value: { status: 'paginated' } })
    if (!result.ok) throw new Error('invalid request')
    expect(decodeNativeDocxPaginatedLayoutForRequest(result.value, floated)).toMatchObject({ ok: true })
    const rows = result.value.pages[0]!.table_rows!
    // The frame replaces w:tblInd with the margin box, and 12 twips displaces
    // the float 600 milli-points below the top it would have had inline.
    expect([rows[0]!.x_millipoints, rows[0]!.y_millipoints]).toEqual([1_000, 1_600])
    // Every row moves with the float; nothing else about the table changes.
    expect(rows.map((row) => row.y_millipoints - inlineRows.value.pages[0]!.table_rows![row.row_ordinal]!.y_millipoints)).toEqual(rows.map(() => 600))

    // A frame that would lift the float off its own page is refused, not clamped.
    const offPage = structuredClone(floated)
    offPage.document.body.blocks[0]!.table!.floating_position!.y_twips = -12
    expect(paginateNativeDocxV1(offPage)).toMatchObject({ ok: true, value: { status: 'refused' } })

    const centred = structuredClone(floated)
    centred.document.body.blocks[0]!.table!.floating_position!.x_alignment = 'center'
    delete centred.document.body.blocks[0]!.table!.floating_position!.y_twips
    const centredResult = paginateNativeDocxV1(centred)
    if (!centredResult.ok) throw new Error('invalid request')
    // Body box 20,000 wide from x 1,000; the 15,000-wide table centres on it.
    expect(centredResult.value.pages[0]!.table_rows![0]!.x_millipoints).toBe(3_500)
  })

  it('floats a table over a balanced two-column section and moves both columns below it', () => {
    const request = balancedFloatFixture()
    const result = paginateNativeDocxV1(request)
    expect(result).toMatchObject({ ok: true, value: { status: 'paginated' } })
    if (!result.ok) throw new Error('invalid request')
    expect(decodeNativeDocxPaginatedLayoutForRequest(result.value, request)).toMatchObject({ ok: true })
    const page = result.value.pages[0]!
    expect(result.value.pages).toHaveLength(1)
    // Body box 20,000 from x 1,000, two 7,500 columns with a 5,000 gap.
    expect(page.columns.map((column) => [column.x_millipoints, column.width_millipoints])).toEqual([[1_000, 7_500], [13_500, 7_500]])

    // The four paragraphs balance 2/2 as if the float were not there, so the
    // float's own anchor is the bottom of the second one: 12,000 into column 0,
    // lifted 2,000 by w:tblpY="-40".
    const float = page.table_rows!
    expect(float).toHaveLength(1)
    expect([float[0]!.x_millipoints, float[0]!.y_millipoints]).toEqual([1_000, 11_000])
    // The authored 18,000 width overhangs both 7,500 columns and is kept.
    expect(float[0]!.width_millipoints).toBe(18_000)
    const bottom = float[0]!.y_millipoints + float[0]!.height_millipoints

    const placed = new Map(page.lines.filter((line) => !line.table_cell_id).map((line) => [line.paragraph_id, line]))
    // The first line of each column is above the float and does not move.
    expect([placed.get('paragraph:a1')!.x_millipoints, placed.get('paragraph:a1')!.y_millipoints]).toEqual([1_000, 1_000])
    expect([placed.get('paragraph:a3')!.x_millipoints, placed.get('paragraph:a3')!.y_millipoints]).toEqual([13_500, 1_000])
    // The second line of each column meets the float and resumes at its bottom,
    // in the column the float is anchored in and in the one it only spans.
    expect(placed.get('paragraph:a2')!.y_millipoints).toBe(bottom)
    expect(placed.get('paragraph:a4')!.y_millipoints).toBe(bottom)

    // A float that clears the text disturbs nothing.
    const clear = paginateNativeDocxV1(balancedFloatFixture(240))
    if (!clear.ok) throw new Error('invalid request')
    const clearLines = new Map(clear.value.pages[0]!.lines.filter((line) => !line.table_cell_id).map((line) => [line.paragraph_id, line.y_millipoints]))
    expect([...clearLines.values()]).toEqual([1_000, 7_000, 1_000, 7_000])
  })

  it('still refuses a table that has to flow through a multi-column section', () => {
    const request = balancedFloatFixture()
    delete request.document.body.blocks[2]!.table!.floating_position
    expect(paginateNativeDocxV1(request)).toMatchObject({ ok: true, value: { status: 'refused', diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'body-table-unsupported' })]) } })

    // So is a frame whose anchor this version cannot state.
    const unplaceable = balancedFloatFixture()
    unplaceable.document.body.blocks[2]!.table!.floating_position!.vertical_anchor = 'page'
    expect(paginateNativeDocxV1(unplaceable)).toMatchObject({ ok: true, value: { status: 'refused' } })
  })

  it('discloses a w:tblpPr frame it cannot place and keeps the table inline', () => {
    const request = fixture()
    narrow(request, 300, 40)
    const inlineY = paginateNativeDocxV1(request)
    if (!inlineY.ok) throw new Error('invalid request')
    const baseline = inlineY.value.pages[0]!.table_rows!.map((row) => [row.x_millipoints, row.y_millipoints])

    const deferred = structuredClone(request)
    deferred.document.body.blocks[0]!.table!.floating_position = {
      horizontal_anchor: 'margin', vertical_anchor: 'page', y_twips: 200,
      left_from_text_twips: 0, right_from_text_twips: 0, top_from_text_twips: 0, bottom_from_text_twips: 0,
    }
    const result = paginateNativeDocxV1(deferred)
    expect(result).toMatchObject({ ok: true, value: { status: 'paginated' } })
    if (!result.ok) throw new Error('invalid request')
    expect(result.value.pages[0]!.table_rows!.map((row) => [row.x_millipoints, row.y_millipoints])).toEqual(baseline)
    expect(result.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({
      code: 'body-table-unsupported', severity: 'deferred', scope_id: 'table:1',
    })]))
  })

  it('fragments natural rows at complete lines while repeating headers and preserving source coverage', () => {
    const request = fixture(20_000), table = request.document.body.blocks[0]!.table!
    table.rows[0]!.repeat_header = true
    table.rows[1]!.cant_split = false
    const paragraph = request.shaped_lines.paragraphs[1]!, originalLine = paragraph.lines[0]!
    paragraph.lines = Array.from({ length: 8 }, (_, ordinal) => ({ ...structuredClone(originalLine), id: `line:${paragraph.paragraph_id}:${ordinal}`, ordinal }))
    paragraph.block_advance_millipoints = 48_000
    const original = JSON.stringify(request), result = paginateNativeDocxV1(request)
    expect(result).toMatchObject({ ok: true, value: { status: 'paginated' } })
    if (!result.ok) throw new Error('invalid request')
    expect(decodeNativeDocxPaginatedLayoutForRequest(result.value, request)).toMatchObject({ ok: true })
    const fragments = result.value.pages.flatMap(page => page.table_rows ?? [])
    expect(fragments.length).toBeGreaterThan(1)
    expect(fragments.reduce((sum, fragment) => sum + fragment.height_millipoints, 0)).toBe(49_000)
    expect(result.value.pages.flatMap(page => page.lines.filter(line => line.paragraph_id === paragraph.paragraph_id).map(line => line.source_line_ordinal))).toEqual([0,1,2,3,4,5,6,7])
    expect(result.value.pages.every(page => page.lines.some(line => line.paragraph_id === 'paragraph:cell:1'))).toBe(true)
    const tampered = structuredClone(result.value)
    tampered.pages[1]!.table_rows![0]!.source_y_millipoints += 1
    expect(decodeNativeDocxPaginatedLayoutForRequest(tampered, request).ok).toBe(false)
    expect(JSON.stringify(request)).toBe(original)
    // A widow group is atomic even when individual lines would fit.
    const tooShort = structuredClone(request)
    tooShort.document.sections[0]!.page.height_twips = 420
    expect(paginateNativeDocxV1(tooShort)).toMatchObject({ ok: true, value: { status: 'refused', pages: [] } })
    request.resolved_layout.paragraphs[1]!.properties.keep_lines = true
    expect(paginateNativeDocxV1(request)).toMatchObject({ ok: true, value: { status: 'refused', pages: [] } })
  })
  it('refuses ambiguous split-row constraints before producing table geometry', () => {
    for (const property of ['keep_next', 'page_break_before'] as const) {
      const request = fixture()
      request.document.body.blocks[0]!.table!.rows[1]!.cant_split = false
      request.resolved_layout.paragraphs[1]!.properties[property] = true
      expect(qualifyNativeDocxTablesV1(request.document, request.resolved_layout)).toMatchObject({ status: 'refused', tables: [] })
    }
    const header = fixture()
    header.document.body.blocks[0]!.table!.rows[0]!.repeat_header = true
    delete header.document.body.blocks[0]!.table!.rows[0]!.cant_split
    expect(qualifyNativeDocxTablesV1(header.document, header.resolved_layout)).toMatchObject({ status: 'refused', tables: [] })
  })
  it('resolves explicit percent widths proportionally against the owning section without changing source', () => {
    const request = fixture(14_000), table = request.document.body.blocks[0]!.table!
    delete table.width_twips; table.width_percent_fiftieths = 2500
    for (const paragraph of request.shaped_lines.paragraphs) paragraph.lines[0]!.available_width_millipoints = 9_000
    const original = JSON.stringify(request)
    const result = qualifyNativeDocxTablesV1(request.document, request.resolved_layout)
    expect(result.status).toBe('qualified')
    if (result.status !== 'qualified') throw new Error('percentage refused')
    expect(result.tables[0]).toMatchObject({ width_millipoints: 10_000, grid_widths_millipoints: [10_000], width_policy: { name: 'fixed-grid-percent-exact-twips-v1', section_id: 'section:1', container_width_twips: 400, percent_fiftieths: 2500, source_grid_widths_twips: [400] } })
    expect([...result.paragraph_widths.values()]).toEqual([9_000,9_000])
    const paginated = paginateNativeDocxV1(request)
    expect(paginated).toMatchObject({ ok: true, value: { status: 'paginated' } })
    if (!paginated.ok) throw new Error('pagination failed')
    expect(decodeNativeDocxPaginatedLayoutForRequest(paginated.value, request).ok).toBe(true)
    expect(JSON.stringify(request)).toBe(original)
  })

  it('refuses non-integral percent projection and conflicting source grid/cell preferences', () => {
    for (const mutate of [
      (table: NonNullable<NativeDocxPaginationRequestV1['document']['body']['blocks'][number]['table']>) => { table.width_percent_fiftieths = 3333 },
      (table: NonNullable<NativeDocxPaginationRequestV1['document']['body']['blocks'][number]['table']>) => { table.rows[0]!.cells[0]!.width_twips = 399 },
      (table: NonNullable<NativeDocxPaginationRequestV1['document']['body']['blocks'][number]['table']>) => { table.width_twips = 400 },
    ]) {
      const request = fixture(), table = request.document.body.blocks[0]!.table!
      delete table.width_twips; table.width_percent_fiftieths = 2500; mutate(table)
      expect(qualifyNativeDocxTablesV1(request.document,request.resolved_layout)).toMatchObject({ status: 'refused', tables: [] })
    }
  })
  it('repeats a leading header exactly once on each continuation page with unique placement identities', () => {
    const request = fixture(14_000)
    const table = request.document.body.blocks[0]!.table!
    table.rows[0]!.repeat_header = true
    const row = structuredClone(table.rows[1]!)
    row.id = 'row:3'; row.cells[0]!.id = 'cell:3'; row.cells[0]!.paragraphs[0]!.id = 'paragraph:cell:3'
    table.rows.push(row)
    request.resolved_layout.paragraphs.push({ ...structuredClone(request.resolved_layout.paragraphs[1]!), paragraph_id: 'paragraph:cell:3' })
    const shaped = structuredClone(request.shaped_lines.paragraphs[1]!)
    shaped.paragraph_id = 'paragraph:cell:3'; shaped.lines[0]!.id = 'line:paragraph:cell:3:0'
    request.shaped_lines.paragraphs.push(shaped)
    const result = paginateNativeDocxV1(request)
    expect(result.ok && result.value.status === 'paginated' ? result.value.pages.map((page) => page.lines.map((line) => line.paragraph_id)) : result).toEqual([['paragraph:cell:1', 'paragraph:cell:2'], ['paragraph:cell:1', 'paragraph:cell:3']])
    if (!result.ok || result.value.status !== 'paginated') throw new Error('pagination refused')
    expect(decodeNativeDocxPaginatedLayoutForRequest(result.value, request)).toMatchObject({ ok: true })
    expect(new Set(result.value.pages.flatMap((page) => page.lines.map((line) => line.id))).size).toBe(4)
    expect(result.value.pages[1]!.lines[0]).toMatchObject({ repeated_table_header: true, y_millipoints: 1_500 })
    for (const mutate of [
      (value: typeof result.value) => { value.pages[1]!.lines[0]!.y_millipoints += 1 },
      (value: typeof result.value) => { delete value.pages[1]!.lines[0]!.repeated_table_header },
      (value: typeof result.value) => { value.pages[1]!.lines.shift(); value.pages[1]!.paragraph_slices.shift() },
    ]) {
      const changed = structuredClone(result.value); mutate(changed)
      expect(decodeNativeDocxPaginatedLayoutForRequest(changed, request).ok).toBe(false)
    }
  })

  it('refuses orphan headers, non-prefix headers and repeating vertical merges atomically', () => {
    for (const configure of [
      (request: NativeDocxPaginationRequestV1) => { request.document.body.blocks[0]!.table!.rows[0]!.repeat_header = true },
      (request: NativeDocxPaginationRequestV1) => { request.document.body.blocks[0]!.table!.rows[1]!.repeat_header = true },
      (request: NativeDocxPaginationRequestV1) => { const row = request.document.body.blocks[0]!.table!.rows[0]!; row.repeat_header = true; row.cells[0]!.vertical_merge = 'restart' },
    ]) {
      const request = fixture(10_000); configure(request)
      expect(paginateNativeDocxV1(request)).toMatchObject({ ok: true, value: { status: 'refused', pages: [] } })
    }
  })

  it('keeps a multi-row header prefix with the next body row and replays all header rows', () => {
    const request = fixture(21_000)
    appendRow(request, 3); appendRow(request, 4)
    const table = request.document.body.blocks[0]!.table!
    table.rows[0]!.repeat_header = true; table.rows[1]!.repeat_header = true
    const result = paginateNativeDocxV1(request)
    expect(result.ok && result.value.status === 'paginated' ? result.value.pages.map((page) => page.lines.map((line) => line.paragraph_id)) : result).toEqual([
      ['paragraph:cell:1', 'paragraph:cell:2', 'paragraph:cell:3'],
      ['paragraph:cell:1', 'paragraph:cell:2', 'paragraph:cell:4'],
    ])
    if (!result.ok) throw new Error('pagination failed')
    expect(decodeNativeDocxPaginatedLayoutForRequest(result.value, request).ok).toBe(true)
    const wrongCell = structuredClone(result.value)
    wrongCell.pages[1]!.lines[0]!.table_cell_id = 'cell:wrong'
    wrongCell.pages[1]!.paragraph_slices[0]!.table_cell_id = 'cell:wrong'
    expect(decodeNativeDocxPaginatedLayoutForRequest(wrongCell, request).ok).toBe(false)
  })

  it('refuses a later body row that fits alone but not together with its repeated headers', () => {
    const request = fixture(21_000)
    appendRow(request, 3)
    const table = request.document.body.blocks[0]!.table!
    table.rows[0]!.repeat_header = true
    table.rows[2]!.height_twips = 400; table.rows[2]!.height_rule = 'exact'
    expect(paginateNativeDocxV1(request)).toMatchObject({ ok: true, value: { status: 'refused', pages: [], diagnostics: [expect.objectContaining({ message: expect.stringContaining('Repeated headers and the next indivisible row') })] } })
  })
  it('derives a source-ordered 2x2 grid without guessing widths or shared cells', () => {
    const request = fixture()
    const table = request.document.body.blocks[0]!.table!
    table.grid_widths_twips = [200, 200]
    for (const [rowIndex, row] of table.rows.entries()) {
      row.cells[0]!.width_twips = 200
      const sourceParagraph = row.cells[0]!.paragraphs[0]!
      const secondParagraph = structuredClone(sourceParagraph)
      secondParagraph.id = `paragraph:cell:${rowIndex + 1}:second`
      secondParagraph.anchor = anchor(`/w:document[1]/w:body[1]/w:tbl[1]/w:tr[${rowIndex + 1}]/w:tc[2]/w:p[1]`, 111 + rowIndex * 40, 119 + rowIndex * 40)
      const secondCell = { ...structuredClone(row.cells[0]!), id: `cell:${rowIndex + 1}:second`, anchor: anchor(`/w:document[1]/w:body[1]/w:tbl[1]/w:tr[${rowIndex + 1}]/w:tc[2]`, 110 + rowIndex * 40, 120 + rowIndex * 40), paragraphs: [secondParagraph] }
      delete secondCell.shading_rgb
      row.cells.push(secondCell)
      request.resolved_layout.paragraphs.push({ paragraph_id: secondParagraph.id, applied_styles: [], properties: {}, paragraph_mark_properties: {} })
    }
    const qualified = qualifyNativeDocxTablesV1(request.document, request.resolved_layout)
    expect(qualified.status === 'qualified' ? qualified.tables[0]!.rows.map((row) => row.cells.map((cell) => [cell.column_ordinal, cell.width_millipoints, cell.content_width_millipoints])) : qualified).toEqual([[[0, 10_000, 9_000], [1, 10_000, 9_000]], [[0, 10_000, 9_000], [1, 10_000, 9_000]]])
  })

  it('qualifies exact fixed grids, derives cell widths/row heights, and has a golden canonical hash', () => {
    const request = fixture()
    const qualified = qualifyNativeDocxTablesV1(request.document, request.resolved_layout)
    expect(qualified.status).toBe('qualified')
    if (qualified.status !== 'qualified') return
    expect([...qualified.paragraph_widths.values()]).toEqual([19_000, 19_000])
    expect(layoutNativeDocxTableRowsV1(qualified.tables[0]!, request.shaped_lines)?.map((row) => row.height_millipoints)).toEqual([7_000, 7_000])
    expect(qualified.sha256).toBe('sha256:3072722fea3ca71712b89cdb8994f81d28ca8f7b5eebd27813f6b119d9ae214c')
    expect(nativeDocxTableProjectionSha256V1([])).toBe('sha256:4f53cda18c2baa0c0354bb5f9a3ecbe5ed12ab4d8e11ba873c2f11161202b945')
    expect(nativeDocxTableProjectionSha256V1(qualified.tables)).toBe(qualified.sha256)
  })

  it('moves indivisible rows to fresh pages and is byte-for-byte deterministic', () => {
    const request = fixture(10_000)
    const first = paginateNativeDocxV1(request)
    const second = paginateNativeDocxV1(structuredClone(request))
    expect(first).toEqual(second)
    expect(JSON.stringify(first)).toBe(JSON.stringify(second))
    expect(first.ok && first.value.status === 'paginated' ? first.value.pages.map((page) => page.lines.map((line) => line.paragraph_id)) : first).toEqual([['paragraph:cell:1'], ['paragraph:cell:2']])
  })

  it('qualifies horizontal spans, vertical merges, atLeast row height, and simple style ids with explicit borders', () => {
    const span = fixture()
    const spanned = span.document.body.blocks[0]!.table!
    spanned.grid_widths_twips = [200, 200]
    spanned.width_twips = 400
    for (const row of spanned.rows) {
      row.cells[0]!.grid_span = 2
      row.cells[0]!.width_twips = 400
    }
    expect(qualifyNativeDocxTablesV1(span.document, span.resolved_layout).status).toBe('qualified')

    const merged = fixture()
    const mergeTable = merged.document.body.blocks[0]!.table!
    mergeTable.rows[0]!.cells[0]!.vertical_merge = 'restart'
    mergeTable.rows[1]!.cells[0]!.vertical_merge = 'continue'
    mergeTable.rows[1]!.cells[0]!.paragraphs[0]!.runs = []
    const qualifiedMerge = qualifyNativeDocxTablesV1(merged.document, merged.resolved_layout)
    expect(qualifiedMerge.status).toBe('qualified')
    if (qualifiedMerge.status === 'qualified') {
      const geometry = layoutNativeDocxTableRowsV1(qualifiedMerge.tables[0]!, merged.shaped_lines)
      expect(geometry?.[0]?.cells[0]?.row_span).toBe(2)
      expect(geometry?.[0]?.cells[0]?.height_millipoints).toBe(14_000)
    }

    const tall = fixture()
    tall.document.body.blocks[0]!.table!.rows[0]!.height_twips = 400
    tall.document.body.blocks[0]!.table!.rows[0]!.height_rule = 'atLeast'
    const qualifiedTall = qualifyNativeDocxTablesV1(tall.document, tall.resolved_layout)
    expect(qualifiedTall.status).toBe('qualified')
    if (qualifiedTall.status === 'qualified') {
      expect(layoutNativeDocxTableRowsV1(qualifiedTall.tables[0]!, tall.shaped_lines)?.[0]?.height_millipoints).toBe(20_000)
    }

    const styled = fixture()
    styled.document.body.blocks[0]!.table!.table_style_id = 'PlainBorders'
    styled.resolved_layout.tables[0]!.style_id = 'PlainBorders'
    expect(qualifyNativeDocxTablesV1(styled.document, styled.resolved_layout).status).toBe('qualified')
  })

  it('refuses over-wide spans, continue content, conditional styles, width drift, and splittable rows without partial geometry', () => {
    for (const mutate of [
      (request: NativeDocxPaginationRequestV1) => { request.document.body.blocks[0]!.table!.rows[0]!.cells[0]!.grid_span = 2 },
      (request: NativeDocxPaginationRequestV1) => { request.document.body.blocks[0]!.table!.rows[0]!.cells[0]!.vertical_merge = 'continue' },
      (request: NativeDocxPaginationRequestV1) => {
        request.document.body.blocks[0]!.table!.table_style_id = 'TableGrid'
        request.resolved_layout.tables[0]!.style_id = 'TableGrid'
        request.resolved_layout.diagnostics.push({ code: 'CONDITIONAL_TABLE_STYLE_PRESERVED', severity: 'unsupported', scope_id: 'table:1', preservation: 'preserve-verbatim', message: 'first-row effects' })
      },
      (request: NativeDocxPaginationRequestV1) => { request.document.body.blocks[0]!.table!.width_twips = 401 },
      (request: NativeDocxPaginationRequestV1) => { const row = request.document.body.blocks[0]!.table!.rows[0]!; row.cant_split = false; row.height_twips = 400; row.height_rule = 'atLeast' },
      (request: NativeDocxPaginationRequestV1) => { request.resolved_layout.diagnostics.push({ code: 'NESTED_TABLE_PASSTHROUGH', severity: 'unsupported', scope_id: 'table:1', preservation: 'preserve-verbatim', message: 'nested table retained' }) },
    ]) {
      const request = fixture()
      mutate(request)
      const qualified = qualifyNativeDocxTablesV1(request.document, request.resolved_layout)
      expect(qualified.status).toBe('refused')
      expect(qualified.tables).toEqual([])
      expect(qualified.paragraph_widths.size).toBe(0)
    }
  })

  it('refuses table extents whose bounded components overflow when combined', () => {
    const request = fixture()
    const table = request.document.body.blocks[0]!.table!
    table.indent_twips = 20_000_000_000
    table.width_twips = 1
    table.grid_widths_twips = [1]
    for (const row of table.rows) row.cells[0]!.width_twips = 1
    const qualified = qualifyNativeDocxTablesV1(request.document, request.resolved_layout)
    expect(qualified.status).toBe('refused')
    expect(qualified.tables).toEqual([])
  })

  it('changes the canonical hash for source-order and paint-semantic changes but ignores object key insertion order', () => {
    const request = fixture()
    const base = qualifyNativeDocxTablesV1(request.document, request.resolved_layout)
    if (base.status !== 'qualified') throw new Error('fixture must qualify')
    const shaded = fixture(); shaded.document.body.blocks[0]!.table!.rows[0]!.cells[0]!.shading_rgb = 'DDEEFE'
    const border = fixture(); border.document.body.blocks[0]!.table!.borders!.top!.size_eighth_points = 9
    const reversed = fixture(); reversed.document.body.blocks[0]!.table!.rows.reverse()
    for (const changed of [shaded, border, reversed]) {
      const value = qualifyNativeDocxTablesV1(changed.document, changed.resolved_layout)
      expect(value.status === 'qualified' ? value.sha256 : '').not.toBe(base.sha256)
    }
    const reordered = structuredClone(request)
    const source = reordered.document.body.blocks[0]!.table!
    reordered.document.body.blocks[0]!.table = Object.fromEntries(Object.entries(source).reverse()) as typeof source
    const same = qualifyNativeDocxTablesV1(reordered.document, reordered.resolved_layout)
    expect(same.status === 'qualified' ? same.sha256 : '').toBe(base.sha256)
  })

  it('fails closed on cyclic, negative-zero, and resource-unbounded hash inputs', () => {
    const request = fixture()
    const qualified = qualifyNativeDocxTablesV1(request.document, request.resolved_layout)
    if (qualified.status !== 'qualified') throw new Error('fixture must qualify')
    const cyclic = structuredClone(qualified.tables) as any
    cyclic[0].table.borders.top = cyclic[0].table.borders
    expect(() => nativeDocxTableProjectionSha256V1(cyclic)).toThrow(/canonical|wire/)
    const negativeZero = structuredClone(qualified.tables) as any
    negativeZero[0].table.indent_twips = -0
    expect(() => nativeDocxTableProjectionSha256V1(negativeZero)).toThrow(/canonical|wire/)
    expect(() => nativeDocxTableProjectionSha256V1(Array.from({ length: 1_001 }, () => qualified.tables[0]!) as any)).toThrow(/bounded/)
  })

  function tableGridEvidence(table: NonNullable<NativeDocxPaginationRequestV1['document']['body']['blocks'][number]['table']>): NativeDocxAutomaticBorderEvidenceV1 {
    const path = '/w:styles[1]/w:style[5]/w:tblPr[1]/w:tblBorders[1]'
    return {
      policy: DOCX_AUTO_BORDER_POLICY, read_only: true, package_sha256: HASH, page_background: 'absent-on-white-preview', background_rgb: 'FFFFFF',
      source_part: 'word/styles.xml', source_path: path, source_sha256: HASH,
      borders: {
        top: { style: 'single', size_eighth_points: 4, color_rgb: '000000' }, right: { style: 'single', size_eighth_points: 4, color_rgb: '000000' },
        bottom: { style: 'single', size_eighth_points: 4, color_rgb: '000000' }, left: { style: 'single', size_eighth_points: 4, color_rgb: '000000' },
        inside_horizontal: { style: 'single', size_eighth_points: 4, color_rgb: '000000' }, inside_vertical: { style: 'single', size_eighth_points: 4, color_rgb: '000000' },
      },
      automatic_edges: ['top', 'right', 'bottom', 'left', 'inside_horizontal', 'inside_vertical'],
      cell_ids: table.rows.flatMap((row) => row.cells.map((cell) => cell.id)),
      source_diagnostics: [{ code: 'TABLE_STYLE_EFFECTS_PRESERVED', scope_id: table.id, part_name: 'word/styles.xml', path: '/w:styles[1]/w:style[5]/w:tblPr[1]' }],
    }
  }

  function tableGridAutofitRequest(): NativeDocxPaginationRequestV1 {
    const request = fixture()
    const table = request.document.body.blocks[0]!.table!
    const page = request.document.sections[0]!.page
    page.width_twips = 11_906
    page.height_twips = 16_838
    page.orientation = 'portrait'
    page.margins = { top_twips: 1_417, right_twips: 1_417, bottom_twips: 1_417, left_twips: 1_417, header_twips: 708, footer_twips: 708, gutter_twips: 0 }
    page.column_spacing_twips = 708
    delete table.width_twips
    delete table.layout
    delete table.alignment
    delete table.indent_twips
    delete table.cell_margins
    delete table.borders
    table.table_style_id = 'TableGrid'
    table.grid_widths_twips = [1_510, 1_511]
    for (const [rowIndex, row] of table.rows.entries()) {
      delete row.cells[0]!.shading_rgb
      row.cells[0]!.width_twips = 1_510
      row.cells[0]!.paragraphs[0]!.runs = [{ id: `run:cell:${rowIndex + 1}`, anchor: anchor(`/w:document[1]/w:body[1]/w:tbl[1]/w:tr[${rowIndex + 1}]/w:tc[1]/w:r[1]`, 101 + rowIndex, 109 + rowIndex), kind: 'text', text: `${rowIndex + 1}` }]
      const emptyParagraph = structuredClone(row.cells[0]!.paragraphs[0]!)
      emptyParagraph.id = `paragraph:cell:${rowIndex + 1}:empty`
      emptyParagraph.anchor = anchor(`/w:document[1]/w:body[1]/w:tbl[1]/w:tr[${rowIndex + 1}]/w:tc[2]/w:p[1]`, 200 + rowIndex, 210 + rowIndex)
      emptyParagraph.runs = []
      row.cells.push({ id: `cell:${rowIndex + 1}:empty`, anchor: anchor(`/w:document[1]/w:body[1]/w:tbl[1]/w:tr[${rowIndex + 1}]/w:tc[2]`, 190 + rowIndex, 220 + rowIndex), width_twips: 1_511, grid_span: 1, vertical_merge: 'none', paragraphs: [emptyParagraph] })
      request.resolved_layout.paragraphs.push({ paragraph_id: emptyParagraph.id, applied_styles: [], properties: { spacing_after_twips: 0, line: 240, line_rule: 'auto' }, paragraph_mark_properties: {} })
      const shapedEmpty = structuredClone(request.shaped_lines.paragraphs[0]!)
      shapedEmpty.paragraph_id = emptyParagraph.id
      shapedEmpty.lines = [{ ...shapedEmpty.lines[0]!, id: `line:${emptyParagraph.id}:0`, fragments: [] }]
      request.shaped_lines.paragraphs.push(shapedEmpty)
      request.shaped_lines.paragraphs[rowIndex]!.lines[0]!.fragments = [{
        id: `fragment:paragraph:cell:${rowIndex + 1}:0:0`, source_kind: 'run', source_id: `run:cell:${rowIndex + 1}`, start_utf16: 0, end_utf16: 1,
        text: `${rowIndex + 1}`, direction: 'ltr', bidi_level: 0, logical_order: 0, script: 'Latn', language: 'fr-FR', whitespace: false,
        advance_inline_millipoints: 5_000, justification_expansion_millipoints: 0, ascent_millipoints: 5_000, descent_millipoints: -1_000, line_gap_millipoints: 0, glyphs: [],
      }]
    }
    request.resolved_layout.tables[0]!.style_id = 'TableGrid'
    request.resolved_layout.tables[0]!.geometry = { layout: 'autofit', alignment: 'left', indent_twips: 0, width_type: 'auto', width_value: 0, cell_margins: { top_twips: 0, right_twips: 108, bottom_twips: 0, left_twips: 108 } }
    request.resolved_layout.diagnostics.push({ code: 'TABLE_STYLE_EFFECTS_PRESERVED', severity: 'unsupported', scope_id: table.id, part_name: 'word/styles.xml', path: '/w:styles[1]/w:style[5]/w:tblPr[1]', preservation: 'preserve-verbatim', message: 'Automatic table borders are preserved' })
    const trailing = paragraph('paragraph:after', 9)
    request.document.body.blocks.push({ kind: 'paragraph', id: trailing.id, paragraph: trailing })
    request.resolved_layout.paragraphs.push({ paragraph_id: trailing.id, applied_styles: [], properties: { spacing_after_twips: 160, line: 259, line_rule: 'auto' }, paragraph_mark_properties: {} })
    return request
  }

  it('preserves authored TableGrid auto-width geometry when automatic-border evidence exact-joins', () => {
    const request = tableGridAutofitRequest()
    const table = request.document.body.blocks[0]!.table!
    request.resolved_layout.tables[0]!.automatic_border_preview = tableGridEvidence(table)
    const original = JSON.stringify(request.document)
    const qualified = qualifyNativeDocxTablesV1(request.document, request.resolved_layout, request.shaped_lines)
    expect(qualified).toMatchObject({
      status: 'qualified',
      tables: [{
        width_millipoints: 151_050,
        grid_widths_millipoints: [75_500, 75_550],
        table: { borders: { top: { style: 'single', size_eighth_points: 4, color_rgb: '000000' } } },
        width_policy: { name: 'source-preferred-nonconflicting-v1', source_grid_widths_twips: [1_510, 1_511], preferred_width_twips: null, container_width_twips: 9_072 },
      }],
    })
    expect(JSON.stringify(request.document)).toBe(original)
    expect(table.borders).toBeUndefined()
    expect(table.layout).toBeUndefined()
    expect(table.width_twips).toBeUndefined()
  })

  it('refuses TableGrid auto-width paint without joining automatic-border evidence, merges, or unsatisfied minima', () => {
    for (const mode of ['missing-preview', 'forged-cells', 'merge', 'wide-word', 'conditional'] as const) {
      const request = tableGridAutofitRequest()
      const table = request.document.body.blocks[0]!.table!
      const evidence = tableGridEvidence(table)
      if (mode !== 'missing-preview') request.resolved_layout.tables[0]!.automatic_border_preview = evidence
      if (mode === 'forged-cells') evidence.cell_ids = ['cell:other']
      if (mode === 'merge') table.rows[0]!.cells[0]!.vertical_merge = 'restart'
      if (mode === 'wide-word') request.shaped_lines.paragraphs[0]!.lines[0]!.fragments[0]!.advance_inline_millipoints = 1_000_000_000
      if (mode === 'conditional') request.resolved_layout.diagnostics.push({ code: 'CONDITIONAL_TABLE_STYLE_PRESERVED', severity: 'unsupported', scope_id: table.id, part_name: 'word/styles.xml', path: '/w:styles[1]/w:style[5]', preservation: 'preserve-verbatim', message: 'first-row effects' })
      expect(qualifyNativeDocxTablesV1(request.document, request.resolved_layout, request.shaped_lines)).toMatchObject({ status: 'refused', tables: [] })
    }
  })

  // tblr-height.docx names the table style `Tabellengitternetz`, which its own
  // styles.xml does not define. ECMA-376 17.7.2 binds w:tblStyle to the w:style
  // whose w:styleId it names, so that reference selects nothing and the table
  // the source states in full has to paint. A style that does exist but whose
  // effects this tier cannot reproduce keeps blocking.
  it('paints a table whose style reference resolves to nothing and still refuses a style that hides effects', () => {
    const dangling = (): NativeDocxPaginationRequestV1 => {
      const request = fixture()
      request.document.body.blocks[0]!.table!.table_style_id = 'Tabellengitternetz'
      request.resolved_layout.source_parts.styles_part = 'word/styles.xml'
      request.resolved_layout.tables[0]!.style_id = 'Tabellengitternetz'
      request.resolved_layout.diagnostics.push({ code: 'MISSING_TABLE_STYLE', severity: 'unsupported', scope_id: 'table:1', part_name: 'word/styles.xml', preservation: 'preserve-verbatim', message: 'The referenced table style is missing and was not guessed' })
      return request
    }
    const qualify = (request: NativeDocxPaginationRequestV1) => qualifyNativeDocxTablesV1(request.document, request.resolved_layout, request.shaped_lines, true)
    expect(qualify(dangling())).toMatchObject({ status: 'qualified' })

    // Every condition that proves the absence carried no formatting is load-bearing.
    for (const mutate of [
      (request: NativeDocxPaginationRequestV1) => { request.resolved_layout.tables[0]!.borders = { top: { style: 'single' as const, size_eighth_points: 8, color_rgb: '112233' } } },
      (request: NativeDocxPaginationRequestV1) => { request.resolved_layout.tables[0]!.cell_shading_rgb = 'FFFF00' },
      (request: NativeDocxPaginationRequestV1) => { request.resolved_layout.tables[0]!.geometry = { layout: 'fixed', alignment: 'left', width_type: 'dxa', width_value: 400, indent_twips: 0, cell_margins: { top_twips: 10, right_twips: 10, bottom_twips: 10, left_twips: 10 } } },
      (request: NativeDocxPaginationRequestV1) => { delete request.resolved_layout.tables[0]!.style_id },
      (request: NativeDocxPaginationRequestV1) => { request.resolved_layout.diagnostics[0]!.part_name = 'word/document.xml' },
      (request: NativeDocxPaginationRequestV1) => { delete request.resolved_layout.diagnostics[0]!.part_name },
      (request: NativeDocxPaginationRequestV1) => { request.resolved_layout.diagnostics[0]!.path = '/w:styles[1]/w:style[5]' },
    ]) {
      const request = dangling()
      mutate(request)
      expect(qualify(request).status).toBe('refused')
    }

    // The same table style existing but carrying unreproducible effects still refuses.
    const effects = dangling()
    effects.resolved_layout.diagnostics.push({ code: 'TABLE_STYLE_EFFECTS_PRESERVED', severity: 'unsupported', scope_id: 'table:1', part_name: 'word/styles.xml', path: '/w:styles[1]/w:style[5]/w:tblPr[1]', preservation: 'preserve-verbatim', message: 'Table-style effects are preserved' })
    expect(qualify(effects).status).toBe('refused')

    // Strict qualification keeps its blanket rule: any resolved-layout
    // diagnostic touching a table still refuses exact paint.
    expect(qualifyNativeDocxTablesV1(dangling().document, dangling().resolved_layout).status).toBe('refused')
  })

  // The same absence reaches the compiler's blocking-diagnostic gate as well,
  // and twice: `resolveTableGeometry` walks the style chain for the geometry
  // the style might have carried, and the first hop of that walk is the named
  // style itself, so its absence is reported a second time in basedOn terms.
  // There is no ancestor and no dropped layer, so neither report blocks a page.
  it('reads a dangling table style reference and its restated basedOn report as render-neutral', () => {
    const resolved = fixture().resolved_layout
    resolved.source_parts.styles_part = 'word/styles.xml'
    resolved.tables[0]!.style_id = 'Tabellengitternetz'
    const missingStyle = { code: 'MISSING_TABLE_STYLE', severity: 'unsupported' as const, scope_id: 'table:1', part_name: 'word/styles.xml', preservation: 'preserve-verbatim' as const, message: 'The referenced table style is missing and was not guessed' }
    const missingAncestor = { ...missingStyle, code: 'MISSING_STYLE_REFERENCE', message: 'The missing basedOn ancestor was ignored; available descendant layers were retained' }
    resolved.diagnostics.push(missingStyle, missingAncestor)
    expect(isRenderNeutralLayoutDiagnostic(missingStyle, resolved)).toBe(true)
    expect(isRenderNeutralLayoutDiagnostic(missingAncestor, resolved)).toBe(true)

    // A missing basedOn ancestor of a style that does exist is a dropped layer,
    // not an unresolvable reference, and keeps blocking.
    const ancestorOnly = fixture().resolved_layout
    ancestorOnly.source_parts.styles_part = 'word/styles.xml'
    ancestorOnly.tables[0]!.style_id = 'TableGrid'
    ancestorOnly.diagnostics.push(missingAncestor)
    expect(isRenderNeutralLayoutDiagnostic(missingAncestor, ancestorOnly)).toBe(false)

    // Any resolved style effect on the table disqualifies both exemptions.
    const shaded = fixture().resolved_layout
    shaded.source_parts.styles_part = 'word/styles.xml'
    shaded.tables[0]!.style_id = 'Tabellengitternetz'
    shaded.tables[0]!.cell_shading_rgb = 'FFFF00'
    shaded.diagnostics.push(missingStyle, missingAncestor)
    expect(isRenderNeutralLayoutDiagnostic(missingStyle, shaded)).toBe(false)
    expect(isRenderNeutralLayoutDiagnostic(missingAncestor, shaded)).toBe(false)
  })

  // The conditional table-style cascade (w:tblStylePr selected by w:tblLook) is
  // resolved per cell by the Go resolver and carried as conditional_cell_shading.
  // A cell's own w:shd still wins; the style's whole-table fill sits below.
  it('paints the fill the conditional table-style cascade resolved for a cell', () => {
    const request = fixture()
    request.resolved_layout.tables[0]!.cell_shading_rgb = 'FFFF00'
    request.resolved_layout.tables[0]!.conditional_cell_shading = [{ cell_id: 'cell:2', shading_rgb: '833C0B' }, { cell_id: 'cell:1', shading_rgb: '7F7F7F' }]
    expect(decodeNativeDocxResolvedLayout(request.resolved_layout).ok).toBe(true)
    const qualified = qualifyNativeDocxTablesV1(request.document, request.resolved_layout)
    if (qualified.status !== 'qualified') throw new Error(`fixture must qualify: ${JSON.stringify(qualified)}`)
    expect(qualified.tables[0]!.rows.map((row) => row.cells[0]!.cell.shading_rgb)).toEqual(['DDEEFF', '833C0B'])
    const geometry = layoutNativeDocxTableRowsV1(qualified.tables[0]!, request.shaped_lines)
    expect(geometry?.map((row) => row.cells[0]!.shading_rgb)).toEqual(['DDEEFF', '833C0B'])
    // A cell the cascade leaves unfilled still takes the whole-table fill.
    const unfilled = fixture()
    unfilled.resolved_layout.tables[0]!.cell_shading_rgb = 'FFFF00'
    unfilled.resolved_layout.tables[0]!.conditional_cell_shading = [{ cell_id: 'cell:1', shading_rgb: '7F7F7F' }]
    const fallback = qualifyNativeDocxTablesV1(unfilled.document, unfilled.resolved_layout)
    expect(fallback.status === 'qualified' ? fallback.tables[0]!.rows[1]!.cells[0]!.cell.shading_rgb : undefined).toBe('FFFF00')
  })

  it('validates conditional cell shading as bounded exact-key entries with unique cell ids and RGB fills', () => {
    for (const [entries, ok] of [
      [[{ cell_id: 'cell:1', shading_rgb: '833C0B' }], true],
      [[], true],
      [[{ cell_id: 'cell:1', shading_rgb: '833C0B' }, { cell_id: 'cell:1', shading_rgb: 'FF0000' }], false],
      [[{ cell_id: 'cell:1', shading_rgb: '833c0b' }], false],
      [[{ cell_id: 'cell:1', shading_rgb: '833C0B', extra: true }], false],
      [[{ cell_id: 'cell:1' }], false],
      [[{ shading_rgb: '833C0B' }], false],
      ['833C0B', false],
    ] as const) {
      const resolved = structuredClone(fixture().resolved_layout) as NativeDocxResolvedLayoutInputV1
      ;(resolved.tables[0] as unknown as Record<string, unknown>).conditional_cell_shading = structuredClone(entries)
      expect(decodeNativeDocxResolvedLayout(resolved).ok, JSON.stringify(entries)).toBe(ok)
    }
    // A resolved conditional fill is a style effect, so a dangling-style
    // exemption that requires the style to have contributed nothing no longer holds.
    const resolved = fixture().resolved_layout
    resolved.source_parts.styles_part = 'word/styles.xml'
    resolved.tables[0]!.style_id = 'Tabellengitternetz'
    resolved.tables[0]!.conditional_cell_shading = [{ cell_id: 'cell:1', shading_rgb: '7F7F7F' }]
    const missingStyle = { code: 'MISSING_TABLE_STYLE', severity: 'unsupported' as const, scope_id: 'table:1', part_name: 'word/styles.xml', preservation: 'preserve-verbatim' as const, message: 'The referenced table style is missing and was not guessed' }
    resolved.diagnostics.push(missingStyle)
    expect(isRenderNeutralLayoutDiagnostic(missingStyle, resolved)).toBe(false)
  })

  it('validates conditional cell borders as four resolved edges per unique cell and treats them as a style effect', () => {
    const edge = { style: 'single' as const, size_eighth_points: 12, color_rgb: 'F4B083' }
    for (const [entries, ok] of [
      [[{ cell_id: 'cell:1', borders: { top: { style: 'none', size_eighth_points: 0 }, bottom: edge } }], true],
      [[{ cell_id: 'cell:1', borders: {} }], true],
      [[{ cell_id: 'cell:1', borders: { top: edge } }, { cell_id: 'cell:1', borders: { top: edge } }], false],
      [[{ cell_id: 'cell:1', borders: { inside_horizontal: edge } }], false],
      [[{ cell_id: 'cell:1', borders: { top: { style: 'double', size_eighth_points: 4, color_rgb: 'F4B083' } } }], false],
      [[{ cell_id: 'cell:1' }], false],
      [[{ borders: { top: edge } }], false],
      ['top', false],
    ] as const) {
      const resolved = structuredClone(fixture().resolved_layout) as NativeDocxResolvedLayoutInputV1
      ;(resolved.tables[0] as unknown as Record<string, unknown>).conditional_cell_borders = structuredClone(entries)
      expect(decodeNativeDocxResolvedLayout(resolved).ok, JSON.stringify(entries)).toBe(ok)
    }
    const resolved = fixture().resolved_layout
    resolved.source_parts.styles_part = 'word/styles.xml'
    resolved.tables[0]!.style_id = 'Tabellengitternetz'
    resolved.tables[0]!.conditional_cell_borders = [{ cell_id: 'cell:1', borders: { top: edge } }]
    const missingStyle = { code: 'MISSING_TABLE_STYLE', severity: 'unsupported' as const, scope_id: 'table:1', part_name: 'word/styles.xml', preservation: 'preserve-verbatim' as const, message: 'The referenced table style is missing and was not guessed' }
    resolved.diagnostics.push(missingStyle)
    expect(isRenderNeutralLayoutDiagnostic(missingStyle, resolved)).toBe(false)
  })
})
