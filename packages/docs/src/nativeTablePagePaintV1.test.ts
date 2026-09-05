import { describe, expect, it } from 'vitest'
import { BIDI_UNICODE_VERSION, NATIVE_BIDI_PROVIDER_ID, NATIVE_BIDI_PROVIDER_REVISION } from '@injoffice/font-metrics/bidi'
import { UNICODE_13_CLASSIFIER_REVISION } from '@injoffice/font-metrics/unicode13'
import { DOCX_NATIVE_PROTOCOL, DOCX_NATIVE_VERSION, type NativeDocxDocumentV1, type NativeDocxParagraphV1 } from './nativeContract.js'
import { DOCX_RESOLVED_LAYOUT_PROTOCOL, DOCX_RESOLVED_LAYOUT_VERSION, type NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'
import { DOCX_SHAPED_LINES_PROTOCOL, DOCX_SHAPED_LINES_VERSION, type NativeDocxShapedLinesV1 } from './nativeShapingLines.js'
import { DOCX_DEFAULT_TAB_STOP_TWIPS, DOCX_PAGINATION_SETTINGS_PROTOCOL, DOCX_PAGINATION_SETTINGS_VERSION } from './nativePaginationSettings.js'
import { DOCX_PAGINATION_REQUEST_PROTOCOL, DOCX_PAGINATION_REQUEST_VERSION, paginateNativeDocxV1, type NativeDocxPaginationRequestV1 } from './nativePaginationV1.js'
import { layoutNativeDocxTableRowsV1, nativeDocxTableProjectionSha256V1, qualifyNativeDocxTablesV1 } from './nativeTablePagePaintV1.js'

const HASH = `sha256:${'a'.repeat(64)}`
const anchor = (path: string, start: number, end: number) => ({ part_name: 'word/document.xml', path, start_byte: start, end_byte: end, xml_sha256: HASH })
const policy = { mode: 'read-only' as const, allowed_operations: [] as const, refusal: { code: 'NATIVE_READ_ONLY', message: 'Source remains authoritative.', preservation: 'refuse-mutation' as const } }

function paragraph(id: string, ordinal: number): NativeDocxParagraphV1 {
  return { id, anchor: anchor(`/w:document[1]/w:body[1]/w:tbl[1]/w:tr[${ordinal + 1}]/w:tc[1]/w:p[1]`, 100 + ordinal * 40, 110 + ordinal * 40), edit_policy: { ...policy, allowed_operations: [] }, properties: {}, runs: [] }
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

describe('bounded native DOCX table page-paint geometry', () => {
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
      (request: NativeDocxPaginationRequestV1) => { request.document.body.blocks[0]!.table!.rows[0]!.cant_split = false },
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
})
