import { readFileSync } from 'node:fs'
import { createHash } from 'node:crypto'
import { createRequire } from 'node:module'
import { describe, expect, it } from 'vitest'
import type { NativeFontManifest } from '@injoffice/font-metrics/layout'
import { createHarfBuzzOutlineProviderV1 } from '@injoffice/font-metrics/harfbuzz'
import { DOCX_NATIVE_PROTOCOL, DOCX_NATIVE_VERSION, type NativeDocxDocumentV1, type NativeDocxTableV1 } from './nativeContract.js'
import { DOCX_RESOLVED_LAYOUT_PROTOCOL, DOCX_RESOLVED_LAYOUT_VERSION, type NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'
import { DOCX_PAGINATION_SETTINGS_PROTOCOL, DOCX_PAGINATION_SETTINGS_VERSION, type NativeDocxPaginationSettingsV1 } from './nativePaginationSettings.js'
import { DOCX_PAGE_PAINT_COMPILER_PROTOCOL, DOCX_PAGE_PAINT_COMPILER_VERSION, prepareNativeDocxPagePaintV1, renderNativeDocxApproximatePagePreviewV1, decodeNativeDocxApproximatePagePreviewV1, type NativeDocxPagePaintPrepareInputV1 } from './nativePagePaintCompilerV1.js'
import { encodeNativeDOCXFontInventoryV1, nativeDOCXCanonicalWireSHA256V1, type NativeDOCXFontInventoryV1 } from './nativeFontInventoryV1.js'
import { decodeNativeDocxApproximateNestedTablesV1, DOCX_APPROXIMATE_NESTED_TABLE_CODE, DOCX_APPROXIMATE_NESTED_TABLE_FONT_CODE, DOCX_APPROXIMATE_NESTED_TABLE_LAYOUT_POLICY, DOCX_APPROXIMATE_NESTED_TABLE_OMITTED_CODE, DOCX_APPROXIMATE_NESTED_TABLE_SIDECAR_REFUSED, DOCX_APPROXIMATE_NESTED_TABLE_TABLE_ID, DOCX_APPROXIMATE_NESTED_TABLE_WARNING, type NativeDocxApproximateNestedTablesV1 } from './nativeApproximateNestedTablesV1.js'
import type { NativeDocxFillTableCellCommandV1, NativeDocxGlyphOutlineRequestV1, NativeDocxStrokeTableBorderCommandV1 } from './nativePagePaintV1.js'

const require = createRequire(import.meta.url)
const FONT_BYTES = new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
const FONT_DIGEST = 'sha256:7da195a74c55bef988d0d48f9508bd5d849425c1770dba5d7bfc6ce9ed848954' as const
const HASH = `sha256:${'a'.repeat(64)}`
const RELATIONSHIPS_HASH = `sha256:${'b'.repeat(64)}`
const SETTINGS_PART = 'word/settings.xml'
const RELATIONSHIPS_PART = 'word/_rels/document.xml.rels'
const REVISION = `rev:${'a'.repeat(32)}`
const POLICY = { mode: 'read-only' as const, allowed_operations: [], refusal: { code: 'NATIVE_READ_ONLY', message: 'Fixture is immutable.', preservation: 'refuse-mutation' as const } }
const OUTER = '/w:document[1]/w:body[1]/w:tbl[1]'
const CELL = `${OUTER}/w:tr[1]/w:tc[1]`
const NESTED = `${CELL}/w:tbl[1]`
const ITEM = 'approximate-nested-table:1:1'

function storedFontDigest(bytes: Uint8Array): string {
  const stored = Uint8Array.from(bytes)
  const key = Uint8Array.from([0xff, 0xee, 0xdd, 0xcc, 0xbb, 0xaa, 0x99, 0x88, 0x77, 0x66, 0x55, 0x44, 0x33, 0x22, 0x11, 0x00])
  for (let index = 0; index < 32 && index < stored.length; index++) stored[index] = stored[index]! ^ key[index % 16]!
  return `sha256:${createHash('sha256').update(stored).digest('hex')}`
}
function anchor(path: string, start: number, end: number) { return { part_name: 'word/document.xml', path, start_byte: start, end_byte: end, xml_sha256: HASH } }
const text = (id: string, path: string, start: number, end: number, value: string) => ({ kind: 'text' as const, id, anchor: anchor(path, start, end), text: value })

/** Outer fixed table whose single cell holds paragraph A, a refused nested
 * table, and paragraph B; embedded DejaVu Sans; legacy-eligible settings. */
function fixture(): { input: NativeDocxPagePaintPrepareInputV1; eligibility: Record<string, unknown> } {
  const a = { id: 'paragraph:a', anchor: anchor(`${CELL}/w:p[1]`, 100, 190), edit_policy: POLICY, properties: {}, runs: [text('run:a', `${CELL}/w:p[1]/w:r[1]`, 110, 180, 'Above')] }
  const b = { id: 'paragraph:b', anchor: anchor(`${CELL}/w:p[2]`, 700, 800), edit_policy: POLICY, properties: {}, runs: [text('run:b', `${CELL}/w:p[2]/w:r[1]`, 710, 790, 'Below')] }
  const table: NativeDocxTableV1 = {
    id: 'table:1', anchor: anchor(OUTER, 90, 900), edit_policy: POLICY,
    width_twips: 9_360, layout: 'fixed', alignment: 'left', indent_twips: 0, grid_widths_twips: [9_360],
    cell_margins: { top_twips: 100, right_twips: 100, bottom_twips: 100, left_twips: 100 },
    borders: { top: { style: 'single', size_eighth_points: 8, color_rgb: '000000' }, right: { style: 'single', size_eighth_points: 8, color_rgb: '000000' }, bottom: { style: 'single', size_eighth_points: 8, color_rgb: '000000' }, left: { style: 'single', size_eighth_points: 8, color_rgb: '000000' } },
    rows: [{ id: 'row:1', anchor: anchor(`${OUTER}/w:tr[1]`, 95, 895), repeat_header: false, cant_split: true, cells: [{ id: 'cell:1', anchor: anchor(CELL, 98, 892), width_twips: 9_360, grid_span: 1, vertical_merge: 'none', paragraphs: [a, b] }] }],
  }
  const document: NativeDocxDocumentV1 = {
    protocol: DOCX_NATIVE_PROTOCOL, version: DOCX_NATIVE_VERSION, document_id: 'document:test', revision: REVISION,
    source: { package_sha256: HASH, main_part: 'word/document.xml' },
    body: { id: 'story:body', kind: 'body', part_name: 'word/document.xml', anchor: anchor('/w:document[1]/w:body[1]', 1, 3_000), blocks: [{ kind: 'table', id: table.id, table }] },
    sections: [{
      id: 'section:1', anchor: anchor('/w:document[1]/w:body[1]/w:sectPr[1]', 2_100, 2_190), starts_at_block_id: table.id, break_type: 'next-page', title_page: false,
      page: { width_twips: 12_240, height_twips: 15_840, orientation: 'portrait', margins: { top_twips: 1_440, right_twips: 1_440, bottom_twips: 1_440, left_twips: 1_440, header_twips: 720, footer_twips: 720, gutter_twips: 0 }, columns: 1, column_spacing_twips: 720, column_layout: 'equal-width', column_definitions: [{ id: 'column:section:1:0', ordinal: 0 }] },
      header_refs: [], footer_refs: [],
    }],
    headers: [], footers: [], notes: [], comment_stories: [], comments: [], capabilities: [],
    passthrough_parts: [
      { part_name: SETTINGS_PART, content_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml', byte_length: 1, sha256: HASH, policy: 'preserve-verbatim' },
      { part_name: RELATIONSHIPS_PART, content_type: 'application/vnd.openxmlformats-package.relationships+xml', byte_length: 1, sha256: RELATIONSHIPS_HASH, policy: 'preserve-verbatim' },
      { part_name: 'word/fontTable.xml', content_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml', byte_length: 1, sha256: HASH, policy: 'preserve-verbatim' },
      { part_name: 'word/_rels/fontTable.xml.rels', content_type: 'application/vnd.openxmlformats-package.relationships+xml', byte_length: 1, sha256: RELATIONSHIPS_HASH, policy: 'preserve-verbatim' },
      { part_name: 'word/fonts/DejaVuSans.odttf', content_type: 'application/vnd.openxmlformats-officedocument.obfuscatedFont', byte_length: FONT_BYTES.byteLength, sha256: storedFontDigest(FONT_BYTES), policy: 'preserve-verbatim' },
    ],
    unsupported: [{ id: 'unsupported:nested', code: 'NESTED_TABLE_OR_CELL_MARKUP', capability: 'table-structure', scope_id: table.id, anchor: anchor(NESTED, 200, 600), preservation: 'refuse-mutation', message: 'Only direct cell paragraphs are modeled; nested content is preserved verbatim' }],
  }
  const mark = { font_family: 'DejaVu Sans', font_size_half_points: 20 }
  const resolved: NativeDocxResolvedLayoutInputV1 = {
    protocol: DOCX_RESOLVED_LAYOUT_PROTOCOL, version: DOCX_RESOLVED_LAYOUT_VERSION, document_id: document.document_id, revision: document.revision,
    source_parts: { main_part: 'word/document.xml', font_table_part: 'word/fontTable.xml' },
    paragraphs: [
      { paragraph_id: a.id, applied_styles: [], properties: { spacing_after_twips: 40 }, paragraph_mark_properties: mark },
      { paragraph_id: b.id, applied_styles: [], properties: {}, paragraph_mark_properties: mark },
    ],
    runs: [
      { run_id: 'run:a', paragraph_id: a.id, applied_paragraph_styles: [], applied_character_styles: [], properties: { ...mark, color: '123456' } },
      { run_id: 'run:b', paragraph_id: b.id, applied_paragraph_styles: [], applied_character_styles: [], properties: { ...mark, color: '123456' } },
    ],
    tables: [{ table_id: table.id }], fonts: [{ name: 'DejaVu Sans' }], diagnostics: [],
  }
  const settings: NativeDocxPaginationSettingsV1 = {
    protocol: DOCX_PAGINATION_SETTINGS_PROTOCOL, version: DOCX_PAGINATION_SETTINGS_VERSION, document_id: document.document_id, revision: document.revision, package_sha256: HASH, main_part: 'word/document.xml',
    relationships_part: RELATIONSHIPS_PART, relationships_sha256: RELATIONSHIPS_HASH, relationship_id: 'rIdSettings', settings_part: SETTINGS_PART, settings_sha256: HASH,
    profile: 'unsupported', default_tab_stop_twips: 720, mirror_margins: false, gutter_at_top: false, even_and_odd_headers: false,
    diagnostics: [{ code: 'COMPATIBILITY_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: SETTINGS_PART, path: '/w:settings[1]/w:compat[1]', preservation: 'preserve-verbatim', message: 'Legacy mode' }],
  }
  const manifest: NativeFontManifest = { version: 1, manifestId: `docx.fonts.${createHash('sha256').update(document.document_id).digest('hex').slice(0, 24)}`, revision: document.revision, faces: [{ faceId: 'font-face:dejavu', family: 'DejaVu Sans', weight: 400, style: 'normal', stretch: 100, source: { kind: 'document', resourceId: `font:${FONT_DIGEST}`, contentDigest: FONT_DIGEST } }], fallbackChains: [] }
  const inventory: NativeDOCXFontInventoryV1 = {
    protocol: 'injoffice.docx.font-inventory', version: 1, document_id: document.document_id, revision: document.revision, package_sha256: HASH, main_part: document.source.main_part, main_sha256: HASH, inventory_sha256: HASH,
    font_table: { part_name: 'word/fontTable.xml', sha256: HASH, main_relationships_part: RELATIONSHIPS_PART, main_relationships_sha256: RELATIONSHIPS_HASH, relationship_id: 'rIdFontTable', relationship_type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable', relationship_target: 'fontTable.xml', font_relationships_part: 'word/_rels/fontTable.xml.rels', font_relationships_sha256: RELATIONSHIPS_HASH },
    families: [{ family_id: 'font-family:dejavu', name: 'DejaVu Sans', faces: [{
      face_id: 'font-face:dejavu', family: 'DejaVu Sans', weight: 400, style: 'normal', stretch: 100,
      source: {
        kind: 'document', face_slot: 'embedRegular', font_table_part: 'word/fontTable.xml', font_table_path: '/w:fonts[1]/w:font[1]/w:embedRegular[1]',
        relationships_part: 'word/_rels/fontTable.xml.rels', relationships_sha256: RELATIONSHIPS_HASH,
        relationship_id: 'rIdFont', relationship_type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font', relationship_target: 'fonts/DejaVuSans.odttf',
        asset_part: 'word/fonts/DejaVuSans.odttf', asset_content_type: 'application/vnd.openxmlformats-officedocument.obfuscatedFont',
        stored_byte_length: FONT_BYTES.byteLength, stored_sha256: storedFontDigest(FONT_BYTES), content_sha256: FONT_DIGEST, resource_id: `font:${FONT_DIGEST}`,
        obfuscation: { algorithm: 'ecma-376-font-obfuscation', font_key: '{00112233-4455-6677-8899-AABBCCDDEEFF}', subsetted: false },
        licensing: { embedding_origin: 'document-package', rights_source: 'sfnt-os2-fstype', rights_status: 'verified', embedding_rights: 'installable', no_subsetting: false, allowed_scope: 'document-only' },
      },
    }] }],
    references: [{ family: 'DejaVu Sans', weight: 400, style: 'normal', scope_ids: ['paragraph:a', 'paragraph:b', 'run:a', 'run:b'] }],
    native_text_manifest: manifest,
    native_text_manifest_sha256: nativeDOCXCanonicalWireSHA256V1(manifest),
  }
  inventory.inventory_sha256 = nativeDOCXCanonicalWireSHA256V1({ ...structuredClone(inventory), inventory_sha256: '' })
  const input: NativeDocxPagePaintPrepareInputV1 = {
    protocol: DOCX_PAGE_PAINT_COMPILER_PROTOCOL, version: DOCX_PAGE_PAINT_COMPILER_VERSION, source_revision: 'git:integration-test',
    outline_provider: { provider_id: 'injoffice.sfnt-outline', provider_revision: 'sfnt-v1' },
    document, resolved_layout: resolved, pagination_settings: settings, font_inventory_json: encodeNativeDOCXFontInventoryV1(inventory),
    font_assets: [{ face_id: 'font-face:dejavu', face_slot: 'embedRegular', resource_id: `font:${FONT_DIGEST}`, content_digest: FONT_DIGEST, collection_index: null, bytes: FONT_BYTES }],
    media_assets: [],
  }
  const eligibility = { protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: settings.document_id, revision: settings.revision, package_sha256: HASH, settings_sha256: settings.settings_sha256, status: 'eligible', legacy_compatibility_mode: 12, reasons: ['Legacy mode'] }
  return { input, eligibility }
}

/** Sidecar mirroring the benchmark markup: 2 x 2 inner table styled with grey borders, a firstRow fill and a bold header cell. */
function sidecar(overrides: Record<string, unknown> = {}, family = 'DejaVu Sans'): NativeDocxApproximateNestedTablesV1 {
  const mark = { font_family: 'DejaVu Sans', font_size_half_points: 20 }
  const cell = (row: number, column: number, start: number, value: string, bold = false) => {
    const id = `${ITEM}:r${row}c${column}`, path = `${NESTED}/w:tr[${row + 1}]/w:tc[${column + 1}]`
    return {
      cell: { id, anchor: anchor(path, start, start + 80), width_twips: column === 0 ? 3_000 : 4_000, grid_span: 1, vertical_merge: 'none' as const, paragraphs: [{ id: `${id}:p0`, anchor: anchor(`${path}/w:p[1]`, start + 2, start + 78), edit_policy: POLICY, properties: {}, runs: [text(`${id}:p0:r0`, `${path}/w:p[1]/w:r[1]`, start + 4, start + 76, value)] }] },
      resolved_paragraph: { paragraph_id: `${id}:p0`, applied_styles: [], properties: { spacing_before_twips: 60, spacing_after_twips: 60 }, paragraph_mark_properties: { ...mark, ...(bold ? { bold: true } : {}) } },
      resolved_run: { run_id: `${id}:p0:r0`, paragraph_id: `${id}:p0`, applied_paragraph_styles: [], applied_character_styles: [], properties: { ...mark, font_family: family, ...(bold ? { bold: true } : {}) } },
    }
  }
  const cells = [cell(0, 0, 212, 'Setting', true), cell(0, 1, 302, 'Instructions', true), cell(1, 0, 402, 'Memory'), cell(1, 1, 492, 'We recommend that you use Dynamic memory.')]
  const border = (size: number) => ({ style: 'single' as const, size_eighth_points: size, color_rgb: '808080' })
  const item = {
    id: ITEM, table_id: 'table:1', cell_id: 'cell:1', diagnostic_ids: ['unsupported:nested'], anchor: anchor(NESTED, 200, 600), preceding_paragraphs: 1, status: 'supported',
    table: {
      id: ITEM, anchor: anchor(NESTED, 200, 600), edit_policy: POLICY, table_style_id: 'Inner', grid_widths_twips: [3_000, 4_000],
      rows: [
        { id: `${ITEM}:r0`, anchor: anchor(`${NESTED}/w:tr[1]`, 210, 390), repeat_header: false, cant_split: false, cells: [cells[0]!.cell, cells[1]!.cell] },
        { id: `${ITEM}:r1`, anchor: anchor(`${NESTED}/w:tr[2]`, 400, 590), repeat_header: false, cant_split: false, cells: [cells[2]!.cell, cells[3]!.cell] },
      ],
    },
    geometry: { layout: 'autofit', alignment: 'left', indent_twips: 360, width_type: 'auto', width_value: 0, cell_margins: { top_twips: 0, right_twips: 86, bottom_twips: 0, left_twips: 86 } },
    style_borders: { top: border(12), left: border(12), bottom: border(12), right: border(12), inside_horizontal: border(6), inside_vertical: border(6) },
    first_row_cell_shading_rgb: 'D9D9D9',
    resolved_paragraphs: cells.map(entry => entry.resolved_paragraph), resolved_runs: cells.map(entry => entry.resolved_run), omitted_runs: 0,
    notes: ['firstRow conditional table-style region applied to the first row (paragraph, run and cell fill properties); its borders and other regions are not applied'],
    ...overrides,
  }
  // The Go sidecar omits empty fields; JSON wire never carries undefined.
  const wire = Object.fromEntries(Object.entries(item).filter(([, value]) => value !== undefined))
  return { protocol: 'injoffice.docx.approximate-nested-tables', version: 1, policy: 'docx.approximate-nested-table-preview-v1', package_sha256: HASH, part_sha256: HASH, items: [wire as never], omitted_count: 0 }
}

function outlineProvider(input: NativeDocxPagePaintPrepareInputV1) {
  const outlines = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
  return { providerId: input.outline_provider.provider_id, providerRevision: input.outline_provider.provider_revision, getGlyphOutline(request: NativeDocxGlyphOutlineRequestV1) { const o = outlines.outline(request.glyph_id); return o.path.length ? { status: 'outlined' as const, ...request, ...o } : { status: 'empty' as const, ...request, units_per_em: o.units_per_em } } }
}

describe('approximate nested tables', () => {
  it('validates the sidecar against the containing cell and its retained refusal', () => {
    const { input } = fixture()
    const document = input.document as NativeDocxDocumentV1
    expect(decodeNativeDocxApproximateNestedTablesV1(sidecar(), document).items).toHaveLength(1)
    expect(() => decodeNativeDocxApproximateNestedTablesV1({ ...sidecar(), package_sha256: `sha256:${'c'.repeat(64)}` }, document)).toThrow(/exact-join/)
    expect(() => decodeNativeDocxApproximateNestedTablesV1(sidecar({ diagnostic_ids: ['unsupported:other'] }), document)).toThrow(/retained nested-table refusal/)
    expect(() => decodeNativeDocxApproximateNestedTablesV1(sidecar({ anchor: anchor(NESTED, 20, 60) }), document)).toThrow(/containing cell/)
    expect(() => decodeNativeDocxApproximateNestedTablesV1(sidecar({ preceding_paragraphs: 2 }), document)).toThrow(/paragraph order/)
    expect(() => decodeNativeDocxApproximateNestedTablesV1(sidecar({ table: { ...sidecar().items[0]!.table, id: 'table:1' } }), document)).toThrow(/bounded grid/)
    const merged = sidecar()
    ;(merged.items[0]!.table!.rows[0]!.cells[0]! as { grid_span: number }).grid_span = 2
    expect(() => decodeNativeDocxApproximateNestedTablesV1(merged, document)).toThrow(/supported subset/)
    expect(() => decodeNativeDocxApproximateNestedTablesV1(sidecar({ status: 'omitted', reason: 'nested-depth-limit', table: undefined, geometry: undefined, style_borders: undefined, first_row_cell_shading_rgb: undefined }), document)).not.toThrow()
  })

  it('paints the inner table inside the outer cell with borders, fills, text and the declared policy, and completes the disclosure', async () => {
    const { input, eligibility } = fixture()
    const before = structuredClone(input)
    const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input), { nestedTables: sidecar() })
    expect(input).toEqual(before)
    expect(paint.status).toBe('painted')
    expect(decodeNativeDocxApproximatePagePreviewV1(paint).ok).toBe(true)
    const commands = paint.pages[0]!.commands
    const outerStrokes = commands.filter((c): c is NativeDocxStrokeTableBorderCommandV1 => c.kind === 'stroke_table_border' && c.table_id === 'table:1')
    expect(outerStrokes.length).toBeGreaterThan(0)
    const outerLeft = Math.min(...outerStrokes.map(c => Math.min(c.x1_millipoints, c.x2_millipoints))), outerRight = Math.max(...outerStrokes.map(c => Math.max(c.x1_millipoints, c.x2_millipoints)))
    const outerTop = Math.min(...outerStrokes.map(c => Math.min(c.y1_millipoints, c.y2_millipoints))), outerBottom = Math.max(...outerStrokes.map(c => Math.max(c.y1_millipoints, c.y2_millipoints)))
    const nestedStrokes = commands.filter((c): c is NativeDocxStrokeTableBorderCommandV1 => c.kind === 'stroke_table_border' && c.table_id === DOCX_APPROXIMATE_NESTED_TABLE_TABLE_ID)
    const nestedFills = commands.filter((c): c is NativeDocxFillTableCellCommandV1 => c.kind === 'fill_table_cell' && c.table_id === DOCX_APPROXIMATE_NESTED_TABLE_TABLE_ID)
    // 2 x 2 cells: shared inner edges are stroked once (as the right/bottom edge of the earlier cell), outer edges from the base style borders.
    expect(nestedStrokes).toHaveLength(12)
    expect(nestedStrokes.every(c => c.stroke_rgb === '808080' && c.width_millipoints > 0)).toBe(true)
    expect(nestedStrokes.filter(c => c.width_millipoints === 12 * 125)).toHaveLength(8)
    expect(nestedStrokes.filter(c => c.width_millipoints === 6 * 125)).toHaveLength(4)
    // The firstRow fill covers both header cells only.
    expect(nestedFills.map(c => [c.row_id, c.fill_rgb])).toEqual([[`${ITEM}:r0`, 'D9D9D9'], [`${ITEM}:r0`, 'D9D9D9']])
    // Inner boxes lie inside the outer cell content box: cell margins 100 twips = 5000 mp, inner indent 360 twips = 18000 mp.
    for (const c of nestedStrokes) {
      expect(Math.min(c.x1_millipoints, c.x2_millipoints)).toBeGreaterThanOrEqual(outerLeft + 5_000 + 18_000)
      expect(Math.max(c.x1_millipoints, c.x2_millipoints)).toBeLessThanOrEqual(outerRight - 5_000)
      expect(Math.min(c.y1_millipoints, c.y2_millipoints)).toBeGreaterThan(outerTop)
      expect(Math.max(c.y1_millipoints, c.y2_millipoints)).toBeLessThan(outerBottom)
    }
    const innerWidth = Math.max(...nestedStrokes.map(c => c.x2_millipoints)) - Math.min(...nestedStrokes.map(c => c.x1_millipoints))
    expect(innerWidth).toBe(7_000 * 50)
    // Glyphs for every inner run attach to the anchor paragraph's line and sit inside the inner table.
    const innerGlyphs = commands.filter(c => c.kind === 'fill_glyph_path' && c.source_id.startsWith(`${ITEM}:`))
    expect(new Set(innerGlyphs.map(c => c.kind === 'fill_glyph_path' && c.source_id)).size).toBe(4)
    const anchorLine = paint.pages[0]!.lines.find(line => line.paragraph_id === 'paragraph:a')!
    expect(innerGlyphs.every(c => c.kind === 'fill_glyph_path' && c.line_id === anchorLine.line_id && anchorLine.command_ids.includes(c.id))).toBe(true)
    const innerTop = Math.min(...nestedStrokes.map(c => c.y1_millipoints))
    expect(innerTop).toBe(anchorLine.y_millipoints + anchorLine.height_millipoints + 40 * 50)
    // Paragraph B follows the reserved extent instead of overlapping the inner table.
    const below = paint.pages[0]!.lines.find(line => line.paragraph_id === 'paragraph:b')!
    expect(below.y_millipoints).toBeGreaterThanOrEqual(Math.max(...nestedStrokes.map(c => c.y2_millipoints)))
    // Fills paint before glyphs, borders after them.
    const firstGlyph = commands.findIndex(c => c.kind === 'fill_glyph_path'), lastGlyph = commands.length - 1 - [...commands].reverse().findIndex(c => c.kind === 'fill_glyph_path')
    expect(nestedFills.every(c => commands.indexOf(c) < firstGlyph)).toBe(true)
    expect(nestedStrokes.every(c => commands.indexOf(c) > lastGlyph)).toBe(true)
    // Declared policy and complete disclosure: the nested refusal is painted, so nothing is omitted.
    expect(paint.reasons).toContain(DOCX_APPROXIMATE_NESTED_TABLE_WARNING)
    expect(paint.reasons.some(r => r.startsWith(`${DOCX_APPROXIMATE_NESTED_TABLE_CODE}: painted 1 of 1`) && r.includes(DOCX_APPROXIMATE_NESTED_TABLE_LAYOUT_POLICY) && r.includes('unsupported:nested') && r.includes('firstRow'))).toBe(true)
    expect(paint.content_status).toBe('complete')
    expect(paint.omitted_content).toEqual([])
    expect(paint.reasons.some(r => r.startsWith(DOCX_APPROXIMATE_NESTED_TABLE_OMITTED_CODE))).toBe(false)
  }, 30000)

  it('fits authored columns wider than the cell, substitutes unavailable inner fonts and discloses both', async () => {
    const { input, eligibility } = fixture()
    const wide = sidecar({}, 'Nowhere Sans')
    wide.items[0]!.table!.grid_widths_twips = [9_000, 12_000]
    for (const row of wide.items[0]!.table!.rows) { row.cells[0]!.width_twips = 9_000; row.cells[1]!.width_twips = 12_000 }
    const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input), { nestedTables: wide })
    expect(paint.status).toBe('painted')
    const nestedStrokes = paint.pages[0]!.commands.filter((c): c is NativeDocxStrokeTableBorderCommandV1 => c.kind === 'stroke_table_border' && c.table_id === DOCX_APPROXIMATE_NESTED_TABLE_TABLE_ID)
    // Content box 9360 - 200 margins - 360 indent = 8800 twips.
    expect(Math.max(...nestedStrokes.map(c => c.x2_millipoints)) - Math.min(...nestedStrokes.map(c => c.x1_millipoints))).toBe(8_800 * 50)
    expect(paint.reasons.some(r => r.includes('fitted from 9000+12000 to'))).toBe(true)
    expect(paint.reasons.some(r => r.startsWith(DOCX_APPROXIMATE_NESTED_TABLE_FONT_CODE) && r.includes('Nowhere Sans') && r.includes('DejaVu Sans'))).toBe(true)
    expect(paint.content_status).toBe('complete')
  }, 30000)

  it('keeps depth-two nesting omitted with disclosure and drops evidence that does not join the source', async () => {
    const { input, eligibility } = fixture()
    const deep = sidecar({ status: 'omitted', reason: 'nested-depth-limit', table: undefined, geometry: undefined, style_borders: undefined, first_row_cell_shading_rgb: undefined, notes: undefined })
    const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input), { nestedTables: deep })
    expect(paint.status).toBe('painted')
    expect(paint.content_status).toBe('partial')
    expect(paint.omitted_content.map(entry => [entry.code, entry.scope_id, entry.path])).toEqual([['NESTED_TABLE_OR_CELL_MARKUP', 'table:1', NESTED]])
    expect(paint.reasons.some(r => r.startsWith(DOCX_APPROXIMATE_NESTED_TABLE_OMITTED_CODE) && r.includes('nested-depth-limit'))).toBe(true)
    expect(paint.reasons).not.toContain(DOCX_APPROXIMATE_NESTED_TABLE_WARNING)
    expect(paint.pages[0]!.commands.some(c => c.kind === 'stroke_table_border' && c.table_id === DOCX_APPROXIMATE_NESTED_TABLE_TABLE_ID)).toBe(false)
    const foreign = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input), { nestedTables: { ...sidecar(), package_sha256: `sha256:${'c'.repeat(64)}` } })
    expect(foreign.reasons).toContain(DOCX_APPROXIMATE_NESTED_TABLE_SIDECAR_REFUSED)
    expect(foreign.content_status).toBe('partial')
    expect(foreign.omitted_content.map(entry => entry.code)).toEqual(['NESTED_TABLE_OR_CELL_MARKUP'])
    const baseline = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input))
    expect(baseline.pages[0]!.commands.length).toBe(foreign.pages[0]!.commands.length)
    await expect(renderNativeDocxApproximatePagePreviewV1(input, { ...eligibility, status: 'ineligible' }, outlineProvider(input), { nestedTables: sidecar() })).rejects.toThrow()
  }, 30000)

  it('never changes strict preparation: the sidecar is not consulted and the strict request hash is identical', async () => {
    const { input } = fixture()
    const strict = async (runtime?: Record<string, unknown>) => {
      try { return (await prepareNativeDocxPagePaintV1(structuredClone(input), runtime as never)).request_sha256 } catch (error) { return `refused:${(error as Error).message}` }
    }
    const without = await strict(), withSidecar = await strict({ nestedTables: sidecar() })
    expect(withSidecar).toBe(without)
    expect(without).not.toBe('')
    if (!without.startsWith('refused:')) {
      const prepared = await prepareNativeDocxPagePaintV1(structuredClone(input), { nestedTables: sidecar() } as never)
      expect(JSON.stringify(prepared.page_paint_request)).not.toContain(ITEM)
    }
  }, 30000)
})
