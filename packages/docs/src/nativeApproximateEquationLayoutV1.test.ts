import { readFileSync } from 'node:fs'
import { createHash } from 'node:crypto'
import { createRequire } from 'node:module'
import { describe, expect, it } from 'vitest'
import type { NativeFontFaceManifest, NativeFontManifest } from '@injoffice/font-metrics/layout'
import { createHarfBuzzOutlineProviderV1 } from '@injoffice/font-metrics/harfbuzz'
import { DOCX_NATIVE_PROTOCOL, DOCX_NATIVE_VERSION, type NativeDocxDocumentV1 } from './nativeContract.js'
import { DOCX_RESOLVED_LAYOUT_PROTOCOL, DOCX_RESOLVED_LAYOUT_VERSION, type NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'
import { DOCX_PAGINATION_SETTINGS_PROTOCOL, DOCX_PAGINATION_SETTINGS_VERSION, type NativeDocxPaginationSettingsV1 } from './nativePaginationSettings.js'
import { DOCX_PAGE_PAINT_COMPILER_PROTOCOL, DOCX_PAGE_PAINT_COMPILER_VERSION, prepareNativeDocxPagePaintV1, renderNativeDocxApproximatePagePreviewV1, decodeNativeDocxApproximatePagePreviewV1, type NativeDocxPagePaintPrepareInputV1 } from './nativePagePaintCompilerV1.js'
import { encodeNativeDOCXFontInventoryV1, nativeDOCXCanonicalWireSHA256V1, type NativeDOCXFontInventoryV1 } from './nativeFontInventoryV1.js'
import { decodeNativeDocxApproximateEquationsV1, selectNativeDocxApproximateMathFaceV1, DOCX_APPROXIMATE_EQUATION_CODE, DOCX_APPROXIMATE_EQUATION_FONT_CODE, DOCX_APPROXIMATE_EQUATION_OMITTED_CODE, DOCX_APPROXIMATE_EQUATION_TABLE_ID, DOCX_APPROXIMATE_EQUATION_WARNING, type NativeDocxApproximateEquationsV1, type NativeDocxApproximateMathNodeV1, type NativeDocxApproximateMathRunV1 } from './nativeApproximateEquationLayoutV1.js'
import type { NativeDocxFillGlyphPathCommandV1, NativeDocxFillTableCellCommandV1, NativeDocxGlyphOutlineRequestV1 } from './nativePagePaintV1.js'

const require = createRequire(import.meta.url)
const FONT_BYTES = new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
const FONT_DIGEST = 'sha256:7da195a74c55bef988d0d48f9508bd5d849425c1770dba5d7bfc6ce9ed848954' as const
const HASH = `sha256:${'a'.repeat(64)}`
const RELATIONSHIPS_HASH = `sha256:${'b'.repeat(64)}`
const SETTINGS_PART = 'word/settings.xml'
const RELATIONSHIPS_PART = 'word/_rels/document.xml.rels'
const REVISION = `rev:${'a'.repeat(32)}`
const MATH_PATH = '/w:document[1]/w:body[1]/w:p[2]/nsf4b2b884:oMathPara[1]'
const INLINE_PATH = '/w:document[1]/w:body[1]/w:p[1]/nsf4b2b884:oMath[1]'

function storedFontDigest(bytes: Uint8Array): string {
  const stored = Uint8Array.from(bytes)
  const key = Uint8Array.from([0xff, 0xee, 0xdd, 0xcc, 0xbb, 0xaa, 0x99, 0x88, 0x77, 0x66, 0x55, 0x44, 0x33, 0x22, 0x11, 0x00])
  for (let index = 0; index < 32 && index < stored.length; index++) stored[index] = stored[index]! ^ key[index % 16]!
  return `sha256:${createHash('sha256').update(stored).digest('hex')}`
}
function anchor(path: string, start: number, end: number) { return { part_name: 'word/document.xml', path, start_byte: start, end_byte: end, xml_sha256: HASH } }

/** Two body paragraphs: text 'A' then an empty equation paragraph, embedded DejaVu Sans. */
function fixture(): NativeDocxPagePaintPrepareInputV1 {
  const policy = { mode: 'read-only' as const, allowed_operations: [], refusal: { code: 'NATIVE_READ_ONLY', message: 'Fixture is immutable.', preservation: 'refuse-mutation' as const } }
  const first = { id: 'paragraph:1', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]', 100, 190), edit_policy: policy, properties: {}, runs: [{ kind: 'text' as const, id: 'run:1', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]', 110, 180), text: 'A' }] }
  const second = { id: 'paragraph:2', anchor: anchor('/w:document[1]/w:body[1]/w:p[2]', 200, 900), edit_policy: policy, properties: {}, runs: [] }
  const document: NativeDocxDocumentV1 = {
    protocol: DOCX_NATIVE_PROTOCOL, version: DOCX_NATIVE_VERSION, document_id: 'document:test', revision: REVISION,
    source: { package_sha256: HASH, main_part: 'word/document.xml' },
    body: { id: 'story:body', kind: 'body', part_name: 'word/document.xml', anchor: anchor('/w:document[1]/w:body[1]', 1, 3_000), blocks: [{ kind: 'paragraph', id: first.id, paragraph: first }, { kind: 'paragraph', id: second.id, paragraph: second }] },
    sections: [{
      id: 'section:1', anchor: anchor('/w:document[1]/w:body[1]/w:sectPr[1]', 2_100, 2_190), starts_at_block_id: first.id, break_type: 'next-page', title_page: false,
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
    unsupported: [{ id: 'unsupported:equation', code: 'UNMODELED_PARAGRAPH_CONTENT', capability: 'run-structure', scope_id: second.id, anchor: anchor(MATH_PATH, 300, 800), preservation: 'refuse-mutation', message: 'Paragraph content outside the v1 run subset is preserved verbatim' }],
  }
  const resolved: NativeDocxResolvedLayoutInputV1 = {
    protocol: DOCX_RESOLVED_LAYOUT_PROTOCOL, version: DOCX_RESOLVED_LAYOUT_VERSION, document_id: document.document_id, revision: document.revision,
    source_parts: { main_part: 'word/document.xml', font_table_part: 'word/fontTable.xml' },
    paragraphs: [
      { paragraph_id: first.id, applied_styles: [], properties: {}, paragraph_mark_properties: { font_family: 'DejaVu Sans', font_size_half_points: 20 } },
      { paragraph_id: second.id, applied_styles: [], properties: {}, paragraph_mark_properties: { font_family: 'DejaVu Sans', font_size_half_points: 20 } },
    ],
    runs: [{ run_id: 'run:1', paragraph_id: first.id, applied_paragraph_styles: [], applied_character_styles: [], properties: { font_family: 'DejaVu Sans', font_size_half_points: 20, color: '123456' } }],
    tables: [], fonts: [{ name: 'DejaVu Sans' }], diagnostics: [],
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
        kind: 'document', face_slot: 'embedRegular', font_table_part: 'word/fontTable.xml', font_table_path: '/w:fonts[1]/w:font[1]/w:embedRegular[1]', relationships_part: 'word/_rels/fontTable.xml.rels', relationships_sha256: RELATIONSHIPS_HASH,
        relationship_id: 'rIdFont', relationship_type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font', relationship_target: 'fonts/DejaVuSans.odttf', asset_part: 'word/fonts/DejaVuSans.odttf', asset_content_type: 'application/vnd.openxmlformats-officedocument.obfuscatedFont',
        stored_byte_length: FONT_BYTES.byteLength, stored_sha256: storedFontDigest(FONT_BYTES), content_sha256: FONT_DIGEST, resource_id: `font:${FONT_DIGEST}`,
        obfuscation: { algorithm: 'ecma-376-font-obfuscation', font_key: '{00112233-4455-6677-8899-AABBCCDDEEFF}', subsetted: false },
        licensing: { embedding_origin: 'document-package', rights_source: 'sfnt-os2-fstype', rights_status: 'verified', embedding_rights: 'installable', no_subsetting: false, allowed_scope: 'document-only' },
      },
    }] }],
    references: [{ family: 'DejaVu Sans', weight: 400, style: 'normal', scope_ids: ['paragraph:1', 'paragraph:2', 'run:1'] }],
    native_text_manifest: manifest, native_text_manifest_sha256: nativeDOCXCanonicalWireSHA256V1(manifest),
  }
  inventory.inventory_sha256 = nativeDOCXCanonicalWireSHA256V1({ ...structuredClone(inventory), inventory_sha256: '' })
  return {
    protocol: DOCX_PAGE_PAINT_COMPILER_PROTOCOL, version: DOCX_PAGE_PAINT_COMPILER_VERSION, source_revision: 'git:integration-test', outline_provider: { provider_id: 'injoffice.sfnt-outline', provider_revision: 'sfnt-v1' },
    document, resolved_layout: resolved, pagination_settings: settings, font_inventory_json: encodeNativeDOCXFontInventoryV1(inventory),
    font_assets: [{ face_id: 'font-face:dejavu', face_slot: 'embedRegular', resource_id: `font:${FONT_DIGEST}`, content_digest: FONT_DIGEST, collection_index: null, bytes: FONT_BYTES }], media_assets: [],
  }
}
function eligibility(input: NativeDocxPagePaintPrepareInputV1) {
  const settings = input.pagination_settings as NativeDocxPaginationSettingsV1
  return { protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: settings.document_id, revision: settings.revision, package_sha256: HASH, settings_sha256: settings.settings_sha256, status: 'eligible', legacy_compatibility_mode: 12, reasons: ['Legacy mode'] }
}
function outlineProvider(input: NativeDocxPagePaintPrepareInputV1) {
  const outlines = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
  return { providerId: input.outline_provider.provider_id, providerRevision: input.outline_provider.provider_revision, getGlyphOutline(request: NativeDocxGlyphOutlineRequestV1) { const o = outlines.outline(request.glyph_id); return o.path.length ? { status: 'outlined' as const, ...request, ...o } : { status: 'empty' as const, ...request, units_per_em: o.units_per_em } } }
}

const RUN: NativeDocxApproximateMathRunV1 = { font_family: 'Cambria Math', font_size_half_points: 20, bold: false, italic: false, color: '4F81BD' }
const text = (value: string, run: Partial<NativeDocxApproximateMathRunV1> = {}): NativeDocxApproximateMathNodeV1 => ({ kind: 'text', text: value, run: { ...RUN, ...run } })
const row = (...children: NativeDocxApproximateMathNodeV1[]): NativeDocxApproximateMathNodeV1 => ({ kind: 'row', children })
/** (x+a)^n = Σ_{k=0}^{n} (n choose k) x^k a^{n-k} */
function binomialTheorem(): NativeDocxApproximateMathNodeV1 {
  return row(
    { kind: 'superscript', run: RUN, children: [row({ kind: 'delimiter', run: RUN, children: [row(text('x+a'))] }), row(text('n'))] },
    text('='),
    { kind: 'nary', run: RUN, chr: '∑', grow: true, children: [row(text('k=0')), row(text('n')), row({ kind: 'delimiter', run: RUN, children: [row({ kind: 'fraction', run: RUN, bar: false, children: [row(text('n')), row(text('k'))] })] }, { kind: 'superscript', run: RUN, children: [row(text('x')), row(text('k'))] }, { kind: 'superscript', run: RUN, children: [row(text('a')), row(text('n-k'))] })] },
  )
}
function sidecar(lines: NativeDocxApproximateMathNodeV1[], overrides: Record<string, unknown> = {}): NativeDocxApproximateEquationsV1 {
  return {
    protocol: 'injoffice.docx.approximate-equations', version: 1, policy: 'docx.approximate-equation-preview-v1', package_sha256: HASH, part_sha256: HASH, omitted_count: 0,
    font_requests: [{ family: 'Cambria Math', weight: 400, style: 'normal' }, { family: 'Cambria Math', weight: 400, style: 'italic' }, { family: 'DejaVu Sans', weight: 400, style: 'italic' }],
    items: [Object.fromEntries(Object.entries({ id: 'approximate-equation:test:1', paragraph_id: 'paragraph:2', diagnostic_ids: ['unsupported:equation'], anchor: anchor(MATH_PATH, 300, 800), status: 'supported', display: true, justification: 'centerGroup', lines, ...overrides }).filter(([, value]) => value !== undefined)) as unknown as NativeDocxApproximateEquationsV1['items'][number]],
  }
}
type Glyph = NativeDocxFillGlyphPathCommandV1
const glyphTop = (glyph: Glyph) => Math.min(...glyph.path.flatMap(p => 'y_millipoints' in p ? [p.y_millipoints] : []))
const glyphBottom = (glyph: Glyph) => Math.max(...glyph.path.flatMap(p => 'y_millipoints' in p ? [p.y_millipoints] : []))
const glyphLeft = (glyph: Glyph) => Math.min(...glyph.path.flatMap(p => 'x_millipoints' in p ? [p.x_millipoints] : []))

describe('approximate OMML equations', () => {
  it('lays the binomial theorem out as a centered display block with recognizable fraction, script and n-ary structure', async () => {
    const input = fixture()
    const before = structuredClone(input)
    const equations = sidecar([binomialTheorem()])
    const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility(input), outlineProvider(input), { equations })
    expect(paint.status).toBe('painted')
    expect(input).toEqual(before)
    expect(decodeNativeDocxApproximatePagePreviewV1(paint).ok).toBe(true)
    const page = paint.pages[0]!
    const line = page.lines.find(entry => entry.paragraph_id === 'paragraph:2')!
    const glyphs = page.commands.filter((c): c is Glyph => c.kind === 'fill_glyph_path' && c.source_id === 'approximate-equation:test:1:run')
    expect(glyphs.length).toBeGreaterThanOrEqual(16)
    for (const glyph of glyphs) { expect(glyph.line_id).toBe(line.line_id); expect(line.command_ids).toContain(glyph.id); expect(glyph.fill_rgb).toBe('4F81BD') }
    expect(page.commands.some(c => c.kind === 'fill_text_highlight')).toBe(false)
    // The equation is centered in the column and sits inside its own reserved line.
    const left = Math.min(...glyphs.map(glyphLeft)), right = Math.max(...glyphs.flatMap(g => g.path.flatMap(p => 'x_millipoints' in p ? [p.x_millipoints] : [])))
    const column = page.columns[0]!
    expect(left).toBeGreaterThan(column.x_millipoints + column.width_millipoints * 0.2)
    expect(right).toBeLessThan(column.x_millipoints + column.width_millipoints * 0.8)
    expect(Math.abs((left + right) / 2 - (column.x_millipoints + column.width_millipoints / 2))).toBeLessThan(6_000)
    expect(Math.min(...glyphs.map(glyphTop))).toBeGreaterThanOrEqual(line.y_millipoints - 1)
    expect(Math.max(...glyphs.map(glyphBottom))).toBeLessThanOrEqual(line.y_millipoints + line.height_millipoints + 1)
    // The display n-ary operator is enlarged, and its limits sit above and below it.
    const sizes = new Set(glyphs.map(g => g.font_size_millipoints))
    expect(sizes.has(10_000)).toBe(true); expect(sizes.has(7_000)).toBe(true); expect(sizes.has(14_000)).toBe(true)
    const sigma = glyphs.find(g => g.font_size_millipoints === 14_000)!
    const scripts = glyphs.filter(g => g.font_size_millipoints === 7_000)
    const above = scripts.filter(s => glyphBottom(s) <= glyphTop(sigma) && Math.abs(glyphLeft(s) - glyphLeft(sigma)) < 12_000)
    const below = scripts.filter(s => glyphTop(s) >= glyphBottom(sigma) && Math.abs(glyphLeft(s) - glyphLeft(sigma)) < 12_000)
    expect(above.length).toBe(1); expect(below.length).toBe(3)
    // Superscripts rise above the baseline glyphs; the noBar binomial paints no rule.
    const base = glyphs.filter(g => g.font_size_millipoints === 10_000)
    const baseBaseline = Math.max(...base.map(glyphBottom))
    const raisedScripts = scripts.filter(s => glyphBottom(s) < baseBaseline - 2_000)
    expect(raisedScripts.length).toBeGreaterThanOrEqual(4)
    expect(page.commands.some(c => c.kind === 'fill_table_cell')).toBe(false)
    expect(paint.reasons).toContain(DOCX_APPROXIMATE_EQUATION_WARNING)
    expect(paint.reasons.some(r => r.startsWith(`${DOCX_APPROXIMATE_EQUATION_CODE}: painted 1 of 1`) && r.includes('aligned center'))).toBe(true)
    expect(paint.reasons.some(r => r.startsWith(`${DOCX_APPROXIMATE_EQUATION_FONT_CODE}: Cambria Math / 400 / italic -> DejaVu Sans / 400 / normal (text-face-fallback, weight/style differ`))).toBe(true)
    // The equation refusal is no longer disclosed as omitted content; the page is not blank.
    expect(paint.omitted_content.some(entry => entry.category === 'equation')).toBe(false)
    expect(paint.unpainted_pages).toEqual([])
  }, 30000)

  it('places the fraction rule on the math axis between numerator and denominator and a radical overbar above the radicand', async () => {
    const input = fixture()
    const equations = sidecar([row({ kind: 'fraction', run: RUN, children: [row(text('a')), row(text('b'))] }, text('+'), { kind: 'radical', run: RUN, degree_hide: true, children: [row(), row(text('2'))] })])
    const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility(input), outlineProvider(input), { equations })
    expect(paint.status).toBe('painted')
    const page = paint.pages[0]!
    const rules = page.commands.filter((c): c is NativeDocxFillTableCellCommandV1 => c.kind === 'fill_table_cell')
    expect(rules).toHaveLength(2)
    expect(rules.every(rule => rule.table_id === DOCX_APPROXIMATE_EQUATION_TABLE_ID && rule.row_id === 'approximate-equation:test:1' && rule.fill_rgb === '4F81BD')).toBe(true)
    const glyphs = page.commands.filter((c): c is Glyph => c.kind === 'fill_glyph_path' && c.source_id === 'approximate-equation:test:1:run')
    // Paint order follows layout order: numerator, denominator, '+', the radical sign, then the radicand.
    const [a, b, plus, sign, two] = ['a', 'b', '+', '√', '2'].map((_, index) => glyphs[index]!)
    expect(glyphs).toHaveLength(5)
    expect(glyphLeft(sign)).toBeLessThan(glyphLeft(two))
    const bar = rules.find(rule => rule.width_millipoints > rule.height_millipoints * 5 && rule.x_millipoints < glyphLeft(plus!))!
    expect(glyphBottom(a!)).toBeLessThan(bar.y_millipoints)
    expect(glyphTop(b!)).toBeGreaterThan(bar.y_millipoints + bar.height_millipoints)
    // The fraction rule sits near the axis of the '+' sign.
    expect(Math.abs(bar.y_millipoints - (glyphTop(plus!) + glyphBottom(plus!)) / 2)).toBeLessThan(1_500)
    // Display fractions keep the text size for numerator and denominator.
    expect(a!.font_size_millipoints).toBe(10_000); expect(b!.font_size_millipoints).toBe(10_000)
    const overbar = rules.find(rule => rule !== bar)!
    expect(overbar.y_millipoints).toBeLessThan(glyphTop(two!))
    expect(overbar.x_millipoints + overbar.width_millipoints).toBeGreaterThanOrEqual(glyphLeft(two!))
    expect(decodeNativeDocxApproximatePagePreviewV1(paint).ok).toBe(true)
  }, 30000)

  it('reserves inline equations in the text line and keeps strict preparation refusing the paragraph content', async () => {
    const input = fixture()
    const document = input.document as NativeDocxDocumentV1
    document.unsupported.push({ id: 'unsupported:inline', code: 'UNMODELED_PARAGRAPH_CONTENT', capability: 'run-structure', scope_id: 'paragraph:1', anchor: anchor(INLINE_PATH, 181, 189), preservation: 'refuse-mutation', message: 'Paragraph content outside the v1 run subset is preserved verbatim' })
    const equations = sidecar([binomialTheorem()])
    equations.items.push({ id: 'approximate-equation:test:2', paragraph_id: 'paragraph:1', diagnostic_ids: ['unsupported:inline'], anchor: anchor(INLINE_PATH, 181, 189), status: 'supported', display: false, lines: [row({ kind: 'subscript', run: RUN, children: [row(text('x')), row(text('1'))] })] })
    const strict = await prepareNativeDocxPagePaintV1(input)
    expect(strict.page_paint_request.pagination_request.shaped_lines.paragraphs.flatMap(p => p.lines.flatMap(l => l.fragments)).some(f => f.source_kind === 'textbox')).toBe(false)
    const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility(input), outlineProvider(input), { equations })
    expect(paint.status).toBe('painted')
    const page = paint.pages[0]!
    const first = page.lines.find(entry => entry.paragraph_id === 'paragraph:1')!
    const textGlyph = page.commands.find((c): c is Glyph => c.kind === 'fill_glyph_path' && c.source_id === 'run:1')!
    const inline = page.commands.filter((c): c is Glyph => c.kind === 'fill_glyph_path' && c.source_id === 'approximate-equation:test:2:run')
    expect(inline.length).toBe(2)
    expect(Math.min(...inline.map(glyphLeft))).toBeGreaterThan(glyphLeft(textGlyph))
    expect(inline.every(g => g.line_id === first.line_id)).toBe(true)
    // The subscript is smaller and lower than its base.
    const [base, sub] = inline
    expect(sub!.font_size_millipoints).toBe(7_000); expect(glyphBottom(sub!)).toBeGreaterThan(glyphBottom(base!))
    // Inline equations keep the paragraph alignment; only the empty display paragraph is centered.
    expect(paint.reasons.some(r => r.includes('approximate-equation:test:2: display'))).toBe(false)
    expect(paint.reasons.some(r => r.includes('painted 2 of 2'))).toBe(true)
    expect(decodeNativeDocxApproximatePagePreviewV1(paint).ok).toBe(true)
  }, 30000)

  it('discloses omitted equations and refuses sidecars that do not exact-join the source document', async () => {
    const input = fixture()
    const omitted = sidecar([], { status: 'omitted', reason: 'unsupported equation element m:eqArr', lines: undefined, justification: undefined })
    const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility(input), outlineProvider(input), { equations: omitted })
    expect(paint.status).toBe('painted')
    expect(paint.reasons.some(r => r.startsWith(`${DOCX_APPROXIMATE_EQUATION_OMITTED_CODE}: approximate-equation:test:1 (unsupported equation element m:eqArr)`))).toBe(true)
    expect(paint.reasons).not.toContain(DOCX_APPROXIMATE_EQUATION_WARNING)
    expect(paint.omitted_content.some(entry => entry.category === 'equation')).toBe(true)
    // A painted equation drops its source refusal from omitted_content; only omitted or unplaced ones keep it.
    const painted = await renderNativeDocxApproximatePagePreviewV1(input, eligibility(input), outlineProvider(input), { equations: sidecar([binomialTheorem()]) })
    expect(painted.omitted_content.some(entry => entry.category === 'equation')).toBe(false)
    const cases: Array<[string, (equations: any) => void]> = [
      ['package', equations => { equations.package_sha256 = `sha256:${'b'.repeat(64)}` }],
      ['paragraph', equations => { equations.items[0].paragraph_id = 'paragraph:1' }],
      ['diagnostic', equations => { equations.items[0].diagnostic_ids = ['unsupported:other'] }],
      ['anchor', equations => { equations.items[0].anchor = anchor(MATH_PATH, 301, 800) }],
      ['unknown field', equations => { equations.extra = true }],
      ['arity', equations => { equations.items[0].lines[0].children[0].children.pop() }],
      ['operator', equations => { equations.items[0].lines[0].children[2].chr = '∑∑' }],
      ['run color', equations => { equations.items[0].lines[0].children[1].run.color = 'blue' }],
      ['font request', equations => { equations.font_requests.push({ family: 'X', weight: 500, style: 'normal' }) }],
      ['display lines', equations => { equations.items[0].display = false; equations.items[0].justification = undefined; equations.items[0].lines.push(row(text('y'))) }],
      ['family bound', equations => { equations.items[0].lines[0].children[1].run.font_family = 'F'.repeat(129) }],
      ['request family bound', equations => { equations.font_requests[0].family = 'F'.repeat(129) }],
      ['reason bound', equations => { equations.items[0] = { ...equations.items[0], status: 'omitted', reason: 'r'.repeat(257) }; delete equations.items[0].lines; delete equations.items[0].justification }],
      ['main part digest', equations => { equations.part_sha256 = `sha256:${'c'.repeat(64)}` }],
    ]
    for (const [name, mutate] of cases) {
      const equations = sidecar([binomialTheorem()])
      mutate(equations)
      expect(() => decodeNativeDocxApproximateEquationsV1(equations, input.document as NativeDocxDocumentV1, HASH), name).toThrow()
      await expect(renderNativeDocxApproximatePagePreviewV1(input, eligibility(input), outlineProvider(input), { equations }), name).rejects.toThrow()
    }
    await expect(renderNativeDocxApproximatePagePreviewV1(input, { ...eligibility(input), status: 'ineligible' }, outlineProvider(input), { equations: sidecar([binomialTheorem()]) })).rejects.toThrow()
  }, 30000)

  it('selects math faces by the declared policy: authored family, declared math substitute, then the text face', () => {
    const face = (family: string, weight = 400, style: 'normal' | 'italic' = 'normal', kind: 'host' | 'document' | 'system' = 'host'): NativeFontFaceManifest => ({ faceId: `face:${family}:${weight}:${style}`, family, weight, style, stretch: 100, source: { kind, resourceId: family, contentDigest: HASH as `sha256:${string}` } })
    const manifest = (faces: NativeFontFaceManifest[]): NativeFontManifest => ({ version: 1, manifestId: 'm', revision: 'r', faces, fallbackChains: [] })
    const requested = { family: 'Cambria Math', weight: 700, style: 'italic' }
    expect(selectNativeDocxApproximateMathFaceV1(manifest([face('Calibri'), face('Cambria Math', 700, 'italic')]), requested, 'Calibri')).toMatchObject({ policy: 'exact', weight_style_match: true, face: { family: 'Cambria Math' } })
    expect(selectNativeDocxApproximateMathFaceV1(manifest([face('Calibri'), face('Cambria Math')]), requested, 'Calibri')).toMatchObject({ policy: 'exact', weight_style_match: false })
    expect(selectNativeDocxApproximateMathFaceV1(manifest([face('Calibri', 700, 'italic'), face('STIX Two Math')]), requested, 'Calibri')).toMatchObject({ policy: 'declared-math-substitute', face: { family: 'STIX Two Math' } })
    expect(selectNativeDocxApproximateMathFaceV1(manifest([face('Calibri'), face('Calibri', 700, 'italic'), face('Arial', 700, 'italic')]), requested, 'Calibri')).toMatchObject({ policy: 'text-face-fallback', weight_style_match: true, face: { family: 'Calibri', weight: 700 } })
    expect(selectNativeDocxApproximateMathFaceV1(manifest([face('Arial')]), requested, 'Calibri')).toMatchObject({ policy: 'text-face-fallback', face: { family: 'Arial' } })
    expect(selectNativeDocxApproximateMathFaceV1(manifest([face('Cambria Math', 400, 'normal', 'system')]), requested, 'Calibri')).toBeUndefined()
  })
})
