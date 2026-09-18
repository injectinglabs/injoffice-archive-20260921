import { readFileSync } from 'node:fs'
import { createHash } from 'node:crypto'
import { createRequire } from 'node:module'
import { describe, expect, it, vi } from 'vitest'
import {qualifyApproximateLegacyTables,DOCX_LEGACY_TABLE_ORIGIN_WARNING,DOCX_TABLE_BORDER_RESERVATION_WARNING,DOCX_TABLE_GRID_FIT_WARNING} from './nativeLegacyTableOriginV1.js'
import type { NativeFontManifest, NativeFontResolver, ResolvedFontFace } from '@injoffice/font-metrics/layout'
import {selectExplicitFontV1} from '@injoffice/font-metrics/layout'
import {renderNativeDocxFontSubstitutionPreviewV1,decodeNativeDocxFontSubstitutionPreviewV1} from './nativePagePaintCompilerV1.js'
import {validateNativeDocxFontPageFieldVariantsV1,validateNativeDocxPageFieldVariantsV1} from './nativePageFieldsV1.js'
import {nativeDocxFontSubstitutionDiagnosticV1,type NativeDocxFontSubstitutionV1} from './nativeFontSubstitutionEvidenceV1.js'
import {qualifyNativeDocxFontCompositionV1} from './nativeFontCompositionV1.js'
import { createHarfBuzzTextShaperV1, createHarfBuzzOutlineProviderV1, inspectHarfBuzzFontMetricsV1 } from '@injoffice/font-metrics/harfbuzz'
import { reorderNativeBidiLineV1 } from '@injoffice/font-metrics/bidi'
import { DOCX_NATIVE_PROTOCOL, DOCX_NATIVE_VERSION, type NativeDocxDocumentV1, type NativeDocxRunV1 } from './nativeContract.js'
import { DOCX_RESOLVED_LAYOUT_PROTOCOL, DOCX_RESOLVED_LAYOUT_VERSION, type NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'
import { DOCX_PAGINATION_SETTINGS_PROTOCOL, DOCX_PAGINATION_SETTINGS_VERSION, type NativeDocxPaginationSettingsV1 } from './nativePaginationSettings.js'
import {
  DOCX_PAGE_PAINT_COMPILER_PROTOCOL,
  DOCX_PAGE_PAINT_COMPILER_VERSION,
  decodeNativeDocxApproximateDrawingShapesV1,
  DOCX_APPROXIMATE_DRAWING_SHAPE_WARNING,
  completeNativeDocxPagePaintV1,
  prepareNativeDocxPagePaintV1,
  renderNativeDocxApproximatePagePreviewV1,
  decodeNativeDocxApproximatePagePreviewV1,
  type NativeDocxPagePaintPrepareInputV1,
} from './nativePagePaintCompilerV1.js'
import { encodeNativeDOCXFontInventoryV1, nativeDOCXCanonicalWireSHA256V1, type NativeDOCXFontInventoryV1 } from './nativeFontInventoryV1.js'
import { decodeNativeDocxPagePaintResourceListV1, qualifyNativeDocxInlineImageV1 } from './nativeImagePagePaintV1.js'
import { paginateNativeDocxV1, paginateNativeDocxApproximateLegacyV1 } from './nativePaginationV1.js'
import { qualifyNativeDocxTablesV1,layoutNativeDocxTableRowsV1, nativeDocxTableProjectionSha256V1 } from './nativeTablePagePaintV1.js'
import { decodeNativeDocxShapedLines } from './nativeShapedLinesContract.js'
import { decodeNativeDocxPagePaintForRequestV1, decodeNativeDocxPagePaintRequestV1, nativeDocxPagePaintShapedLinesSha256V1, decodeNativeDocxApproximateComputedPagePaintV1, nativeDocxPagePaintPaginatedLayoutSha256V1 } from './nativePagePaintV1.js'
import { nativeDocxPageFieldDocumentV1 } from './nativePageFieldsV1.js'
import {deriveNativeSquareWrapPlanV1} from './nativeSquareWrapV1.js'
import {nativeDocxFloatingAnchorOriginsV1, resolveNativeDocxFloatingAnchorV1} from './nativeFloatingAnchorV1.js'
import { renderNativeDocxAutomaticBorderPreviewV1 } from './nativePagePaintCompilerV1.js'
import { NativeDocxPreviewRefusalV1, nativeDocxPreviewRefusalRecordV1, DOCX_PREVIEW_REFUSAL_PROTOCOL, DOCX_PREVIEW_REFUSAL_VERSION } from './nativePreviewRefusalV1.js'
import { projectNativeDocxAutomaticBordersV1, decodeNativeDocxAutomaticBorderPreviewV1 } from './nativeAutomaticBorderPreviewV1.js'
import { DOCX_AUTO_BORDER_POLICY, DOCX_AUTO_BORDER_WARNING } from './nativeAutomaticBorderEvidenceV1.js'
import { DOCX_ABSENT_FONT_SIZE_WARNING, projectNativeDocxAbsentFontSizesV1 } from './nativeAbsentFontSizeV1.js'
import { DOCX_APPROXIMATE_DRAWING_CHART_WARNING, DOCX_APPROXIMATE_DRAWING_CHART_SIDECAR_REFUSED, decodeNativeDocxApproximateDrawingChartsV1 } from './nativeApproximateDrawingChartsV1.js'
import { DOCX_APPROXIMATE_INDENTED_CELL_LINE_WARNING, DOCX_APPROXIMATE_INERT_NOTE_SEPARATOR_WARNING } from './nativePaginationV1.js'
import { DOCX_LATIN_FONT_FALLBACK_WARNING, projectNativeDocxLatinFontFallbacksV1 } from './nativeLatinFontFallbackV1.js'
import { projectNativeDocxAbsentFontFamiliesV1, stripNativeDocxAbsentFontFamiliesV1 } from './nativeAbsentFontFamilyV1.js'
import { DOCX_APPROXIMATE_IMAGE_EXTENT_WARNING } from './nativeApproximateImageExtentV1.js'
import { compileNativeDocxPagePaintV1 } from './nativePagePaintV1.js'

const require = createRequire(import.meta.url)
const FONT_BYTES = new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
const FONT_DIGEST = 'sha256:7da195a74c55bef988d0d48f9508bd5d849425c1770dba5d7bfc6ce9ed848954' as const
const HASH = `sha256:${'a'.repeat(64)}`
const RELATIONSHIPS_HASH = `sha256:${'b'.repeat(64)}`
const SETTINGS_PART = 'word/settings.xml'
const RELATIONSHIPS_PART = 'word/_rels/document.xml.rels'
const REVISION = `rev:${'a'.repeat(32)}`

function storedFontDigest(bytes: Uint8Array): string {
  const key = '00112233445566778899AABBCCDDEEFF'
  const stored = Uint8Array.from(bytes)
  for (let index = 0; index < 32; index += 1) stored[index] ^= Number.parseInt(key.slice((15 - index % 16) * 2, (16 - index % 16) * 2), 16)
  return `sha256:${createHash('sha256').update(stored).digest('hex')}`
}
const PNG_BYTES = new Uint8Array(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=', 'base64'))
const PNG_DIGEST = `sha256:${createHash('sha256').update(PNG_BYTES).digest('hex')}` as `sha256:${string}`

function anchor(path: string, start: number, end: number) {
  return { part_name: 'word/document.xml', path, start_byte: start, end_byte: end, xml_sha256: HASH }
}

function fixture(): NativeDocxPagePaintPrepareInputV1 {
  const paragraph = {
    id: 'paragraph:1',
    anchor: anchor('/w:document[1]/w:body[1]/w:p[1]', 100, 190),
    edit_policy: { mode: 'read-only' as const, allowed_operations: [], refusal: { code: 'NATIVE_READ_ONLY', message: 'Fixture is immutable.', preservation: 'refuse-mutation' as const } },
    properties: {},
    runs: [{ kind: 'text' as const, id: 'run:1', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]', 110, 180), text: 'A' }],
  }
  const document: NativeDocxDocumentV1 = {
    protocol: DOCX_NATIVE_PROTOCOL,
    version: DOCX_NATIVE_VERSION,
    document_id: 'document:test',
    revision: REVISION,
    source: { package_sha256: HASH, main_part: 'word/document.xml' },
    body: { id: 'story:body', kind: 'body', part_name: 'word/document.xml', anchor: anchor('/w:document[1]/w:body[1]', 1, 3_000), blocks: [{ kind: 'paragraph', id: paragraph.id, paragraph }] },
    sections: [{
      id: 'section:1',
      anchor: anchor('/w:document[1]/w:body[1]/w:sectPr[1]', 2_100, 2_190),
      starts_at_block_id: paragraph.id,
      break_type: 'next-page',
      title_page: false,
      page: {
        width_twips: 12_240, height_twips: 15_840, orientation: 'portrait',
        margins: { top_twips: 1_440, right_twips: 1_440, bottom_twips: 1_440, left_twips: 1_440, header_twips: 720, footer_twips: 720, gutter_twips: 0 },
        columns: 1, column_spacing_twips: 720, column_layout: 'equal-width', column_definitions: [{ id: 'column:section:1:0', ordinal: 0 }],
      },
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
    unsupported: [],
  }
  const resolved: NativeDocxResolvedLayoutInputV1 = {
    protocol: DOCX_RESOLVED_LAYOUT_PROTOCOL,
    version: DOCX_RESOLVED_LAYOUT_VERSION,
    document_id: document.document_id,
    revision: document.revision,
    source_parts: { main_part: 'word/document.xml', font_table_part: 'word/fontTable.xml' },
    paragraphs: [{ paragraph_id: paragraph.id, applied_styles: [], properties: {}, paragraph_mark_properties: { font_family: 'DejaVu Sans', font_size_half_points: 20 } }],
    runs: [{ run_id: 'run:1', paragraph_id: paragraph.id, applied_paragraph_styles: [], applied_character_styles: [], properties: { font_family: 'DejaVu Sans', font_size_half_points: 20, color: '123456' } }],
    tables: [], fonts: [{ name: 'DejaVu Sans' }], diagnostics: [],
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
    default_tab_stop_twips: 720,
    mirror_margins: false,
    gutter_at_top: false,
    even_and_odd_headers: false,
    compatibility_mode: 15,
    diagnostics: [],
  }
  const manifest: NativeFontManifest = {
    version: 1,
    manifestId: `docx.fonts.${createHash('sha256').update(document.document_id).digest('hex').slice(0, 24)}`,
    revision: document.revision,
    faces: [{
      faceId: 'font-face:dejavu', family: 'DejaVu Sans', weight: 400, style: 'normal', stretch: 100,
      source: { kind: 'document', resourceId: `font:${FONT_DIGEST}`, contentDigest: FONT_DIGEST },
    }],
    fallbackChains: [],
  }
  const inventory: NativeDOCXFontInventoryV1 = {
    protocol: 'injoffice.docx.font-inventory', version: 1,
    document_id: document.document_id, revision: document.revision, package_sha256: HASH,
    main_part: document.source.main_part, main_sha256: HASH, inventory_sha256: HASH,
    font_table: {
      part_name: 'word/fontTable.xml', sha256: HASH,
      main_relationships_part: RELATIONSHIPS_PART, main_relationships_sha256: RELATIONSHIPS_HASH,
      relationship_id: 'rIdFontTable', relationship_type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable', relationship_target: 'fontTable.xml',
      font_relationships_part: 'word/_rels/fontTable.xml.rels', font_relationships_sha256: RELATIONSHIPS_HASH,
    },
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
    references: [{ family: 'DejaVu Sans', weight: 400, style: 'normal', scope_ids: ['paragraph:1', 'run:1'] }],
    native_text_manifest: manifest,
    native_text_manifest_sha256: nativeDOCXCanonicalWireSHA256V1(manifest),
  }
  inventory.inventory_sha256 = nativeDOCXCanonicalWireSHA256V1({ ...structuredClone(inventory), inventory_sha256: '' })
  return {
    protocol: DOCX_PAGE_PAINT_COMPILER_PROTOCOL,
    version: DOCX_PAGE_PAINT_COMPILER_VERSION,
    source_revision: 'git:integration-test',
    outline_provider: { provider_id: 'injoffice.sfnt-outline', provider_revision: 'sfnt-v1' },
    document,
    resolved_layout: resolved,
    pagination_settings: settings,
    font_inventory_json: encodeNativeDOCXFontInventoryV1(inventory),
    font_assets: [{ face_id: 'font-face:dejavu', face_slot: 'embedRegular', resource_id: `font:${FONT_DIGEST}`, content_digest: FONT_DIGEST, collection_index: null, bytes: FONT_BYTES }],
    media_assets: [],
  }
}

function tableFixture(): NativeDocxPagePaintPrepareInputV1 {
  const input = fixture()
  const document = input.document as NativeDocxDocumentV1
  const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
  const paragraph = document.body.blocks[0]!.paragraph!
  const table = {
    id: 'table:1', anchor: anchor('/w:document[1]/w:body[1]/w:tbl[1]', 90, 200), edit_policy: { mode: 'read-only' as const, allowed_operations: [], refusal: { code: 'NATIVE_READ_ONLY', message: 'Fixture is immutable.', preservation: 'refuse-mutation' as const } },
    width_twips: 9_360, layout: 'fixed' as const, alignment: 'left' as const, indent_twips: 0, grid_widths_twips: [9_360],
    cell_margins: { top_twips: 100, right_twips: 100, bottom_twips: 100, left_twips: 100 },
    borders: { top: { style: 'single' as const, size_eighth_points: 8, color_rgb: '000000' }, right: { style: 'single' as const, size_eighth_points: 8, color_rgb: '000000' }, bottom: { style: 'single' as const, size_eighth_points: 8, color_rgb: '000000' }, left: { style: 'single' as const, size_eighth_points: 8, color_rgb: '000000' } },
    rows: [{ id: 'row:1', anchor: anchor('/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]', 95, 195), repeat_header: false, cant_split: true, cells: [{ id: 'cell:1', anchor: anchor('/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]', 98, 192), width_twips: 9_360, grid_span: 1, vertical_merge: 'none' as const, shading_rgb: 'DDEEFF', paragraphs: [paragraph] }] }],
  }
  document.body.blocks = [{ kind: 'table', id: table.id, table }]
  document.sections[0]!.starts_at_block_id = table.id
  resolved.tables = [{ table_id: table.id }]
  return input
}

function rewriteInventory(input: NativeDocxPagePaintPrepareInputV1, mutate: (inventory: NativeDOCXFontInventoryV1) => void): void {
  const inventory = JSON.parse(input.font_inventory_json) as NativeDOCXFontInventoryV1
  mutate(inventory)
  inventory.inventory_sha256 = ''
  inventory.inventory_sha256 = nativeDOCXCanonicalWireSHA256V1(inventory)
  input.font_inventory_json = encodeNativeDOCXFontInventoryV1(inventory)
}

/** Two next-page sections that differ only in their left and right margins. */
function twoSectionFixture(): NativeDocxPagePaintPrepareInputV1 {
  const input = fixture()
  const document = input.document as NativeDocxDocumentV1
  const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
  const first = document.body.blocks[0]!.paragraph!
  const second = structuredClone(first)
  second.id = 'paragraph:2'
  second.anchor = anchor('/w:document[1]/w:body[1]/w:p[2]', 200, 290)
  second.runs = [{ ...second.runs[0]!, id: 'run:2', anchor: anchor('/w:document[1]/w:body[1]/w:p[2]/w:r[1]', 210, 280), text: 'B' }]
  document.body.blocks.push({ kind: 'paragraph', id: second.id, paragraph: second })
  resolved.paragraphs.push({ ...structuredClone(resolved.paragraphs[0]!), paragraph_id: second.id })
  resolved.runs.push({ ...structuredClone(resolved.runs[0]!), run_id: 'run:2', paragraph_id: second.id })
  rewriteInventory(input, (inventory) => { inventory.references[0]!.scope_ids = ['paragraph:1', 'paragraph:2', 'run:1', 'run:2'] })
  const narrow = structuredClone(document.sections[0]!)
  narrow.id = 'section:2'
  narrow.anchor = anchor('/w:document[1]/w:body[1]/w:sectPr[2]', 2_200, 2_290)
  narrow.starts_at_block_id = second.id
  narrow.page.margins.left_twips = 2_880
  narrow.page.margins.right_twips = 2_880
  narrow.page.column_definitions = [{ id: 'column:section:2:0', ordinal: 0 }]
  document.sections.push(narrow)
  return input
}

function imageFixture(): NativeDocxPagePaintPrepareInputV1 {
  const input = fixture()
  const document = input.document as NativeDocxDocumentV1
  const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
  const paragraph = document.body.blocks[0]!.paragraph!
  paragraph.runs.splice(0, 0, {
    kind: 'drawing', id: 'run:image', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]', 105, 109),
    drawing: {
      id: 'drawing:1', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:drawing[1]', 106, 108),
      relationship_id: 'rImage', media_part: 'word/media/image.png', content_type: 'image/png', placement: 'inline',
      width_emu: 127_000, height_emu: 127_000,
      edit_policy: { mode: 'read-only', allowed_operations: [], refusal: { code: 'EXTRACT_ONLY', message: 'Fixture drawing is immutable.', preservation: 'refuse-mutation' } },
    },
  })
  resolved.runs.splice(0, 0, { run_id: 'run:image', paragraph_id: paragraph.id, applied_paragraph_styles: [], applied_character_styles: [], properties: {} })
  document.passthrough_parts.push({ part_name: 'word/media/image.png', content_type: 'image/png', byte_length: PNG_BYTES.byteLength, sha256: PNG_DIGEST, policy: 'preserve-verbatim' })
  input.media_assets = [{ part_name: 'word/media/image.png', content_type: 'image/png', content_digest: PNG_DIGEST, bytes: PNG_BYTES }]
  return input
}

function combinedImageTableFixture(): NativeDocxPagePaintPrepareInputV1 {
  const input = imageFixture()
  const document = input.document as NativeDocxDocumentV1
  const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
  const paragraph = structuredClone(document.body.blocks[0]!.paragraph!)
  paragraph.id = 'paragraph:table'
  paragraph.anchor = anchor('/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]/w:p[1]', 210, 290)
  paragraph.runs = [{ kind: 'text', id: 'run:table', anchor: anchor('/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]/w:p[1]/w:r[1]', 220, 280), text: 'A' }]
  const table = {
    id: 'table:combined', anchor: anchor('/w:document[1]/w:body[1]/w:tbl[1]', 200, 300), edit_policy: { mode: 'read-only' as const, allowed_operations: [], refusal: { code: 'NATIVE_READ_ONLY', message: 'Fixture is immutable.', preservation: 'refuse-mutation' as const } },
    width_twips: 9_360, layout: 'fixed' as const, alignment: 'left' as const, indent_twips: 0, grid_widths_twips: [9_360],
    cell_margins: { top_twips: 100, right_twips: 100, bottom_twips: 100, left_twips: 100 },
    borders: { top: { style: 'single' as const, size_eighth_points: 8, color_rgb: '000000' }, right: { style: 'single' as const, size_eighth_points: 8, color_rgb: '000000' }, bottom: { style: 'single' as const, size_eighth_points: 8, color_rgb: '000000' }, left: { style: 'single' as const, size_eighth_points: 8, color_rgb: '000000' } },
    rows: [{ id: 'row:combined', anchor: anchor('/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]', 205, 295), repeat_header: false, cant_split: true, cells: [{ id: 'cell:combined', anchor: anchor('/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]', 208, 292), width_twips: 9_360, grid_span: 1, vertical_merge: 'none' as const, shading_rgb: 'DDEEFF', paragraphs: [paragraph] }] }],
  }
  document.body.blocks.push({ kind: 'table', id: table.id, table })
  resolved.paragraphs.push({ paragraph_id: paragraph.id, applied_styles: [], properties: {}, paragraph_mark_properties: { font_family: 'DejaVu Sans', font_size_half_points: 20 } })
  resolved.runs.push({ run_id: 'run:table', paragraph_id: paragraph.id, applied_paragraph_styles: [], applied_character_styles: [], properties: { font_family: 'DejaVu Sans', font_size_half_points: 20, color: '123456' } })
  resolved.tables.push({ table_id: table.id })
  rewriteInventory(input, (inventory) => { inventory.references[0]!.scope_ids = ['paragraph:1', 'paragraph:table', 'run:1', 'run:table'] })
  return input
}

function noteFixture(): NativeDocxPagePaintPrepareInputV1 {
  const input = fixture()
  const document = input.document as NativeDocxDocumentV1
  const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
  const bodyRun = document.body.blocks[0]!.paragraph!.runs[0]!
  bodyRun.kind = 'reference'; delete bodyRun.text
  bodyRun.reference = { kind: 'footnote', target_id: 'story:footnote:1' }
  bodyRun.properties = { character_style_id: 'FootnoteReference' }
  const part = 'word/footnotes.xml'
  const noteAnchor = (path: string, start: number, end: number) => ({ part_name: part, path, start_byte: start, end_byte: end, xml_sha256: HASH })
  const policy = { mode: 'read-only' as const, allowed_operations: [], refusal: { code: 'NATIVE_READ_ONLY', message: 'Fixture note is immutable.', preservation: 'refuse-mutation' as const } }
  const separatorParagraph = {
    id: 'paragraph:footnote-separator', anchor: noteAnchor('/w:footnotes[1]/w:footnote[1]/w:p[1]', 10, 60), edit_policy: policy, properties: {},
    runs: [] as NativeDocxRunV1[],
  }
  const noteParagraph = {
    id: 'paragraph:footnote-1', anchor: noteAnchor('/w:footnotes[1]/w:footnote[2]/w:p[1]', 80, 180), edit_policy: policy, properties: {},
    runs: [
      { kind: 'reference' as const, id: 'run:footnote-label-1', anchor: noteAnchor('/w:footnotes[1]/w:footnote[2]/w:p[1]/w:r[1]', 90, 110), properties: { character_style_id: 'FootnoteReference' }, reference: { kind: 'footnote' as const, target_id: 'story:footnote:1', role: 'label' as const } },
      { kind: 'text' as const, id: 'run:footnote-text-1', anchor: noteAnchor('/w:footnotes[1]/w:footnote[2]/w:p[1]/w:r[2]', 111, 170), text: ' Native note' },
    ],
  }
  document.notes = [
    { id: 'story:footnote-separator', kind: 'footnote', part_name: part, native_story_id: '-1', relationship_id: 'rIdFootnotes', note_role: 'separator', anchor: noteAnchor('/w:footnotes[1]/w:footnote[1]', 1, 70), blocks: [{ kind: 'paragraph', id: separatorParagraph.id, paragraph: separatorParagraph }] },
    { id: 'story:footnote:1', kind: 'footnote', part_name: part, native_story_id: '1', relationship_id: 'rIdFootnotes', note_role: 'content', anchor: noteAnchor('/w:footnotes[1]/w:footnote[2]', 71, 190), blocks: [{ kind: 'paragraph', id: noteParagraph.id, paragraph: noteParagraph }] },
  ]
  const properties = { font_family: 'DejaVu Sans', font_size_half_points: 20, color: '123456' }
  const bodyResolved = resolved.runs.find((run) => run.run_id === bodyRun.id)!
  bodyResolved.character_style_id = 'FootnoteReference'
  bodyResolved.applied_character_styles = ['FootnoteReference']
  for (const paragraph of [separatorParagraph, noteParagraph]) {
    resolved.paragraphs.push({ paragraph_id: paragraph.id, applied_styles: [], properties: {}, paragraph_mark_properties: properties })
    for (const run of paragraph.runs) {
      const reference = run.kind === 'reference'
      resolved.runs.push({
        run_id: run.id,
        paragraph_id: paragraph.id,
        ...(reference ? { character_style_id: 'FootnoteReference' } : {}),
        applied_paragraph_styles: [],
        applied_character_styles: reference ? ['FootnoteReference'] : [],
        properties: { ...properties },
      })
    }
  }
  rewriteInventory(input, (inventory) => {
    inventory.references[0]!.scope_ids.push(separatorParagraph.id, noteParagraph.id, ...separatorParagraph.runs.map((run) => run.id), ...noteParagraph.runs.map((run) => run.id))
    inventory.references[0]!.scope_ids.sort()
  })
  return input
}

function endnoteFixture(): NativeDocxPagePaintPrepareInputV1 {
  const input = noteFixture()
  const document = input.document as NativeDocxDocumentV1
  for (const story of document.notes) {
    story.kind = 'endnote'
    story.id = story.id.replace('footnote', 'endnote')
    story.part_name = 'word/endnotes.xml'
    story.relationship_id = 'rIdEndnotes'
    story.anchor.part_name = story.part_name
    for (const block of story.blocks) if (block.paragraph) {
      block.paragraph.anchor.part_name = story.part_name
      for (const run of block.paragraph.runs) {
        run.anchor.part_name = story.part_name
        if (run.reference) { run.reference.kind = 'endnote'; run.reference.target_id = story.id; run.properties = { ...run.properties, character_style_id: 'EndnoteReference' } }
      }
    }
  }
  for (const run of (input.resolved_layout as NativeDocxResolvedLayoutInputV1).runs) {
    if (run.character_style_id !== 'FootnoteReference') continue
    run.character_style_id = 'EndnoteReference'
    run.applied_character_styles = ['EndnoteReference']
  }
  document.body.blocks[0]!.paragraph!.runs[0]!.reference = { kind: 'endnote', target_id: 'story:endnote:1' }
  document.body.blocks[0]!.paragraph!.runs[0]!.properties = { character_style_id: 'EndnoteReference' }
  return input
}

function continuedEndnotePaintFixture(): NativeDocxPagePaintPrepareInputV1 {
  const input = endnoteFixture()
  const document = input.document as NativeDocxDocumentV1
  const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
  document.sections[0]!.page.margins.bottom_twips = 13_200
  const separator = document.notes[0]!
  const note = document.notes[1]!
  note.anchor.end_byte = 2_000
  const original = note.blocks[0]!.paragraph!
  const originalResolved = resolved.paragraphs.find((entry) => entry.paragraph_id === original.id)!
  originalResolved.properties.keep_lines = true
  const originalRun = original.runs[1]!
  const originalResolvedRun = resolved.runs.find((entry) => entry.run_id === originalRun.id)!
  const scopes: string[] = []
  for (let index = 2; index <= 7; index += 1) {
    const paragraph = structuredClone(original)
    paragraph.id = `paragraph:endnote:${index}`
    paragraph.anchor = { ...paragraph.anchor, path: `/w:endnotes[1]/w:endnote[2]/w:p[${index}]`, start_byte: index * 200, end_byte: index * 200 + 100 }
    const run = { ...structuredClone(originalRun), id: `run:endnote:${index}`, text: ` Endnote paragraph ${index}`, anchor: { ...paragraph.anchor, path: `${paragraph.anchor.path}/w:r[1]`, start_byte: paragraph.anchor.start_byte + 10, end_byte: paragraph.anchor.end_byte - 10 } }
    paragraph.runs = [run]
    note.blocks.push({ kind: 'paragraph', id: paragraph.id, paragraph })
    resolved.paragraphs.push({ ...structuredClone(originalResolved), paragraph_id: paragraph.id })
    resolved.runs.push({ ...structuredClone(originalResolvedRun), paragraph_id: paragraph.id, run_id: run.id })
    scopes.push(paragraph.id, run.id)
  }
  const continuation = structuredClone(separator)
  continuation.id = 'story:endnote:continuation'
  continuation.note_role = 'continuation-separator'
  continuation.native_story_id = '0'
  continuation.blocks[0]!.id = 'paragraph:endnote:continuation'
  continuation.blocks[0]!.paragraph!.id = continuation.blocks[0]!.id
  document.notes.push(continuation)
  resolved.paragraphs.push({ ...structuredClone(resolved.paragraphs.find((entry) => entry.paragraph_id === separator.blocks[0]!.id)!), paragraph_id: continuation.blocks[0]!.id })
  scopes.push(continuation.blocks[0]!.id)
  rewriteInventory(input, (inventory) => { inventory.references[0]!.scope_ids.push(...scopes); inventory.references[0]!.scope_ids.sort() })
  return input
}

function combinedNoteImageTableHeaderFixture(): NativeDocxPagePaintPrepareInputV1 {
  const input = noteFixture()
  const document = input.document as NativeDocxDocumentV1
  const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
  const bodyParagraph = document.body.blocks[0]!.paragraph!
  bodyParagraph.runs.unshift({
    kind: 'drawing', id: 'run:image', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]', 105, 109),
    drawing: {
      id: 'drawing:1', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:drawing[1]', 106, 108),
      relationship_id: 'rImage', media_part: 'word/media/image.png', content_type: 'image/png', placement: 'inline',
      width_emu: 127_000, height_emu: 127_000,
      edit_policy: { mode: 'read-only', allowed_operations: [], refusal: { code: 'EXTRACT_ONLY', message: 'Fixture drawing is immutable.', preservation: 'refuse-mutation' } },
    },
  })
  resolved.runs.unshift({ run_id: 'run:image', paragraph_id: bodyParagraph.id, applied_paragraph_styles: [], applied_character_styles: [], properties: {} })
  document.passthrough_parts.push({ part_name: 'word/media/image.png', content_type: 'image/png', byte_length: PNG_BYTES.byteLength, sha256: PNG_DIGEST, policy: 'preserve-verbatim' })
  input.media_assets = [{ part_name: 'word/media/image.png', content_type: 'image/png', content_digest: PNG_DIGEST, bytes: PNG_BYTES }]

  const tableParagraph = {
    id: 'paragraph:table:combined', anchor: anchor('/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]/w:p[1]', 210, 290),
    edit_policy: { mode: 'read-only' as const, allowed_operations: [], refusal: { code: 'NATIVE_READ_ONLY', message: 'Fixture is immutable.', preservation: 'refuse-mutation' as const } },
    properties: {}, runs: [{ kind: 'text' as const, id: 'run:table:combined', anchor: anchor('/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]/w:p[1]/w:r[1]', 220, 280), text: 'Table' }],
  }
  const table = {
    id: 'table:combined:notes', anchor: anchor('/w:document[1]/w:body[1]/w:tbl[1]', 200, 300), edit_policy: { mode: 'read-only' as const, allowed_operations: [], refusal: { code: 'NATIVE_READ_ONLY', message: 'Fixture is immutable.', preservation: 'refuse-mutation' as const } },
    width_twips: 9_360, layout: 'fixed' as const, alignment: 'left' as const, indent_twips: 0, grid_widths_twips: [9_360],
    cell_margins: { top_twips: 100, right_twips: 100, bottom_twips: 100, left_twips: 100 },
    borders: { top: { style: 'single' as const, size_eighth_points: 8, color_rgb: '000000' }, right: { style: 'single' as const, size_eighth_points: 8, color_rgb: '000000' }, bottom: { style: 'single' as const, size_eighth_points: 8, color_rgb: '000000' }, left: { style: 'single' as const, size_eighth_points: 8, color_rgb: '000000' } },
    rows: [{ id: 'row:combined:notes', anchor: anchor('/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]', 205, 295), repeat_header: false, cant_split: true, cells: [{ id: 'cell:combined:notes', anchor: anchor('/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]', 208, 292), width_twips: 9_360, grid_span: 1, vertical_merge: 'none' as const, shading_rgb: 'DDEEFF', paragraphs: [tableParagraph] }] }],
  }
  document.body.blocks.push({ kind: 'table', id: table.id, table })
  resolved.paragraphs.push({ paragraph_id: tableParagraph.id, applied_styles: [], properties: {}, paragraph_mark_properties: { font_family: 'DejaVu Sans', font_size_half_points: 20 } })
  resolved.runs.push({ run_id: 'run:table:combined', paragraph_id: tableParagraph.id, applied_paragraph_styles: [], applied_character_styles: [], properties: { font_family: 'DejaVu Sans', font_size_half_points: 20, color: '123456' } })
  resolved.tables.push({ table_id: table.id })

  const headerPart = 'word/header1.xml'
  const headerAnchor = (path: string, start: number, end: number) => ({ part_name: headerPart, path, start_byte: start, end_byte: end, xml_sha256: HASH })
  const headerParagraph = {
    id: 'paragraph:header:combined', anchor: headerAnchor('/w:hdr[1]/w:p[1]', 10, 90),
    edit_policy: { mode: 'read-only' as const, allowed_operations: [], refusal: { code: 'NATIVE_READ_ONLY', message: 'Fixture is immutable.', preservation: 'refuse-mutation' as const } },
    properties: {}, runs: [{ kind: 'text' as const, id: 'run:header:combined', anchor: headerAnchor('/w:hdr[1]/w:p[1]/w:r[1]', 20, 80), text: 'Header' }],
  }
  document.headers.push({ id: 'story:header:combined', kind: 'header', part_name: headerPart, anchor: headerAnchor('/w:hdr[1]', 1, 100), blocks: [{ kind: 'paragraph', id: headerParagraph.id, paragraph: headerParagraph }] })
  document.sections[0]!.header_refs.push({ kind: 'default', story_id: 'story:header:combined', relationship_id: 'rIdHeader' })
  resolved.paragraphs.push({ paragraph_id: headerParagraph.id, applied_styles: [], properties: {}, paragraph_mark_properties: { font_family: 'DejaVu Sans', font_size_half_points: 20 } })
  resolved.runs.push({ run_id: 'run:header:combined', paragraph_id: headerParagraph.id, applied_paragraph_styles: [], applied_character_styles: [], properties: { font_family: 'DejaVu Sans', font_size_half_points: 20, color: '123456' } })

  rewriteInventory(input, (inventory) => {
    inventory.references[0]!.scope_ids.push(tableParagraph.id, 'run:table:combined', headerParagraph.id, 'run:header:combined')
    inventory.references[0]!.scope_ids.sort()
  })
  return input
}

/** The qualified floating record for one drawing, or a loud failure. */
function qualifiedFloating(document: NativeDocxDocumentV1, drawing: NativeDocxDocumentV1['body']['blocks'][number]['paragraph'] extends undefined ? never : NonNullable<NativeDocxDocumentV1['body']['blocks'][number]['paragraph']>['runs'][number]['drawing']) {
  const run = document.body.blocks.flatMap(block => block.paragraph?.runs ?? []).find(candidate => candidate.drawing === drawing)
  const qualified = qualifyNativeDocxInlineImageV1(document, run!.id, drawing!)
  if (!qualified.ok || !qualified.value.floating) throw new Error(`floating anchor did not qualify: ${qualified.ok ? 'not floating' : qualified.message}`)
  return qualified.value.floating
}

describe('native DOCX page-paint compiler v1', () => {
  it.each([true,false])('replays border reservation in glyphs, following paragraphs and page breaks, cant_split %s',async(cantSplit)=>{
    const input=tableFixture(),document=input.document as NativeDocxDocumentV1,resolved=input.resolved_layout as NativeDocxResolvedLayoutInputV1,settings=input.pagination_settings as NativeDocxPaginationSettingsV1
    const table=document.body.blocks[0]!.table!
    table.rows[0]!.cant_split=cantSplit
    table.borders!.top!.size_eighth_points=4;table.borders!.bottom!.size_eighth_points=4;table.borders!.inside_horizontal={...table.borders!.top!}
    table.cell_margins!.top_twips=0;table.cell_margins!.bottom_twips=0
    const following=structuredClone(table.rows[0]!.cells[0]!.paragraphs[0]!);following.id='paragraph:following';following.runs[0]!.id='run:following'
    following.anchor=anchor('/w:document[1]/w:body[1]/w:p[1]',201,299);following.runs[0]!.anchor=anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]',210,290)
    document.body.blocks.push({kind:'paragraph',id:following.id,paragraph:following})
    resolved.paragraphs.push({...structuredClone(resolved.paragraphs[0]!),paragraph_id:following.id})
    resolved.runs.push({...structuredClone(resolved.runs[0]!),run_id:following.runs[0]!.id,paragraph_id:following.id})
    rewriteInventory(input,i=>{i.references[0]!.scope_ids.push(following.id,following.runs[0]!.id);i.references[0]!.scope_ids.sort()})
    settings.profile='unsupported';delete settings.compatibility_mode
    settings.diagnostics=[{code:'COMPATIBILITY_SETTING_UNSUPPORTED',severity:'unsupported',part_name:SETTINGS_PART,path:'/w:settings[1]/w:compat[1]',preservation:'preserve-verbatim',message:'Legacy mode'}]
    const eligibility={protocol:'injoffice.docx.approximation-eligibility',version:1,document_id:settings.document_id,revision:settings.revision,package_sha256:HASH,settings_sha256:settings.settings_sha256,status:'eligible',legacy_compatibility_mode:12,reasons:['Legacy mode']}
    const outlines=createHarfBuzzOutlineProviderV1({bytes:FONT_BYTES,contentDigest:FONT_DIGEST})
    const provider={providerId:input.outline_provider.provider_id,providerRevision:input.outline_provider.provider_revision,getGlyphOutline(request:import('./nativePagePaintV1.js').NativeDocxGlyphOutlineRequestV1){const o=outlines.outline(request.glyph_id);return o.path.length?{status:'outlined' as const,...request,...o}:{status:'empty' as const,...request,units_per_em:o.units_per_em}}}
    const before=structuredClone(input),paint=await renderNativeDocxApproximatePagePreviewV1(input,eligibility,provider),baseline=await renderNativeDocxApproximatePagePreviewV1(input,{...eligibility,legacy_compatibility_mode:14},provider)
    expect(paint.status).toBe('painted');expect(baseline.status).toBe('painted')
    expect(paint.reasons).toContain(DOCX_TABLE_BORDER_RESERVATION_WARNING)
    expect(paint.reasons).toContain(DOCX_TABLE_GRID_FIT_WARNING)
    expect(paint.table_width_policy).toBe('approximate-authored-grid-fitted-v1')
    expect(decodeNativeDocxApproximatePagePreviewV1({...paint,reasons:paint.reasons.filter(r=>r!==DOCX_TABLE_GRID_FIT_WARNING)}).ok).toBe(false)
    expect(decodeNativeDocxApproximatePagePreviewV1({...paint,table_width_policy:'unknown'}).ok).toBe(false)
    const topBorder=(value:typeof paint)=>value.pages[0]!.commands.find(c=>c.kind==='stroke_table_border'&&c.edge==='top')
    expect(topBorder(paint)).toEqual(topBorder(baseline))
    const fill=(value:typeof paint)=>{const command=value.pages[0]!.commands.find(c=>c.kind==='fill_table_cell');if(!command||command.kind!=='fill_table_cell')throw new Error('Expected table fill');return command}
    expect(fill(paint)).toEqual({...fill(baseline),height_millipoints:fill(baseline).height_millipoints+500})
    expect(decodeNativeDocxApproximatePagePreviewV1({...paint,reasons:paint.reasons.filter(r=>r!==DOCX_TABLE_BORDER_RESERVATION_WARNING)}).ok).toBe(false)
    const glyphs=paint.pages[0]!.commands.filter(c=>c.kind==='fill_glyph_path'),oldGlyphs=baseline.pages[0]!.commands.filter(c=>c.kind==='fill_glyph_path')
    expect(glyphs.length).toBeGreaterThan(1)
    for(const [i,g]of glyphs.entries())expect(g.path).toEqual(oldGlyphs[i]!.path.map(part=>Object.fromEntries(Object.entries(part).map(([k,v])=>[k,k.endsWith('y_millipoints')?(v as number)+500:v]))))
    const strict=await prepareNativeDocxPagePaintV1(input),request=structuredClone(strict.page_paint_request)
    expect(request.paginated_layout.status).toBe('refused')
    const q=qualifyApproximateLegacyTables(document,resolved,request.pagination_request.shaped_lines,eligibility)
    if(q.status!=='qualified')throw new Error('Expected policy table')
    request.paginated_layout=paginateNativeDocxApproximateLegacyV1(request.pagination_request,eligibility).layout
    request.integrity.paginated_layout_sha256=nativeDocxPagePaintPaginatedLayoutSha256V1(request.paginated_layout);request.integrity.table_projection_sha256=q.sha256
    expect(decodeNativeDocxApproximateComputedPagePaintV1(request,eligibility).fidelity).toBe('approximate')
    const forged=structuredClone(request);forged.paginated_layout.pages[0]!.lines[0]!.y_millipoints-=500;forged.integrity.paginated_layout_sha256=nativeDocxPagePaintPaginatedLayoutSha256V1(forged.paginated_layout)
    expect(()=>decodeNativeDocxApproximateComputedPagePaintV1(forged,eligibility)).toThrow()
    const short=structuredClone(request.pagination_request),oldLayout=paginateNativeDocxApproximateLegacyV1(short,{...eligibility,legacy_compatibility_mode:14}).layout
    const last=oldLayout.pages[0]!.lines.at(-1)!,page=short.document.sections[0]!.page
    page.margins.bottom_twips=page.height_twips-Math.ceil((last.y_millipoints+last.height_millipoints)/50)
    const shortBaseline=paginateNativeDocxApproximateLegacyV1(short,{...eligibility,legacy_compatibility_mode:14}).layout
    expect(shortBaseline.pages,JSON.stringify(shortBaseline.diagnostics)).toHaveLength(1)
    expect(paginateNativeDocxApproximateLegacyV1(short,eligibility).layout.pages).toHaveLength(2)
    const fieldInput=structuredClone(input),fieldDocument=fieldInput.document as NativeDocxDocumentV1
    const field=fieldDocument.body.blocks[1]!.paragraph!.runs[0]!;field.page_field='PAGE';field.text=''
    const fieldPaint=await renderNativeDocxApproximatePagePreviewV1(fieldInput,eligibility,provider)
    expect(fieldPaint.status).toBe('painted')
    expect(fieldPaint.rendering_provenance.body_field_source_sha256).toBeDefined()
    expect(topBorder(fieldPaint)).toEqual(topBorder(paint))
    if(cantSplit){
      const wrapped=structuredClone(fieldInput),wrappedDocument=wrapped.document as NativeDocxDocumentV1,t=wrappedDocument.body.blocks[0]!.table!
      t.layout='autofit';t.width_twips=1000;t.rows[0]!.cells[0]!.paragraphs[0]!.runs[0]!.text='A A A A A A A A A A A A A A A A A A A A'
      const modernInput={...wrapped,pagination_settings:fixture().pagination_settings},modern=await prepareNativeDocxPagePaintV1(modernInput),shaped=modern.page_paint_request.pagination_request.shaped_lines
      expect(shaped.paragraphs.find(p=>p.paragraph_id===t.rows[0]!.cells[0]!.paragraphs[0]!.id)!.lines.length).toBeGreaterThan(1)
      const finalTables=qualifyApproximateLegacyTables(wrappedDocument,wrapped.resolved_layout as NativeDocxResolvedLayoutInputV1,shaped,eligibility)
      if(finalTables.status!=='qualified')throw new Error('Expected wrapped autofit')
      expect(finalTables.tables[0]!.border_reservation_policy).toBeUndefined()
      expect((await renderNativeDocxApproximatePagePreviewV1(wrapped,eligibility,provider)).status).toBe('painted')
    }
    expect(input).toEqual(before)
  },20000)
  it.each([2,4,8])('reserves authored %s eighth-point horizontal borders only in legacy approximation',async size=>{
    const input=tableFixture(),document=input.document as NativeDocxDocumentV1,resolved=input.resolved_layout as NativeDocxResolvedLayoutInputV1
    const table=document.body.blocks[0]!.table!
    table.borders!.top!.size_eighth_points=size;table.borders!.bottom!.size_eighth_points=size
    table.borders!.inside_horizontal={...table.borders!.top!}
    // Unequal padding discriminates the content top from the bottom limit.
    table.cell_margins!.top_twips=40;table.cell_margins!.bottom_twips=80
    const before=structuredClone(input),prepared=await prepareNativeDocxPagePaintV1(input),request=prepared.page_paint_request.pagination_request
    const strict=qualifyNativeDocxTablesV1(document,resolved,request.shaped_lines),approx=qualifyApproximateLegacyTables(document,resolved,request.shaped_lines,{legacy_compatibility_mode:12})
    if(strict.status!=='qualified'||approx.status!=='qualified')throw new Error('Expected tables')
    const a=layoutNativeDocxTableRowsV1(approx.tables[0]!,request.shaped_lines)![0]!,b=layoutNativeDocxTableRowsV1(strict.tables[0]!,request.shaped_lines)![0]!
    expect(a.height_millipoints-b.height_millipoints).toBe(size*125)
    expect(a.cells[0]!.content_y_millipoints-b.cells[0]!.content_y_millipoints).toBe(size*125)
    expect(a.cells[0]!.content_height_millipoints).toBe(b.cells[0]!.content_height_millipoints)
    expect(a.cells[0]!.width_millipoints).toBe(b.cells[0]!.width_millipoints)
    expect(approx.sha256).not.toBe(strict.sha256)
    expect(strict.tables[0]!.border_reservation_policy).toBeUndefined()
    expect(qualifyApproximateLegacyTables(document,resolved,request.shaped_lines,{legacy_compatibility_mode:14})).toEqual(strict)
    const taller=structuredClone(request.shaped_lines);taller.paragraphs[0]!.lines[0]!.line_height_millipoints+=2000
    expect(layoutNativeDocxTableRowsV1(approx.tables[0]!,taller)![0]!.height_millipoints-a.height_millipoints).toBe(2000)
    const multiline=structuredClone(request.shaped_lines);multiline.paragraphs[0]!.lines.push({...multiline.paragraphs[0]!.lines[0]!,id:'line:extra',ordinal:1})
    const outOfProfile=qualifyApproximateLegacyTables(document,resolved,multiline,{legacy_compatibility_mode:12})
    if(outOfProfile.status!=='qualified')throw new Error('Expected unchanged strict table geometry')
    expect(outOfProfile.tables[0]!.border_reservation_policy).toBeUndefined()
    for(const change of ['missing','unequal','none','exact','minimum','merged'] as const){
      const copy=structuredClone(document),t=copy.body.blocks[0]!.table!
      if(change==='missing')delete t.borders!.inside_horizontal
      if(change==='unequal')t.borders!.bottom!.size_eighth_points+=1
      if(change==='none')t.borders!.bottom={style:'none',size_eighth_points:0}
      if(change==='exact'||change==='minimum'){t.rows[0]!.height_rule=change==='exact'?'exact':'atLeast';t.rows[0]!.height_twips=1000}
      if(change==='merged')t.rows[0]!.cells[0]!.vertical_merge='restart'
      const q=qualifyApproximateLegacyTables(copy,resolved,request.shaped_lines,{legacy_compatibility_mode:12})
      if(q.status==='qualified')expect(q.tables[0]!.border_reservation_policy).toBeUndefined()
    }
    expect(input).toEqual(before)
  },15000)
  it('applies the declared host family only to a package that states no font anywhere', () => {
    const input = fixture()
    const document = input.document as NativeDocxDocumentV1
    const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    const paragraph = document.body.blocks[0]!.paragraph!
    const policy = { kind: 'host-default-family-v1', family: 'Aptos' }
    const references = (layout: NativeDocxResolvedLayoutInputV1) => [...layout.runs.map(run => run.properties), ...layout.paragraphs.map(p => p.paragraph_mark_properties)]
      .filter(properties => properties.font_family !== undefined)
      .map(properties => ({ family: properties.font_family!, weight: properties.bold === true ? 700 : 400, style: properties.italic === true ? 'italic' : 'normal' }))
    const facts = [
      { scope_kind: 'paragraph-mark' as const, scope_id: paragraph.id, part_name: paragraph.anchor.part_name, path: paragraph.anchor.path, package_sha256: HASH },
      { scope_kind: 'run' as const, scope_id: paragraph.runs[0]!.id, part_name: paragraph.runs[0]!.anchor.part_name, path: paragraph.runs[0]!.anchor.path, package_sha256: HASH },
    ]
    // A package that states a family keeps using it: the whole-package precondition refuses.
    expect(() => projectNativeDocxAbsentFontFamiliesV1(document, resolved, facts, policy, references, () => true)).toThrow('no font reference at all')
    const partial = structuredClone(resolved)
    delete partial.runs[0]!.properties.font_family
    expect(() => projectNativeDocxAbsentFontFamiliesV1(document, partial, [facts[1]!], policy, references, () => true)).toThrow('no font reference at all')
    // A resolved family is never overridden even if a caller claims no references.
    expect(() => projectNativeDocxAbsentFontFamiliesV1(document, resolved, [facts[0]!], policy, () => [], () => true)).toThrow('override')
    const fontless = structuredClone(resolved)
    delete fontless.paragraphs[0]!.paragraph_mark_properties.font_family
    for (const run of fontless.runs) delete run.properties.font_family
    const projected = projectNativeDocxAbsentFontFamiliesV1(document, fontless, facts, policy, references, () => true)
    expect(projected.applied).toEqual(facts.map(fact => ({ ...fact, chosen_family: 'Aptos' })))
    expect(projected.resolved.runs[0]!.properties.font_family).toBe('Aptos')
    expect(projected.resolved.paragraphs[0]!.paragraph_mark_properties.font_family).toBe('Aptos')
    expect(fontless.runs[0]!.properties.font_family).toBeUndefined()
    // An unattested host face is recorded nowhere and the scope stays unshaped.
    expect(projectNativeDocxAbsentFontFamiliesV1(document, fontless, facts, policy, references, () => false).applied).toEqual([])
    // The strict view removes the projection so the strict inventory still joins.
    expect(stripNativeDocxAbsentFontFamiliesV1(projected.resolved, facts)).toEqual(fontless)
    expect(stripNativeDocxAbsentFontFamiliesV1(fontless, [])).toBe(fontless)
    expect(() => projectNativeDocxAbsentFontFamiliesV1(document, fontless, facts, { kind: 'host-default-family-v1', family: 'Calibri' }, references, () => true)).toThrow('explicit')
    expect(() => projectNativeDocxAbsentFontFamiliesV1(document, fontless, [{ ...facts[1]!, path: '/w:document[1]/w:body[1]/w:p[9]/w:r[1]' }], policy, references, () => true)).toThrow('scope anchor')
    expect(projectNativeDocxAbsentFontFamiliesV1(document, resolved, [], policy, references, () => true).applied).toEqual([])
  })
  it('limits note host-size policy to clean empty reserved separator paragraphs', () => {
    const input = noteFixture()
    const document = input.document as NativeDocxDocumentV1
    const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    const paragraph = document.notes[0]!.blocks[0]!.paragraph!
    const target = resolved.paragraphs.find(p => p.paragraph_id === paragraph.id)!
    target.paragraph_mark_properties = { ...target.paragraph_mark_properties }
    delete target.paragraph_mark_properties.font_size_half_points
    const fact = { scope_kind: 'paragraph-mark' as const, scope_id: paragraph.id, part_name: paragraph.anchor.part_name, path: paragraph.anchor.path, package_sha256: HASH }
    const shape = 'sizeless-document-defaults' as const
    const policy = { kind: 'host-default-size-v1', half_points: 20 }
    expect(projectNativeDocxAbsentFontSizesV1(document, resolved, [fact], policy, shape).applied).toEqual([{ ...fact, chosen_half_points: 20 }])
    expect(target.paragraph_mark_properties.font_size_half_points).toBeUndefined()
    const content = structuredClone(document)
    content.notes[0]!.note_role = 'content'; content.notes[0]!.native_story_id = '3'
    expect(() => projectNativeDocxAbsentFontSizesV1(content, resolved, [fact], policy, shape)).toThrow('scope anchor')
    const malformed = structuredClone(document)
    malformed.unsupported.push({ id: 'unsupported:note', code: 'UNMODELED_NOTE_MARKUP', capability: 'notes', scope_id: document.notes[0]!.id, anchor: paragraph.anchor, preservation: 'preserve-verbatim', message: 'Malformed instruction remains refused' })
    expect(() => projectNativeDocxAbsentFontSizesV1(malformed, resolved, [fact], policy, shape)).toThrow('scope anchor')
  })
  it('applies a declared host size only to source-proven omissions in the existing approximate envelope', async () => {
    const input = fixture()
    const document = input.document as NativeDocxDocumentV1
    const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    const settings = input.pagination_settings as NativeDocxPaginationSettingsV1
    const paragraph = document.body.blocks[0]!.paragraph!
    paragraph.runs[0]!.text = ''
    delete resolved.paragraphs[0]!.paragraph_mark_properties!.font_size_half_points
    settings.profile = 'unsupported'
    delete settings.compatibility_mode
    settings.diagnostics = [{ code: 'COMPATIBILITY_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: SETTINGS_PART, path: '/w:settings[1]/w:compat[1]', preservation: 'preserve-verbatim', message: 'Legacy mode 12' }]
    const absent = [{ scope_kind: 'paragraph-mark' as const, scope_id: paragraph.id, part_name: paragraph.anchor.part_name, path: paragraph.anchor.path, package_sha256: HASH }]
    const eligibility = { protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: settings.document_id, revision: settings.revision, package_sha256: HASH, settings_sha256: settings.settings_sha256, status: 'eligible', legacy_compatibility_mode: 12, reasons: ['Legacy mode 12 uses current layout'], absent_font_sizes: absent, absent_font_size_shape: 'sizeless-document-defaults' as const }
    const provider = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
    const outline = { providerId: input.outline_provider.provider_id, providerRevision: input.outline_provider.provider_revision, getGlyphOutline(request: import('./nativePagePaintV1.js').NativeDocxGlyphOutlineRequestV1) { const value = provider.outline(request.glyph_id); return value.path.length ? { status: 'outlined' as const, ...request, ...value } : { status: 'empty' as const, ...request, units_per_em: value.units_per_em } } }
    const original = structuredClone(input)
    const withoutPolicy = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outline)
    expect(withoutPolicy.status).toBe('refused')
    const result = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outline, { fontSizePolicy: { kind: 'host-default-size-v1', half_points: 20 } })
    expect(result.status).toBe('painted')
    expect(result.protocol).toBe('injoffice.docx.approximate-page-preview')
    expect(result.approximated_font_sizes).toEqual([{ ...absent[0], chosen_half_points: 20 }])
    expect(result.source_absent_font_sizes).toEqual(absent)
    expect(result.reasons).toContain(DOCX_ABSENT_FONT_SIZE_WARNING)
    expect(input).toEqual(original)
    expect(decodeNativeDocxApproximatePagePreviewV1(result).ok).toBe(true)
    for (const mutation of [{ approximated_font_sizes: undefined }, { source_absent_font_sizes: [] }, { reasons: result.reasons.filter(r => r !== DOCX_ABSENT_FONT_SIZE_WARNING) }, { approximated_font_sizes: [{ ...absent[0], chosen_half_points: 24 }] }, { source_absent_font_size_shape: undefined }, { source_absent_font_size_shape: 'absent-document-defaults' }]) expect(decodeNativeDocxApproximatePagePreviewV1({ ...result, ...mutation }).ok).toBe(false)
    expect(() => projectNativeDocxAbsentFontSizesV1(document, original.resolved_layout, [{ ...absent[0]!, path: '/wrong' }], { kind: 'host-default-size-v1', half_points: 20 }, 'sizeless-document-defaults')).toThrow()
    const authored = structuredClone(resolved); authored.paragraphs[0]!.paragraph_mark_properties!.font_size_half_points = 20
    expect(() => projectNativeDocxAbsentFontSizesV1(document, authored, absent, { kind: 'host-default-size-v1', half_points: 20 }, 'sizeless-document-defaults')).toThrow('override')
    const unrelated = structuredClone(resolved)
    unrelated.diagnostics.push({ code: 'INVALID_FONT_SIZE', severity: 'unsupported', scope_id: 'run:1', part_name: 'word/document.xml', path: '/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:rPr[1]/w:sz[1]', preservation: 'preserve-verbatim', message: 'A malformed independent authored run size must remain refused' })
    expect(projectNativeDocxAbsentFontSizesV1(document, unrelated, absent, { kind: 'host-default-size-v1', half_points: 20 }, 'sizeless-document-defaults').resolved.diagnostics).toEqual(unrelated.diagnostics)
    // The other shape's Word-derived size, and the host's old invented 11 pt,
    // are both refused for a package whose defaults record states no size.
    expect(() => projectNativeDocxAbsentFontSizesV1(document, resolved, absent, { kind: 'host-default-size-v1', half_points: 24 }, 'sizeless-document-defaults')).toThrow('proven source shape')
    expect(() => projectNativeDocxAbsentFontSizesV1(document, resolved, absent, { kind: 'host-default-size-v1', half_points: 22 }, 'sizeless-document-defaults')).toThrow('proven source shape')
    expect(() => projectNativeDocxAbsentFontSizesV1(document, resolved, absent, { kind: 'host-default-size-v1', half_points: 20 }, undefined)).toThrow('proven source shape')
    const tableInput = tableFixture()
    const tableResolved = tableInput.resolved_layout as NativeDocxResolvedLayoutInputV1
    delete tableResolved.paragraphs[0]!.paragraph_mark_properties!.font_size_half_points
    expect(projectNativeDocxAbsentFontSizesV1(tableInput.document, tableResolved, absent, { kind: 'host-default-size-v1', half_points: 20 }, 'sizeless-document-defaults').applied).toEqual([{ ...absent[0], chosen_half_points: 20 }])
    expect(() => projectNativeDocxAbsentFontSizesV1(tableInput.document, tableResolved, [{ ...absent[0]!, path: '/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]/w:p[2]' }], { kind: 'host-default-size-v1', half_points: 20 }, 'sizeless-document-defaults')).toThrow('scope anchor')
    const strict = await prepareNativeDocxPagePaintV1(input)
    expect(strict.page_paint_request.paginated_layout.status).toBe('refused')
  }, 20000)
  function autoBorderFixture() {
    const input = tableFixture()
    const document = input.document as NativeDocxDocumentV1
    const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    const table = document.body.blocks[0]!.table!
    const borders = table.borders!
    delete table.borders
    delete table.rows[0]!.cells[0]!.shading_rgb
    const path = `${table.anchor.path}/w:tblPr[1]/w:tblBorders[1]`
    const diagnostic = { code: 'UNMODELED_TABLE_PROPERTY', scope_id: table.id, part_name: 'word/document.xml', path }
    document.unsupported.push({ id: 'unsupported:auto', code: diagnostic.code, scope_id: table.id, capability: 'table-properties', anchor: anchor(path, 91, 94), preservation: 'preserve-verbatim', message: 'Automatic border color remains unsupported by strict source paint' })
    resolved.tables[0]!.automatic_border_preview = {
      policy: DOCX_AUTO_BORDER_POLICY, read_only: true, package_sha256: HASH,
      page_background: 'absent-on-white-preview', background_rgb: 'FFFFFF', source_part: 'word/document.xml', source_path: path, source_sha256: HASH,
      borders, automatic_edges: ['top', 'right', 'bottom', 'left'], cell_ids: ['cell:1'], source_diagnostics: [diagnostic],
    }
    return input
  }

  it.each([[false,true],[true,true],[false,false],[true,false]])('applies source-qualified legacy origin only in approximate paint, automatic borders %s, unsplit %s',async (automatic,cantSplit)=>{
    const input=automatic?autoBorderFixture():tableFixture(),document=input.document as NativeDocxDocumentV1,resolved=input.resolved_layout as NativeDocxResolvedLayoutInputV1,settings=input.pagination_settings as NativeDocxPaginationSettingsV1
    settings.profile='unsupported';delete settings.compatibility_mode
    settings.diagnostics=[{code:'COMPATIBILITY_SETTING_UNSUPPORTED',severity:'unsupported',part_name:SETTINGS_PART,path:'/w:settings[1]/w:compat[1]',preservation:'preserve-verbatim',message:'Legacy12'}]
    const table=document.body.blocks[0]!.table!,part=table.anchor.part_name,base=table.anchor.path+'/w:tblPr[1]'
    table.rows[0]!.cant_split=cantSplit!
    const fact={table_id:table.id,package_sha256:HASH,indent_twips:0,left_margin_twips:100,source_indent:{part_name:part,path:base+'/w:tblInd[1]',sha256:HASH},source_margin:{part_name:part,path:base+'/w:tblCellMar[1]/w:left[1]',sha256:HASH}}
    const eligibility={protocol:'injoffice.docx.approximation-eligibility',version:1,document_id:settings.document_id,revision:settings.revision,package_sha256:HASH,settings_sha256:settings.settings_sha256,status:'eligible',legacy_compatibility_mode:12,reasons:['Legacy12'],legacy_table_origins:[fact]}
    const outlines=createHarfBuzzOutlineProviderV1({bytes:FONT_BYTES,contentDigest:FONT_DIGEST})
    const provider={providerId:input.outline_provider.provider_id,providerRevision:input.outline_provider.provider_revision,getGlyphOutline(request:import('./nativePagePaintV1.js').NativeDocxGlyphOutlineRequestV1){const o=outlines.outline(request.glyph_id);return o.path.length?{status:'outlined' as const,...request,...o}:{status:'empty' as const,...request,units_per_em:o.units_per_em}}}
    const before=structuredClone(input)
    const {legacy_table_origins: _origins,...baselineEligibility}=eligibility
    const render=(e:unknown)=>automatic?renderNativeDocxAutomaticBorderPreviewV1(input,provider,undefined,e):renderNativeDocxApproximatePagePreviewV1(input,e,provider)
    const baseline=await render(baselineEligibility),shifted=await render(eligibility)
    expect(shifted.status).toBe('painted');expect(baseline.status).toBe('painted')
    expect(shifted.reasons).toContain(DOCX_LEGACY_TABLE_ORIGIN_WARNING)
    const border=(p:typeof shifted)=>p.pages[0]!.commands.find(c=>c.kind==='stroke_table_border')!
    const a=border(shifted),b=border(baseline)
    expect(a).toMatchObject({...b,x1_millipoints:(b as any).x1_millipoints-5000,x2_millipoints:(b as any).x2_millipoints-5000})
    expect(shifted.pages.map(p=>[p.width_millipoints,p.height_millipoints])).toEqual(baseline.pages.map(p=>[p.width_millipoints,p.height_millipoints]))
    for(const [index,command] of shifted.pages[0]!.commands.entries()){
      const original=baseline.pages[0]!.commands[index]!
      if(command.kind==='fill_glyph_path'&&original.kind==='fill_glyph_path'){
        expect(command.path).toEqual(original.path.map(part=>Object.fromEntries(Object.entries(part).map(([key,value])=>[key,key.endsWith('x_millipoints')?(value as number)-5000:value]))))
      }else if(command.kind==='stroke_table_border'&&original.kind==='stroke_table_border'){
        expect(command).toEqual({...original,x1_millipoints:original.x1_millipoints-5000,x2_millipoints:original.x2_millipoints-5000})
      }else if(command.kind==='fill_table_cell'&&original.kind==='fill_table_cell'){
        expect(command).toEqual({...original,x_millipoints:original.x_millipoints-5000})
      }
    }
    expect(shifted.rendering_provenance.table_projection.sha256).not.toBe(baseline.rendering_provenance.table_projection.sha256)
    expect(input).toEqual(before)
    const decode=automatic?decodeNativeDocxAutomaticBorderPreviewV1:decodeNativeDocxApproximatePagePreviewV1
    expect(decode(shifted).ok).toBe(true)
    expect(decode({...shifted,table_border_layout_policy:'unknown'}).ok).toBe(false)
    expect(shifted.table_width_policy).toBe('approximate-authored-grid-fitted-v1')
    expect(decode({...shifted,table_width_policy:'unknown'}).ok).toBe(false)
    expect(decode({...shifted,reasons:shifted.reasons.filter(r=>r!==DOCX_TABLE_GRID_FIT_WARNING)}).ok).toBe(false)
    const oldEnvelope=structuredClone(shifted);delete oldEnvelope.table_border_layout_policy;delete oldEnvelope.table_width_policy;oldEnvelope.reasons=oldEnvelope.reasons.filter(r=>r!==DOCX_TABLE_BORDER_RESERVATION_WARNING&&r!==DOCX_TABLE_GRID_FIT_WARNING)
    expect(decode(oldEnvelope).ok).toBe(true)
    // 256 bounded source reasons plus all optional host-policy warnings fit.
    const maximumReasons=[...Array.from({length:255},(_,i)=>`Source reason ${i}`),...shifted.reasons,'Optional host font-size policy warning']
    expect(decode({...shifted,reasons:maximumReasons}).ok).toBe(true)
    expect(decode({...shifted,reasons:[...maximumReasons,...Array(301).fill('excess')]}).ok).toBe(false)
    expect(decode({...shifted,reasons:shifted.reasons.filter(r=>r!==DOCX_LEGACY_TABLE_ORIGIN_WARNING)}).ok).toBe(false)
    for(const changed of [{...fact,left_margin_twips:101},{...fact,table_id:'other'},{...fact,source_margin:{...fact.source_margin,path:'/wrong'}}])await expect(render({...eligibility,legacy_table_origins:[changed]})).rejects.toThrow()
    await expect(render({...eligibility,legacy_compatibility_mode:14})).rejects.toThrow()
    if(!automatic){
      const strict=await prepareNativeDocxPagePaintV1(input)
      expect(strict.page_paint_request.paginated_layout.status).toBe('refused')
      const strictTables=qualifyNativeDocxTablesV1(document,resolved,strict.page_paint_request.pagination_request.shaped_lines)
      const approximateTables=qualifyApproximateLegacyTables(document,resolved,strict.page_paint_request.pagination_request.shaped_lines,eligibility)
      expect(strictTables).toMatchObject({status:'qualified',tables:[{x_millipoints:0}]})
      expect(approximateTables).toMatchObject({status:'qualified',tables:[{x_millipoints:-5000}]})
      if(approximateTables.status!=='qualified')throw new Error('Expected approximate table')
      expect((await render({...eligibility,legacy_table_origins:[]})).status).toBe('painted')
      const computed=structuredClone(strict.page_paint_request)
      computed.paginated_layout=paginateNativeDocxApproximateLegacyV1(computed.pagination_request,eligibility).layout
      computed.integrity.paginated_layout_sha256=nativeDocxPagePaintPaginatedLayoutSha256V1(computed.paginated_layout)
      computed.integrity.table_projection_sha256=approximateTables.sha256
      expect(decodeNativeDocxApproximateComputedPagePaintV1(computed,eligibility).fidelity).toBe('approximate')
      expect(decodeNativeDocxPagePaintRequestV1(computed).ok).toBe(false)
      if(!cantSplit){
        expect(computed.paginated_layout.pages[0]!.table_rows?.length).toBeGreaterThan(0)
        for(const [key,delta] of [['x_millipoints',-1],['width_millipoints',1],['y_millipoints',-1000000]] as const){
          const wrongRow=structuredClone(computed)
          wrongRow.paginated_layout.pages[0]!.table_rows![0]![key]+=delta
          wrongRow.integrity.paginated_layout_sha256=nativeDocxPagePaintPaginatedLayoutSha256V1(wrongRow.paginated_layout)
          expect(()=>decodeNativeDocxApproximateComputedPagePaintV1(wrongRow,eligibility)).toThrow()
        }
        expect(()=>decodeNativeDocxApproximateComputedPagePaintV1(computed,baselineEligibility)).toThrow()
      }
      const forged=structuredClone(computed);forged.paginated_layout.pages[0]!.lines[0]!.x_millipoints+=1
      forged.integrity.paginated_layout_sha256=nativeDocxPagePaintPaginatedLayoutSha256V1(forged.paginated_layout)
      expect(()=>decodeNativeDocxApproximateComputedPagePaintV1(forged,eligibility)).toThrow()
      const changedDigest={...eligibility,legacy_table_origins:[{...fact,source_indent:{...fact.source_indent,sha256:`sha256:${'f'.repeat(64)}`}}]}
      const digestProjection=qualifyApproximateLegacyTables(document,resolved,strict.page_paint_request.pagination_request.shaped_lines,changedDigest)
      if(digestProjection.status!=='qualified')throw new Error('Expected declared digest projection')
      expect(digestProjection.sha256).not.toBe(approximateTables.sha256)
      expect(()=>decodeNativeDocxApproximateComputedPagePaintV1(computed,changedDigest)).toThrow()
      const underflow=structuredClone(document);underflow.sections[0]!.page.margins.left_twips=0
      expect(()=>qualifyApproximateLegacyTables(underflow,resolved,strict.page_paint_request.pagination_request.shaped_lines,eligibility)).toThrow('page bounds')
    }
  },20000)
  it.each([undefined, 12] as const)('renders source-qualified automatic borders only as a distinct white-background approximation (legacy %s)', async (mode) => {
    const input = autoBorderFixture()
    const settings = input.pagination_settings as NativeDocxPaginationSettingsV1
    let eligibility: unknown
    if (mode !== undefined) {
      settings.profile = 'unsupported'
      delete settings.compatibility_mode
      settings.diagnostics = [{ code: 'COMPATIBILITY_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: SETTINGS_PART, path: '/w:settings[1]/w:compat[1]', preservation: 'preserve-verbatim', message: 'Legacy Word mode 12 requires different semantics' }]
      eligibility = { protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: settings.document_id, revision: settings.revision, package_sha256: settings.package_sha256, settings_sha256: settings.settings_sha256, status: 'eligible', legacy_compatibility_mode: mode, reasons: ['Legacy Word mode 12 is approximated using current layout'] }
    }
    const original = structuredClone(input)
    const outlines = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
    const result = await renderNativeDocxAutomaticBorderPreviewV1(input, {
      providerId: input.outline_provider.provider_id, providerRevision: input.outline_provider.provider_revision,
      getGlyphOutline(request) {
        const outline = outlines.outline(request.glyph_id)
        return outline.path.length ? { status: 'outlined' as const, ...request, ...outline } : { status: 'empty' as const, ...request, units_per_em: outline.units_per_em }
      },
    }, undefined, eligibility)
    expect(result.status).toBe('painted')
    expect(result).toMatchObject({ fidelity: 'approximate', read_only: true, page_background_rgb: 'FFFFFF' })
    expect(result.reasons).toContain(DOCX_AUTO_BORDER_WARNING)
    expect(result.source_diagnostics.document).toHaveLength(1)
    expect(result.pages[0]!.commands.some(command => command.kind === 'stroke_table_border')).toBe(true)
    expect(input).toEqual(original)
    expect(decodeNativeDocxAutomaticBorderPreviewV1(result).ok).toBe(true)
    if (mode !== undefined) expect(decodeNativeDocxAutomaticBorderPreviewV1({...result,reasons:result.reasons.filter(reason=>!reason.includes('natural ascent at the top'))}).ok).toBe(false)
    const cyclic: any = {}; cyclic.self = cyclic
    const huge: any = { reasons: Array(100_001).fill('x') }
    for (const hostile of [cyclic, huge]) {
      const clone = vi.spyOn(globalThis, 'structuredClone')
      try {
        expect(decodeNativeDocxAutomaticBorderPreviewV1({ ...result, legacy_eligibility: hostile }).ok).toBe(false)
        expect(clone.mock.calls.some(([value]) => value === hostile || (value as any)?.legacy_eligibility === hostile)).toBe(false)
      } finally { clone.mockRestore() }
    }
    if (mode !== undefined) expect(decodeNativeDocxAutomaticBorderPreviewV1({ ...result, legacy_eligibility: undefined }).ok).toBe(false)
    for (const mutation of [
      { reasons: [] }, { page_background_rgb: '000000' }, { approximated_render_properties: [] },
      { source_diagnostics: { document: [], resolved: [] } },
      { source: { ...result.source, package_sha256: `sha256:${'f'.repeat(64)}` } },
      { fidelity: 'exact' }, { unexpected: true },
    ]) expect(decodeNativeDocxAutomaticBorderPreviewV1({ ...result, ...mutation }).ok).toBe(false)
    const strict = await prepareNativeDocxPagePaintV1(input)
    expect(strict.page_paint_request.paginated_layout).toMatchObject({ status: 'refused', pages: [] })
  }, 20000)

  it('refuses forged automatic-border evidence and keeps unrelated source diagnostics', () => {
    const mutations: ((document: NativeDocxDocumentV1, resolved: NativeDocxResolvedLayoutInputV1) => void)[] = [
      (_, r) => { r.tables[0]!.automatic_border_preview!.package_sha256 = `sha256:${'f'.repeat(64)}` },
      (_, r) => { r.tables[0]!.automatic_border_preview!.cell_ids = ['cell:other'] },
      (_, r) => { r.tables[0]!.automatic_border_preview!.source_diagnostics[0]!.code = 'UNKNOWN_SOURCE' },
      (_, r) => { r.tables[0]!.automatic_border_preview!.source_path = '/w:document[1]/w:body[1]/w:tbl[2]/w:tblPr[1]/w:tblBorders[1]' },
      (d) => { d.body.blocks[0]!.table!.rows[0]!.cells[0]!.shading_rgb = '000000' },
      (d) => { d.body.blocks[0]!.table!.rows[0]!.cells[0]!.grid_span = 2 },
      (d) => { d.body.blocks[0]!.table!.borders = { top: { style: 'single', size_eighth_points: 8, color_rgb: 'FF0000' } } },
      (_, r) => { r.tables[0]!.automatic_border_preview!.borders.top!.color_rgb = 'FFFFFF' },
      (_, r) => { r.tables[0]!.automatic_border_preview!.automatic_edges.push('top') },
    ]
    for (const mutate of mutations) {
      const input = autoBorderFixture()
      mutate(input.document as NativeDocxDocumentV1, input.resolved_layout as NativeDocxResolvedLayoutInputV1)
      expect(() => projectNativeDocxAutomaticBordersV1(input.document, input.resolved_layout)).toThrow()
    }
    const input = autoBorderFixture(), doc = input.document as NativeDocxDocumentV1
    doc.unsupported.push({ ...doc.unsupported[0]!, id: 'unsupported:other', code: 'UNKNOWN_SOURCE' })
    expect(projectNativeDocxAutomaticBordersV1(doc, input.resolved_layout).document.unsupported.map(d => d.code)).toEqual(['UNKNOWN_SOURCE'])
  })

  // The policy needs a white page behind the table, and only a drawing can put
  // ink there. A header, footer, note or comment that draws nothing leaves the
  // page white, so its existence alone must not refuse the projection.
  it('asks every story for drawings instead of excluding it for existing', () => {
    const storyPart = 'word/header1.xml'
    const at = (path: string) => ({ part_name: storyPart, path, start_byte: 1, end_byte: 100, xml_sha256: HASH })
    const story = (drawing: boolean) => {
      const run = drawing
        ? {
            kind: 'drawing' as const, id: 'run:story:1', anchor: at('/w:hdr[1]/w:p[1]/w:r[1]'),
            drawing: {
              id: 'drawing:story:1', anchor: at('/w:hdr[1]/w:p[1]/w:r[1]/w:drawing[1]'),
              relationship_id: 'rStoryImage', media_part: 'word/media/image.png', content_type: 'image/png', placement: 'inline' as const,
              width_emu: 127_000, height_emu: 127_000,
              edit_policy: { mode: 'read-only' as const, allowed_operations: [], refusal: { code: 'EXTRACT_ONLY', message: 'Fixture drawing is immutable.', preservation: 'refuse-mutation' as const } },
            },
          }
        : { kind: 'text' as const, id: 'run:story:1', anchor: at('/w:hdr[1]/w:p[1]/w:r[1]'), text: 'Header' }
      const paragraph = {
        id: 'paragraph:story:1', anchor: at('/w:hdr[1]/w:p[1]'),
        edit_policy: { mode: 'read-only' as const, allowed_operations: [], refusal: { code: 'NATIVE_READ_ONLY', message: 'Fixture is immutable.', preservation: 'refuse-mutation' as const } },
        properties: {}, runs: [run],
      }
      return { id: 'story:header:1', kind: 'header' as const, part_name: storyPart, anchor: at('/w:hdr[1]'), blocks: [{ kind: 'paragraph' as const, id: paragraph.id, paragraph }] }
    }
    for (const drawing of [false, true]) {
      const input = autoBorderFixture(), document = input.document as NativeDocxDocumentV1
      document.headers.push(story(drawing) as unknown as NativeDocxDocumentV1['headers'][number])
      document.passthrough_parts.push({ part_name: 'word/media/image.png', content_type: 'image/png', byte_length: PNG_BYTES.byteLength, sha256: PNG_DIGEST, policy: 'preserve-verbatim' })
      const run = () => projectNativeDocxAutomaticBordersV1(document, input.resolved_layout)
      if (drawing) expect(run).toThrow('exclude drawings')
      else expect(run().facts.length).toBe(1)
    }
  })

  // Measured against a genuine Microsoft Word 16.112.4 render of
  // NumberedList.docx at 96 DPI: a fixed (exact / at-least) line box puts the
  // surplus leading above the text and seats the descent on the box bottom,
  // where an automatic (multiple) box keeps the ascent at the top.
  it.each(['exact', 'atLeast'] as const)('seats an expanded %s line box on its bottom while strict still refuses it', async (rule) => {
    const input = fixture(), resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    const outlines = createHarfBuzzOutlineProviderV1({bytes:FONT_BYTES,contentDigest:FONT_DIGEST})
    const provider = {providerId:input.outline_provider.provider_id,providerRevision:input.outline_provider.provider_revision,getGlyphOutline(request:any){const outline=outlines.outline(request.glyph_id);return outline.path.length ? {status:'outlined' as const,...request,...outline} : {status:'empty' as const,...request,units_per_em:outline.units_per_em}}}
    const naturalPrepared = await prepareNativeDocxPagePaintV1(fixture())
    const {compileNativeDocxPagePaintV1} = await import('./nativePagePaintV1.js')
    const natural = await compileNativeDocxPagePaintV1(naturalPrepared.page_paint_request,provider)
    expect(natural.ok && natural.value.status).toBe('painted')
    if (!natural.ok || natural.value.status !== 'painted') return
    const naturalLine = natural.value.pages[0]!.lines[0]!
    const shapedLine = naturalPrepared.page_paint_request.pagination_request.shaped_lines.paragraphs[0]!.lines[0]!
    // Ask for a box a whole natural line taller than the shaped one.
    const twips = Math.ceil(shapedLine.line_height_millipoints * 2 / 50)
    const expanded = twips * 50
    resolved.paragraphs[0]!.properties = {...resolved.paragraphs[0]!.properties, line_rule: rule, line: twips}
    const original = structuredClone(input)
    const prepared = await prepareNativeDocxPagePaintV1(input)
    const strict = await compileNativeDocxPagePaintV1(prepared.page_paint_request,provider)
    expect(strict.ok && strict.value.status).toBe('refused')
    const settings = input.pagination_settings as NativeDocxPaginationSettingsV1
    settings.profile='unsupported';delete settings.compatibility_mode
    settings.diagnostics=[{code:'COMPATIBILITY_SETTING_UNSUPPORTED',severity:'unsupported',part_name:SETTINGS_PART,path:'/w:settings[1]/w:compat[1]',preservation:'preserve-verbatim',message:'Legacy layout'}]
    const eligibility={protocol:'injoffice.docx.approximation-eligibility',version:1,document_id:settings.document_id,revision:settings.revision,package_sha256:settings.package_sha256,settings_sha256:settings.settings_sha256,status:'eligible',legacy_compatibility_mode:12,reasons:['Legacy layout approximation']}
    const approximate = await renderNativeDocxApproximatePagePreviewV1(input,eligibility,provider)
    expect(approximate.status).toBe('painted')
    if (approximate.status !== 'painted') return
    const painted = approximate.pages[0]!.lines[0]!
    expect(painted.height_millipoints).toBe(expanded)
    // Descent seated on the box bottom, not the ascent on the box top.
    expect(painted.baseline_y_millipoints - painted.y_millipoints).toBe(expanded + shapedLine.descent_millipoints)
    expect(painted.baseline_y_millipoints - painted.y_millipoints).toBeGreaterThan(naturalLine.baseline_y_millipoints - naturalLine.y_millipoints)
    expect(approximate.reasons.some(reason=>reason.includes('seats the descent on the box bottom'))).toBe(true)
    expect(decodeNativeDocxApproximatePagePreviewV1(approximate).ok).toBe(true)
    expect(input.resolved_layout).toEqual(original.resolved_layout)
    expect(input.document).toEqual(original.document)
    // The strict lane is byte-identical whether or not the approximate lane ran.
    const strictAfter = await compileNativeDocxPagePaintV1((await prepareNativeDocxPagePaintV1(structuredClone(original))).page_paint_request,provider)
    expect(JSON.stringify(strictAfter)).toBe(JSON.stringify(strict))
  }, 30000)

  it.each([120, 480])('keeps strict line-box guards and discloses approximate expanded baseline placement (%s)', async (line) => {
    const input = fixture(), resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    resolved.paragraphs[0]!.properties = {...resolved.paragraphs[0]!.properties, line_rule:'auto', line}
    const original = structuredClone(input)
    const outlines = createHarfBuzzOutlineProviderV1({bytes:FONT_BYTES,contentDigest:FONT_DIGEST})
    const provider = {providerId:input.outline_provider.provider_id,providerRevision:input.outline_provider.provider_revision,getGlyphOutline(request:any){const outline=outlines.outline(request.glyph_id);return outline.path.length ? {status:'outlined' as const,...request,...outline} : {status:'empty' as const,...request,units_per_em:outline.units_per_em}}}
    const prepared = await prepareNativeDocxPagePaintV1(input)
    const {compileNativeDocxPagePaintV1} = await import('./nativePagePaintV1.js')
    const strict = await compileNativeDocxPagePaintV1(prepared.page_paint_request,provider)
    expect(strict.ok && strict.value.status).toBe('refused')
    const settings = input.pagination_settings as NativeDocxPaginationSettingsV1
    settings.profile='unsupported';delete settings.compatibility_mode
    settings.diagnostics=[{code:'COMPATIBILITY_SETTING_UNSUPPORTED',severity:'unsupported',part_name:SETTINGS_PART,path:'/w:settings[1]/w:compat[1]',preservation:'preserve-verbatim',message:'Legacy layout'}]
    const eligibility={protocol:'injoffice.docx.approximation-eligibility',version:1,document_id:settings.document_id,revision:settings.revision,package_sha256:settings.package_sha256,settings_sha256:settings.settings_sha256,status:'eligible',legacy_compatibility_mode:12,reasons:['Legacy layout approximation']}
    const approximate = await renderNativeDocxApproximatePagePreviewV1(input,eligibility,provider)
    expect(approximate.status).toBe(line>240?'painted':'refused')
    if (line > 240) {
      const naturalPrepared = await prepareNativeDocxPagePaintV1(fixture())
      const natural = await compileNativeDocxPagePaintV1(naturalPrepared.page_paint_request,provider)
      expect(natural.ok && natural.value.status).toBe('painted')
      if (natural.ok && natural.value.status === 'painted') {
        const a = approximate.pages[0]!.lines[0]!, b = natural.value.pages[0]!.lines[0]!
        expect(a.baseline_y_millipoints - a.y_millipoints).toBe(b.baseline_y_millipoints - b.y_millipoints)
        expect(a.height_millipoints).toBeGreaterThan(b.height_millipoints)
      }
    }
    expect(approximate.reasons.some(reason=>reason.includes('natural ascent at the top'))).toBe(true)
    expect(decodeNativeDocxApproximatePagePreviewV1({...approximate,reasons:approximate.reasons.filter(reason=>!reason.includes('natural ascent at the top'))}).ok).toBe(false)
    expect(input.resolved_layout).toEqual(original.resolved_layout)
    expect(input.document).toEqual(original.document)
  })

  it('paints remaining glyphs when approximate omits paragraph-scoped run diagnostics', async () => {
    const input = fixture()
    const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    const settings = input.pagination_settings as NativeDocxPaginationSettingsV1
    resolved.diagnostics.push(
      { code: 'PARTIAL_RUN_PROPERTIES', severity: 'unsupported', scope_id: 'paragraph:1', part_name: 'word/document.xml', path: '/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:rPr[1]', preservation: 'preserve-verbatim', message: 'Only the conservative v1 run-property subset is exposed' },
      { code: 'FONT_MATCHING_METADATA_PRESERVED', severity: 'unsupported', scope_id: resolved.document_id, part_name: 'word/fontTable.xml', path: '/w:fonts[1]/w:font[1]/w:embedRegular[1]', preservation: 'preserve-verbatim', message: 'Validated font matching metadata is preserved' },
    )
    settings.profile = 'unsupported'
    delete settings.compatibility_mode
    settings.diagnostics = [{ code: 'COMPATIBILITY_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: SETTINGS_PART, path: '/w:settings[1]/w:compat[1]', preservation: 'preserve-verbatim', message: 'Legacy layout' }]
    const eligibility = { protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: settings.document_id, revision: settings.revision, package_sha256: settings.package_sha256, settings_sha256: settings.settings_sha256, status: 'eligible' as const, legacy_compatibility_mode: 12 as const, reasons: ['Legacy layout approximation'] }
    const outlines = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
    const provider = { providerId: input.outline_provider.provider_id, providerRevision: input.outline_provider.provider_revision, getGlyphOutline(request: import('./nativePagePaintV1.js').NativeDocxGlyphOutlineRequestV1) { const outline = outlines.outline(request.glyph_id); return outline.path.length ? { status: 'outlined' as const, ...request, ...outline } : { status: 'empty' as const, ...request, units_per_em: outline.units_per_em } } }
    const strict = await prepareNativeDocxPagePaintV1(input)
    expect(strict.page_paint_request.pagination_request.shaped_lines.paragraphs).toEqual([])
    const approximate = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, provider)
    expect(approximate.status).toBe('painted')
    expect(approximate.pages[0]!.commands.some(command => command.kind === 'fill_glyph_path')).toBe(true)
  })

  it('paints an off-lattice picture in the approximate lane, discloses the rounding, and keeps strict refusing', async () => {
    // FigureAsLabelPicture.docx / graphic-object-fliph.docx / lvlPicBulletId.docx:
    // the authored extent is not a whole number of milli-points, so the exact
    // slice declines the picture over at most 0.005 pt and the page paints
    // around it. The approximate lane paints it at the nearest milli-point.
    const input = imageFixture()
    const document = input.document as NativeDocxDocumentV1
    const drawing = document.body.blocks[0]!.paragraph!.runs[0]!.drawing!
    drawing.width_emu = 2_751_151
    drawing.height_emu = 2_063_363
    const settings = input.pagination_settings as NativeDocxPaginationSettingsV1
    settings.profile = 'unsupported'
    delete settings.compatibility_mode
    settings.diagnostics = [{ code: 'COMPATIBILITY_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: SETTINGS_PART, path: '/w:settings[1]/w:compat[1]', preservation: 'preserve-verbatim', message: 'Legacy layout' }]
    const eligibility = { protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: settings.document_id, revision: settings.revision, package_sha256: settings.package_sha256, settings_sha256: settings.settings_sha256, status: 'eligible' as const, legacy_compatibility_mode: 12 as const, reasons: ['Legacy layout approximation'] }
    const outlines = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
    const provider = { providerId: input.outline_provider.provider_id, providerRevision: input.outline_provider.provider_revision, getGlyphOutline(request: import('./nativePagePaintV1.js').NativeDocxGlyphOutlineRequestV1) { const outline = outlines.outline(request.glyph_id); return outline.path.length ? { status: 'outlined' as const, ...request, ...outline } : { status: 'empty' as const, ...request, units_per_em: outline.units_per_em } } }
    const original = structuredClone(input)
    // Strict: the picture is unqualified, so no page exists at all.
    expect(qualifyNativeDocxInlineImageV1(document, 'run:image', drawing).ok).toBe(false)
    const strict = await compileNativeDocxPagePaintV1((await prepareNativeDocxPagePaintV1(structuredClone(input))).page_paint_request, provider)
    expect(strict.ok && strict.value.status).toBe('refused')
    const approximate = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, provider)
    expect(approximate.status).toBe('painted')
    const painted = approximate.pages[0]!.commands.find(command => command.kind === 'paint_inline_image')
    if (!painted || painted.kind !== 'paint_inline_image') throw new Error('Expected the picture to paint')
    expect(painted.width_millipoints).toBe(216_630)
    expect(painted.height_millipoints).toBe(162_470)
    // The asset the command names is the authored one, byte for byte.
    expect(approximate.resources).toMatchObject([{ part_name: 'word/media/image.png', content_digest: PNG_DIGEST, byte_length: PNG_BYTES.byteLength }])
    expect(painted.asset_id).toBe(approximate.resources[0]!.id)
    // Recorded, never silent.
    expect(approximate.approximated_image_extents).toEqual([
      { run_id: 'run:image', drawing_id: 'drawing:1', part_name: 'word/document.xml', path: '/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:drawing[1]', field: 'width_emu', source_emu: 2_751_151, painted_emu: 2_751_201 },
      { run_id: 'run:image', drawing_id: 'drawing:1', part_name: 'word/document.xml', path: '/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:drawing[1]', field: 'height_emu', source_emu: 2_063_363, painted_emu: 2_063_369 },
    ])
    expect(approximate.reasons).toContain(DOCX_APPROXIMATE_IMAGE_EXTENT_WARNING)
    expect(decodeNativeDocxApproximatePagePreviewV1(approximate).ok).toBe(true)
    // Neither half of the disclosure may be dropped.
    expect(decodeNativeDocxApproximatePagePreviewV1({ ...approximate, reasons: approximate.reasons.filter(reason => reason !== DOCX_APPROXIMATE_IMAGE_EXTENT_WARNING) }).ok).toBe(false)
    const { approximated_image_extents: _dropped, ...withoutFacts } = approximate
    expect(decodeNativeDocxApproximatePagePreviewV1(withoutFacts).ok).toBe(false)
    expect(decodeNativeDocxApproximatePagePreviewV1({ ...approximate, approximated_image_extents: [{ ...approximate.approximated_image_extents![0]!, painted_emu: 2_751_151 }] }).ok).toBe(false)
    // The caller's input, and the strict lane reading it, are untouched.
    expect(input).toEqual(original)
    const strictAfter = await compileNativeDocxPagePaintV1((await prepareNativeDocxPagePaintV1(structuredClone(original))).page_paint_request, provider)
    expect(JSON.stringify(strictAfter)).toBe(JSON.stringify(strict))
  }, 20000)

  it('leaves an exact picture extent unrounded and undisclosed', async () => {
    const input = imageFixture()
    const settings = input.pagination_settings as NativeDocxPaginationSettingsV1
    settings.profile = 'unsupported'
    delete settings.compatibility_mode
    settings.diagnostics = [{ code: 'COMPATIBILITY_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: SETTINGS_PART, path: '/w:settings[1]/w:compat[1]', preservation: 'preserve-verbatim', message: 'Legacy layout' }]
    const eligibility = { protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: settings.document_id, revision: settings.revision, package_sha256: settings.package_sha256, settings_sha256: settings.settings_sha256, status: 'eligible' as const, legacy_compatibility_mode: 12 as const, reasons: ['Legacy layout approximation'] }
    const outlines = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
    const provider = { providerId: input.outline_provider.provider_id, providerRevision: input.outline_provider.provider_revision, getGlyphOutline(request: import('./nativePagePaintV1.js').NativeDocxGlyphOutlineRequestV1) { const outline = outlines.outline(request.glyph_id); return outline.path.length ? { status: 'outlined' as const, ...request, ...outline } : { status: 'empty' as const, ...request, units_per_em: outline.units_per_em } } }
    const approximate = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, provider)
    expect(approximate.status).toBe('painted')
    expect(approximate.approximated_image_extents).toBeUndefined()
    expect(approximate.reasons).not.toContain(DOCX_APPROXIMATE_IMAGE_EXTENT_WARNING)
    expect(approximate.pages[0]!.commands.some(command => command.kind === 'paint_inline_image')).toBe(true)
  }, 20000)

  it('joins automatic-border whole-part hashes to exactly one preserved source part when available', () => {
    const input = autoBorderFixture()
    const document = input.document as NativeDocxDocumentV1
    const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    document.passthrough_parts.push({ part_name: 'word/document.xml', content_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml', byte_length: 3000, sha256: HASH, policy: 'preserve-verbatim' })
    // Strict source wire forbids representing a modeled story as passthrough.
    expect(() => projectNativeDocxAutomaticBordersV1(document, resolved)).toThrow('valid original')
    document.passthrough_parts.pop()
    document.unsupported = []
    const evidence = resolved.tables[0]!.automatic_border_preview!
    evidence.source_part = 'word/styles.xml'
    evidence.source_path = '/w:styles[1]/w:style[1]/w:tblPr[1]/w:tblBorders[1]'
    evidence.source_diagnostics = [{ code: 'TABLE_STYLE_EFFECTS_PRESERVED', scope_id: 'table:1', part_name: evidence.source_part, path: '/w:styles[1]/w:style[1]/w:tblPr[1]' }]
    resolved.source_parts.styles_part = evidence.source_part
    resolved.diagnostics = [{ ...evidence.source_diagnostics[0]!, severity: 'unsupported', preservation: 'preserve-verbatim', message: 'Automatic border source remains strict-refused' }]
    expect(() => projectNativeDocxAutomaticBordersV1(document, resolved)).toThrow('digest')
    document.passthrough_parts.push({ part_name: evidence.source_part, content_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml', byte_length: 1000, sha256: HASH, policy: 'preserve-verbatim' })
    expect(() => projectNativeDocxAutomaticBordersV1(document, resolved)).not.toThrow()
    document.passthrough_parts.at(-1)!.sha256 = `sha256:${'f'.repeat(64)}`
    expect(() => projectNativeDocxAutomaticBordersV1(document, resolved)).toThrow('digest')
    document.passthrough_parts.at(-1)!.sha256 = HASH
    document.passthrough_parts.push({ ...document.passthrough_parts.at(-1)! })
    expect(() => projectNativeDocxAutomaticBordersV1(document, resolved)).toThrow()
  })
  // Each case includes a real HarfBuzz cold start. Bound it independently of
  // the default five-second unit-test timeout on shared CI runners.
  it.each([12, 14] as const)('renders approximate mode %s through real HarfBuzz without changing strict preparation', async (mode) => {
      const input = fixture()
      const settings = input.pagination_settings as NativeDocxPaginationSettingsV1
      settings.profile = 'unsupported'
      delete settings.compatibility_mode
      settings.diagnostics = [{ code: 'COMPATIBILITY_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: SETTINGS_PART, path: '/w:settings[1]/w:compat[1]', preservation: 'preserve-verbatim', message: `Legacy Word mode ${mode} requires different semantics` }]
      const original = structuredClone(input)
      const eligibility = { protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: settings.document_id, revision: settings.revision, package_sha256: settings.package_sha256, settings_sha256: settings.settings_sha256, status: 'eligible', legacy_compatibility_mode: mode, reasons: [`Legacy Word mode ${mode} is approximated using current layout`] }
      const outlines = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
      const approximate = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, {
        providerId: input.outline_provider.provider_id, providerRevision: input.outline_provider.provider_revision,
        getGlyphOutline(request) {
          const outline = outlines.outline(request.glyph_id)
          return outline.path.length ? { status: 'outlined' as const, ...request, ...outline } : { status: 'empty' as const, ...request, units_per_em: outline.units_per_em }
        },
      })
      expect(approximate.status).toBe('painted')
      expect(approximate.pages.length).toBeGreaterThan(0)
      expect(approximate.pages[0]!.commands.length).toBeGreaterThan(0)
      expect(decodeNativeDocxApproximatePagePreviewV1(approximate).ok).toBe(true)
      expect(input).toEqual(original)
      const strict = await prepareNativeDocxPagePaintV1(input)
      expect(strict.page_paint_request.paginated_layout).toMatchObject({ status: 'refused', pages: [] })
      expect(strict.outline_requests).toEqual([])
  }, 15_000)
  function hostFixture() {
    const input = fixture()
    const original = JSON.parse(input.font_inventory_json) as NativeDOCXFontInventoryV1
    const manifest = structuredClone(original.native_text_manifest!)
    manifest.faces[0]!.source.kind = 'host'
    const candidate = manifest.faces[0]!
    const face: ResolvedFontFace = { faceId: candidate.faceId, family: candidate.family, weight: 400, style: 'normal', stretch: 100, sourceKind: 'host', resourceId: candidate.source.resourceId, contentDigest: FONT_DIGEST, resolution: 'exact', matchedFamily: candidate.family }
    const metrics = inspectHarfBuzzFontMetricsV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
    const resolver: NativeFontResolver = {
      providerId: 'test.host-fonts', providerRevision: 'v1',
      resolve: () => ({ status: 'resolved', face, attemptedFaceIds: [face.faceId], decisions: [] }),
      load: () => ({ face, bytes: Uint8Array.from(FONT_BYTES), metrics }),
    }
    input.font_assets = []
    const document = input.document as NativeDocxDocumentV1
    document.passthrough_parts = document.passthrough_parts.filter(p => p.part_name !== 'word/_rels/fontTable.xml.rels' && !p.part_name.startsWith('word/fonts/'))
    rewriteInventory(input, inventory => {
      inventory.families.forEach(f => { f.faces = [] })
      delete inventory.font_table!.font_relationships_part
      delete inventory.font_table!.font_relationships_sha256
      delete inventory.native_text_manifest
      delete inventory.native_text_manifest_sha256
    })
    return { input, fonts: { manifest, resolver } }
  }

  // word/settings.xml is optional in the format. Without it the pagination
  // settings carry no relationship evidence, and the font-inventory join used to
  // compare a real part name against undefined and refuse the whole document.
  it('paints a document whose package has no settings.xml, and still joins the inventory to the package', async () => {
    const input = fixture()
    const settings = input.pagination_settings as NativeDocxPaginationSettingsV1
    settings.profile = 'absent-default'
    delete settings.relationships_part
    delete settings.relationships_sha256
    delete settings.relationship_id
    delete settings.settings_part
    delete settings.settings_sha256
    delete settings.compatibility_mode
    // The regression: this threw 'does not exact-join pagination settings' and
    // refused the document outright, because a real part name never equals the
    // absent settings evidence. It must now compile.
    await expect(prepareNativeDocxPagePaintV1(input)).resolves.toBeDefined()

    // The binding is still joined against the document's passthrough inventory:
    // a font table whose hash does not match the package is still refused.
    const forged = fixture()
    const forgedSettings = forged.pagination_settings as NativeDocxPaginationSettingsV1
    forgedSettings.profile = 'absent-default'
    delete forgedSettings.relationships_part
    delete forgedSettings.relationships_sha256
    delete forgedSettings.relationship_id
    delete forgedSettings.settings_part
    delete forgedSettings.settings_sha256
    delete forgedSettings.compatibility_mode
    rewriteInventory(forged, inventory => { inventory.font_table!.sha256 = `sha256:${'d'.repeat(64)}` })
    await expect(prepareNativeDocxPagePaintV1(forged)).rejects.toThrow('does not exact-join the native document passthrough inventory')
  }, 15_000)

  it('renders a non-embedded-font document with explicit content-addressed host fonts', async () => {
    const { input, fonts } = hostFixture()
    const before = JSON.stringify(input)
    const prepared = await prepareNativeDocxPagePaintV1(input, { fonts })
    expect(prepared.outline_requests.length).toBeGreaterThan(0)
    expect(prepared.providers.resolver_id).toBe('test.host-fonts')
    expect(JSON.stringify(input)).toBe(before)
    await expect(prepareNativeDocxPagePaintV1(input)).rejects.toThrow(/configure explicit host fonts/)
  })

  // word/fontTable.xml is optional in the format, exactly as word/settings.xml
  // is. A minimal package such as sdt_after_section_break.docx stores only
  // [Content_Types].xml, _rels/.rels and word/document.xml, and requiring a
  // font-table package binding refused every such document outright.
  it('paints a document whose package has no fontTable.xml', async () => {
    const { input, fonts } = hostFixture()
    const document = input.document as NativeDocxDocumentV1
    const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    delete resolved.source_parts.font_table_part
    document.passthrough_parts = document.passthrough_parts.filter((part) => part.part_name !== 'word/fontTable.xml')
    rewriteInventory(input, (inventory) => { delete inventory.font_table; inventory.families = [] })
    // The regression: this threw 'font inventory has no font-table package
    // binding' before any layout decision was taken. It must now compile.
    const prepared = await prepareNativeDocxPagePaintV1(input, { fonts })
    expect(prepared.outline_requests.length).toBeGreaterThan(0)

    // An inventory that describes a family has read a font table, so it must
    // still carry the package binding that proves where it read it.
    const forged = hostFixture()
    const forgedResolved = forged.input.resolved_layout as NativeDocxResolvedLayoutInputV1
    delete forgedResolved.source_parts.font_table_part
    rewriteInventory(forged.input, (inventory) => { delete inventory.font_table })
    await expect(prepareNativeDocxPagePaintV1(forged.input, { fonts: forged.fonts })).rejects.toThrow('font families require a font-table binding')
  })

  // Word wraps a body paragraph at the column width of the section that owns
  // it. The preview used to shape every paragraph at one width and refuse a
  // document whose sections disagreed; it now maps each section's paragraphs to
  // its own column width, which the line core already takes per paragraph.
  it('shapes each section\'s body paragraphs at that section\'s own column width', async () => {
    const input = twoSectionFixture()
    const prepared = await prepareNativeDocxPagePaintV1(input)
    const shaped = prepared.page_paint_request.pagination_request.shaped_lines
    // 12240 - 1440 - 1440 twips, and 12240 - 2880 - 2880 twips, in milli-points.
    expect(shaped.available_width_millipoints).toBe(468_000)
    expect(Object.fromEntries(shaped.paragraphs.map((paragraph) => [paragraph.paragraph_id, paragraph.lines.map((line) => line.available_width_millipoints)])))
      .toEqual({ 'paragraph:1': [468_000], 'paragraph:2': [324_000] })
    expect(prepared.page_paint_request.paginated_layout.status).toBe('paginated')
  })

  // The shaped record does not restate its shaping width per paragraph, so
  // pagination re-derives it from the offered line intervals. A section whose
  // body lines were offered more than its own column never joins this geometry.
  it('refuses a section whose body lines were shaped wider than its own column', async () => {
    const input = twoSectionFixture()
    const prepared = await prepareNativeDocxPagePaintV1(input)
    const request = structuredClone(prepared.page_paint_request.pagination_request)
    request.document.sections[1]!.page.margins.left_twips = 4_320
    request.document.sections[1]!.page.margins.right_twips = 4_320
    const layout = paginateNativeDocxV1(request)
    expect(layout.ok && layout.value.status).toBe('refused')
    expect(layout.ok && layout.value.diagnostics.map((entry) => entry.code)).toContain('section-width-mismatch')
  })

  it('refuses approximate pagination policies in the font-substitution legacy preview instead of applying them undisclosed', async () => {
    const {input,fonts}=hostFixture()
    const document=input.document as NativeDocxDocumentV1,resolved=input.resolved_layout as NativeDocxResolvedLayoutInputV1,settings=input.pagination_settings as NativeDocxPaginationSettingsV1
    resolved.fonts[0]!.name='Missing Family'
    for(const run of resolved.runs)run.properties.font_family='Missing Family'
    for(const paragraph of resolved.paragraphs)paragraph.paragraph_mark_properties!.font_family='Missing Family'
    rewriteInventory(input,inventory=>{inventory.families[0]!.name='Missing Family';inventory.references.forEach(reference=>{reference.family='Missing Family'})})
    // Word 2010 separator stories with unmodeled markup and no note reference anywhere.
    const part='word/footnotes.xml'
    for(const [index,role] of (['separator','continuation-separator'] as const).entries()){
      const first={id:`paragraph:${role}:1`,anchor:{part_name:part,path:`/w:footnotes[1]/w:footnote[${index+1}]/w:p[1]`,start_byte:10+index*100,end_byte:40+index*100,xml_sha256:HASH},edit_policy:document.body.blocks[0]!.paragraph!.edit_policy,properties:{},runs:[]}
      const story={id:`story:footnote:${role}`,kind:'footnote' as const,native_story_id:role==='separator'?'-1':'0',relationship_id:'rIdFootnotes',note_role:role,part_name:part,anchor:{part_name:part,path:`/w:footnotes[1]/w:footnote[${index+1}]`,start_byte:1+index*100,end_byte:70+index*100,xml_sha256:HASH},blocks:[{kind:'paragraph' as const,id:first.id,paragraph:first}]}
      document.notes.push(story as never)
      document.unsupported.push({id:`unsupported:${role}`,code:'UNMODELED_NOTE_MARKUP',capability:'notes',scope_id:story.id,anchor:story.anchor,preservation:'preserve-verbatim',message:'Reserved note separator stories must contain exactly one matching instruction leaf and no visible text'} as never)
      resolved.paragraphs.push({paragraph_id:first.id,applied_styles:[],properties:{},paragraph_mark_properties:{font_family:'Missing Family',font_size_half_points:20}})
    }
    rewriteInventory(input,inventory=>{for(const reference of inventory.references)if(reference.weight===400&&reference.style==='normal')reference.scope_ids=[...reference.scope_ids,'paragraph:separator:1','paragraph:continuation-separator:1'].sort()})
    settings.profile='unsupported';delete settings.compatibility_mode
    settings.diagnostics=[{code:'COMPATIBILITY_SETTING_UNSUPPORTED',severity:'unsupported',part_name:SETTINGS_PART,path:'/w:settings[1]/w:compat[1]',preservation:'preserve-verbatim',message:'Legacy12'}]
    const eligibility={protocol:'injoffice.docx.approximation-eligibility',version:1,document_id:settings.document_id,revision:settings.revision,package_sha256:HASH,settings_sha256:settings.settings_sha256,status:'eligible',legacy_compatibility_mode:12,reasons:['Legacy12']}
    const policy={version:1 as const,mappings:[{sourceFamily:'Missing Family',targetFamily:'DejaVu Sans',weight:400 as const,style:'normal' as const}]}
    const metrics=inspectHarfBuzzFontMetricsV1({bytes:FONT_BYTES,contentDigest:FONT_DIGEST})
    fonts.resolver.resolve=({run})=>{const selection=selectExplicitFontV1(fonts.manifest,run,policy);if(!selection)throw new Error('No configured face');return {status:'resolved',face:selection.face,attemptedFaceIds:[selection.face.faceId],decisions:[]}}
    fonts.resolver.load=face=>({face,bytes:Uint8Array.from(FONT_BYTES),metrics})
    const configured={...fonts,substitutionPolicy:policy}
    const provider=createHarfBuzzOutlineProviderV1({bytes:FONT_BYTES,contentDigest:FONT_DIGEST})
    const outlines={providerId:input.outline_provider.provider_id,providerRevision:input.outline_provider.provider_revision,getGlyphOutline(request:any){const outline=provider.outline(request.glyph_id);return outline.path.length?{status:'outlined' as const,...request,...outline}:{status:'empty' as const,...request,units_per_em:outline.units_per_em}}}
    const inv=JSON.parse(input.font_inventory_json)
    const composition={source_document:input.document,source_resolved_layout:input.resolved_layout,source_pagination_settings:input.pagination_settings,source_font_inventory_json:input.font_inventory_json,font_descriptor_eligibility:{protocol:'injoffice.docx.font-substitution-eligibility',version:1,document_id:inv.document_id,revision:inv.revision,package_sha256:inv.package_sha256,font_table:inv.font_table??null,facts:[]},legacy_eligibility:eligibility}
    const fontPreview=await renderNativeDocxFontSubstitutionPreviewV1(input,outlines,{fonts:configured,composition})
    expect(fontPreview.status).toBe('refused')
    expect(fontPreview.diagnostics.map(d=>d.message).join(' ')).toContain('Font preview does not apply approximate pagination policies')
    expect(fontPreview.diagnostics.map(d=>d.message).join(' ')).toContain(DOCX_APPROXIMATE_INERT_NOTE_SEPARATOR_WARNING)
    expect(fontPreview.reasons).not.toContain(DOCX_APPROXIMATE_INERT_NOTE_SEPARATOR_WARNING)
  },20000)

  it('renders missing source fonts only through the distinct read-only substitution envelope', async () => {
    const {input,fonts}=hostFixture()
    const resolved=input.resolved_layout as NativeDocxResolvedLayoutInputV1
    resolved.fonts[0]!.name='Missing Family'
    resolved.runs[0]!.properties.font_family='Missing Family'
    resolved.paragraphs[0]!.paragraph_mark_properties!.font_family='Missing Family'
    ;(input.document as NativeDocxDocumentV1).body.blocks[0]!.paragraph!.runs[0]!.text='Actual source 123.'
    rewriteInventory(input,inventory=>{
      inventory.families[0]!.name='Missing Family'
      inventory.references.forEach(reference=>{reference.family='Missing Family'})
    })
    const policy={version:1 as const,mappings:[{sourceFamily:'Missing Family',targetFamily:'DejaVu Sans',weight:400 as const,style:'normal' as const}]}
    const metrics=inspectHarfBuzzFontMetricsV1({bytes:FONT_BYTES,contentDigest:FONT_DIGEST})
    fonts.resolver.resolve=({run})=>{
      const selection=selectExplicitFontV1(fonts.manifest,run,policy)
      if(!selection)throw new Error('No configured face')
      const face=selection.face
      return {status:'resolved',face,attemptedFaceIds:[face.faceId],decisions:[]}
    }
    fonts.resolver.load=face=>({face,bytes:Uint8Array.from(FONT_BYTES),metrics})
    const configured={...fonts,substitutionPolicy:policy}
    const provider=createHarfBuzzOutlineProviderV1({bytes:FONT_BYTES,contentDigest:FONT_DIGEST})
    const outlines={providerId:input.outline_provider.provider_id,providerRevision:input.outline_provider.provider_revision,getGlyphOutline(request:any){
      const outline=provider.outline(request.glyph_id)
      return outline.path.length?{status:'outlined' as const,...request,...outline}:{status:'empty' as const,...request,units_per_em:outline.units_per_em}
    }}
    const original=structuredClone(input)
    const result=await renderNativeDocxFontSubstitutionPreviewV1(input,outlines,{fonts:configured})
    expect(result.status).toBe('painted')
    expect(result.pages[0]!.commands.length).toBeGreaterThan(0)
    expect(result).toMatchObject({protocol:'injoffice.docx.font-substitution-preview',fidelity:'approximate',read_only:true})
    expect(result.substitutions).toEqual([expect.objectContaining({source_id:'run:1',source_role:'run',source_family:'Missing Family',selected_family:'DejaVu Sans',font_digest:FONT_DIGEST})])
    expect(decodeNativeDocxFontSubstitutionPreviewV1(result)).toEqual(result)
    const inv=JSON.parse(input.font_inventory_json),composition={source_document:input.document,source_resolved_layout:input.resolved_layout,source_pagination_settings:input.pagination_settings,source_font_inventory_json:input.font_inventory_json,font_descriptor_eligibility:{protocol:'injoffice.docx.font-substitution-eligibility',version:1,document_id:inv.document_id,revision:inv.revision,package_sha256:inv.package_sha256,font_table:inv.font_table??null,facts:[]}}
    expect(qualifyNativeDocxFontCompositionV1(composition).document).toEqual(input.document)
    const composed=await renderNativeDocxFontSubstitutionPreviewV1(input,outlines,{fonts:configured,composition})
    expect(composed.status).toBe('painted');expect(composed.composition).toEqual(composition)
    for(const mutate of [
      (v:any)=>{v.source_document.revision='stale'},
      (v:any)=>{v.source_font_inventory_json+=' '.repeat(3*1024*1024)},
      (v:any)=>{v.source_resolved_layout.fonts[0].name='forged'},
      (v:any)=>{v.font_descriptor_eligibility.package_sha256=`sha256:${'f'.repeat(64)}`},
      (v:any)=>{v.automatic_borders=true},
      (v:any)=>{v.font_size_policy={kind:'host-default-size-v1',half_points:20}},
      (v:any)=>{v.unknown_policy=true},
    ]){const forged=structuredClone(composition);mutate(forged);await expect(renderNativeDocxFontSubstitutionPreviewV1(input,outlines,{fonts:configured,composition:forged})).rejects.toThrow()}
    for(const mutate of [(v:any)=>{v.composition_sha256=HASH},(v:any)=>{delete v.composition},(v:any)=>{v.composition.source_document.revision='stale'}]){const forged=structuredClone(composed);mutate(forged);expect(()=>decodeNativeDocxFontSubstitutionPreviewV1(forged)).toThrow()}
    expect(input).toEqual(original)
    // Blank PAGE uses paragraph-mark metrics in the base shape; its numeric
    // variant uses a run face. Both role records must survive the envelope.
    const withHeader=structuredClone(input),headerDocument=withHeader.document as NativeDocxDocumentV1,headerResolved=withHeader.resolved_layout as NativeDocxResolvedLayoutInputV1
    const header=structuredClone(headerDocument.body.blocks[0]!.paragraph!)
    header.id='paragraph:header';header.anchor={...anchor('/w:hdr[1]/w:p[1]',10,900),part_name:'word/header1.xml'}
    header.runs=[{kind:'text',id:'run:header',anchor:{...header.anchor,path:'/w:hdr[1]/w:p[1]/w:fldSimple[1]/w:r[1]/w:t[1]',start_byte:20,end_byte:80},text:'',page_field:'PAGE'}]
    headerDocument.headers.push({id:'story:header',kind:'header',part_name:'word/header1.xml',anchor:{...header.anchor,path:'/w:hdr[1]',start_byte:1,end_byte:1000},blocks:[{kind:'paragraph',id:header.id,paragraph:header}]})
    headerDocument.sections[0]!.header_refs.push({kind:'default',story_id:'story:header',relationship_id:'rIdHeader'})
    headerResolved.paragraphs.push({...structuredClone(headerResolved.paragraphs[0]!),paragraph_id:header.id,properties:{alignment:'right'}})
    headerResolved.runs.push({...structuredClone(headerResolved.runs[0]!),run_id:'run:header',paragraph_id:header.id})
    rewriteInventory(withHeader,inventory=>{inventory.references[0]!.scope_ids.push(header.id,'run:header');inventory.references[0]!.scope_ids.sort()})
    const headerBefore=structuredClone(withHeader),headerResult=await renderNativeDocxFontSubstitutionPreviewV1(withHeader,outlines,{fonts:configured})
    expect(headerResult.status).toBe('painted')
    expect(headerResult.substitutions).toEqual(expect.arrayContaining([
      expect.objectContaining({source_id:'paragraph:header',source_role:'paragraph-mark'}),
      expect.objectContaining({source_id:'run:header',source_role:'run'}),
    ]))
    expect(headerResult.pages[0]!.lines.some(l=>l.region==='header')).toBe(true)
    expect(withHeader).toEqual(headerBefore)
    await expect(prepareNativeDocxPagePaintV1(withHeader,{fonts:configured})).rejects.toThrow(/strict|substitution/i)
    const exactHeader=structuredClone(withHeader),exactResolved=exactHeader.resolved_layout as NativeDocxResolvedLayoutInputV1
    exactResolved.fonts[0]!.name='DejaVu Sans'
    exactResolved.runs.forEach(r=>{r.properties.font_family='DejaVu Sans'})
    exactResolved.paragraphs.forEach(p=>{p.paragraph_mark_properties!.font_family='DejaVu Sans'})
    rewriteInventory(exactHeader,inventory=>{inventory.families[0]!.name='DejaVu Sans';inventory.references.forEach(r=>{r.family='DejaVu Sans'})})
    const exactPrepared=await prepareNativeDocxPagePaintV1(exactHeader,{fonts}),request=structuredClone(exactPrepared.page_paint_request.pagination_request),variants=structuredClone(exactPrepared.page_paint_request.page_field_variants!)
    request.resolved_layout.runs.forEach(r=>{r.properties.font_family='Missing Family'})
    request.resolved_layout.paragraphs.forEach(p=>{p.paragraph_mark_properties!.font_family='Missing Family'})
    const attach=(shaped:typeof request.shaped_lines)=>{
      const records:NativeDocxFontSubstitutionV1[]=[]
      for(const p of shaped.paragraphs)for(const l of p.lines){
        const sources=l.fragments.length?l.fragments.map(f=>({id:f.source_id,role:f.source_kind as 'run'})):[{id:p.paragraph_id,role:'paragraph-mark' as const}]
        for(const source of sources)if(!records.some(r=>r.source_id===source.id&&r.source_role===source.role))records.push({source_id:source.id,source_role:source.role,source_family:'Missing Family',selected_family:'DejaVu Sans',face_id:fonts.manifest.faces[0]!.faceId,font_digest:FONT_DIGEST,weight:400,style:'normal'})
      }
      shaped.font_substitutions=records;shaped.diagnostics.push(...records.map(nativeDocxFontSubstitutionDiagnosticV1))
    }
    attach(request.shaped_lines);variants.forEach(v=>attach(v.shaped_lines))
    const variantPolicy={manifest:fonts.manifest,policy},layout=exactPrepared.page_paint_request.paginated_layout
    expect(validateNativeDocxFontPageFieldVariantsV1(request,layout,variants,variantPolicy)).toHaveLength(1)
    expect(()=>validateNativeDocxPageFieldVariantsV1(request,layout,variants)).toThrow()
    for(const mutate of [
      (v:typeof variants)=>{v[0]!.shaped_lines.font_substitutions=[]},
      (v:typeof variants)=>{v[0]!.shaped_lines.font_substitutions![0]!.font_digest=HASH},
      (v:typeof variants)=>{v[0]!.shaped_lines.diagnostics[0]!.message+=' forged'},
      (v:typeof variants)=>{v[0]!.shaped_lines.paragraphs[0]!.lines[0]!.line_height_millipoints+=1},
      (v:typeof variants)=>{v[0]!.shaped_lines.providers.shaper_revision='changed'},
    ]){const forged=structuredClone(variants);mutate(forged);expect(()=>validateNativeDocxFontPageFieldVariantsV1(request,layout,forged,variantPolicy)).toThrow()}
    header.runs[0]!.text='999'
    await expect(renderNativeDocxFontSubstitutionPreviewV1(withHeader,outlines,{fonts:configured})).rejects.toThrow()
    await expect(prepareNativeDocxPagePaintV1(input,{fonts:configured})).rejects.toThrow(/strict|substitution/i)
    for(const mutate of [
      (value:any)=>{value.source.revision='rev:other'},
      (value:any)=>{value.policy_sha256=HASH},
      (value:any)=>{value.substitutions.push(structuredClone(value.substitutions[0]))},
      (value:any)=>{value.substitutions[0].selected_family='Another family'},
      (value:any)=>{value.reasons=[]},
      (value:any)=>{value.read_only=false},
      (value:any)=>{value.substitutions[0].font_digest=HASH},
      (value:any)=>{value.substitutions[0].face_id='forged-face'},
      (value:any)=>{value.selected_font_manifest.faces[0].source.contentDigest=HASH},
    ]){
      const corrupt=structuredClone(result);mutate(corrupt)
      expect(()=>decodeNativeDocxFontSubstitutionPreviewV1(corrupt)).toThrow()
    }
    const paragraph=(input.document as NativeDocxDocumentV1).body.blocks[0]!.paragraph!
    paragraph.runs[0]!.text='א'
    expect(await renderNativeDocxFontSubstitutionPreviewV1(input,outlines,{fonts:configured})).toMatchObject({status:'refused',pages:[]})
    paragraph.runs[0]!.text='A'
    resolved.runs[0]!.properties.rtl=true
    await expect(renderNativeDocxFontSubstitutionPreviewV1(input,outlines,{fonts:configured})).rejects.toThrow(/left-to-right/)
    paragraph.runs=[];resolved.runs=[]
    rewriteInventory(input,inventory=>{inventory.references.forEach(reference=>{reference.scope_ids=['paragraph:1']})})
    const empty=await renderNativeDocxFontSubstitutionPreviewV1(input,outlines,{fonts:configured})
    expect(empty.status).toBe('painted')
    expect(empty.substitutions).toEqual([expect.objectContaining({source_id:'paragraph:1',source_role:'paragraph-mark',source_family:'Missing Family',font_digest:FONT_DIGEST})])
  })

  it('refuses stale, missing, corrupt, and overriding host font evidence', async () => {
    const stale = hostFixture(); stale.fonts.manifest.revision = 'rev:other'
    await expect(prepareNativeDocxPagePaintV1(stale.input, { fonts: stale.fonts })).rejects.toThrow(/revision/)
    const missing = hostFixture(); missing.fonts.manifest.faces = []
    await expect(prepareNativeDocxPagePaintV1(missing.input, { fonts: missing.fonts })).rejects.toThrow(/manifest|face/)
    const corrupt = hostFixture(), load = corrupt.fonts.resolver.load
    corrupt.fonts.resolver.load = async face => {
      const resource = await load(face)
      if ('bytes' in resource) resource.bytes[100] ^= 1
      return resource
    }
    await expect(prepareNativeDocxPagePaintV1(corrupt.input, { fonts: corrupt.fonts })).rejects.toThrow(/exact-join/)
    const override = hostFixture()
    await expect(prepareNativeDocxPagePaintV1(fixture(), { fonts: override.fonts })).rejects.toThrow(/override an embedded/)
  })
  it.each(['single', 'double', 'words'] as const)('paints font-bound %s underlines and rejects decoration tampering', async (style) => {
    const input = fixture()
    ;(input.document as NativeDocxDocumentV1).body.blocks[0]!.paragraph!.runs[0]!.properties = { underline: style }
    ;(input.document as NativeDocxDocumentV1).body.blocks[0]!.paragraph!.runs[0]!.text = 'A A'
    ;(input.resolved_layout as NativeDocxResolvedLayoutInputV1).runs[0]!.properties.underline = style
    const prepared = await prepareNativeDocxPagePaintV1(input)
    const provider = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: prepared.outline_requests.map(request => {
      const outline = provider.outline(request.glyph_id)
      return outline.path.length ? { status: 'outlined' as const, ...request, ...outline } : { status: 'empty' as const, ...request, units_per_em: outline.units_per_em }
    }) })
    expect(completed.page_paint_output.status, JSON.stringify(completed.page_paint_output)).toBe('painted')
    if (completed.page_paint_output.status !== 'painted') throw new Error('underline refused')
    const page = completed.page_paint_output.pages[0]!
    const strokes = page.commands.filter(command => command.kind === 'stroke_text_underline')
    expect(strokes).toHaveLength(style === 'double' ? 6 : style === 'words' ? 2 : 3)
    expect(page.commands.at(-1)?.kind).toBe('stroke_text_underline')
    const fragment = prepared.page_paint_request.pagination_request.shaped_lines.paragraphs[0]!.lines[0]!.fragments[0]!
    expect(strokes[0]).toMatchObject({ y1_millipoints: page.lines[0]!.baseline_y_millipoints - fragment.underline_position_millipoints!, width_millipoints: fragment.underline_thickness_millipoints })
    for (const patch of [{ stroke_rgb: 'FF0000' }, { width_millipoints: 1 }, { y1_millipoints: 1, y2_millipoints: 1 }, { source_id: 'run:other' }]) {
      const tampered = structuredClone(completed.page_paint_output)
      Object.assign(tampered.pages[0]!.commands.find(command => command.kind === 'stroke_text_underline')!, patch)
      expect(decodeNativeDocxPagePaintForRequestV1(tampered, completed.page_paint_request, completed.page_paint_request.outline_provider).ok).toBe(false)
    }
    const missing = structuredClone(completed.page_paint_output)
    missing.pages[0]!.commands.pop(); missing.pages[0]!.lines[0]!.command_ids.pop()
    expect(decodeNativeDocxPagePaintForRequestV1(missing, completed.page_paint_request, completed.page_paint_request.outline_provider).ok).toBe(false)
  })
  it('scales real-font subscript/superscript advances and outlines with source-bound OS/2 metrics', async () => {
    const plain = fixture(); (plain.document as NativeDocxDocumentV1).body.blocks[0]!.paragraph!.runs[0]!.text = '2'
    const base = await prepareNativeDocxPagePaintV1(plain)
    const original = base.page_paint_request.pagination_request.shaped_lines.paragraphs[0]!.lines[0]!.fragments[0]!
    const provider = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
    for (const alignment of ['subscript','superscript'] as const) {
      const input = structuredClone(plain)
      ;(input.document as NativeDocxDocumentV1).body.blocks[0]!.paragraph!.runs[0]!.properties = { vertical_alignment: alignment }
      ;(input.resolved_layout as NativeDocxResolvedLayoutInputV1).runs[0]!.properties.vertical_alignment = alignment
      const prepared = await prepareNativeDocxPagePaintV1(input)
      const fragment = prepared.page_paint_request.pagination_request.shaped_lines.paragraphs[0]!.lines[0]!.fragments[0]!
      expect(fragment.script_transform).toMatchObject({ kind: alignment, font_sha256: FONT_DIGEST, units_per_em: 2048 })
      expect(fragment.script_transform!.x_size).not.toBe(fragment.script_transform!.y_size)
      expect(fragment.advance_inline_millipoints).toBeLessThan(original.advance_inline_millipoints)
      expect(Math.sign(fragment.glyphs[0]!.offset_y_millipoints)).toBe(alignment === 'superscript' ? 1 : -1)
      const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: prepared.outline_requests.map((request) => ({ status: 'outlined' as const, ...request, ...provider.outline(request.glyph_id) })) })
      expect(completed.page_paint_output.status).toBe('painted')
      expect(decodeNativeDocxPagePaintForRequestV1(completed.page_paint_output, completed.page_paint_request, completed.page_paint_request.outline_provider).ok).toBe(true)
      for (const patch of [undefined, { ...fragment.script_transform!, kind: alignment === 'subscript' ? 'superscript' : 'subscript' }, { ...fragment.script_transform!, font_sha256: `sha256:${'f'.repeat(64)}` }]) {
        const invalid = structuredClone(prepared.page_paint_request)
        const changed = invalid.pagination_request.shaped_lines.paragraphs[0]!.lines[0]!.fragments[0]!
        if (patch === undefined) delete changed.script_transform; else changed.script_transform = patch as typeof changed.script_transform
        invalid.integrity.shaped_lines_sha256 = nativeDocxPagePaintShapedLinesSha256V1(invalid.pagination_request.shaped_lines)
        expect(decodeNativeDocxPagePaintRequestV1(invalid).ok).toBe(false)
      }
      for (const decoration of [{ highlight:'yellow' },{ underline:'single' as const }]) {
        const decorated = structuredClone(input)
        Object.assign((decorated.resolved_layout as NativeDocxResolvedLayoutInputV1).runs[0]!.properties, decoration)
        const refused = await prepareNativeDocxPagePaintV1(decorated)
        expect(refused.page_paint_request.paginated_layout.status).toBe('refused')
      }
    }
  })
  it.each([undefined,0,7])('converges body PAGE/NUMPAGES with decimal restart %s while preserving the original source and read-only policy', async (start) => {
    const input=fixture(), document=input.document as NativeDocxDocumentV1, resolved=input.resolved_layout as NativeDocxResolvedLayoutInputV1
    if(start!==undefined) document.sections[0]!.page_number_start=start
    const first=document.body.blocks[0]!.paragraph!
    first.runs[0]!.page_field='NUMPAGES';first.runs[0]!.text=''
    const second=structuredClone(first)
    second.id='paragraph:second';second.anchor=anchor('/w:document[1]/w:body[1]/w:p[2]',200,290);second.properties.page_break_before=true
    second.runs[0]!.id='run:second';second.runs[0]!.anchor=anchor('/w:document[1]/w:body[1]/w:p[2]/w:r[1]/w:t[1]',210,280);second.runs[0]!.page_field='PAGE'
    document.body.blocks.push({kind:'paragraph',id:second.id,paragraph:second})
    resolved.paragraphs.push({...structuredClone(resolved.paragraphs[0]!),paragraph_id:second.id,properties:{page_break_before:true}})
    resolved.runs.push({...structuredClone(resolved.runs[0]!),run_id:'run:second',paragraph_id:second.id})
    rewriteInventory(input,inventory=>{inventory.references[0]!.scope_ids.push(second.id,'run:second');inventory.references[0]!.scope_ids.sort()})
    const before=JSON.stringify(document),prepared=await prepareNativeDocxPagePaintV1(input),request=prepared.page_paint_request
    expect(JSON.stringify(document)).toBe(before)
    expect(request.body_field_source).toEqual(document)
    expect(request.pagination_request.document.body.blocks.map(block=>block.paragraph!.runs[0]!.text)).toEqual(['2',String((start??1)+1)])
    expect(request.pagination_request.document.body.blocks.every(block=>block.paragraph!.edit_policy.mode==='read-only')).toBe(true)
    const provider=createHarfBuzzOutlineProviderV1({bytes:FONT_BYTES,contentDigest:FONT_DIGEST})
    const completed=await completeNativeDocxPagePaintV1({prepared,outline_results:prepared.outline_requests.map(outline=>({status:'outlined' as const,...outline,...provider.outline(outline.glyph_id)}))})
    expect(completed.page_paint_output.status).toBe('painted')
    expect(decodeNativeDocxPagePaintForRequestV1(completed.page_paint_output,request,request.outline_provider).ok).toBe(true)
    for(const mutate of [
      (r:typeof request)=>{delete r.body_field_source},
      (r:typeof request)=>{delete r.body_field_source; delete r.integrity.body_field_source_sha256},
      (r:typeof request)=>{r.body_field_source!.body.blocks[0]!.paragraph!.runs[0]!.text='999'},
      (r:typeof request)=>{r.pagination_request.document.body.blocks[0]!.paragraph!.runs[0]!.text='1'},
      (r:typeof request)=>{r.body_field_source!.body.anchor.xml_sha256=`sha256:${'0'.repeat(64)}`},
    ]){const invalid=structuredClone(request);mutate(invalid);expect(decodeNativeDocxPagePaintRequestV1(invalid).ok).toBe(false)}
  })
  it.each([12,14] as const)('converges approximate mode %s body PAGE/NUMPAGES without altering strict sources', async mode => {
    const input=fixture(), document=input.document as NativeDocxDocumentV1, resolved=input.resolved_layout as NativeDocxResolvedLayoutInputV1
    document.sections[0]!.page_number_start=7
    const first=document.body.blocks[0]!.paragraph!
    first.runs[0]!.page_field='NUMPAGES';first.runs[0]!.text=''
    const second=structuredClone(first)
    second.id='paragraph:second';second.anchor=anchor('/w:document[1]/w:body[1]/w:p[2]',200,290);second.properties.page_break_before=true
    second.runs[0]!.id='run:second';second.runs[0]!.anchor=anchor('/w:document[1]/w:body[1]/w:p[2]/w:r[1]/w:t[1]',210,280);second.runs[0]!.page_field='PAGE'
    document.body.blocks.push({kind:'paragraph',id:second.id,paragraph:second})
    resolved.paragraphs.push({...structuredClone(resolved.paragraphs[0]!),paragraph_id:second.id,properties:{page_break_before:true}})
    resolved.runs.push({...structuredClone(resolved.runs[0]!),run_id:'run:second',paragraph_id:second.id})
    rewriteInventory(input,inventory=>{inventory.references[0]!.scope_ids.push(second.id,'run:second');inventory.references[0]!.scope_ids.sort()})
    const reference=await prepareNativeDocxPagePaintV1(structuredClone(input))
    expect(reference.page_paint_request.pagination_request.document.body.blocks.map(block=>block.paragraph!.runs[0]!.text)).toEqual(['2','8'])
    const settings=input.pagination_settings as NativeDocxPaginationSettingsV1
    settings.profile='unsupported';delete settings.compatibility_mode
    settings.diagnostics=[{code:'COMPATIBILITY_SETTING_UNSUPPORTED',severity:'unsupported',part_name:SETTINGS_PART,path:'/w:settings[1]/w:compat[1]/w:compatSetting[1]',preservation:'preserve-verbatim',message:`Legacy mode ${mode}`}]
    const eligibility={protocol:'injoffice.docx.approximation-eligibility',version:1,document_id:settings.document_id,revision:settings.revision,package_sha256:settings.package_sha256,settings_sha256:settings.settings_sha256,status:'eligible',legacy_compatibility_mode:mode,reasons:[`Legacy mode ${mode}`]}
    const original=structuredClone(input)
    const outlines=createHarfBuzzOutlineProviderV1({bytes:FONT_BYTES,contentDigest:FONT_DIGEST})
    const approximate=await renderNativeDocxApproximatePagePreviewV1(input,eligibility,{providerId:input.outline_provider.provider_id,providerRevision:input.outline_provider.provider_revision,getGlyphOutline(request){const outline=outlines.outline(request.glyph_id);return outline.path.length?{status:'outlined' as const,...request,...outline}:{status:'empty' as const,...request,units_per_em:outline.units_per_em}}})
    expect(approximate).toMatchObject({status:'painted',fidelity:'approximate',read_only:true})
    expect(approximate.pages).toHaveLength(2)
    const expectedGlyphs=reference.page_paint_request.pagination_request.shaped_lines.paragraphs.flatMap(paragraph=>paragraph.lines.flatMap(line=>line.fragments.flatMap(fragment=>fragment.glyphs.map(glyph=>glyph.glyph_id))))
    expect(approximate.pages.flatMap(page=>page.commands.filter(command=>command.kind==='fill_glyph_path').map(command=>command.glyph_id))).toEqual(expectedGlyphs)
    expect(approximate.rendering_provenance.pagination_settings).toEqual(settings)
    expect(approximate.rendering_provenance.body_field_source_sha256).toBeDefined()
    expect(decodeNativeDocxApproximatePagePreviewV1(approximate).ok).toBe(true)
    expect(input).toEqual(original)
    const computed=structuredClone(reference.page_paint_request)
    computed.pagination_request.pagination_settings=structuredClone(settings)
    computed.paginated_layout=paginateNativeDocxApproximateLegacyV1(computed.pagination_request,eligibility).layout
    computed.integrity.paginated_layout_sha256=nativeDocxPagePaintPaginatedLayoutSha256V1(computed.paginated_layout)
    expect(decodeNativeDocxPagePaintRequestV1(computed).ok).toBe(false)
    expect(decodeNativeDocxApproximateComputedPagePaintV1(computed,eligibility).fidelity).toBe('approximate')
    for(const mutate of [
      (request:typeof computed)=>{delete request.body_field_source},
      (request:typeof computed)=>{request.pagination_request.document.body.blocks[0]!.paragraph!.runs[0]!.text='9'},
      (request:typeof computed)=>{request.paginated_layout.pages[0]!.lines[0]!.x_millipoints+=1;request.integrity.paginated_layout_sha256=nativeDocxPagePaintPaginatedLayoutSha256V1(request.paginated_layout)},
    ]){const invalid=structuredClone(computed);mutate(invalid);expect(()=>decodeNativeDocxApproximateComputedPagePaintV1(invalid,eligibility)).toThrow()}
    expect(()=>decodeNativeDocxApproximateComputedPagePaintV1(computed,undefined)).toThrow('explicit eligibility')
    await expect(prepareNativeDocxPagePaintV1(input)).rejects.toThrow('bounded successful pagination')
  },15_000)
  it('keeps interior square-image islands unsupported in approximate field layout', async () => {
    const input=imageFixture(), document=input.document as NativeDocxDocumentV1
    const paragraph=document.body.blocks[0]!.paragraph!
    Object.assign(paragraph.runs[0]!.drawing!,{placement:'floating',x_emu:2540000,y_emu:914400,width_emu:1270000,height_emu:635000,horizontal_relative_from:'page',vertical_relative_from:'page',wrap:'square',floating_layer:'front',stacking_order:7})
    const field=paragraph.runs.find(run=>run.kind==='text')!;field.page_field='NUMPAGES';field.text=''
    const settings=input.pagination_settings as NativeDocxPaginationSettingsV1
    settings.profile='unsupported';delete settings.compatibility_mode
    settings.diagnostics=[{code:'COMPATIBILITY_SETTING_UNSUPPORTED',severity:'unsupported',part_name:SETTINGS_PART,path:'/w:settings[1]',preservation:'preserve-verbatim',message:'Legacy mode12'}]
    const eligibility={protocol:'injoffice.docx.approximation-eligibility',version:1,document_id:settings.document_id,revision:settings.revision,package_sha256:settings.package_sha256,settings_sha256:settings.settings_sha256,status:'eligible',legacy_compatibility_mode:12,reasons:['Legacy mode12']}
    await expect(renderNativeDocxApproximatePagePreviewV1(input,eligibility, {providerId:input.outline_provider.provider_id,providerRevision:input.outline_provider.provider_revision,getGlyphOutline(){throw new Error('must not request glyphs')}})).rejects.toThrow('two text intervals')
  })
  it.each([undefined,12,14] as const)('derives repeated header/footer PAGE/NUMPAGES with approximate mode %s', async mode => {
    const input = fixture()
    const document = input.document as NativeDocxDocumentV1
    const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    if(mode!==undefined)document.sections[0]!.page_number_start=7
    const second = structuredClone(document.body.blocks[0]!.paragraph!)
    second.id = 'paragraph:second'; second.anchor = anchor('/w:document[1]/w:body[1]/w:p[2]', 200, 290)
    second.properties.page_break_before = true
    second.runs[0]!.id = 'run:second'; second.runs[0]!.anchor = anchor('/w:document[1]/w:body[1]/w:p[2]/w:r[1]', 210, 280)
    document.body.blocks.push({ kind: 'paragraph', id: second.id, paragraph: second })
    resolved.paragraphs.push({ ...structuredClone(resolved.paragraphs[0]!), paragraph_id: second.id, properties: { page_break_before: true } })
    resolved.runs.push({ ...structuredClone(resolved.runs[0]!), run_id: 'run:second', paragraph_id: second.id })
    const scopes = [second.id, 'run:second']
    for (const region of ['header', 'footer'] as const) {
      const paragraph = structuredClone(document.body.blocks[0]!.paragraph!)
      const root = region === 'header' ? 'hdr' : 'ftr'
      paragraph.id = `paragraph:${region}`
      paragraph.anchor = { ...anchor(`/w:${root}[1]/w:p[1]`, 10, 900), part_name: `word/${region}1.xml` }
      paragraph.runs = ['PAGE', 'NUMPAGES'].map((instruction, index) => ({ kind: 'text', id: `run:${region}:${index}`, anchor: { ...paragraph.anchor, path: `/w:${root}[1]/w:p[1]/w:fldSimple[${index + 1}]/w:r[1]/w:t[1]`, start_byte: 20 + index * 100, end_byte: 80 + index * 100 }, text: '', page_field: instruction as 'PAGE' | 'NUMPAGES' }))
      document[region === 'header' ? 'headers' : 'footers'].push({ id: `story:${region}`, kind: region, part_name: `word/${region}1.xml`, anchor: { ...paragraph.anchor, path: `/w:${root}[1]`, start_byte: 1, end_byte: 1000 }, blocks: [{ kind: 'paragraph', id: paragraph.id, paragraph }] })
      document.sections[0]![region === 'header' ? 'header_refs' : 'footer_refs'].push({ kind: 'default', story_id: `story:${region}`, relationship_id: `rId${region}` })
      resolved.paragraphs.push({ ...structuredClone(resolved.paragraphs[0]!), paragraph_id: paragraph.id, properties: { alignment: 'right' } })
      for (const run of paragraph.runs) resolved.runs.push({ ...structuredClone(resolved.runs[0]!), run_id: run.id, paragraph_id: paragraph.id })
      scopes.push(paragraph.id, ...paragraph.runs.map((run) => run.id))
    }
    rewriteInventory(input, (inventory) => { inventory.references[0]!.scope_ids.push(...scopes); inventory.references[0]!.scope_ids.sort() })
    if(mode===14){const field=document.body.blocks[0]!.paragraph!.runs[0]!;field.page_field='NUMPAGES';field.text=''}
    const prepared = await prepareNativeDocxPagePaintV1(input)
    expect(prepared.page_paint_request.page_field_variants).toHaveLength(2)
    for (const [index, variant] of prepared.page_paint_request.page_field_variants!.entries()) for (const region of ['header', 'footer']) {
      const text = variant.shaped_lines.paragraphs.find((p) => p.story_kind === region)!.lines.flatMap((line) => line.fragments).map((fragment) => fragment.text).join('')
      expect(text).toBe(`${index + (mode===undefined?1:7)}2`)
    }
    const provider = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: prepared.outline_requests.map((request) => ({ status: 'outlined' as const, ...request, ...provider.outline(request.glyph_id) })) })
    expect(completed.page_paint_output.status).toBe('painted')
    expect(decodeNativeDocxPagePaintForRequestV1(completed.page_paint_output, completed.page_paint_request, completed.page_paint_request.outline_provider).ok).toBe(true)
    if(mode!==undefined){
      const settings=input.pagination_settings as NativeDocxPaginationSettingsV1
      settings.profile='unsupported';delete settings.compatibility_mode
      settings.diagnostics=[{code:'COMPATIBILITY_SETTING_UNSUPPORTED',severity:'unsupported',part_name:SETTINGS_PART,path:'/w:settings[1]',preservation:'preserve-verbatim',message:`Legacy mode${mode}`}]
      const eligibility={protocol:'injoffice.docx.approximation-eligibility',version:1,document_id:settings.document_id,revision:settings.revision,package_sha256:settings.package_sha256,settings_sha256:settings.settings_sha256,status:'eligible',legacy_compatibility_mode:mode,reasons:[`Legacy mode${mode}`]}
      const before=structuredClone(input)
      const approximate=await renderNativeDocxApproximatePagePreviewV1(input,eligibility,{providerId:input.outline_provider.provider_id,providerRevision:input.outline_provider.provider_revision,getGlyphOutline(request){const outline=provider.outline(request.glyph_id);return outline.path.length?{status:'outlined' as const,...request,...outline}:{status:'empty' as const,...request,units_per_em:outline.units_per_em}}})
      expect(approximate).toMatchObject({status:'painted',fidelity:'approximate',read_only:true})
      expect(approximate.pages).toEqual(completed.page_paint_output.pages)
      expect(approximate.rendering_provenance.pagination_settings).toEqual(settings)
      expect(decodeNativeDocxApproximatePagePreviewV1(approximate).ok).toBe(true)
      expect(input).toEqual(before)
      const cached=structuredClone(input)
      ;(cached.document as NativeDocxDocumentV1).headers[0]!.blocks[0]!.paragraph!.runs[0]!.text='999'
      await expect(renderNativeDocxApproximatePagePreviewV1(cached,eligibility,{providerId:input.outline_provider.provider_id,providerRevision:input.outline_provider.provider_revision,getGlyphOutline(){throw new Error('must not request glyphs')}})).rejects.toThrow('cached text is not authoritative')
      const computed=structuredClone(prepared.page_paint_request)
      computed.pagination_request.pagination_settings=structuredClone(settings)
      computed.paginated_layout=paginateNativeDocxApproximateLegacyV1(computed.pagination_request,eligibility).layout
      computed.integrity.paginated_layout_sha256=nativeDocxPagePaintPaginatedLayoutSha256V1(computed.paginated_layout)
      expect(decodeNativeDocxPagePaintRequestV1(computed).ok).toBe(false)
      expect(decodeNativeDocxApproximateComputedPagePaintV1(computed,eligibility).fidelity).toBe('approximate')
      const missing=structuredClone(computed);delete missing.page_field_variants
      missing.integrity.shaped_lines_sha256=nativeDocxPagePaintShapedLinesSha256V1(missing.pagination_request.shaped_lines)
      expect(()=>decodeNativeDocxApproximateComputedPagePaintV1(missing,eligibility)).toThrow()
      const reversed=structuredClone(computed);reversed.page_field_variants!.reverse()
      reversed.integrity.shaped_lines_sha256=nativeDocxPagePaintShapedLinesSha256V1(reversed.pagination_request.shaped_lines,reversed.page_field_variants)
      expect(()=>decodeNativeDocxApproximateComputedPagePaintV1(reversed,eligibility)).toThrow()
    }
    for (const mutate of [
      (request: typeof prepared.page_paint_request) => { delete request.page_field_variants },
      (request: typeof prepared.page_paint_request) => { request.page_field_variants!.reverse() },
      (request: typeof prepared.page_paint_request) => { request.page_field_variants![0]!.shaped_lines = structuredClone(request.page_field_variants![1]!.shaped_lines) },
    ]) {
      const invalid = structuredClone(prepared.page_paint_request); mutate(invalid)
      invalid.integrity.shaped_lines_sha256 = nativeDocxPagePaintShapedLinesSha256V1(invalid.pagination_request.shaped_lines, invalid.page_field_variants)
      expect(decodeNativeDocxPagePaintRequestV1(invalid).ok).toBe(false)
    }
    if(mode===undefined){
      const stale = structuredClone(document); stale.headers[0]!.blocks[0]!.paragraph!.runs[0]!.text = '999'
      expect(() => nativeDocxPageFieldDocumentV1(stale, 0, 2)).toThrow(/cached results/)
      expect(() => nativeDocxPageFieldDocumentV1(document, 0, 65)).toThrow(/bounded/)
      const bodyField = structuredClone(document); bodyField.body.blocks[0]!.paragraph!.runs[0]!.page_field = 'PAGE'
      expect(() => nativeDocxPageFieldDocumentV1(bodyField, 0, 2)).toThrow(/Body fields/)
    }
  },15_000)
  it('exact-joins source-attested inline image flips and half-turns without changing layout extents', async () => {
    for (const rotation_degrees of [0, 90, 180, 270] as const) for (const flip_horizontal of [false, true]) for (const flip_vertical of [false, true]) {
      const input = imageFixture()
      const drawing = (input.document as NativeDocxDocumentV1).body.blocks[0]!.paragraph!.runs.find((run) => run.drawing)!.drawing!
      Object.assign(drawing, { rotation_degrees, flip_horizontal, flip_vertical })
      drawing.source_crop = { left: 12345, top: 2500, right: 5000, bottom: 100 }
      const prepared = await prepareNativeDocxPagePaintV1(input)
      const provider = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
      const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: prepared.outline_requests.map((request) => ({ status: 'outlined' as const, ...request, ...provider.outline(request.glyph_id) })) })
      expect(completed.page_paint_output.status).toBe('painted')
      if (completed.page_paint_output.status !== 'painted') throw new Error('image refused')
      const command = completed.page_paint_output.pages[0]!.commands.find((command) => command.kind === 'paint_inline_image')!
      expect(command).toMatchObject({ width_millipoints: 10_000, height_millipoints: 10_000, transform: { rotation_degrees, flip_horizontal, flip_vertical } })
      expect(decodeNativeDocxPagePaintForRequestV1(completed.page_paint_output, completed.page_paint_request, completed.page_paint_request.outline_provider).ok).toBe(true)
      if (command.kind !== 'paint_inline_image') throw new Error('missing image')
      command.source_crop.left += 1
      expect(decodeNativeDocxPagePaintForRequestV1(completed.page_paint_output, completed.page_paint_request, completed.page_paint_request.outline_provider).ok).toBe(false)
      command.source_crop.left -= 1
      command.transform.flip_horizontal = !flip_horizontal
      expect(decodeNativeDocxPagePaintForRequestV1(completed.page_paint_output, completed.page_paint_request, completed.page_paint_request.outline_provider).ok).toBe(false)
    }
  })
  it('paints page-relative floating images outside text flow with source-bound front/behind layers', async () => {
    for (const layer of ['front', 'behind'] as const) {
      const input = imageFixture()
      const drawing = (input.document as NativeDocxDocumentV1).body.blocks[0]!.paragraph!.runs[0]!.drawing!
      Object.assign(drawing, { placement: 'floating', x_emu: 914400, y_emu: 1270000, horizontal_relative_from: 'page', vertical_relative_from: 'page', wrap: 'none', floating_layer: layer, stacking_order: 7 })
      const prepared = await prepareNativeDocxPagePaintV1(input)
      const fragment = prepared.page_paint_request.pagination_request.shaped_lines.paragraphs[0]!.lines[0]!.fragments.find(fragment => fragment.source_kind === 'image')!
      expect(fragment).toMatchObject({ advance_inline_millipoints: 0, ascent_millipoints: 0, descent_millipoints: 0, glyphs: [] })
      const provider = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
      const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: prepared.outline_requests.map(request => ({ status: 'outlined' as const, ...request, ...provider.outline(request.glyph_id) })) })
      expect(completed.page_paint_output.status).toBe('painted')
      if (completed.page_paint_output.status !== 'painted') throw new Error('floating refused')
      const page = completed.page_paint_output.pages[0]!
      const command = layer === 'behind' ? page.commands[0]! : page.commands.at(-1)!
      expect(command).toMatchObject({ kind: 'paint_floating_image', x_millipoints: 72000, y_millipoints: 100000, width_millipoints: 10000, height_millipoints: 10000, layer, stacking_order: 7 })
      expect(decodeNativeDocxPagePaintForRequestV1(completed.page_paint_output,completed.page_paint_request,completed.page_paint_request.outline_provider).ok).toBe(true)
      if (command.kind !== 'paint_floating_image') throw new Error('missing floating image')
      command.x_millipoints += 10
      expect(decodeNativeDocxPagePaintForRequestV1(completed.page_paint_output,completed.page_paint_request,completed.page_paint_request.outline_provider).ok).toBe(false)
      command.x_millipoints -= 10
      command.stacking_order += 1
      expect(decodeNativeDocxPagePaintForRequestV1(completed.page_paint_output,completed.page_paint_request,completed.page_paint_request.outline_provider).ok).toBe(false)
      command.stacking_order -= 1
      page.commands.reverse()
      expect(decodeNativeDocxPagePaintForRequestV1(completed.page_paint_output,completed.page_paint_request,completed.page_paint_request.outline_provider).ok).toBe(false)
    }
  })
  it('wraps complete source lines around a page-edge square image and restores full width below it', async () => {
    for (const edge of ['left','right'] as const) {
      const input=imageFixture(),doc=input.document as NativeDocxDocumentV1,paragraph=doc.body.blocks[0]!.paragraph!,drawing=paragraph.runs[0]!.drawing!
      paragraph.runs[1]!.text='Square wrapped source text continues after the picture. '.repeat(8)
      if(edge==='left'){
        const field=structuredClone(paragraph.runs[1]!)
        field.id='run:square-field';field.text='';field.page_field='NUMPAGES';field.anchor=anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[3]/w:fldSimple[1]',150,170)
        paragraph.runs.push(field)
        const resolved=input.resolved_layout as NativeDocxResolvedLayoutInputV1
        resolved.runs.push({...structuredClone(resolved.runs.find(run=>run.run_id===paragraph.runs[1]!.id)!),run_id:field.id})
        rewriteInventory(input,inventory=>{inventory.references[0]!.scope_ids.push(field.id);inventory.references[0]!.scope_ids.sort()})
      }
      Object.assign(drawing,{placement:'floating',x_emu:edge==='left'?914400:5588000,y_emu:914400,width_emu:1270000,height_emu:635000,horizontal_relative_from:'page',vertical_relative_from:'page',wrap:'square',floating_layer:'front',stacking_order:7})
      const before=JSON.stringify(input),prepared=await prepareNativeDocxPagePaintV1(input),request=prepared.page_paint_request,layout=request.paginated_layout
      expect(JSON.stringify(input)).toBe(before)
      if(layout.status!=='paginated')throw new Error('square source refused')
      const lines=layout.pages[0]!.lines,overlap=lines.filter(line=>line.y_millipoints<122000),below=lines.filter(line=>line.y_millipoints>=122000)
      expect(overlap.length).toBeGreaterThan(1);expect(below.length).toBeGreaterThan(0)
      for(const line of overlap){expect(line.x_millipoints).toBe(edge==='left'?172000:72000);expect(line.width_millipoints).toBeLessThanOrEqual(368000)}
      for(const line of below)expect(line.x_millipoints).toBe(72000)
      expect(decodeNativeDocxPagePaintRequestV1(request).ok).toBe(true)
      if(edge==='left'){
        expect(request.body_field_source).toEqual(doc)
        expect(request.pagination_request.document.body.blocks[0]!.paragraph!.runs[2]!.text).toBe(String(layout.pages.length))
      }
      const provider=createHarfBuzzOutlineProviderV1({bytes:FONT_BYTES,contentDigest:FONT_DIGEST}),completed=await completeNativeDocxPagePaintV1({prepared,outline_results:prepared.outline_requests.map(request=>{const outline=provider.outline(request.glyph_id);return outline.path.length?{status:'outlined' as const,...request,...outline}:{status:'empty' as const,...request,units_per_em:outline.units_per_em}})})
      expect(completed.page_paint_output.status).toBe('painted')
      const tampered=structuredClone(request),changedDrawing=tampered.pagination_request.document.body.blocks[0]!.paragraph!.runs[0]!.drawing!
      if(edge==='left')changedDrawing.width_emu=1397000
      else changedDrawing.x_emu=5461000
      expect(decodeNativeDocxPagePaintRequestV1(tampered).ok).toBe(false)
      for(const story of ['header','footer','footnote','endnote','comment'] as const){
        const forged=structuredClone(request.pagination_request.shaped_lines),paragraph=forged.paragraphs[0]!
        paragraph.story_kind=story
        paragraph.lines[0]!.exclusion_start_millipoints=paragraph.lines[0]!.inline_offset_millipoints
        expect(decodeNativeDocxShapedLines(forged).ok).toBe(false)
      }
    }
  }, 30_000)
  it.each(['left','right'] as const)('uses source-verified %s square exclusions in approximate current-policy layout', async edge => {
    const input=imageFixture(),doc=input.document as NativeDocxDocumentV1,paragraph=doc.body.blocks[0]!.paragraph!,drawing=paragraph.runs[0]!.drawing!
    paragraph.runs[1]!.text='Square wrapped source text continues after the picture. '.repeat(8)
    if(edge==='left'){
      const field=structuredClone(paragraph.runs[1]!)
      field.id='run:square-field';field.text='';field.page_field='NUMPAGES';field.anchor=anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[3]/w:fldSimple[1]',150,170)
      paragraph.runs.push(field)
      const resolved=input.resolved_layout as NativeDocxResolvedLayoutInputV1
      resolved.runs.push({...structuredClone(resolved.runs.find(run=>run.run_id===paragraph.runs[1]!.id)!),run_id:field.id})
      rewriteInventory(input,inventory=>{inventory.references[0]!.scope_ids.push(field.id);inventory.references[0]!.scope_ids.sort()})
    }
    Object.assign(drawing,{placement:'floating',x_emu:edge==='left'?914400:5588000,y_emu:914400,width_emu:1270000,height_emu:635000,horizontal_relative_from:'page',vertical_relative_from:'page',wrap:'square',floating_layer:'front',stacking_order:7})
    const reference=await prepareNativeDocxPagePaintV1(structuredClone(input))
    const settings=input.pagination_settings as NativeDocxPaginationSettingsV1
    settings.profile='unsupported';delete settings.compatibility_mode
    settings.diagnostics=[{code:'COMPATIBILITY_SETTING_UNSUPPORTED',severity:'unsupported',part_name:SETTINGS_PART,path:'/w:settings[1]',preservation:'preserve-verbatim',message:'Legacy mode14'}]
    const eligibility={protocol:'injoffice.docx.approximation-eligibility',version:1,document_id:settings.document_id,revision:settings.revision,package_sha256:settings.package_sha256,settings_sha256:settings.settings_sha256,status:'eligible',legacy_compatibility_mode:14,reasons:['Legacy mode14']}
    const before=structuredClone(input),outlines=createHarfBuzzOutlineProviderV1({bytes:FONT_BYTES,contentDigest:FONT_DIGEST})
    const outline=(request:Parameters<import('./nativePagePaintV1.js').NativeDocxGlyphOutlineProviderV1['getGlyphOutline']>[0])=>{const path=outlines.outline(request.glyph_id);return path.path.length?{status:'outlined' as const,...request,...path}:{status:'empty' as const,...request,units_per_em:path.units_per_em}}
    const approximate=await renderNativeDocxApproximatePagePreviewV1(input,eligibility,{providerId:input.outline_provider.provider_id,providerRevision:input.outline_provider.provider_revision,getGlyphOutline:outline})
    const strict=await completeNativeDocxPagePaintV1({prepared:reference,outline_results:reference.outline_requests.map(outline)})
    expect(approximate.status).toBe('painted')
    expect(approximate.pages).toEqual(strict.page_paint_output.pages)
    expect(decodeNativeDocxApproximatePagePreviewV1(approximate).ok).toBe(true)
    expect(input).toEqual(before)
    const computed=structuredClone(reference.page_paint_request)
    computed.pagination_request.pagination_settings=structuredClone(settings)
    computed.paginated_layout=paginateNativeDocxApproximateLegacyV1(computed.pagination_request,eligibility).layout
    computed.integrity.paginated_layout_sha256=nativeDocxPagePaintPaginatedLayoutSha256V1(computed.paginated_layout)
    expect(decodeNativeDocxApproximateComputedPagePaintV1(computed,eligibility).fidelity).toBe('approximate')
    expect(decodeNativeDocxPagePaintRequestV1(computed).ok).toBe(false)
    const stale=structuredClone(computed)
    const staleDrawing=stale.pagination_request.document.body.blocks[0]!.paragraph!.runs[0]!.drawing!
    if(edge==='left')staleDrawing.width_emu+=127000
    else staleDrawing.x_emu!-=127000
    expect(()=>decodeNativeDocxApproximateComputedPagePaintV1(stale,eligibility)).toThrow()
  },15_000)
  it('independently rejects stale square geometry, interior islands, and blocked lines', async () => {
    const input=imageFixture(),source=input.document as NativeDocxDocumentV1,drawing=source.body.blocks[0]!.paragraph!.runs[0]!.drawing!
    Object.assign(drawing,{placement:'floating',x_emu:914400,y_emu:914400,width_emu:1270000,height_emu:635000,horizontal_relative_from:'page',vertical_relative_from:'page',wrap:'none',floating_layer:'front',stacking_order:7})
    const prepared=await prepareNativeDocxPagePaintV1(input),request=prepared.page_paint_request,p=request.pagination_request,d=p.document.body.blocks[0]!.paragraph!.runs[0]!.drawing!
    d.wrap='square'
    const plan=deriveNativeSquareWrapPlanV1(p.document,p.resolved_layout,p.shaped_lines,request.paginated_layout)
    expect(plan['paragraph:1']![0]).toEqual({start_millipoints:100000,width_millipoints:368000})
    expect(()=>deriveNativeSquareWrapPlanV1(p.document,p.resolved_layout,p.shaped_lines,request.paginated_layout,true)).toThrow('source-derived exclusion')
    expect(decodeNativeDocxPagePaintRequestV1(request).ok).toBe(false)
    d.x_emu=2540000;expect(()=>deriveNativeSquareWrapPlanV1(p.document,p.resolved_layout,p.shaped_lines,request.paginated_layout)).toThrow('two text intervals')
    d.x_emu=914400;d.width_emu=5943600
    expect(deriveNativeSquareWrapPlanV1(p.document,p.resolved_layout,p.shaped_lines,request.paginated_layout)['paragraph:1']![0]!.end_millipoints).toBeGreaterThan(0)
  })
  it('resolves a column/paragraph-relative square anchor and needs its wrap distances to do it', async () => {
    // The Word fixture this mirrors (anchor-position.docx) puts a picture 0.089"
    // inside the column with 0.09" wrap distance on each side. The bare picture
    // box is an interior island; the box widened by distL/distR touches the body
    // edge, which is the single interval Word actually lays out.
    const input=imageFixture(),source=input.document as NativeDocxDocumentV1,drawing=source.body.blocks[0]!.paragraph!.runs[0]!.drawing!
    Object.assign(drawing,{placement:'floating',x_emu:914400,y_emu:914400,width_emu:1270000,height_emu:635000,horizontal_relative_from:'page',vertical_relative_from:'page',wrap:'none',floating_layer:'front',stacking_order:7})
    const prepared=await prepareNativeDocxPagePaintV1(input),request=prepared.page_paint_request,p=request.pagination_request
    const d=p.document.body.blocks[0]!.paragraph!.runs[0]!.drawing!
    const body=request.paginated_layout.pages[0]!.body_box,paragraphTop=request.paginated_layout.pages[0]!.lines[0]!.y_millipoints
    Object.assign(d,{wrap:'square',horizontal_relative_from:'column',vertical_relative_from:'paragraph',x_emu:81280,y_emu:-4445,wrap_distance_left_emu:114300,wrap_distance_right_emu:114300})
    // x resolves against the body box, y against the anchoring paragraph, and the
    // interval starts at the picture's right edge plus distR.
    const resolvedX=body.x_millipoints+6400,resolvedRight=resolvedX+100000+9000
    expect(deriveNativeSquareWrapPlanV1(p.document,p.resolved_layout,p.shaped_lines,request.paginated_layout)['paragraph:1']![0])
      .toEqual({start_millipoints:resolvedRight-body.x_millipoints,width_millipoints:body.x_millipoints+body.width_millipoints-resolvedRight})
    expect(resolveNativeDocxFloatingAnchorV1(qualifiedFloating(p.document,d),100000,request.paginated_layout.pages[0]!,nativeDocxFloatingAnchorOriginsV1(request.paginated_layout).get('paragraph:1')!))
      .toEqual({x_millipoints:resolvedX,y_millipoints:paragraphTop-350,exclusion_left_millipoints:resolvedX-9000,exclusion_right_millipoints:resolvedRight})
    // Drop only the wrap distances and the very same anchor becomes an island.
    delete d.wrap_distance_left_emu;delete d.wrap_distance_right_emu
    expect(()=>deriveNativeSquareWrapPlanV1(p.document,p.resolved_layout,p.shaped_lines,request.paginated_layout)).toThrow('two text intervals')
  })
  it('refuses a resolved anchor whose paragraph is not placed once on a single-column page', async () => {
    const input=imageFixture(),source=input.document as NativeDocxDocumentV1,drawing=source.body.blocks[0]!.paragraph!.runs[0]!.drawing!
    Object.assign(drawing,{placement:'floating',x_emu:914400,y_emu:914400,width_emu:1270000,height_emu:635000,horizontal_relative_from:'page',vertical_relative_from:'page',wrap:'none',floating_layer:'front',stacking_order:7})
    const prepared=await prepareNativeDocxPagePaintV1(input),request=prepared.page_paint_request,p=request.pagination_request
    const d=p.document.body.blocks[0]!.paragraph!.runs[0]!.drawing!
    Object.assign(d,{wrap:'square',horizontal_relative_from:'column',vertical_relative_from:'paragraph',x_emu:0,y_emu:0,wrap_distance_left_emu:114300,wrap_distance_right_emu:114300})
    const page=request.paginated_layout.pages[0]!,origins=nativeDocxFloatingAnchorOriginsV1(request.paginated_layout)
    const floating=qualifiedFloating(p.document,d)
    // A paragraph split across pages has no single origin and is not indexed.
    if(request.paginated_layout.status!=='paginated')throw new Error('fixture must paginate')
    const straddled={...request.paginated_layout,pages:[page,{...page,id:'page:2',lines:page.lines.map(line=>({...line,id:`${line.id}:2`}))}]}
    expect(nativeDocxFloatingAnchorOriginsV1(straddled).has('paragraph:1')).toBe(false)
    // Multi-column bodies have no single column origin to resolve against.
    expect(()=>resolveNativeDocxFloatingAnchorV1(floating,100000,{...page,columns:[page.columns[0]!,page.columns[0]!]},origins.get('paragraph:1')!)).toThrow('single-column')
    expect(()=>resolveNativeDocxFloatingAnchorV1(floating,100000,{...page,id:'page:other'},origins.get('paragraph:1')!)).toThrow('does not own its paragraph')
  })
  it('explicitly refuses square wrapping combined with fixed or autofit table flow', async () => {
    for(const layout of ['fixed','autofit'] as const){
      const input=combinedImageTableFixture(),document=input.document as NativeDocxDocumentV1
      document.body.blocks[1]!.table!.layout=layout
      Object.assign(document.body.blocks[0]!.paragraph!.runs[0]!.drawing!,{placement:'floating',x_emu:914400,y_emu:914400,width_emu:1270000,height_emu:635000,horizontal_relative_from:'page',vertical_relative_from:'page',wrap:'square',floating_layer:'front',stacking_order:7})
      await expect(prepareNativeDocxPagePaintV1(input)).rejects.toThrow('without tables or notes')
    }
  })
  it('reports a refused layout instead of deriving a square-wrap plan from it', async () => {
    const input = imageFixture(), document = input.document as NativeDocxDocumentV1
    const settings = input.pagination_settings as NativeDocxPaginationSettingsV1
    Object.assign(document.body.blocks[0]!.paragraph!.runs[0]!.drawing!, { placement: 'floating', x_emu: 914_400, y_emu: 914_400, width_emu: 1_270_000, height_emu: 635_000, horizontal_relative_from: 'page', vertical_relative_from: 'page', wrap: 'square', floating_layer: 'front', stacking_order: 7 })
    // Anything pagination refuses leaves no placed line to wrap around. The
    // caller is owed that refusal, not a throw from the wrap solver.
    settings.profile = 'unsupported'
    delete settings.compatibility_mode
    settings.diagnostics = [{ code: 'COMPATIBILITY_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: SETTINGS_PART, path: '/w:settings[1]/w:compat[1]', preservation: 'preserve-verbatim', message: 'Legacy mode 12' }]
    const prepared = await prepareNativeDocxPagePaintV1(input)
    expect(prepared.page_paint_request.paginated_layout.status).toBe('refused')
  })
  it('refuses ambiguous stacking, non-body anchors and excessive floating counts', () => {
    const input = imageFixture(), document = input.document as NativeDocxDocumentV1
    const paragraph = document.body.blocks[0]!.paragraph!, run = paragraph.runs[0]!, drawing = run.drawing!
    Object.assign(drawing, { placement: 'floating', x_emu: 0, y_emu: 0, horizontal_relative_from: 'page', vertical_relative_from: 'page', wrap: 'none', floating_layer: 'front', stacking_order: 7 })
    expect(qualifyNativeDocxInlineImageV1(document,run.id,drawing).ok).toBe(true)
    paragraph.runs.push({ ...run, id: 'run:duplicate', drawing: { ...drawing, id: 'drawing:duplicate' } })
    expect(qualifyNativeDocxInlineImageV1(document,run.id,drawing)).toMatchObject({ ok: false, code: 'unsupported-image' })
    paragraph.runs.pop()
    paragraph.runs.shift()
    expect(qualifyNativeDocxInlineImageV1(document,run.id,drawing)).toMatchObject({ ok: false, code: 'unsupported-image' })
    paragraph.runs = Array.from({ length: 129 }, (_,index) => ({ ...run, id: index ? `run:float:${index}` : run.id, drawing: { ...drawing, id: index ? `drawing:float:${index}` : drawing.id, stacking_order: index } }))
    expect(qualifyNativeDocxInlineImageV1(document,run.id,paragraph.runs[0]!.drawing!)).toMatchObject({ ok: false, code: 'resource-limit' })
  })
  it('refuses floating wrapping, unqualified positioning, missing layering and off-page extents', async () => {
    for (const mutation of [ { wrap: 'tight' }, { horizontal_relative_from: 'leftMargin' }, { vertical_relative_from: 'line' }, { horizontal_relative_from: 'paragraph' }, { vertical_relative_from: 'column' }, { floating_layer: undefined }, { x_emu: -127 }, { x_emu: 1 }, { y_emu: 127000000 } ]) {
      const input = imageFixture()
      Object.assign((input.document as NativeDocxDocumentV1).body.blocks[0]!.paragraph!.runs[0]!.drawing!, { placement: 'floating', x_emu: 914400, y_emu: 1270000, horizontal_relative_from: 'page', vertical_relative_from: 'page', wrap: 'none', floating_layer: 'front', stacking_order: 7 },mutation)
      let outcome = 'unknown'
      try {
        const prepared = await prepareNativeDocxPagePaintV1(input)
        const provider = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
        const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: prepared.outline_requests.map(request => ({ status: 'outlined' as const, ...request, ...provider.outline(request.glyph_id) })) })
        outcome = completed.page_paint_output.status
      } catch (error) { expect(error).toBeInstanceOf(Error); outcome = 'refused' }
      expect(outcome).toBe('refused')
    }
  })
  it('paints text highlight behind real glyphs and rejects color/geometry/coverage tampering', async () => {
    const input = fixture()
    ;(input.document as NativeDocxDocumentV1).body.blocks[0]!.paragraph!.runs[0]!.properties = { highlight: 'yellow' }
    ;(input.document as NativeDocxDocumentV1).body.blocks[0]!.paragraph!.runs[0]!.text = 'AA'
    ;(input.resolved_layout as NativeDocxResolvedLayoutInputV1).runs[0]!.properties.highlight = 'yellow'
    const paragraph = (input.document as NativeDocxDocumentV1).body.blocks[0]!.paragraph!
    paragraph.anchor.end_byte = 240
    paragraph.runs.push({ ...structuredClone(paragraph.runs[0]!), id: 'run:plain', properties: {}, text: 'A', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[2]/w:t[1]', 200, 230) })
    const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    resolved.runs.push({ ...structuredClone(resolved.runs[0]!), run_id: 'run:plain', properties: { ...resolved.runs[0]!.properties, highlight: 'none' } })
    rewriteInventory(input, (inventory) => { inventory.references[0]!.scope_ids.push('run:plain'); inventory.references[0]!.scope_ids.sort() })
    const prepared = await prepareNativeDocxPagePaintV1(input)
    const provider = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: prepared.outline_requests.map((request) => ({ status: 'outlined' as const, ...request, ...provider.outline(request.glyph_id) })) })
    expect(completed.page_paint_output.status).toBe('painted')
    if (completed.page_paint_output.status !== 'painted') throw new Error('highlight refused')
    const page = completed.page_paint_output.pages[0]!
    const fragment = prepared.page_paint_request.pagination_request.shaped_lines.paragraphs[0]!.lines[0]!.fragments[0]!
    expect(page.commands[0]).toMatchObject({ kind: 'fill_text_highlight', fill_rgb: 'FFFF00', x_millipoints: page.lines[0]!.x_millipoints, y_millipoints: page.lines[0]!.baseline_y_millipoints - fragment.ascent_millipoints, width_millipoints: fragment.advance_inline_millipoints, height_millipoints: fragment.ascent_millipoints - fragment.descent_millipoints })
    expect(page.commands.map((command) => command.kind)).toEqual(['fill_text_highlight', 'fill_text_highlight', 'fill_glyph_path', 'fill_glyph_path', 'fill_glyph_path'])
    expect(page.commands.at(-1)).toMatchObject({ source_id: 'run:plain' })
    for (const patch of [{ fill_rgb: 'FF0000' }, { x_millipoints: 1 }, { width_millipoints: 1 }, { source_id: 'run:other' }]) {
      const tampered = structuredClone(completed.page_paint_output)
      Object.assign(tampered.pages[0]!.commands[0]!, patch)
      expect(decodeNativeDocxPagePaintForRequestV1(tampered, completed.page_paint_request, completed.page_paint_request.outline_provider).ok).toBe(false)
    }
    const missing = structuredClone(completed.page_paint_output)
    missing.pages[0]!.commands.shift(); missing.pages[0]!.lines[0]!.command_ids.shift()
    expect(decodeNativeDocxPagePaintForRequestV1(missing, completed.page_paint_request, completed.page_paint_request.outline_provider).ok).toBe(false)
    const reordered = structuredClone(completed.page_paint_output)
    reordered.pages[0]!.commands.reverse(); reordered.pages[0]!.lines[0]!.command_ids.reverse()
    expect(decodeNativeDocxPagePaintForRequestV1(reordered, completed.page_paint_request, completed.page_paint_request.outline_provider).ok).toBe(false)
  })
  it('keeps highlight backgrounds behind neighboring overhanging glyph ink', async () => {
    const input = fixture()
    ;(input.document as NativeDocxDocumentV1).body.blocks[0]!.paragraph!.runs[0]!.text = 'AA'
    ;(input.resolved_layout as NativeDocxResolvedLayoutInputV1).runs[0]!.properties.highlight = 'cyan'
    const prepared = await prepareNativeDocxPagePaintV1(input)
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: prepared.outline_requests.map((request) => ({ status: 'outlined' as const, ...request, units_per_em: 2048, path: [{ kind: 'move_to' as const, x: 0, y: 0 }, { kind: 'line_to' as const, x: 3000, y: 0 }, { kind: 'line_to' as const, x: 3000, y: 1000 }, { kind: 'close_path' as const }] })) })
    expect(completed.page_paint_output.status).toBe('painted')
    if (completed.page_paint_output.status !== 'painted') throw new Error('highlight refused')
    const commands = completed.page_paint_output.pages[0]!.commands
    expect(commands.map((command) => command.kind)).toEqual(['fill_text_highlight', 'fill_text_highlight', 'fill_glyph_path', 'fill_glyph_path'])
    const background = commands[0]!, glyph = commands[2]!
    if (background.kind !== 'fill_text_highlight' || glyph.kind !== 'fill_glyph_path') throw new Error('missing commands')
    expect(glyph.path.some((point) => 'x_millipoints' in point && point.x_millipoints > background.x_millipoints + background.width_millipoints)).toBe(true)
  })
  it('deterministically joins real HarfBuzz shaping, pagination, and all-or-nothing page paint', async () => {
    const first = await prepareNativeDocxPagePaintV1(fixture())
    const second = await prepareNativeDocxPagePaintV1(fixture())
    expect(JSON.stringify(second)).toBe(JSON.stringify(first))
    expect(first.outline_requests).toHaveLength(1)
    expect(first.page_paint_request.paginated_layout.status).toBe('paginated')
    expect(first.providers).toMatchObject({ shaper_id: 'injoffice.harfbuzzjs', shaper_revision: expect.stringMatching(/^sha256:[0-9a-f]{64}$/) })
    const outline = first.outline_requests[0]!
    const completed = await completeNativeDocxPagePaintV1({ prepared: first, outline_results: [{
      status: 'outlined', face: outline.face, glyph_id: outline.glyph_id, units_per_em: 2_048,
      path: [{ kind: 'move_to', x: 0, y: 0 }, { kind: 'line_to', x: 1_000, y: 0 }, { kind: 'line_to', x: 1_000, y: 1_000 }, { kind: 'close_path' }],
    }] })
    expect(completed.page_paint_output.status).toBe('painted')
    if (completed.page_paint_output.status === 'painted') {
      expect(completed.page_paint_output.pages[0]).toEqual(expect.objectContaining({
        section_ids: ['section:1'],
        columns: [expect.objectContaining({ id: 'column:section:1:0', section_id: 'section:1', ordinal: 0 })],
        lines: [expect.objectContaining({ section_id: 'section:1', column_id: 'column:section:1:0', column_ordinal: 0 })],
      }))
      expect(completed.page_paint_output.provenance.paginated_layout.sha256).toBe(first.page_paint_request.integrity.paginated_layout_sha256)
    }
    expect(completed).toMatchObject({ canonical_request_validated: true, canonical_output_validated: true, outline_coverage_complete: true })
  })

  it('shapes a selected static header at full body width, repeats it over multi-column pages, and binds its canonical plan hash', async () => {
    const input = fixture()
    const headerPart = 'word/header1.xml'
    const headerAnchor = (path: string, start: number, end: number) => ({ part_name: headerPart, path, start_byte: start, end_byte: end, xml_sha256: HASH })
    const headerParagraph = {
      id: 'paragraph:header', anchor: headerAnchor('/w:hdr[1]/w:p[1]', 10, 90),
      edit_policy: { mode: 'read-only' as const, allowed_operations: [], refusal: { code: 'NATIVE_READ_ONLY', message: 'Fixture is immutable.', preservation: 'refuse-mutation' as const } },
      properties: {}, runs: [{ kind: 'text' as const, id: 'run:header', anchor: headerAnchor('/w:hdr[1]/w:p[1]/w:r[1]/w:t[1]', 20, 80), text: 'Header' }],
    }
    const document = structuredClone(input.document) as NativeDocxDocumentV1
    document.sections[0]!.page.columns = 2
    document.sections[0]!.page.column_spacing_twips = 720
    document.sections[0]!.page.column_definitions = [
      { id: 'column:section:1:0', ordinal: 0 },
      { id: 'column:section:1:1', ordinal: 1 },
    ]
    const secondBodyParagraph = structuredClone(document.body.blocks[0]!.paragraph!)
    secondBodyParagraph.id = 'paragraph:2'
    secondBodyParagraph.anchor = anchor('/w:document[1]/w:body[1]/w:p[2]', 200, 290)
    secondBodyParagraph.runs[0]!.id = 'run:2'
    secondBodyParagraph.runs[0]!.anchor = anchor('/w:document[1]/w:body[1]/w:p[2]/w:r[1]', 210, 280)
    secondBodyParagraph.runs[0]!.text = 'B'
    document.body.blocks.push({ kind: 'paragraph', id: secondBodyParagraph.id, paragraph: secondBodyParagraph })
    document.headers.push({ id: 'story:header', kind: 'header', part_name: headerPart, anchor: headerAnchor('/w:hdr[1]', 1, 100), blocks: [{ kind: 'paragraph', id: headerParagraph.id, paragraph: headerParagraph }] })
    document.sections[0]!.header_refs.push({ kind: 'default', story_id: 'story:header', relationship_id: 'rIdHeader' })
    const secondSection = structuredClone(document.sections[0]!)
    secondSection.id = 'section:2'
    secondSection.anchor = anchor('/w:document[1]/w:body[1]/w:sectPr[2]', 2_200, 2_290)
    secondSection.starts_at_block_id = secondBodyParagraph.id
    secondSection.break_type = 'next-page'
    secondSection.page.column_definitions = [
      { id: 'column:section:2:0', ordinal: 0 },
      { id: 'column:section:2:1', ordinal: 1 },
    ]
    document.sections.push(secondSection)
    input.document = document
    const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    resolved.paragraphs.push({ paragraph_id: headerParagraph.id, applied_styles: [], properties: {}, paragraph_mark_properties: { font_family: 'DejaVu Sans', font_size_half_points: 20 } })
    resolved.runs.push({ run_id: 'run:header', paragraph_id: headerParagraph.id, applied_paragraph_styles: [], applied_character_styles: [], properties: { font_family: 'DejaVu Sans', font_size_half_points: 20, color: '123456' } })
    resolved.paragraphs.push({ ...structuredClone(resolved.paragraphs[0]!), paragraph_id: secondBodyParagraph.id })
    resolved.runs.push({ ...structuredClone(resolved.runs[0]!), run_id: 'run:2', paragraph_id: secondBodyParagraph.id })
    rewriteInventory(input, (inventory) => { inventory.references[0]!.scope_ids = ['paragraph:1', 'paragraph:2', 'paragraph:header', 'run:1', 'run:2', 'run:header'] })

    const prepared = await prepareNativeDocxPagePaintV1(input)
    expect(prepared.page_paint_request.pagination_request.shaped_lines.available_width_millipoints).toBe(216_000)
    expect(prepared.page_paint_request.pagination_request.shaped_lines.paragraphs.find((paragraph) => paragraph.paragraph_id === 'paragraph:header')?.lines[0]?.available_width_millipoints).toBe(468_000)
    expect(prepared.page_paint_request.paginated_layout).toEqual(expect.objectContaining({ status: 'paginated', diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'header-footer-selection-deferred' })]) }))
    const results = prepared.outline_requests.map((outline) => ({
      status: 'outlined' as const, face: outline.face, glyph_id: outline.glyph_id, units_per_em: 2_048,
      path: [{ kind: 'move_to' as const, x: 0, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 1_000 }, { kind: 'close_path' as const }],
    }))
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: results })
    expect(completed.page_paint_output).toEqual(expect.objectContaining({ status: 'painted', provenance: expect.objectContaining({ header_footer_layout: expect.objectContaining({ protocol: 'injoffice.docx.header-footer-layout', sha256: expect.stringMatching(/^sha256:/) }) }) }))
    if (completed.page_paint_output.status === 'painted') {
      expect(completed.page_paint_output.pages.length).toBeGreaterThan(1)
      for (const page of completed.page_paint_output.pages) expect(page.lines).toEqual(expect.arrayContaining([expect.objectContaining({ paragraph_id: 'paragraph:header', region: 'header', y_millipoints: 36_000 })]))
      expect(new Set(completed.page_paint_output.pages.flatMap((page) => page.lines.filter((line) => line.region === 'header').map((line) => line.section_id)))).toEqual(new Set(['section:1', 'section:2']))
      const commandIDs = completed.page_paint_output.pages.flatMap((page) => page.commands.map((command) => command.id))
      expect(new Set(commandIDs).size).toBe(commandIDs.length)
      expect(completed.page_paint_output.pages.flatMap((page) => page.lines.filter((line) => line.region === 'header')).every((line) => line.column_id === undefined && line.column_ordinal === undefined)).toBe(true)
    }
  }, 15_000)

  it('deterministically paints a digest-bound exact-EMU inline PNG and emits canonical request/output hashes', async () => {
    const first = await prepareNativeDocxPagePaintV1(imageFixture())
    const second = await prepareNativeDocxPagePaintV1(imageFixture())
    expect(second).toEqual(first)
    expect(first.request_sha256).toMatch(/^sha256:[0-9a-f]{64}$/)
    expect(first.page_paint_request.media_assets).toMatchObject([{ part_name: 'word/media/image.png', content_digest: PNG_DIGEST, width_px: 1, height_px: 1 }])
    const outline = first.outline_requests[0]!
    const completed = await completeNativeDocxPagePaintV1({ prepared: first, outline_results: [{
      status: 'outlined', face: outline.face, glyph_id: outline.glyph_id, units_per_em: 2_048,
      path: [{ kind: 'move_to', x: 0, y: 0 }, { kind: 'line_to', x: 1_000, y: 0 }, { kind: 'line_to', x: 1_000, y: 1_000 }, { kind: 'close_path' }],
    }] })
    expect(completed.canonical_request_sha256).toBe(first.request_sha256)
    expect(completed.canonical_output_sha256).toMatch(/^sha256:[0-9a-f]{64}$/)
    expect(completed.page_paint_output.status).toBe('painted')
    if (completed.page_paint_output.status !== 'painted') return
    expect(completed.page_paint_output.resources).toEqual(first.page_paint_request.media_assets)
    expect(completed.page_paint_output.pages.flatMap((page) => page.commands).find((command) => command.kind === 'paint_inline_image')).toMatchObject({
      drawing_id: 'drawing:1', x_millipoints: 72_000, width_millipoints: 10_000, height_millipoints: 10_000,
      source_crop: { left: 0, top: 0, right: 0, bottom: 0 }, transform: { rotation_degrees: 0, flip_horizontal: false, flip_vertical: false },
    })
  })

  it('reserves exact inline effect extents without resizing image content', async () => {
    const input = imageFixture(), doc = input.document as NativeDocxDocumentV1
    doc.body.blocks[0]!.paragraph!.runs[0]!.drawing!.inline_effect_extent_emu = { left: 12_700, top: 25_400, right: 38_100, bottom: 50_800 }
    const prepared = await prepareNativeDocxPagePaintV1(input)
    const fragment = prepared.page_paint_request.pagination_request.shaped_lines.paragraphs[0]!.lines[0]!.fragments.find(f => f.source_kind === 'image')!
    expect(fragment).toMatchObject({ advance_inline_millipoints: 14_000, ascent_millipoints: 12_000, descent_millipoints: -4_000 })
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: prepared.outline_requests.map(outline => ({ status: 'outlined' as const, face: outline.face, glyph_id: outline.glyph_id, units_per_em: 2_048, path: [{ kind: 'move_to' as const, x: 0, y: 0 }, { kind: 'line_to' as const, x: 1000, y: 0 }, { kind: 'line_to' as const, x: 1000, y: 1000 }, { kind: 'close_path' as const }] })) })
    expect(completed.page_paint_output.status).toBe('painted')
    expect(completed.page_paint_output.pages.flatMap(p => p.commands).find(c => c.kind === 'paint_inline_image')).toMatchObject({ x_millipoints: 73_000, width_millipoints: 10_000, height_millipoints: 10_000 })
    const invalid = imageFixture()
    const invalidDrawing = (invalid.document as NativeDocxDocumentV1).body.blocks[0]!.paragraph!.runs[0]!.drawing!
    invalidDrawing.inline_effect_extent_emu = { left: 1, top: 0, right: 0, bottom: 0 }
    expect(qualifyNativeDocxInlineImageV1(invalid.document as NativeDocxDocumentV1, 'run:image', invalidDrawing)).toMatchObject({ ok: false, code: 'unsupported-image', message: expect.stringContaining('Inline effect extents') })
    const invalidPrepared = await prepareNativeDocxPagePaintV1(invalid)
    expect(invalidPrepared.page_paint_request.paginated_layout.status).toBe('refused')
    expect(invalidPrepared.page_paint_request.media_assets).toEqual([])
    const overflow = imageFixture(), overflowDoc = overflow.document as NativeDocxDocumentV1
    const drawing = overflowDoc.body.blocks[0]!.paragraph!.runs[0]!.drawing!
    drawing.width_emu = 12_700_000_000
    drawing.inline_effect_extent_emu = {left:127,top:0,right:0,bottom:0}
    expect(qualifyNativeDocxInlineImageV1(overflowDoc,'run:image',drawing)).toMatchObject({ok:false,code:'resource-limit'})
  })

  it('orders an inline image as an attested UAX #9 object inside an RTL paragraph', async () => {
    const input = imageFixture()
    const document = input.document as NativeDocxDocumentV1
    const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    document.body.blocks[0]!.paragraph!.runs[1]!.text = 'אב'
    resolved.paragraphs[0]!.properties = { bidi: true, alignment: 'start' }
    resolved.runs[1]!.properties = { ...resolved.runs[1]!.properties, language: 'he-IL' }

    const prepared = await prepareNativeDocxPagePaintV1(input)
    const shaped = prepared.page_paint_request.pagination_request.shaped_lines.paragraphs[0]!
    expect(shaped.direction).toBe('rtl')
    expect(shaped.lines[0]!.fragments.find((fragment) => fragment.source_kind === 'image')).toMatchObject({ source_id: 'run:image', direction: 'rtl', bidi_level: 1 })
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: prepared.outline_requests.map((outline) => ({
      status: 'outlined' as const, face: outline.face, glyph_id: outline.glyph_id, units_per_em: 2_048,
      path: [{ kind: 'move_to' as const, x: 0, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 1_000 }, { kind: 'close_path' as const }],
    })) })
    expect(completed.page_paint_output.status).toBe('painted')
    if (completed.page_paint_output.status !== 'painted') return
    const image = completed.page_paint_output.pages.flatMap((page) => page.commands).find((command) => command.kind === 'paint_inline_image')
    expect(image?.id).toMatch(/^paint:placed:.*:image$/)
  })

  it('rejects coordinated RTL image level/order tampering and same-level source text replay tampering', async () => {
    const input = imageFixture()
    const document = input.document as NativeDocxDocumentV1
    const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    document.body.blocks[0]!.paragraph!.runs[1]!.text = 'אב'
    resolved.paragraphs[0]!.properties = { bidi: true, alignment: 'start' }
    resolved.runs[1]!.properties = { ...resolved.runs[1]!.properties, language: 'he-IL' }
    const prepared = await prepareNativeDocxPagePaintV1(input)

    const tampered = structuredClone(prepared.page_paint_request.pagination_request)
    const line = tampered.shaped_lines.paragraphs[0]!.lines[0]!
    const image = line.fragments.find((fragment) => fragment.source_kind === 'image')!
    image.bidi_level = 0
    image.direction = 'ltr'
    const logical = [...line.fragments].sort((left, right) => left.logical_order - right.logical_order)
    const order = reorderNativeBidiLineV1(logical.map((fragment) => fragment.bidi_level), 1, logical.map((fragment) => fragment.whitespace))
    expect(order.ok).toBe(true)
    if (!order.ok) return
    line.fragments = order.value.visualToLogical.map((index) => logical[index]!)
    line.fragments.forEach((fragment, visualIndex) => { fragment.id = `fragment:${tampered.shaped_lines.paragraphs[0]!.paragraph_id}:0:${visualIndex}` })
    line.logical_to_visual = [...order.value.logicalToVisual]
    const tamperedDecoded = decodeNativeDocxShapedLines(tampered.shaped_lines)
    expect(tamperedDecoded.ok, JSON.stringify(tamperedDecoded)).toBe(true)
    expect(paginateNativeDocxV1(tampered)).toEqual(expect.objectContaining({ ok: false, issues: expect.arrayContaining([expect.objectContaining({ code: 'BROKEN_REFERENCE' })]) }))

    const textTampered = structuredClone(prepared.page_paint_request.pagination_request)
    const hebrew = textTampered.shaped_lines.paragraphs[0]!.lines[0]!.fragments.find((fragment) => fragment.source_kind === 'run')!
    hebrew.text = hebrew.text === 'א' ? 'ב' : 'א'
    expect(decodeNativeDocxShapedLines(textTampered.shaped_lines).ok).toBe(true)
    expect(paginateNativeDocxV1(textTampered)).toEqual(expect.objectContaining({ ok: false, issues: expect.arrayContaining([expect.objectContaining({ message: expect.stringContaining('exactly replay') })]) }))
  })

  it('atomically joins attested fonts, a native table, and digest-bound image media in one page-paint request', async () => {
    const prepared = await prepareNativeDocxPagePaintV1(combinedImageTableFixture())
    const outlined = prepared.outline_requests.map((outline) => ({
      status: 'outlined' as const, face: outline.face, glyph_id: outline.glyph_id, units_per_em: 2_048,
      path: [{ kind: 'move_to' as const, x: 0, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 1_000 }, { kind: 'close_path' as const }],
    }))
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: outlined })
    expect(completed.page_paint_output.status).toBe('painted')
    if (completed.page_paint_output.status !== 'painted') return
    expect(completed.page_paint_output.provenance).toMatchObject({
      font_manifest: { sha256: expect.stringMatching(/^sha256:/) },
      table_projection: { sha256: expect.stringMatching(/^sha256:/) },
      media_assets: { sha256: expect.stringMatching(/^sha256:/) },
    })
    expect(new Set(completed.page_paint_output.pages.flatMap((page) => page.commands.map((command) => command.kind)))).toEqual(new Set(['paint_inline_image', 'fill_glyph_path', 'fill_table_cell', 'stroke_table_border']))

    const refused = await completeNativeDocxPagePaintV1({ prepared, outline_results: outlined.map((entry, index) => index === 0 ? { status: 'refused' as const, face: entry.face, glyph_id: entry.glyph_id, code: 'missing-glyph' as const, message: 'intentional combined-fixture refusal' } : entry) })
    expect(refused.page_paint_output).toMatchObject({ status: 'refused', resources: [], pages: [] })

    const mediaTamper = combinedImageTableFixture()
    mediaTamper.media_assets = [{ ...mediaTamper.media_assets[0]!, content_type: 'image/jpeg', content_digest: `sha256:${'f'.repeat(64)}` }]
    await expect(prepareNativeDocxPagePaintV1(mediaTamper)).rejects.toThrow(/exact-join|exactly cover|exactly match|digest/i)
  })

  it('fails closed on media digest drift, malformed dimensions, oversized bytes, and non-exact EMU geometry', async () => {
    const drift = imageFixture()
    drift.media_assets = [{ ...drift.media_assets[0]!, bytes: Uint8Array.from([...PNG_BYTES.slice(0, -1), PNG_BYTES.at(-1)! ^ 1]) }]
    await expect(prepareNativeDocxPagePaintV1(drift)).rejects.toThrow(/content digest/)

    const lengthDrift = imageFixture()
    ;(lengthDrift.document as NativeDocxDocumentV1).passthrough_parts.at(-1)!.byte_length += 1
    await expect(prepareNativeDocxPagePaintV1(lengthDrift)).rejects.toThrow(/exact-join/)

    const malformed = imageFixture()
    const broken = Uint8Array.from(PNG_BYTES)
    broken[16] = 0; broken[17] = 0; broken[18] = 0; broken[19] = 0
    const brokenDigest = `sha256:${createHash('sha256').update(broken).digest('hex')}` as `sha256:${string}`
    ;(malformed.document as NativeDocxDocumentV1).passthrough_parts.at(-1)!.sha256 = brokenDigest
    malformed.media_assets = [{ ...malformed.media_assets[0]!, content_digest: brokenDigest, bytes: broken }]
    await expect(prepareNativeDocxPagePaintV1(malformed)).rejects.toThrow(/supported static/)

    const oversized = imageFixture()
    const huge = new Uint8Array(16 * 1024 * 1024 + 1)
    const hugeDigest = `sha256:${createHash('sha256').update(huge).digest('hex')}` as `sha256:${string}`
    ;(oversized.document as NativeDocxDocumentV1).passthrough_parts.at(-1)!.sha256 = hugeDigest
    ;(oversized.document as NativeDocxDocumentV1).passthrough_parts.at(-1)!.byte_length = huge.byteLength
    oversized.media_assets = [{ ...oversized.media_assets[0]!, content_digest: hugeDigest, bytes: huge }]
    await expect(prepareNativeDocxPagePaintV1(oversized)).rejects.toThrow(/bounded page-paint budget/)

    const inexact = imageFixture()
    ;(inexact.document as NativeDocxDocumentV1).body.blocks[0]!.paragraph!.runs[0]!.drawing!.width_emu = 127_001
    inexact.media_assets = []
    const prepared = await prepareNativeDocxPagePaintV1(inexact)
    expect(prepared.page_paint_request.paginated_layout.status).toBe('refused')
    expect(prepared.outline_requests).toEqual([])

    const duplicate = imageFixture()
    duplicate.media_assets = [duplicate.media_assets[0]!, duplicate.media_assets[0]!]
    await expect(prepareNativeDocxPagePaintV1(duplicate)).rejects.toThrow(/supplied more than once/)

    const missing = imageFixture()
    missing.media_assets = []
    await expect(prepareNativeDocxPagePaintV1(missing)).rejects.toThrow(/must cover every qualified unique inline picture part/)

    const vector = imageFixture()
    ;(vector.document as NativeDocxDocumentV1).body.blocks[0]!.paragraph!.runs[0]!.drawing!.content_type = 'image/svg+xml'
    ;(vector.document as NativeDocxDocumentV1).passthrough_parts.at(-1)!.content_type = 'image/svg+xml'
    vector.media_assets = []
    const vectorPrepared = await prepareNativeDocxPagePaintV1(vector)
    expect(vectorPrepared.page_paint_request.paginated_layout).toMatchObject({ status: 'refused', pages: [] })
    expect(vectorPrepared.page_paint_request.media_assets).toEqual([])
  })

  it('paints a page whose supplier offers raster parts no qualified inline picture names', async () => {
    // A package supplier walks the whole model for PNG/JPEG parts; this module
    // qualifies only inline run drawings. numbering.xml picture bullets and
    // refused drawings therefore arrive as assets that paint nothing.
    const bullets = imageFixture()
    ;(bullets.document as NativeDocxDocumentV1).passthrough_parts.push({ part_name: 'word/media/bullet.png', content_type: 'image/png', byte_length: PNG_BYTES.byteLength, sha256: PNG_DIGEST, policy: 'preserve-verbatim' })
    bullets.media_assets = [...bullets.media_assets, { part_name: 'word/media/bullet.png', content_type: 'image/png', content_digest: PNG_DIGEST, bytes: PNG_BYTES }]
    const prepared = await prepareNativeDocxPagePaintV1(bullets)
    expect(prepared.page_paint_request.media_assets.map((asset) => asset.part_name)).toEqual(['word/media/image.png'])
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: prepared.outline_requests.map((outline) => ({
      status: 'outlined' as const, face: outline.face, glyph_id: outline.glyph_id, units_per_em: 2_048,
      path: [{ kind: 'move_to' as const, x: 0, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 1_000 }, { kind: 'close_path' as const }],
    })) })
    expect(completed.page_paint_output.status).toBe('painted')
    expect(completed.page_paint_output.resources.map((resource) => resource.part_name)).toEqual(['word/media/image.png'])

    // The same holds when the only drawing naming that part is itself refused.
    const refusedDrawing = imageFixture()
    ;(refusedDrawing.document as NativeDocxDocumentV1).body.blocks[0]!.paragraph!.runs[0]!.drawing!.width_emu = 127_001
    const refusedPrepared = await prepareNativeDocxPagePaintV1(refusedDrawing)
    expect(refusedPrepared.page_paint_request.media_assets).toEqual([])

    // An over-collected part is dropped, never joined loosely: a supplied part
    // that does name a qualified picture still has to match it byte for byte.
    const tampered = imageFixture()
    tampered.media_assets = [{ ...tampered.media_assets[0]!, content_digest: `sha256:${'c'.repeat(64)}` }]
    await expect(prepareNativeDocxPagePaintV1(tampered)).rejects.toThrow(/exact-join one qualified native picture/)
  })

  it('normalizes legal ASCII MIME case while rejecting non-canonical OPC media part names', async () => {
    const upper = imageFixture()
    const upperDocument = upper.document as NativeDocxDocumentV1
    upperDocument.body.blocks[0]!.paragraph!.runs[0]!.drawing!.content_type = 'IMAGE/PNG'
    upperDocument.passthrough_parts.at(-1)!.content_type = 'IMAGE/PNG'
    upper.media_assets = [{ ...upper.media_assets[0]!, content_type: 'IMAGE/PNG' }]
    const prepared = await prepareNativeDocxPagePaintV1(upper)
    expect(prepared.page_paint_request.media_assets[0]!.content_type).toBe('image/png')

    const resource = prepared.page_paint_request.media_assets[0]!
    for (const part_name of ['word/media/image%3F.png', 'word/media/image%23.png', 'word/media/image%01.png', 'word/media/image%2E', 'word/media/image.']) {
      expect(() => decodeNativeDocxPagePaintResourceListV1([{ ...resource, part_name }]), part_name).toThrow(/identity is malformed/)
    }
  })

  it('paints a selected header PNG through the same inline image command as body pictures', async () => {
    const input = fixture()
    const document = input.document as NativeDocxDocumentV1
    const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    const headerPart = 'word/header1.xml'
    const headerAnchor = (path: string, start: number, end: number) => ({ part_name: headerPart, path, start_byte: start, end_byte: end, xml_sha256: HASH })
    const drawingRun = {
      kind: 'drawing' as const, id: 'run:header:image', anchor: headerAnchor('/w:hdr[1]/w:p[1]/w:r[1]', 20, 80),
      drawing: {
        id: 'drawing:header:image', anchor: headerAnchor('/w:hdr[1]/w:p[1]/w:r[1]/w:drawing[1]', 25, 75),
        relationship_id: 'rHeaderImage', media_part: 'word/media/header.png', content_type: 'image/png', placement: 'inline' as const,
        width_emu: 127_000, height_emu: 127_000,
        edit_policy: { mode: 'read-only' as const, allowed_operations: [], refusal: { code: 'EXTRACT_ONLY', message: 'Fixture drawing is immutable.', preservation: 'refuse-mutation' as const } },
      },
    }
    const headerParagraph = {
      id: 'paragraph:header:image', anchor: headerAnchor('/w:hdr[1]/w:p[1]', 10, 90),
      edit_policy: { mode: 'read-only' as const, allowed_operations: [], refusal: { code: 'NATIVE_READ_ONLY', message: 'Fixture is immutable.', preservation: 'refuse-mutation' as const } },
      properties: {}, runs: [drawingRun],
    }
    document.headers.push({ id: 'story:header:image', kind: 'header', part_name: headerPart, anchor: headerAnchor('/w:hdr[1]', 1, 100), blocks: [{ kind: 'paragraph', id: headerParagraph.id, paragraph: headerParagraph }] })
    document.sections[0]!.header_refs.push({ kind: 'default', story_id: 'story:header:image', relationship_id: 'rIdHeader' })
    document.passthrough_parts.push(
      { part_name: 'word/_rels/header1.xml.rels', content_type: 'application/vnd.openxmlformats-package.relationships+xml', byte_length: 1, sha256: HASH, policy: 'preserve-verbatim' },
      { part_name: 'word/media/header.png', content_type: 'image/png', byte_length: PNG_BYTES.byteLength, sha256: PNG_DIGEST, policy: 'preserve-verbatim' },
    )
    resolved.paragraphs.push({ paragraph_id: headerParagraph.id, applied_styles: [], properties: {}, paragraph_mark_properties: {} })
    resolved.runs.push({ run_id: drawingRun.id, paragraph_id: headerParagraph.id, applied_paragraph_styles: [], applied_character_styles: [], properties: {} })
    input.media_assets = [{ part_name: 'word/media/header.png', content_type: 'image/png', content_digest: PNG_DIGEST, bytes: PNG_BYTES }]

    const prepared = await prepareNativeDocxPagePaintV1(input)
    expect(prepared.page_paint_request.media_assets).toEqual(expect.arrayContaining([expect.objectContaining({ part_name: 'word/media/header.png', content_type: 'image/png' })]))
    const results = prepared.outline_requests.map((outline) => ({
      status: 'outlined' as const, face: outline.face, glyph_id: outline.glyph_id, units_per_em: 2_048,
      path: [{ kind: 'move_to' as const, x: 0, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 1_000 }, { kind: 'close_path' as const }],
    }))
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: results })
    expect(completed.page_paint_output.status).toBe('painted')
    if (completed.page_paint_output.status !== 'painted') return
    expect(completed.page_paint_output.pages.flatMap((page) => page.commands).find((command) => command.kind === 'paint_inline_image')).toMatchObject({
      kind: 'paint_inline_image', drawing_id: 'drawing:header:image', width_millipoints: 10_000, height_millipoints: 10_000,
    })
  })

  it('carries real Hebrew, neutrals, and numerals through visual-order page paint', async () => {
    const input = fixture()
    const paragraph = (input.document as NativeDocxDocumentV1).body.blocks[0]!.paragraph!
    paragraph.runs[0]!.text = 'אב 12'
    const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    resolved.paragraphs[0]!.properties = { bidi: true, alignment: 'start' }
    resolved.runs[0]!.properties = { ...resolved.runs[0]!.properties, language: 'he-IL' }
    const prepared = await prepareNativeDocxPagePaintV1(input)
    const shaped = prepared.page_paint_request.pagination_request.shaped_lines.paragraphs[0]!
    expect(shaped.direction).toBe('rtl')
    expect(shaped.lines[0]!.fragments.map((fragment) => fragment.text).join('')).toBe('12 בא')
    expect(shaped.lines[0]!.fragments.map((fragment) => fragment.direction)).toEqual(['ltr', 'ltr', 'rtl', 'rtl', 'rtl'])
    expect(prepared.providers).toMatchObject({ bidi_id: 'injoffice.bidi-js', bidi_unicode_version: '13.0.0' })
    const results = prepared.outline_requests.map((outline) => ({
      status: 'outlined' as const, face: outline.face, glyph_id: outline.glyph_id, units_per_em: 2_048,
      path: [{ kind: 'move_to' as const, x: 0, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 1_000 }, { kind: 'close_path' as const }],
    }))
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: results })
    expect(completed.page_paint_output.status).toBe('painted')
  })

  it('discards image resources and pages when a later glyph provider result refuses', async () => {
    const prepared = await prepareNativeDocxPagePaintV1(imageFixture())
    const outline = prepared.outline_requests[0]!
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: [{ status: 'refused', face: outline.face, glyph_id: outline.glyph_id, code: 'missing-glyph', message: 'fixture glyph absent' }] })
    expect(completed.page_paint_output).toMatchObject({ status: 'refused', resources: [], pages: [], diagnostics: [{ code: 'missing-glyph' }] })
    expect(completed.canonical_output_sha256).toMatch(/^sha256:[0-9a-f]{64}$/)
  })

  it('fails closed for font tampering, missing coverage, duplicate coverage, and provider drift', async () => {
    const tampered = fixture()
    tampered.font_assets[0]!.bytes = Uint8Array.from(tampered.font_assets[0]!.bytes)
    tampered.font_assets[0]!.bytes[0] ^= 0xff
    await expect(prepareNativeDocxPagePaintV1(tampered)).rejects.toThrow(/digest/)

    const prepared = await prepareNativeDocxPagePaintV1(fixture())
    await expect(completeNativeDocxPagePaintV1({ prepared, outline_results: [] })).rejects.toThrow(/exactly cover/)
    const outline = prepared.outline_requests[0]!
    const result = { status: 'outlined' as const, face: outline.face, glyph_id: outline.glyph_id, units_per_em: 2_048, path: [{ kind: 'move_to' as const, x: 0, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 1_000 }, { kind: 'close_path' as const }] }
    await expect(completeNativeDocxPagePaintV1({ prepared, outline_results: [result, result] })).rejects.toThrow()
    const drifted = structuredClone(prepared)
    drifted.page_paint_request.outline_provider.provider_revision = 'drifted'
    await expect(completeNativeDocxPagePaintV1({ prepared: drifted, outline_results: [result] })).rejects.toThrow()
  })

  it('paints qualified table shading, cell text, and deterministic borders from one canonical pipeline', async () => {
    const prepared = await prepareNativeDocxPagePaintV1(tableFixture())
    expect(prepared.page_paint_request.paginated_layout.status).toBe('paginated')
    expect(prepared.outline_requests).toHaveLength(1)
    const outline = prepared.outline_requests[0]!
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: [{ status: 'outlined', face: outline.face, glyph_id: outline.glyph_id, units_per_em: 2_048, path: [{ kind: 'move_to', x: 0, y: 0 }, { kind: 'line_to', x: 1_000, y: 0 }, { kind: 'line_to', x: 1_000, y: 1_000 }, { kind: 'close_path' }] }] })
    expect(completed.page_paint_output.status).toBe('painted')
    if (completed.page_paint_output.status !== 'painted') return
    expect(completed.page_paint_output.pages[0]!.commands.map((command) => command.kind)).toEqual(['fill_table_cell', 'fill_glyph_path', 'stroke_table_border', 'stroke_table_border', 'stroke_table_border', 'stroke_table_border'])
    expect(completed.page_paint_output.provenance.table_projection.sha256).toMatch(/^sha256:[0-9a-f]{64}$/)
  })
  it('strict requests attest the strict table projection for approximate-only-qualifiable tables', async () => {
    const input = tableFixture(), document = input.document as NativeDocxDocumentV1
    const table = document.body.blocks[0]!.table!
    delete table.layout; delete table.alignment; delete table.indent_twips; delete table.width_twips; delete table.cell_margins
    const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    const strict = qualifyNativeDocxTablesV1(document, resolved)
    expect(strict.status).toBe('refused')
    expect(qualifyApproximateLegacyTables(document, resolved, undefined, { legacy_compatibility_mode: 14 }).status).toBe('qualified')
    expect(qualifyApproximateLegacyTables(document, resolved, undefined)).toEqual(strict)
    const prepared = await prepareNativeDocxPagePaintV1(input)
    expect(prepared.page_paint_request.paginated_layout.status).toBe('refused')
    expect(prepared.page_paint_request.paginated_layout.diagnostics.some(diagnostic => diagnostic.code === 'body-table-unsupported')).toBe(true)
    expect(prepared.page_paint_request.integrity.table_projection_sha256).toBe(nativeDocxTableProjectionSha256V1([]))
    expect(decodeNativeDocxPagePaintRequestV1(prepared.page_paint_request).ok).toBe(true)
  })
  it.each([false, true])('autofits unequal text columns with resolved geometry %s and replays the width policy', async (inherited) => {
    const input = tableFixture(), document = input.document as NativeDocxDocumentV1, resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    const table = document.body.blocks[0]!.table!, left = table.rows[0]!.cells[0]!
    table.layout = 'autofit'; table.width_twips = 4_000; table.grid_widths_twips = [4_680, 4_680]
    left.width_twips = 4_680; left.paragraphs[0]!.runs[0]!.text = 'ID'
    const right = structuredClone(left), paragraph = right.paragraphs[0]!, oldParagraph = paragraph.id, oldRun = paragraph.runs[0]!.id
    right.id += ':right'; paragraph.id += ':right'; paragraph.runs[0]!.id += ':right'; paragraph.runs[0]!.text = 'Longer content '.repeat(12)
    table.rows[0]!.cells.push(right)
    resolved.paragraphs.push({ ...structuredClone(resolved.paragraphs.find(p=>p.paragraph_id===oldParagraph)!), paragraph_id: paragraph.id })
    resolved.runs.push({ ...structuredClone(resolved.runs.find(r=>r.run_id===oldRun)!), paragraph_id:paragraph.id, run_id:paragraph.runs[0]!.id })
    rewriteInventory(input, inventory=>{inventory.references[0]!.scope_ids.push(paragraph.id,paragraph.runs[0]!.id);inventory.references[0]!.scope_ids.sort()})
    if (inherited) {
      resolved.tables[0]!.geometry = { layout: 'autofit', alignment: 'left', indent_twips: table.indent_twips!, width_type: 'dxa', width_value: table.width_twips!, cell_margins: { ...table.cell_margins! } }
      delete table.layout; delete table.alignment; delete table.indent_twips; delete table.width_twips; delete table.cell_margins
    }
    const original = JSON.stringify(document), prepared = await prepareNativeDocxPagePaintV1(input)
    const provider = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: prepared.outline_requests.map(request=>{const outline=provider.outline(request.glyph_id);return outline.path.length?{status:'outlined' as const,...request,...outline}:{status:'empty' as const,...request,units_per_em:outline.units_per_em}}) })
    if(completed.page_paint_output.status!=='painted')throw new Error(JSON.stringify(completed.page_paint_output))
    const fills=completed.page_paint_output.pages[0]!.commands.filter(command=>command.kind==='fill_table_cell')
    expect(fills).toHaveLength(2)
    expect(fills[0]!.width_millipoints).toBeLessThan(fills[1]!.width_millipoints)
    expect(fills.reduce((sum,fill)=>sum+fill.width_millipoints,0)).toBe(200_000)
    const finalShaped=prepared.page_paint_request.pagination_request.shaped_lines
    const qualified=qualifyNativeDocxTablesV1(document,resolved,finalShaped)
    expect(qualified).toMatchObject({status:'qualified',tables:[{width_policy:{name:'shaped-content-minmax-v1',source_grid_widths_twips:[4680,4680],preferred_width_twips:4000}}]})
    expect(qualifyNativeDocxTablesV1(document,resolved,{...finalShaped,revision:'stale'})).toMatchObject({status:'refused',tables:[]})
    expect(prepared.page_paint_request.pagination_request.shaped_lines.paragraphs[1]!.lines.length).toBeGreaterThan(1)
    expect(decodeNativeDocxPagePaintForRequestV1(completed.page_paint_output,completed.page_paint_request,completed.page_paint_request.outline_provider).ok).toBe(true)
    expect(JSON.stringify(document)).toBe(original)
    const tampered=structuredClone(completed.page_paint_output)
    const fill=tampered.pages[0]!.commands.find(command=>command.kind==='fill_table_cell')!
    if(fill.kind==='fill_table_cell')fill.width_millipoints+=50
    expect(decodeNativeDocxPagePaintForRequestV1(tampered,completed.page_paint_request,completed.page_paint_request.outline_provider).ok).toBe(false)
    // Auto/omitted preferred widths shrink to intrinsic max content, capped by
    // the source section; this is not equal-grid scaling in disguise.
    delete table.width_twips
    if (inherited) { resolved.tables[0]!.geometry!.width_type='auto'; resolved.tables[0]!.geometry!.width_value=0 }
    for(const cell of table.rows[0]!.cells)delete cell.width_twips
    const auto=await prepareNativeDocxPagePaintV1(input)
    expect(auto.page_paint_request.paginated_layout.status).toBe('paginated')
    const automatic=qualifyNativeDocxTablesV1(document,resolved,auto.page_paint_request.pagination_request.shaped_lines)
    expect(automatic.status).toBe('qualified')
    if(automatic.status==='qualified') {
      expect(automatic.tables[0]!.width_millipoints).toBeLessThanOrEqual(468_000)
      expect(automatic.tables[0]!.grid_widths_millipoints[0]).toBeLessThan(automatic.tables[0]!.grid_widths_millipoints[1]!)
    }
  }, 15_000)
  it.each([false,true])('preserves agreed source widths and an empty column with inherited geometry %s', async inherited=>{
    const input=tableFixture(),document=input.document as NativeDocxDocumentV1,resolved=input.resolved_layout as NativeDocxResolvedLayoutInputV1
    const table=document.body.blocks[0]!.table!,left=table.rows[0]!.cells[0]!
    table.layout='autofit';delete table.width_twips;table.grid_widths_twips=[2000,3000];left.width_twips=2000;left.paragraphs[0]!.runs[0]!.text='ID'
    const right=structuredClone(left),p=right.paragraphs[0]!,oldP=p.id,oldR=p.runs[0]!.id
    right.id+=':empty';right.width_twips=3000;p.id+=':empty';p.runs[0]!.id+=':empty';p.runs[0]!.text='';table.rows[0]!.cells.push(right)
    resolved.paragraphs.push({...structuredClone(resolved.paragraphs.find(x=>x.paragraph_id===oldP)!),paragraph_id:p.id})
    resolved.runs.push({...structuredClone(resolved.runs.find(x=>x.run_id===oldR)!),paragraph_id:p.id,run_id:p.runs[0]!.id})
    rewriteInventory(input,i=>{i.references[0]!.scope_ids.push(p.id,p.runs[0]!.id);i.references[0]!.scope_ids.sort()})
    if(inherited){
      resolved.tables[0]!.geometry={layout:'autofit',alignment:'left',indent_twips:table.indent_twips!,width_type:'auto',width_value:0,cell_margins:{...table.cell_margins!}}
      delete table.layout;delete table.alignment;delete table.indent_twips;delete table.cell_margins
    }
    const before=JSON.stringify(input),prepared=await prepareNativeDocxPagePaintV1(input),shaped=prepared.page_paint_request.pagination_request.shaped_lines
    const q=qualifyNativeDocxTablesV1(document,resolved,shaped)
    expect(q).toMatchObject({status:'qualified',tables:[{width_millipoints:250000,grid_widths_millipoints:[100000,150000],width_policy:{name:'source-preferred-nonconflicting-v1',source_grid_widths_twips:[2000,3000],source_cell_widths_twips:[[2000,3000]],preferred_width_twips:null}}]})
    const provider=createHarfBuzzOutlineProviderV1({bytes:FONT_BYTES,contentDigest:FONT_DIGEST})
    const completed=await completeNativeDocxPagePaintV1({prepared,outline_results:prepared.outline_requests.map(request=>{const o=provider.outline(request.glyph_id);return o.path.length?{status:'outlined' as const,...request,...o}:{status:'empty' as const,...request,units_per_em:o.units_per_em}})})
    if(completed.page_paint_output.status!=='painted'||q.status!=='qualified')throw new Error('Source-preferred paint refused')
    expect(completed.page_paint_output.pages[0]!.commands.filter(c=>c.kind==='fill_table_cell').map(c=>c.width_millipoints)).toEqual([100000,150000])
    expect(completed.page_paint_output.provenance.table_projection.sha256).toBe(q.sha256)
    expect(decodeNativeDocxPagePaintForRequestV1(completed.page_paint_output,completed.page_paint_request,completed.page_paint_request.outline_provider).ok).toBe(true)
    expect(JSON.stringify(input)).toBe(before)
    expect(qualifyNativeDocxTablesV1(document,{...resolved,revision:'stale'},shaped).status).toBe('refused')
    expect(qualifyNativeDocxTablesV1(document,resolved,{...shaped,document_id:'foreign'}).status).toBe('refused')
    const tampered=structuredClone(completed.page_paint_output)
    tampered.provenance.table_projection.sha256='sha256:'+'0'.repeat(64)
    expect(decodeNativeDocxPagePaintForRequestV1(tampered,completed.page_paint_request,completed.page_paint_request.outline_provider).ok).toBe(false)
    // A cell that states no absolute preferred width states nothing to
    // conflict with (ECMA-376 17.4.72), so the authored grid still governs.
    const omitted=structuredClone(document);delete omitted.body.blocks[0]!.table!.rows[0]!.cells[1]!.width_twips
    const omittedQualified=qualifyNativeDocxTablesV1(omitted,resolved,shaped)
    expect(omittedQualified).toMatchObject({status:'qualified',tables:[{width_millipoints:250000,grid_widths_millipoints:[100000,150000],width_policy:{name:'source-preferred-nonconflicting-v1',source_cell_widths_twips:[[2000,null]]}}]})
    for(const change of ['conflicting','explicit-table','section-overflow','content-overflow'] as const){
      const d=structuredClone(document),r=structuredClone(resolved),s=structuredClone(shaped),t=d.body.blocks[0]!.table!
      if(change==='conflicting')t.rows[0]!.cells[1]!.width_twips=2999
      if(change==='explicit-table')t.width_twips=5000
      if(change==='section-overflow'){t.grid_widths_twips=[10000,10000];for(const cell of t.rows[0]!.cells)cell.width_twips=10000}
      if(change==='content-overflow')s.paragraphs[0]!.lines[0]!.fragments[0]!.advance_inline_millipoints=150000
      const fallback=qualifyNativeDocxTablesV1(d,r,s);expect(fallback.status).toBe('qualified')
      if(fallback.status==='qualified'){expect(fallback.tables[0]!.width_policy?.name).toBe('shaped-content-minmax-v1');expect(fallback.sha256).not.toBe(q.sha256)}
    }
  },15000)
  /** tdf117297_tableStyle.docx and tdf118812_tableStyles-comprehensive.docx.
   * The intrinsic-width probe used to refuse the whole document over any
   * shaping diagnostic at all. A paint-only resolved-layout diagnostic is
   * propagated under its own code precisely because it does not change a
   * shaping advance, so it cannot change an intrinsic width; and a measurement
   * the approximate lane cannot take selects its declared authored-grid policy
   * instead of discarding every other block on the page. */
  it.each(['paint-only','blocking'] as const)('measures intrinsic table widths through a %s resolved diagnostic outside the table', async (kind) => {
    const input = tableFixture(), document = input.document as NativeDocxDocumentV1, resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    const table = document.body.blocks[0]!.table!
    table.layout = 'autofit'; delete table.width_twips
    const paragraph = structuredClone(table.rows[0]!.cells[0]!.paragraphs[0]!), oldParagraph = paragraph.id, oldRun = paragraph.runs[0]!.id
    paragraph.id = 'paragraph:beside-autofit'; paragraph.runs[0]!.id = 'run:beside-autofit'
    document.body.blocks.push({ kind: 'paragraph', id: paragraph.id, paragraph })
    resolved.paragraphs.push({ ...structuredClone(resolved.paragraphs.find((entry) => entry.paragraph_id === oldParagraph)!), paragraph_id: paragraph.id })
    resolved.runs.push({ ...structuredClone(resolved.runs.find((entry) => entry.run_id === oldRun)!), paragraph_id: paragraph.id, run_id: paragraph.runs[0]!.id })
    rewriteInventory(input, (inventory) => { inventory.references[0]!.scope_ids.push(paragraph.id, paragraph.runs[0]!.id); inventory.references[0]!.scope_ids.sort() })
    resolved.diagnostics.push({
      code: kind === 'paint-only' ? 'THEME_COLOR_PRESERVED' : 'SCRIPT_FONT_PRESERVED', severity: 'unsupported', scope_id: paragraph.runs[0]!.id,
      part_name: 'word/document.xml', path: '/w:document[1]/w:body[1]/w:p[1]/w:r[1]', preservation: 'preserve-verbatim', message: 'Preserved for a future painter',
    })
    const before = JSON.stringify(document)
    // A paint-only diagnostic is propagated by shaping under a code that states
    // it does not change an advance, so the intrinsic probe stays usable and
    // the table keeps its authored grid. A blocking one really does leave the
    // paragraph unshaped, and strict paint still owes the caller that refusal.
    if (kind === 'blocking') {
      await expect(prepareNativeDocxPagePaintV1(input)).rejects.toThrow('Content autofit measurement refused')
      expect(JSON.stringify(document)).toBe(before)
      return
    }
    const prepared = await prepareNativeDocxPagePaintV1(input)
    expect(prepared.page_paint_request.paginated_layout.status).toBe('paginated')
    expect(qualifyNativeDocxTablesV1(document, resolved, prepared.page_paint_request.pagination_request.shaped_lines)).toMatchObject({
      status: 'qualified', tables: [{ width_millipoints: 9_360 * 50, width_policy: { name: 'source-preferred-nonconflicting-v1' } }],
    })
    expect(JSON.stringify(document)).toBe(before)
  }, 15000)

  it('composes content autofit with a pagination-dependent body field without changing either source', async () => {
    const input = tableFixture(), document = input.document as NativeDocxDocumentV1, resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    const table = document.body.blocks[0]!.table!
    table.layout = 'autofit'
    const paragraph = structuredClone(table.rows[0]!.cells[0]!.paragraphs[0]!), oldParagraph = paragraph.id, oldRun = paragraph.runs[0]!.id
    paragraph.id = 'paragraph:after-autofit'; paragraph.runs[0]!.id = 'run:after-autofit'
    paragraph.properties.page_break_before = true
    paragraph.runs[0]!.text = ''; paragraph.runs[0]!.page_field = 'PAGE'
    document.body.blocks.push({kind:'paragraph',id:paragraph.id,paragraph})
    resolved.paragraphs.push({...structuredClone(resolved.paragraphs.find(p=>p.paragraph_id===oldParagraph)!),paragraph_id:paragraph.id,properties:{page_break_before:true}})
    resolved.runs.push({...structuredClone(resolved.runs.find(r=>r.run_id===oldRun)!),paragraph_id:paragraph.id,run_id:paragraph.runs[0]!.id})
    rewriteInventory(input,inventory=>{inventory.references[0]!.scope_ids.push(paragraph.id,paragraph.runs[0]!.id);inventory.references[0]!.scope_ids.sort()})
    const before=JSON.stringify(document),prepared=await prepareNativeDocxPagePaintV1(input),request=prepared.page_paint_request
    expect(JSON.stringify(document)).toBe(before)
    expect(request.body_field_source).toEqual(document)
    expect(request.pagination_request.document.body.blocks[1]!.paragraph!.runs[0]!.text).toBe('2')
    expect(request.paginated_layout.status).toBe('paginated')
    expect(decodeNativeDocxPagePaintRequestV1(request).ok).toBe(true)
  })
  it('refuses unsatisfied content minima and unsupported autofit spacing/merge policies atomically', async () => {
    for (const mode of ['wide-word', 'indent', 'merge', 'percent'] as const) {
      const input=tableFixture(), document=input.document as NativeDocxDocumentV1, resolved=input.resolved_layout as NativeDocxResolvedLayoutInputV1
      const table=document.body.blocks[0]!.table!
      table.layout='autofit'
      if(mode==='wide-word')table.rows[0]!.cells[0]!.paragraphs[0]!.runs[0]!.text='W'.repeat(100)
      if(mode==='indent')resolved.paragraphs[0]!.properties.indent_start_twips=100
      if(mode==='merge')table.rows[0]!.cells[0]!.vertical_merge='restart'
      if(mode==='percent'){delete table.width_twips;table.width_percent_fiftieths=2500}
      await expect(prepareNativeDocxPagePaintV1(input)).rejects.toThrow(/autofit/i)
    }
  })
  it('paints source-bound natural row fragments without duplicating glyphs or full-row shading', async () => {
    const input = tableFixture(), document = input.document as NativeDocxDocumentV1
    const table = document.body.blocks[0]!.table!
    table.rows[0]!.cant_split = false
    table.width_twips = 4_680; table.grid_widths_twips = [4_680]; table.rows[0]!.cells[0]!.width_twips = 4_680
    table.rows[0]!.cells[0]!.paragraphs[0]!.runs[0]!.text = 'a '.repeat(200)
    document.sections[0]!.page.height_twips = 4_480
    document.sections[0]!.page.orientation = 'landscape'
    const original = JSON.stringify(document), prepared = await prepareNativeDocxPagePaintV1(input)
    if (prepared.page_paint_request.paginated_layout.status !== 'paginated') throw new Error(JSON.stringify(prepared.page_paint_request.paginated_layout))
    const provider = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: prepared.outline_requests.map(request => { const outline = provider.outline(request.glyph_id); return outline.path.length ? { status: 'outlined' as const, ...request, ...outline } : { status: 'empty' as const, ...request, units_per_em: outline.units_per_em } }) })
    if (completed.page_paint_output.status !== 'painted') throw new Error(JSON.stringify(completed.page_paint_output))
    const pages = completed.page_paint_output.pages
    expect(pages.length).toBeGreaterThan(1)
    for (const page of pages) {
      const fill = page.commands.find(command => command.kind === 'fill_table_cell')!
      expect(fill).toMatchObject({ kind: 'fill_table_cell', y_millipoints: 72_000 })
      if (fill.kind === 'fill_table_cell') expect(fill.height_millipoints).toBeLessThanOrEqual(80_000)
    }
    expect(new Set(pages.flatMap(page => page.lines.map(line => line.line_id))).size).toBe(prepared.page_paint_request.pagination_request.shaped_lines.paragraphs[0]!.lines.length)
    expect(decodeNativeDocxPagePaintForRequestV1(completed.page_paint_output, completed.page_paint_request, completed.page_paint_request.outline_provider).ok).toBe(true)
    expect(JSON.stringify(document)).toBe(original)
  })
  it('reflows cell shaping to the exact percentage table width and binds the width policy in provenance', async () => {
    const input = tableFixture(), table = (input.document as NativeDocxDocumentV1).body.blocks[0]!.table!
    delete table.width_twips; table.width_percent_fiftieths = 2500
    const original = JSON.stringify(input.document)
    const prepared = await prepareNativeDocxPagePaintV1(input)
    expect(prepared.page_paint_request.pagination_request.shaped_lines.paragraphs[0]!.lines[0]!.available_width_millipoints).toBe(224_000)
    const provider = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: prepared.outline_requests.map((request) => ({ status:'outlined' as const,...request,...provider.outline(request.glyph_id) })) })
    expect(completed.page_paint_output.status).toBe('painted')
    if(completed.page_paint_output.status!=='painted')throw new Error('percentage paint refused')
    expect(completed.page_paint_output.pages[0]!.commands.find(command=>command.kind==='fill_table_cell')).toMatchObject({ width_millipoints:234_000 })
    expect(decodeNativeDocxPagePaintForRequestV1(completed.page_paint_output,completed.page_paint_request,completed.page_paint_request.outline_provider).ok).toBe(true)
    expect(JSON.stringify(input.document)).toBe(original)
  })

  it('paints source-bound repeated table headings, shading and borders on continuation pages', async () => {
    const input = tableFixture()
    const document = input.document as NativeDocxDocumentV1
    const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    const table = document.body.blocks[0]!.table!
    table.rows[0]!.repeat_header = true
    for (let index = 2; index <= 4; index += 1) {
      const row = structuredClone(table.rows[0]!)
      row.id = `row:${index}`; row.repeat_header = false
      row.cells[0]!.id = `cell:${index}`
      const paragraph = row.cells[0]!.paragraphs[0]!
      const oldParagraph = paragraph.id, oldRun = paragraph.runs[0]!.id
      paragraph.id = `paragraph:table:${index}`; paragraph.runs[0]!.id = `run:table:${index}`
      resolved.paragraphs.push({ ...structuredClone(resolved.paragraphs.find((entry) => entry.paragraph_id === oldParagraph)!), paragraph_id: paragraph.id })
      resolved.runs.push({ ...structuredClone(resolved.runs.find((entry) => entry.run_id === oldRun)!), paragraph_id: paragraph.id, run_id: paragraph.runs[0]!.id })
      table.rows.push(row)
      rewriteInventory(input, (inventory) => { inventory.references[0]!.scope_ids.push(paragraph.id, paragraph.runs[0]!.id); inventory.references[0]!.scope_ids.sort() })
    }
    // Real adjacent cells overlap vertically but occupy distinct source-bound
    // cell flows; generic paragraph overlap rejection must not reject a grid.
    table.grid_widths_twips = [4_680, 4_680]
    for (const row of table.rows) {
      row.cells[0]!.width_twips = 4_680
      const cell = structuredClone(row.cells[0]!)
      cell.id += ':right'
      const paragraph = cell.paragraphs[0]!, oldParagraph = paragraph.id, oldRun = paragraph.runs[0]!.id
      paragraph.id += ':right'; paragraph.runs[0]!.id += ':right'
      resolved.paragraphs.push({ ...structuredClone(resolved.paragraphs.find((entry) => entry.paragraph_id === oldParagraph)!), paragraph_id: paragraph.id })
      resolved.runs.push({ ...structuredClone(resolved.runs.find((entry) => entry.run_id === oldRun)!), paragraph_id: paragraph.id, run_id: paragraph.runs[0]!.id })
      row.cells.push(cell)
      rewriteInventory(input, (inventory) => { inventory.references[0]!.scope_ids.push(paragraph.id, paragraph.runs[0]!.id); inventory.references[0]!.scope_ids.sort() })
    }
    // Natural 10 pt font metrics plus 10 pt margins; 50 pt body fits two rows.
    document.sections[0]!.page.height_twips = 3_880
    document.sections[0]!.page.orientation = 'landscape'
    const prepared = await prepareNativeDocxPagePaintV1(input)
    const provider = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: prepared.outline_requests.map((request) => ({ status: 'outlined' as const, ...request, ...provider.outline(request.glyph_id) })) })
    expect(completed.page_paint_output.status).toBe('painted')
    if (completed.page_paint_output.status !== 'painted') throw new Error(JSON.stringify(completed.page_paint_output))
    const pages = completed.page_paint_output.pages
    expect(pages.length).toBeGreaterThan(1)
    const headerParagraph = table.rows[0]!.cells[0]!.paragraphs[0]!.id
    for (const page of pages) {
      expect(page.lines[0]!.paragraph_id).toBe(headerParagraph)
      expect(page.commands.some((command) => command.kind === 'fill_table_cell' && command.row_id === 'row:1')).toBe(true)
      expect(page.commands.some((command) => command.kind === 'stroke_table_border' && command.row_id === 'row:1')).toBe(true)
    }
    const allIDs = pages.flatMap((page) => page.commands.map((command) => command.id))
    expect(new Set(allIDs).size).toBe(allIDs.length)
    expect(decodeNativeDocxPagePaintForRequestV1(completed.page_paint_output, completed.page_paint_request, completed.page_paint_request.outline_provider).ok).toBe(true)
    const tampered = structuredClone(completed.page_paint_output)
    const repeatedFill = tampered.pages[1]!.commands.find((command) => command.kind === 'fill_table_cell')!
    if (repeatedFill.kind === 'fill_table_cell') repeatedFill.y_millipoints += 1
    expect(decodeNativeDocxPagePaintForRequestV1(tampered, completed.page_paint_request, completed.page_paint_request.outline_provider).ok).toBe(false)
  })

  it('anchors table paint to the first placed line of a multiline cell paragraph', async () => {
    const input = tableFixture()
    const document = input.document as NativeDocxDocumentV1
    document.body.blocks[0]!.table!.rows[0]!.cells[0]!.paragraphs[0]!.runs[0]!.text = 'A '.repeat(80)
    const prepared = await prepareNativeDocxPagePaintV1(input)
    expect(prepared.page_paint_request.paginated_layout.status).toBe('paginated')
    if (prepared.page_paint_request.paginated_layout.status !== 'paginated') return
    expect(prepared.page_paint_request.paginated_layout.pages[0]!.lines.length).toBeGreaterThan(1)
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: prepared.outline_requests.map((outline) => ({
      status: 'outlined' as const, face: outline.face, glyph_id: outline.glyph_id, units_per_em: 2_048,
      path: [{ kind: 'move_to' as const, x: 0, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 1_000 }, { kind: 'close_path' as const }],
    })) })
    expect(completed.page_paint_output.status).toBe('painted')
    if (completed.page_paint_output.status !== 'painted') return
    const page = completed.page_paint_output.pages[0]!
    const fill = page.commands.find((command) => command.kind === 'fill_table_cell')
    expect(fill?.y_millipoints).toBe(page.body_box.y_millipoints)
  })

  it('refuses cell-border conflicts before requesting any glyph outlines', async () => {
    const input = tableFixture()
    ;(input.document as NativeDocxDocumentV1).body.blocks[0]!.table!.rows[0]!.cells[0]!.borders = {
      top: { style: 'single', size_eighth_points: 8, color_rgb: '000000' },
    }
    const prepared = await prepareNativeDocxPagePaintV1(input)
    expect(prepared.page_paint_request.paginated_layout.status).toBe('refused')
    expect(prepared.outline_requests).toEqual([])
  })

  it('refuses qualified tables in multi-column flow atomically before outline requests', async () => {
    const input = tableFixture()
    const section = (input.document as NativeDocxDocumentV1).sections[0]!
    section.page.columns = 2
    section.page.column_spacing_twips = 20
    section.page.column_definitions = [
      { id: 'section:1:column:0', ordinal: 0 },
      { id: 'section:1:column:1', ordinal: 1 },
    ]
    await expect(prepareNativeDocxPagePaintV1(input)).rejects.toThrow(/table content.*multi-column/)
  })

  it('refuses the #87 manifest-substitution exploit and stale or tampered inventory attestations with no output', async () => {
    const substituted = fixture() as any
    substituted.font_manifest = {
      version: 1, manifestId: 'caller:unrelated', revision: substituted.document.revision,
      faces: [{ faceId: 'font-face:caller', family: 'DejaVu Sans', weight: 400, style: 'normal', stretch: 100, source: { kind: 'document', resourceId: `font:sha256:${'f'.repeat(64)}`, contentDigest: `sha256:${'f'.repeat(64)}` } }], fallbackChains: [],
    }
    substituted.font_assets = [{ face_id: 'font-face:caller', face_slot: 'embedRegular', resource_id: `font:sha256:${'f'.repeat(64)}`, content_digest: `sha256:${'f'.repeat(64)}`, collection_index: null, bytes: FONT_BYTES }]
    await expect(prepareNativeDocxPagePaintV1(substituted)).rejects.toThrow(/input identity/)

    const stale = fixture() as any
    stale.document.revision = `rev:${'b'.repeat(32)}`
    await expect(prepareNativeDocxPagePaintV1(stale)).rejects.toThrow(/exact-join/)

    const licenseTamper = fixture()
    rewriteInventory(licenseTamper, (inventory) => { inventory.families[0]!.faces[0]!.source.licensing.embedding_rights = 'editable' })
    await expect(prepareNativeDocxPagePaintV1(licenseTamper)).rejects.toThrow(/license flags/)

    const relationshipTamper = fixture()
    rewriteInventory(relationshipTamper, (inventory) => { inventory.families[0]!.faces[0]!.source.relationship_type = 'http://purl.oclc.org/ooxml/officeDocument/relationships/font' })
    await expect(prepareNativeDocxPagePaintV1(relationshipTamper)).rejects.toThrow(/relationship/)

    const relationshipHashTamper = fixture()
    rewriteInventory(relationshipHashTamper, (inventory) => { inventory.font_table!.main_relationships_sha256 = `sha256:${'c'.repeat(64)}` })
    await expect(prepareNativeDocxPagePaintV1(relationshipHashTamper)).rejects.toThrow(/exact-join/)

    const fontTableHashTamper = fixture()
    rewriteInventory(fontTableHashTamper, (inventory) => { inventory.font_table!.sha256 = `sha256:${'c'.repeat(64)}` })
    await expect(prepareNativeDocxPagePaintV1(fontTableHashTamper)).rejects.toThrow(/exact-join/)

    const storedHashTamper = fixture()
    rewriteInventory(storedHashTamper, (inventory) => { inventory.families[0]!.faces[0]!.source.stored_sha256 = `sha256:${'c'.repeat(64)}` })
    await expect(prepareNativeDocxPagePaintV1(storedHashTamper)).rejects.toThrow(/exact-join/)

    const incomplete = fixture()
    incomplete.font_assets = []
    await expect(prepareNativeDocxPagePaintV1(incomplete)).rejects.toThrow(/exactly cover/)

    const duplicateAsset = fixture() as any
    duplicateAsset.font_assets.push({ ...duplicateAsset.font_assets[0], bytes: Uint8Array.from(duplicateAsset.font_assets[0].bytes) })
    await expect(prepareNativeDocxPagePaintV1(duplicateAsset)).rejects.toThrow(/exactly cover|duplicate/)

    const partialReference = fixture() as any
    partialReference.resolved_layout.runs[0].properties.bold = true
    partialReference.resolved_layout.paragraphs[0].paragraph_mark_properties.bold = true
    rewriteInventory(partialReference, (inventory) => { inventory.references[0]!.weight = 700 })
    await expect(prepareNativeDocxPagePaintV1(partialReference)).rejects.toThrow(/every authored font reference/)

    const sourcePartMismatch = fixture() as any
    sourcePartMismatch.resolved_layout.source_parts.font_table_part = 'word/other-fontTable.xml'
    await expect(prepareNativeDocxPagePaintV1(sourcePartMismatch)).rejects.toThrow(/exact-join/)

    const targetNormalization = fixture()
    rewriteInventory(targetNormalization, (inventory) => { inventory.families[0]!.faces[0]!.source.relationship_target = 'fonts/%2E/DejaVuSans.odttf' })
    await expect(prepareNativeDocxPagePaintV1(targetNormalization)).rejects.toThrow(/relationship target/)

    const wrongResource = fixture()
    wrongResource.font_assets[0]!.resource_id = `font:sha256:${'f'.repeat(64)}`
    await expect(prepareNativeDocxPagePaintV1(wrongResource)).rejects.toThrow(/exact-join/)

    const unknown = fixture()
    unknown.font_inventory_json = unknown.font_inventory_json.replace('{"protocol":', '{"unknown":true,"protocol":')
    await expect(prepareNativeDocxPagePaintV1(unknown)).rejects.toThrow(/unknown/)

    const duplicateField = fixture()
    duplicateField.font_inventory_json = duplicateField.font_inventory_json.replace('{"protocol":', '{"protocol":"injoffice.docx.font-inventory","protocol":')
    await expect(prepareNativeDocxPagePaintV1(duplicateField)).rejects.toThrow(/canonical Go/)

    const invalidBeforeShaper = fixture() as any
    invalidBeforeShaper.document = {}
    let shaperCreations = 0
    await expect(prepareNativeDocxPagePaintV1(invalidBeforeShaper, { createShaper: () => { shaperCreations += 1; throw new Error('must not be called') } })).rejects.toThrow(/document is invalid/)
    expect(shaperCreations).toBe(0)

    const fontBeforeBidi = fixture()
    ;(fontBeforeBidi.document as NativeDocxDocumentV1).body.blocks[0]!.paragraph!.runs[0]!.text = '\u0870'
    fontBeforeBidi.font_assets[0]!.bytes = Uint8Array.from(fontBeforeBidi.font_assets[0]!.bytes)
    fontBeforeBidi.font_assets[0]!.bytes[0] ^= 0xff
    await expect(prepareNativeDocxPagePaintV1(fontBeforeBidi)).rejects.toThrow(/digest/)
  })

  it('refuses a stale or mislabeled injected HarfBuzz engine before shaping', async () => {
    const stale = createHarfBuzzTextShaperV1({ sourceRevision: 'git:stale-engine' })
    await expect(prepareNativeDocxPagePaintV1(fixture(), { createShaper: () => stale })).rejects.toThrow(/source revision/)

    const exact = createHarfBuzzTextShaperV1({ sourceRevision: 'git:integration-test' })
    const mislabeled = Object.freeze({
      ...exact,
      provenance: Object.freeze({ ...exact.provenance, provider_revision: `sha256:${'f'.repeat(64)}` }),
    })
    await expect(prepareNativeDocxPagePaintV1(fixture(), { createShaper: () => mislabeled })).rejects.toThrow(/provenance/)

    const fake = Object.freeze({
      ...exact,
      providerId: 'attacker.fake-shaper',
      provenance: Object.freeze({ ...exact.provenance, provider_id: 'attacker.fake-shaper' }),
    })
    await expect(prepareNativeDocxPagePaintV1(fixture(), { createShaper: () => fake as never })).rejects.toThrow(/provenance/)

    const selfAttestedForgery = Object.freeze({ ...exact, provenance: Object.freeze({ ...exact.provenance }), shape: exact.shape.bind(exact) })
    await expect(prepareNativeDocxPagePaintV1(fixture(), { createShaper: () => selfAttestedForgery })).rejects.toThrow(/provenance/)
  })

  it('carries exact multi-column identity through shaping, pagination, and paint', async () => {
    const input = fixture()
    const document = input.document as NativeDocxDocumentV1
    const firstSection = document.sections[0]!
    firstSection.page.columns = 2
    firstSection.page.column_spacing_twips = 720
    firstSection.page.column_definitions = [
      { id: 'column:section:1:0', ordinal: 0 },
      { id: 'column:section:1:1', ordinal: 1 },
    ]
    const secondParagraph = structuredClone(document.body.blocks[0]!.paragraph!)
    secondParagraph.id = 'paragraph:2'
    secondParagraph.anchor = anchor('/w:document[1]/w:body[1]/w:p[2]', 200, 290)
    secondParagraph.runs[0]!.id = 'run:2'
    secondParagraph.runs[0]!.anchor = anchor('/w:document[1]/w:body[1]/w:p[2]/w:r[1]', 210, 280)
    secondParagraph.runs[0]!.text = 'B'
    document.body.blocks.push({ kind: 'paragraph', id: secondParagraph.id, paragraph: secondParagraph })
    const secondSection = structuredClone(firstSection)
    secondSection.id = 'section:2'
    secondSection.anchor = anchor('/w:document[1]/w:body[1]/w:sectPr[2]', 2_200, 2_290)
    secondSection.starts_at_block_id = secondParagraph.id
    secondSection.break_type = 'next-column'
    secondSection.page.column_definitions = [
      { id: 'column:section:2:0', ordinal: 0 },
      { id: 'column:section:2:1', ordinal: 1 },
    ]
    document.sections.push(secondSection)
    const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    resolved.paragraphs.push({ ...structuredClone(resolved.paragraphs[0]!), paragraph_id: secondParagraph.id })
    resolved.runs.push({ ...structuredClone(resolved.runs[0]!), run_id: 'run:2', paragraph_id: secondParagraph.id })
    rewriteInventory(input, (inventory) => { inventory.references[0]!.scope_ids.push('paragraph:2', 'run:2'); inventory.references[0]!.scope_ids.sort() })
    const prepared = await prepareNativeDocxPagePaintV1(input)
    expect(prepared.page_paint_request.paginated_layout.status).toBe('paginated')
    if (prepared.page_paint_request.paginated_layout.status !== 'paginated') return
    expect(prepared.page_paint_request.paginated_layout.pages[0]!.section_ids).toEqual(['section:1', 'section:2'])
    expect(prepared.page_paint_request.paginated_layout.pages[0]!.lines.map((line) => [line.section_id, line.column_id, line.column_ordinal])).toEqual([
      ['section:1', 'column:section:1:0', 0],
      ['section:2', 'column:section:2:1', 1],
    ])
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: prepared.outline_requests.map((outline) => ({
      status: 'outlined' as const, face: outline.face, glyph_id: outline.glyph_id, units_per_em: 2_048,
      path: [{ kind: 'move_to', x: 0, y: 0 }, { kind: 'line_to', x: 1_000, y: 0 }, { kind: 'line_to', x: 1_000, y: 1_000 }, { kind: 'close_path' }],
    })) })
    expect(completed.page_paint_output.status).toBe('painted')
    if (completed.page_paint_output.status === 'painted') {
      expect(completed.page_paint_output.pages[0]!.section_ids).toEqual(['section:1', 'section:2'])
      expect(completed.page_paint_output.pages[0]!.lines.map((line) => [line.section_id, line.column_id, line.column_ordinal])).toEqual([
        ['section:1', 'column:section:1:0', 0],
        ['section:2', 'column:section:2:1', 1],
      ])
    }
  })
  it('shapes, balances, and paints a wrapped paragraph across equal-width columns', async () => {
    const input = fixture()
    const document = input.document as NativeDocxDocumentV1
    const section = document.sections[0]!
    section.page.columns = 2
    section.page.column_spacing_twips = 720
    section.page.column_definitions = [
      { id: 'column:section:1:0', ordinal: 0 }, { id: 'column:section:1:1', ordinal: 1 },
    ]
    document.body.blocks[0]!.paragraph!.runs[0]!.text = 'A '.repeat(200)
    const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    resolved.paragraphs[0]!.properties.widow_control = false
    const prepared = await prepareNativeDocxPagePaintV1(input)
    const layout = prepared.page_paint_request.paginated_layout
    expect(layout.status).toBe('paginated')
    if (layout.status !== 'paginated') return
    expect(layout.pages).toHaveLength(1)
    const lines = layout.pages[0]!.lines
    expect(lines.length).toBeGreaterThan(2)
    const counts = [0, 1].map((ordinal) => lines.filter((line) => line.column_ordinal === ordinal).length)
    expect(counts).toEqual([Math.ceil(lines.length / 2), Math.floor(lines.length / 2)])
    expect(layout.pages[0]!.paragraph_slices).toHaveLength(2)
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: prepared.outline_requests.map((outline) => ({
      status: 'outlined' as const, face: outline.face, glyph_id: outline.glyph_id, units_per_em: 2_048,
      path: [{ kind: 'move_to', x: 0, y: 0 }, { kind: 'line_to', x: 1_000, y: 0 }, { kind: 'line_to', x: 1_000, y: 1_000 }, { kind: 'close_path' }],
    })) })
    expect(completed.page_paint_output.status).toBe('painted')
    if (completed.page_paint_output.status === 'painted') {
      expect(completed.page_paint_output.pages[0]!.lines.map((line) => line.column_id)).toEqual(lines.map((line) => line.column_id))
    }
  })

  it('paints a multi-character source note label as the clusters the shaper produced', async () => {
    // Word's decimalZero prints the first note as "01" and lowerRoman prints the
    // second as "ii". A marker longer than one character shapes into more than
    // one visual cluster, so the painter has to hold each fragment to its own
    // slice of the marker and require the fragments to cover the whole label.
    const input = noteFixture()
    ;(input.document as NativeDocxDocumentV1).note_numbering = [{ kind: 'footnote', format: 'decimalZero', labels: ['01'] }]
    const prepared = await prepareNativeDocxPagePaintV1(input)
    const markerFragments = prepared.page_paint_request.pagination_request.shaped_lines.paragraphs
      .flatMap((paragraph) => paragraph.lines).flatMap((line) => line.fragments)
      .filter((fragment) => fragment.source_id === 'run:1' || fragment.source_id === 'run:footnote-label-1')
    expect(markerFragments.length).toBeGreaterThan(0)
    expect(markerFragments.map((fragment) => fragment.text).join('')).toBe('0101')
    const outlineResults = prepared.outline_requests.map((outline) => ({
      status: 'outlined' as const, face: outline.face, glyph_id: outline.glyph_id, units_per_em: 2_048,
      path: [{ kind: 'move_to' as const, x: 0, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 1_000 }, { kind: 'close_path' as const }],
    }))
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: outlineResults })
    expect(completed.page_paint_output.status, JSON.stringify(completed.page_paint_output.diagnostics)).toBe('painted')
  })

  it('composes exact note shaping, bottom placement, paint, and paginated-layout attestation', async () => {
    const prepared = await prepareNativeDocxPagePaintV1(noteFixture())
    expect(prepared.page_paint_request.paginated_layout.status).toBe('paginated')
    if (prepared.page_paint_request.paginated_layout.status !== 'paginated') return
    expect(prepared.page_paint_request.paginated_layout.pages[0]!.note_stories?.map((entry) => entry.note_role)).toEqual(['separator', 'content'])
    expect(prepared.page_paint_request.paginated_layout.pages[0]!.note_stories).toEqual(expect.arrayContaining([
      expect.objectContaining({ section_id: 'section:1', column_id: 'column:section:1:0', column_ordinal: 0, lines: expect.arrayContaining([expect.objectContaining({ section_id: 'section:1', column_id: 'column:section:1:0', column_ordinal: 0 })]) }),
    ]))
    const referenceGlyphs = prepared.page_paint_request.pagination_request.shaped_lines.paragraphs
      .flatMap((paragraph) => paragraph.lines)
      .flatMap((line) => line.fragments)
      .filter((fragment) => fragment.source_id === 'run:1' || fragment.source_id === 'run:footnote-label-1')
      .flatMap((fragment) => fragment.glyphs)
    expect(referenceGlyphs.length).toBeGreaterThan(0)
    expect(referenceGlyphs.every((glyph) => glyph.offset_y_millipoints === 0)).toBe(true)
    expect(prepared.page_paint_request.integrity.paginated_layout_sha256).toMatch(/^sha256:[0-9a-f]{64}$/)
    const outlineResults = prepared.outline_requests.map((outline) => ({
      status: 'outlined' as const, face: outline.face, glyph_id: outline.glyph_id, units_per_em: 2_048,
      path: [{ kind: 'move_to' as const, x: 0, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 1_000 }, { kind: 'close_path' as const }],
    }))
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: outlineResults })
    expect(completed.page_paint_output.status, JSON.stringify(completed.page_paint_output)).toBe('painted')
    if (completed.page_paint_output.status === 'painted') {
      expect(completed.page_paint_output.pages[0]!.lines.length).toBeGreaterThan(1)
      expect(completed.page_paint_output.pages[0]!.lines).toEqual(expect.arrayContaining([
        expect.objectContaining({ region: 'footnote', section_id: 'section:1', column_id: 'column:section:1:0', column_ordinal: 0 }),
      ]))
      expect(decodeNativeDocxPagePaintForRequestV1(completed.page_paint_output, prepared.page_paint_request, prepared.page_paint_request.outline_provider).ok).toBe(true)
      const provenanceDrift = structuredClone(completed.page_paint_output)
      provenanceDrift.pages[0]!.lines.find((line) => line.region === 'footnote')!.column_id = 'column:drifted'
      expect(decodeNativeDocxPagePaintForRequestV1(provenanceDrift, prepared.page_paint_request, prepared.page_paint_request.outline_provider).ok).toBe(false)
    }

    const drifted = structuredClone(prepared)
    drifted.page_paint_request.integrity.paginated_layout_sha256 = `sha256:${'0'.repeat(64)}`
    await expect(completeNativeDocxPagePaintV1({ prepared: drifted, outline_results: outlineResults })).rejects.toThrow(/paginated_layout_sha256/)
  })

  it('keeps decimal note markers inside the pinned paragraph bidi plan', async () => {
    const input = noteFixture()
    const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    resolved.paragraphs.find((paragraph) => paragraph.paragraph_id === 'paragraph:1')!.properties.bidi = true
    const prepared = await prepareNativeDocxPagePaintV1(input)
    const marker = prepared.page_paint_request.pagination_request.shaped_lines.paragraphs
      .find((paragraph) => paragraph.paragraph_id === 'paragraph:1')!.lines[0]!.fragments
      .find((fragment) => fragment.source_id === 'run:1')!
    expect(marker).toMatchObject({ text: '1', direction: 'ltr', bidi_level: 2, logical_order: 0 })
    expect(prepared.providers).toMatchObject({ bidi_id: 'injoffice.bidi-js', bidi_unicode_version: '13.0.0' })
  })

  it('atomically composes notes, digest-bound images, tables, headers, attested fonts, HarfBuzz, and pagination', async () => {
    const prepared = await prepareNativeDocxPagePaintV1(combinedNoteImageTableHeaderFixture())
    expect(prepared.providers).toMatchObject({ shaper_id: 'injoffice.harfbuzzjs', shaper_revision: expect.stringMatching(/^sha256:/) })
    expect(prepared.page_paint_request.integrity).toMatchObject({
      font_manifest_sha256: expect.stringMatching(/^sha256:/),
      table_projection_sha256: expect.stringMatching(/^sha256:/),
      media_assets_sha256: expect.stringMatching(/^sha256:/),
      paginated_layout_sha256: expect.stringMatching(/^sha256:/),
    })
    expect(prepared.page_paint_request.paginated_layout.status).toBe('paginated')
    const outlineResults = prepared.outline_requests.map((outline) => ({
      status: 'outlined' as const, face: outline.face, glyph_id: outline.glyph_id, units_per_em: 2_048,
      path: [{ kind: 'move_to' as const, x: 0, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 1_000 }, { kind: 'close_path' as const }],
    }))
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: outlineResults })
    expect(completed.page_paint_output.status, JSON.stringify(completed.page_paint_output)).toBe('painted')
    if (completed.page_paint_output.status !== 'painted') return
    expect(completed.page_paint_output.resources).toHaveLength(1)
    expect(completed.page_paint_output.provenance).toMatchObject({
      paginated_layout: { sha256: prepared.page_paint_request.integrity.paginated_layout_sha256 },
      media_assets: { sha256: prepared.page_paint_request.integrity.media_assets_sha256 },
      font_manifest: { sha256: prepared.page_paint_request.integrity.font_manifest_sha256 },
    })
    const lines = completed.page_paint_output.pages.flatMap((page) => page.lines)
    expect(lines).toEqual(expect.arrayContaining([
      expect.objectContaining({ paragraph_id: 'paragraph:header:combined' }),
      expect.objectContaining({ paragraph_id: 'paragraph:footnote-1' }),
      expect.objectContaining({ paragraph_id: 'paragraph:table:combined' }),
    ]))
    expect(new Set(completed.page_paint_output.pages.flatMap((page) => page.commands.map((command) => command.kind)))).toEqual(new Set(['paint_inline_image', 'fill_glyph_path', 'fill_table_cell', 'stroke_table_border', 'stroke_note_separator']))

    const refused = await completeNativeDocxPagePaintV1({ prepared, outline_results: outlineResults.map((entry, index) => index === 0 ? { status: 'refused' as const, face: entry.face, glyph_id: entry.glyph_id, code: 'missing-glyph' as const, message: 'late combined refusal' } : entry) })
    expect(refused.page_paint_output).toMatchObject({ status: 'refused', resources: [], pages: [] })

    for (const field of ['media_assets_sha256', 'paginated_layout_sha256'] as const) {
      const drifted = structuredClone(prepared)
      drifted.page_paint_request.integrity[field] = `sha256:${'0'.repeat(64)}`
      await expect(completeNativeDocxPagePaintV1({ prepared: drifted, outline_results: outlineResults })).rejects.toThrow(new RegExp(field))
    }
  })

  it('keeps a dormant endnote sentinel inert and refuses its unsupported semantics only on activation', async () => {
    const input = continuedEndnotePaintFixture()
    const document = input.document as NativeDocxDocumentV1
    const continuation = document.notes.at(-1)!
    document.unsupported.push({ id: 'unsupported:continuation', code: 'UNMODELED_NOTE_MARKUP', capability: 'notes', scope_id: continuation.id, preservation: 'refuse-mutation', message: 'Unsupported continuation source.' })
    document.sections[0]!.page.margins.bottom_twips = 1_440
    const fitting = await prepareNativeDocxPagePaintV1(input)
    expect(fitting.page_paint_request.paginated_layout.status).toBe('paginated')
    expect(fitting.page_paint_request.pagination_request.shaped_lines.paragraphs.some((paragraph) => paragraph.story_id === continuation.id)).toBe(false)
    document.sections[0]!.page.margins.bottom_twips = 13_200
    const active = await prepareNativeDocxPagePaintV1(input)
    expect(active.page_paint_request.paginated_layout).toEqual(expect.objectContaining({ status: 'refused', pages: [], sections: [] }))
  })

  it.each(['endnote', 'footnote'] as const)('lazily shapes and paints source-bound %s continuation with a full-width rule', async (kind) => {
    const input = continuedEndnotePaintFixture()
    if (kind === 'footnote') {
      const replace = (text: string) => text.replaceAll('endnote', 'footnote').replaceAll('Endnote', 'Footnote')
      input.document = JSON.parse(replace(JSON.stringify(input.document)))
      input.resolved_layout = JSON.parse(replace(JSON.stringify(input.resolved_layout)))
      input.font_inventory_json = replace(input.font_inventory_json)
      rewriteInventory(input, inventory => { for (const reference of inventory.references) reference.scope_ids.sort() })
    }
    const before = structuredClone(input)
    const prepared = await prepareNativeDocxPagePaintV1(input)
    const layout = prepared.page_paint_request.paginated_layout
    expect(layout.status, JSON.stringify(layout.diagnostics)).toBe('paginated')
    if (layout.status !== 'paginated') return
    expect(layout.pages.length).toBeGreaterThan(1)
    expect(layout.pages[0]!.note_stories![0]!.note_role).toBe('separator')
    expect(layout.pages.slice(1).every((page) => page.note_stories![0]!.note_role === 'continuation-separator')).toBe(true)
    expect(layout.pages.flatMap((page) => page.note_stories!.filter((story) => story.note_role === 'content').flatMap((story) => story.lines))).toHaveLength(7)
    expect(input).toEqual(before)
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: prepared.outline_requests.map((outline) => ({
      status: 'outlined' as const, face: outline.face, glyph_id: outline.glyph_id, units_per_em: 2_048,
      path: [{ kind: 'move_to', x: 0, y: 0 }, { kind: 'line_to', x: 1_000, y: 0 }, { kind: 'line_to', x: 1_000, y: 1_000 }, { kind: 'close_path' }],
    })) })
    expect(completed.page_paint_output.status, JSON.stringify(completed.page_paint_output)).toBe('painted')
    if (completed.page_paint_output.status !== 'painted') return
    const rules = completed.page_paint_output.pages.map((page) => page.commands.find((command) => command.kind === 'stroke_note_separator')!)
    expect(rules[0]!.x2_millipoints - rules[0]!.x1_millipoints).toBe(144_000)
    expect(rules.slice(1).every((rule) => rule.x2_millipoints - rule.x1_millipoints === layout.pages[0]!.body_box.width_millipoints)).toBe(true)
    const forged = structuredClone(completed.page_paint_output)
    const forgedRule = forged.pages[1]!.commands.find((command) => command.kind === 'stroke_note_separator')!
    forgedRule.x2_millipoints -= 1
    expect(decodeNativeDocxPagePaintForRequestV1(forged, completed.page_paint_request, completed.page_paint_request.outline_provider).ok).toBe(false)
  })

  it('paints the same qualified exact subset for endnotes', async () => {
    const prepared = await prepareNativeDocxPagePaintV1(endnoteFixture())
    expect(prepared.page_paint_request.paginated_layout.status).toBe('paginated')
    if (prepared.page_paint_request.paginated_layout.status !== 'paginated') return
    expect(prepared.page_paint_request.paginated_layout.pages.at(-1)!.note_stories?.map((entry) => entry.story_kind)).toEqual(['endnote', 'endnote'])
    const outlineResults = prepared.outline_requests.map((outline) => ({
      status: 'outlined' as const, face: outline.face, glyph_id: outline.glyph_id, units_per_em: 2_048,
      path: [{ kind: 'move_to' as const, x: 0, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 0 }, { kind: 'line_to' as const, x: 1_000, y: 1_000 }, { kind: 'close_path' as const }],
    }))
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: outlineResults })
    expect(completed.page_paint_output.status).toBe('painted')
  })

  it('keeps an unused continuation sentinel inert and refuses only when continuation is actually required', async () => {
    const input = noteFixture()
    const document = input.document as NativeDocxDocumentV1
    document.notes.push({
      id: 'story:footnote-continuation', kind: 'footnote', part_name: 'word/footnotes.xml', native_story_id: '0', relationship_id: 'rIdFootnotes', note_role: 'continuation-separator',
      anchor: { part_name: 'word/footnotes.xml', path: '/w:footnotes[1]/w:footnote[3]', start_byte: 191, end_byte: 230, xml_sha256: HASH }, blocks: [{
        kind: 'paragraph', id: 'paragraph:footnote-continuation', paragraph: {
          id: 'paragraph:footnote-continuation',
          anchor: { part_name: 'word/footnotes.xml', path: '/w:footnotes[1]/w:footnote[3]/w:p[1]', start_byte: 200, end_byte: 220, xml_sha256: HASH },
          edit_policy: document.notes[0]!.blocks[0]!.paragraph!.edit_policy,
          properties: {},
          runs: [],
        },
      }],
    })
    ;(input.resolved_layout as NativeDocxResolvedLayoutInputV1).paragraphs.push({
      paragraph_id: 'paragraph:footnote-continuation', applied_styles: [], properties: {},
      paragraph_mark_properties: { font_family: 'DejaVu Sans', font_size_half_points: 20, color: '123456' },
    })
    rewriteInventory(input, (inventory) => {
      inventory.references[0]!.scope_ids.push('paragraph:footnote-continuation')
      inventory.references[0]!.scope_ids.sort()
    })
    document.unsupported.push({
      id: 'unsupported:inert-continuation', code: 'UNMODELED_CONTINUATION_SENTINEL', capability: 'notes', scope_id: 'story:footnote-continuation',
      preservation: 'refuse-mutation', message: 'Dormant continuation markup remains inert for a fitting note.',
    })
    const prepared = await prepareNativeDocxPagePaintV1(input)
    expect(prepared.page_paint_request.paginated_layout).toEqual(expect.objectContaining({ status: 'paginated' }))

    const requiresContinuation = noteFixture()
    ;(requiresContinuation.document as NativeDocxDocumentV1).notes.push(structuredClone(document.notes.at(-1)!))
    ;(requiresContinuation.resolved_layout as NativeDocxResolvedLayoutInputV1).paragraphs.push({
      paragraph_id: 'paragraph:footnote-continuation', applied_styles: [], properties: {},
      paragraph_mark_properties: { font_family: 'DejaVu Sans', font_size_half_points: 20, color: '123456' },
    })
    rewriteInventory(requiresContinuation, (inventory) => {
      inventory.references[0]!.scope_ids.push('paragraph:footnote-continuation')
      inventory.references[0]!.scope_ids.sort()
    })
    ;(requiresContinuation.document as NativeDocxDocumentV1).sections[0]!.page.height_twips = 4_000
    ;(requiresContinuation.document as NativeDocxDocumentV1).sections[0]!.page.orientation = 'landscape'
    const contentRun = (requiresContinuation.document as NativeDocxDocumentV1).notes.find((story) => story.note_role === 'content')!.blocks[0]!.paragraph!.runs.find((run) => run.kind === 'text')!
    contentRun.text = 'continued note content '.repeat(20)
    const continued = await prepareNativeDocxPagePaintV1(requiresContinuation)
    const layout = continued.page_paint_request.paginated_layout
    expect(layout.status).toBe('paginated')
    if(layout.status!=='paginated')return
    expect(layout.pages.length).toBeGreaterThan(1)
    const originalLines=continued.page_paint_request.pagination_request.shaped_lines.paragraphs.find(p=>p.paragraph_id==='paragraph:footnote-1')!.lines
    expect(layout.pages.flatMap(p=>p.note_stories!.filter(n=>n.note_role==='content').flatMap(n=>n.lines.map(l=>l.line_id)))).toEqual(originalLines.map(l=>l.id))
    const provider=createHarfBuzzOutlineProviderV1({bytes:FONT_BYTES,contentDigest:FONT_DIGEST})
    const completed = await completeNativeDocxPagePaintV1({prepared:continued,outline_results:continued.outline_requests.map(request=>{
      const outline=provider.outline(request.glyph_id)
      return outline.path.length?{status:'outlined' as const,...request,...outline}:{status:'empty' as const,...request,units_per_em:outline.units_per_em}
    })})
    expect(completed.page_paint_output.status).toBe('painted')
    expect(completed.page_paint_output.pages.flatMap(p=>p.commands).filter(c=>c.kind==='fill_glyph_path'&&c.source_id==='run:footnote-label-1')).toHaveLength(1)
    ;(requiresContinuation.document as NativeDocxDocumentV1).unsupported.push(structuredClone(document.unsupported.at(-1)!))
    const refused=await prepareNativeDocxPagePaintV1(requiresContinuation)
    expect(refused.page_paint_request.paginated_layout).toMatchObject({status:'refused',pages:[]})
    expect(refused.outline_requests).toEqual([])
  }, 15_000)
})

describe('unequal whole-paragraph column compiler', () => {
  function unequalInput(): NativeDocxPagePaintPrepareInputV1 {
    const input = fixture()
    const document = input.document as NativeDocxDocumentV1
    const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    const original = structuredClone(document.body.blocks[0]!)
    document.body.blocks = Array.from({ length: 9 }, (_, index) => {
      const block = structuredClone(original)
      block.id = `paragraph:${index + 1}`
      block.paragraph!.id = block.id
      block.paragraph!.runs[0]!.id = `run:${index + 1}`
      block.paragraph!.runs[0]!.text = 'Column paragraph has enough source text to wrap at different authored widths.'
      return block
    })
    resolved.paragraphs = document.body.blocks.map((block) => ({ ...structuredClone(resolved.paragraphs[0]!), paragraph_id: block.id, properties: { keep_lines: true } }))
    resolved.runs = document.body.blocks.map((block) => ({ ...structuredClone(resolved.runs[0]!), run_id: block.paragraph!.runs[0]!.id, paragraph_id: block.id }))
    const section = document.sections[0]!
    section.page.margins.bottom_twips = 12000
    section.page.columns = 2
    section.page.column_layout = 'explicit'
    section.page.column_spacing_twips = 0
    section.page.column_definitions = [
      { id: 'column:section:1:0', ordinal: 0, width_twips: 3000, space_after_twips: 720 },
      { id: 'column:section:1:1', ordinal: 1, width_twips: 5640, space_after_twips: 0 },
    ]
    document.unsupported = [{ id: 'unsupported:unequal', code: 'UNEQUAL_SECTION_COLUMNS', capability: 'sections', scope_id: section.id, anchor: section.anchor, preservation: 'refuse-mutation', message: 'Unequal column widths require per-column shaping outside the exact v1 slice' }]
    ;(input.pagination_settings as NativeDocxPaginationSettingsV1).no_column_balance = true
    rewriteInventory(input, (inventory) => { inventory.references[0]!.scope_ids = [...resolved.paragraphs.map((p) => p.paragraph_id), ...resolved.runs.map((r) => r.run_id)].sort() })
    return input
  }

  // Real HarfBuzz shaping at both widths plus repeated source/paint replay
  // exceeds the default 5s budget on shared CI runners.
  it('shapes both widths, paints actual outlines, and binds unused candidate evidence', async () => {
    const input = unequalInput()
    const before = structuredClone(input)
    const prepared = await prepareNativeDocxPagePaintV1(input)
    const request = prepared.page_paint_request
    expect(request.paginated_layout.status).toBe('paginated')
    expect(request.pagination_request.column_shaped_lines?.map((candidate) => candidate.available_width_millipoints)).toEqual([150000, 282000])
    const candidates = request.pagination_request.column_shaped_lines!
    expect(candidates[0].paragraphs[0]!.lines.length).toBeGreaterThan(candidates[1].paragraphs[0]!.lines.length)
    expect(prepared.outline_requests.length).toBeGreaterThan(0)
    const outlines = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: prepared.outline_requests.map((requested) => {
      const outline = outlines.outline(requested.glyph_id)
      return outline.path.length ? { ...requested, ...outline, status: 'outlined' as const } : { ...requested, units_per_em: outline.units_per_em, status: 'empty' as const }
    }) })
    expect(completed.page_paint_output.status, JSON.stringify(completed.page_paint_output)).toBe('painted')
    expect(decodeNativeDocxPagePaintForRequestV1(completed.page_paint_output, request, request.outline_provider).ok).toBe(true)
    expect(input).toEqual(before)
    const tampered = structuredClone(request)
    tampered.pagination_request.column_shaped_lines![1].paragraphs[0]!.lines[0]!.fragments[0]!.face_id = 'font-face:forged'
    tampered.integrity.shaped_lines_sha256 = nativeDocxPagePaintShapedLinesSha256V1(tampered.pagination_request.shaped_lines, undefined, tampered.pagination_request.column_shaped_lines)
    expect(decodeNativeDocxPagePaintRequestV1(tampered).ok).toBe(false)
    const oldHash = structuredClone(request)
    oldHash.integrity.shaped_lines_sha256 = nativeDocxPagePaintShapedLinesSha256V1(oldHash.pagination_request.shaped_lines)
    expect(decodeNativeDocxPagePaintRequestV1(oldHash).ok).toBe(false)
  }, 15_000)

  it('refuses multiline paragraphs without keepLines and paragraphs too tall at either width', async () => {
    const missing = unequalInput()
    delete (missing.resolved_layout as NativeDocxResolvedLayoutInputV1).paragraphs[0]!.properties.keep_lines
    await expect(prepareNativeDocxPagePaintV1(missing)).rejects.toThrow('whole-paragraph candidates')
    const oversized = unequalInput()
    ;(oversized.document as NativeDocxDocumentV1).sections[0]!.page.margins.bottom_twips = 14200
    await expect(prepareNativeDocxPagePaintV1(oversized)).rejects.toThrow('whole-paragraph candidates')
  }, 15_000)
})

describe('whole footnote reservation compiler', () => {
  function reflowInput(prefix: number, following: number, height = 90000, members = 1): NativeDocxPagePaintPrepareInputV1 {
    const input = noteFixture()
    const document = input.document as NativeDocxDocumentV1
    const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    const reference = document.body.blocks[0]!
    const paragraphs = Array.from({ length: prefix + following }, (_, index) => {
      const block = structuredClone(reference)
      block.id = `paragraph:following:${index}`; block.paragraph!.id = block.id
      block.paragraph!.runs = [{ kind: 'text', id: `run:following:${index}`, anchor: block.paragraph!.runs[0]!.anchor, text: 'Ordinary whole body paragraph' }]
      resolved.paragraphs.push({ paragraph_id: block.id, applied_styles: [], properties: {}, paragraph_mark_properties: { font_family: 'DejaVu Sans', font_size_half_points: 20 } })
      resolved.runs.push({ run_id: block.paragraph!.runs[0]!.id, paragraph_id: block.id, applied_paragraph_styles: [], applied_character_styles: [], properties: { font_family: 'DejaVu Sans', font_size_half_points: 20 } })
      return block
    })
    document.body.blocks = [...paragraphs.slice(0, prefix), reference, ...paragraphs.slice(prefix)]
    document.sections[0]!.starts_at_block_id = document.body.blocks[0]!.id
    document.sections[0]!.page.margins.bottom_twips = 15840 - 1440 - height / 50
    const note = document.notes[1]!.blocks[0]!.paragraph!
    note.runs[1]!.text = 'The complete footnote remains with the reference and uses one kept paragraph with enough source text to wrap at the authored width. '.repeat(2)
    resolved.paragraphs.find((paragraph) => paragraph.paragraph_id === note.id)!.properties.keep_lines = true
    for (let index = 1; index < members; index += 1) {
      const paragraph = structuredClone(note)
      paragraph.id = `${note.id}:member:${index}`
      paragraph.runs = [{ ...structuredClone(note.runs[1]!), id: `run:${paragraph.id}`, text: 'Another authored kept paragraph.' }]
      document.notes[1]!.blocks.push({ kind: 'paragraph', id: paragraph.id, paragraph })
      resolved.paragraphs.push({ ...structuredClone(resolved.paragraphs.find((entry) => entry.paragraph_id === note.id)!), paragraph_id: paragraph.id })
      resolved.runs.push({ ...structuredClone(resolved.runs.find((entry) => entry.run_id === note.runs[1]!.id)!), run_id: paragraph.runs[0]!.id, paragraph_id: paragraph.id })
    }
    if (members > 1) for (const [index, block] of document.notes[1]!.blocks.entries()) resolved.paragraphs.find((entry) => entry.paragraph_id === block.id)!.properties.keep_next = index < members - 1
    rewriteInventory(input, (inventory) => { inventory.references[0]!.scope_ids = [...resolved.paragraphs.map((paragraph) => paragraph.paragraph_id), ...resolved.runs.map((run) => run.run_id)].sort() })
    return input
  }
  it('paints every authored footnote-chain paragraph with one label after body reflow', async () => {
    const input = reflowInput(5, 2, 90000, 2), before = structuredClone(input)
    const prepared = await prepareNativeDocxPagePaintV1(input)
    const layout = prepared.page_paint_request.paginated_layout
    expect(layout.status, JSON.stringify(layout.diagnostics)).toBe('paginated')
    const note = layout.pages[1]!.note_stories![1]!
    expect([...new Set(note.lines.map((line) => line.paragraph_id))]).toEqual((input.document as NativeDocxDocumentV1).notes[1]!.blocks.map((block) => block.id))
    expect(layout.pages[0]!.note_stories ?? []).toHaveLength(0)
    const outlines = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: prepared.outline_requests.map((request) => {
      const outline = outlines.outline(request.glyph_id)
      return outline.path.length ? { ...request, ...outline, status: 'outlined' as const } : { ...request, units_per_em: outline.units_per_em, status: 'empty' as const }
    }) })
    expect(completed.page_paint_output.status).toBe('painted')
    expect(decodeNativeDocxPagePaintForRequestV1(completed.page_paint_output, completed.page_paint_request, completed.page_paint_request.outline_provider).ok).toBe(true)
    const forged = structuredClone(completed.page_paint_request)
    forged.paginated_layout.pages[1]!.note_stories![1]!.lines.pop()
    expect(decodeNativeDocxPagePaintRequestV1(forged).ok).toBe(false)
    expect(input).toEqual(before)
  }, 15_000)
  it.each([[130000, [1, 2]], [70000, [1]]])('paints two globally numbered note chains with a shared page separator (%s)', async (height, firstPageNumbers) => {
    const input = reflowInput(1, 4, height as number, 2)
    const document = input.document as NativeDocxDocumentV1, resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    const first = document.notes[1]!
    first.blocks[0]!.paragraph!.runs[1]!.text = ' First note.'
    const second = structuredClone(first)
    second.id = 'story:footnote:second'; second.native_story_id = '2'
    for (const block of second.blocks) {
      const oldID = block.id; block.id += ':second'; block.paragraph!.id = block.id
      resolved.paragraphs.push({ ...structuredClone(resolved.paragraphs.find((entry) => entry.paragraph_id === oldID)!), paragraph_id: block.id })
      for (const run of block.paragraph!.runs) {
        const oldRunID = run.id; run.id += ':second'
        if (run.reference) run.reference.target_id = second.id
        resolved.runs.push({ ...structuredClone(resolved.runs.find((entry) => entry.run_id === oldRunID)!), run_id: run.id, paragraph_id: block.id })
      }
    }
    document.notes.push(second)
    const reference = structuredClone(document.body.blocks[1]!), oldParagraph = reference.id
    reference.id += ':second'; reference.paragraph!.id = reference.id
    resolved.paragraphs.push({ ...structuredClone(resolved.paragraphs.find((entry) => entry.paragraph_id === oldParagraph)!), paragraph_id: reference.id })
    for (const run of reference.paragraph!.runs) {
      const oldID = run.id; run.id += ':second'
      if (run.reference) run.reference.target_id = second.id
      resolved.runs.push({ ...structuredClone(resolved.runs.find((entry) => entry.run_id === oldID)!), run_id: run.id, paragraph_id: reference.id })
    }
    document.body.blocks.splice(2, 0, reference)
    rewriteInventory(input, (inventory) => { inventory.references[0]!.scope_ids = [...resolved.paragraphs.map((entry) => entry.paragraph_id), ...resolved.runs.map((entry) => entry.run_id)].sort() })
    const before = structuredClone(input), prepared = await prepareNativeDocxPagePaintV1(input)
    const layout = prepared.page_paint_request.paginated_layout
    expect(layout.status, JSON.stringify(layout.diagnostics)).toBe('paginated')
    expect(layout.pages[0]!.note_stories?.filter((note) => note.note_role === 'content').map((note) => note.number)).toEqual(firstPageNumbers)
    expect(layout.pages.flatMap((page) => page.note_stories ?? []).filter((note) => note.note_role === 'content').map((note) => note.number)).toEqual([1, 2])
    for (const page of layout.pages) if (page.note_stories?.length) expect(page.note_stories.filter((note) => note.note_role === 'separator')).toHaveLength(1)
    const outlines = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
    const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: prepared.outline_requests.map((request) => {
      const outline = outlines.outline(request.glyph_id)
      return outline.path.length ? { ...request, ...outline, status: 'outlined' as const } : { ...request, units_per_em: outline.units_per_em, status: 'empty' as const }
    }) })
    expect(completed.page_paint_output.status).toBe('painted')
    expect(decodeNativeDocxPagePaintForRequestV1(completed.page_paint_output, completed.page_paint_request, completed.page_paint_request.outline_provider).ok).toBe(true)
    expect(input).toEqual(before)
  }, 15_000)
  // Two complete real-outline/replay scenarios share this integration test.
  it('paints references and whole notes after later body or the reference itself moves', async () => {
    for (const [prefix, following, notePage] of [[1, 6, 0], [5, 1, 1]]) {
      const input = reflowInput(prefix!, following!)
      const before = structuredClone(input)
      const prepared = await prepareNativeDocxPagePaintV1(input)
      const layout = prepared.page_paint_request.paginated_layout
      expect(layout.status, JSON.stringify(layout.diagnostics)).toBe('paginated')
      expect(layout.pages.length).toBeGreaterThan(1)
      expect(layout.pages[notePage!]!.note_stories?.some((note) => note.note_role === 'content')).toBe(true)
      const outlines = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
      const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: prepared.outline_requests.map((request) => {
        const outline = outlines.outline(request.glyph_id)
        return outline.path.length ? { ...request, ...outline, status: 'outlined' as const } : { ...request, units_per_em: outline.units_per_em, status: 'empty' as const }
      }) })
      expect(completed.page_paint_output.status, JSON.stringify(completed.page_paint_output)).toBe('painted')
      expect(decodeNativeDocxPagePaintForRequestV1(completed.page_paint_output, completed.page_paint_request, completed.page_paint_request.outline_provider).ok).toBe(true)
      expect(input).toEqual(before)
    }
  }, 15_000)
  it('keeps an oversized reference-note pair refused with no pages', async () => {
    const prepared = await prepareNativeDocxPagePaintV1(reflowInput(0, 1, 30000))
    expect(prepared.page_paint_request.paginated_layout).toEqual(expect.objectContaining({ status: 'refused', pages: [] }))
  })
})

describe('source-anchored textbox page composition',()=>{
 function textboxFixture(){
  const input=fixture(),document=input.document as NativeDocxDocumentV1,p=document.body.blocks[0]!.paragraph!
  p.anchor.end_byte=2000
  p.runs[0]!.anchor=anchor(p.anchor.path+'/w:r[2]',1500,1800)
  const root=p.anchor.path+'/w:r[1]/w:drawing[1]/wp:anchor[1]',a=anchor(root+'/a:graphic[1]/a:graphicData[1]',200,900)
  document.unsupported=[{id:'shape:1',code:'PICTURE_GRAPHIC_REQUIRED',capability:'drawings',scope_id:p.id,anchor:a,preservation:'refuse-mutation',message:'Drawing preserved'}]
  const evidence:import('./nativeTextboxGeometryPreviewV1.js').NativeDocxTextboxGeometryEvidenceV1={items:[{owner:{package_sha256:HASH,part_sha256:HASH,paragraph_id:p.id,diagnostic_id:'shape:1',anchor:a,kind:'drawingml',status:'supported',paragraphs:['Page rectangle'],reason:''},geometry:{width_emu:2743200,height_emu:914400,insets_emu:[91440,91440,91440,91440],fill_rgb:'FFF2CC',line_rgb:'204060',line_width_emu:12700,font_family:'DejaVu Sans',font_size_half_points:24,text_rgb:'102030'},page_anchor:{policy:'page-offset-no-wrap-v1',source_anchor:anchor(root,110,1000),horizontal_anchor:anchor(root+'/wp:positionH[1]/wp:posOffset[1]',120,135),vertical_anchor:anchor(root+'/wp:positionV[1]/wp:posOffset[1]',140,155),x_emu:914400,y_emu:1828800}}],omitted_count:0}
  return {input,document,evidence}
 }
 function multipleTextboxFixture(separatePages=false){
  const f=textboxFixture(),p=f.document.body.blocks[0]!.paragraph!,second=structuredClone(f.evidence.items[0]!)
  f.document.body.anchor.end_byte=10000;f.document.sections[0]!.anchor=anchor(f.document.sections[0]!.anchor.path,9500,9600)
  const shift=separatePages?4000:2000,root=separatePages?p.anchor.path.replace('/w:p[1]','/w:p[2]'):p.anchor.path
  second.owner.diagnostic_id='shape:2';second.owner.paragraphs=['Second rectangle'];second.geometry!.fill_rgb='DDEEFF'
  for(const a of [second.owner.anchor,second.page_anchor!.source_anchor,second.page_anchor!.horizontal_anchor,second.page_anchor!.vertical_anchor]){a.path=a.path.replace(p.anchor.path,root);if(!separatePages)a.path=a.path.replace('/w:r[1]','/w:r[2]');a.start_byte+=shift;a.end_byte+=shift}
  second.page_anchor!.x_emu=3657600;second.page_anchor!.y_emu=2743200
  if(separatePages){
   const next=structuredClone(p);next.id='paragraph:second';next.anchor={...next.anchor,path:root,start_byte:p.anchor.start_byte+shift,end_byte:p.anchor.end_byte+shift}
   next.runs[0]!.id='run:second';next.runs[0]!.anchor={...next.runs[0]!.anchor,path:root+'/w:r[2]',start_byte:1500+shift,end_byte:1800+shift}
   f.document.body.blocks.push({kind:'paragraph',id:next.id,paragraph:next});second.owner.paragraph_id=next.id
   f.document.sections[0]!.page.margins.bottom_twips=14000
   const resolved=f.input.resolved_layout as NativeDocxResolvedLayoutInputV1
   resolved.paragraphs.push({...structuredClone(resolved.paragraphs[0]!),paragraph_id:next.id,properties:{}})
   resolved.runs.push({...structuredClone(resolved.runs[0]!),run_id:'run:second',paragraph_id:next.id})
   rewriteInventory(f.input,i=>{i.references[0]!.scope_ids.push(next.id,'run:second');i.references[0]!.scope_ids.sort()})
  } else {p.anchor.end_byte=4000;p.runs[0]!.anchor={...p.runs[0]!.anchor,path:p.anchor.path+'/w:r[3]',start_byte:3500,end_byte:3800}}
  f.document.unsupported.push({...structuredClone(f.document.unsupported[0]!),id:'shape:2',scope_id:second.owner.paragraph_id,anchor:second.owner.anchor})
  f.evidence.items.push(second)
  return f
 }
 const provider={providerId:'injoffice.sfnt-outline',providerRevision:'sfnt-v1',getGlyphOutline(request:import('./nativePagePaintV1.js').NativeDocxGlyphOutlineRequestV1){
  const outline=createHarfBuzzOutlineProviderV1({bytes:FONT_BYTES,contentDigest:FONT_DIGEST}).outline(request.glyph_id)
  return outline.path.length?{status:'outlined' as const,...request,...outline}:{status:'empty' as const,...request,units_per_em:outline.units_per_em}
 }}
 it('composes actual-font body and textbox paint without removing original source diagnostics',async()=>{
  const {renderNativeDocxTextboxPagePreviewV1:render}=await import('./nativeTextboxPageCompilerV1.js')
  const {decodeNativeDocxTextboxPagePreviewV1:decode}=await import('./nativeTextboxPagePreviewV1.js')
  const {input,document,evidence}=textboxFixture(),before=structuredClone({input,evidence}),result=await render(input,evidence,FONT_BYTES,provider)
  expect(result.body_paint.pages).toHaveLength(1);expect(result.textbox).toMatchObject({page_id:result.body_paint.pages[0]!.id,x_millipoints:72000,y_millipoints:144000,paint:{status:'supported'}})
  expect(result.textbox.paint.paths.length).toBeGreaterThan(0);expect(result.source_diagnostics).toEqual(document.unsupported)
  expect(decode(document,evidence,result,FONT_DIGEST)).toEqual(result)
  expect(await render(input,evidence,FONT_BYTES,provider)).toEqual(result);expect({input,evidence}).toEqual(before)
  for(const mutate of [(v:typeof result)=>{v.textbox.page_id='page:other'},(v:typeof result)=>{v.textbox.x_millipoints++},(v:typeof result)=>{v.source_diagnostics=[]},(v:typeof result)=>{v.source_sha256=RELATIONSHIPS_HASH},(v:typeof result)=>{v.textbox.paint.input_sha256=RELATIONSHIPS_HASH},(v:typeof result)=>{v.body_paint.provenance.revision='rev:other'},(v:typeof result)=>{v.textbox.paint.paths[0]='M 0 0 L 1 1'}]){
   const forged=structuredClone(result);mutate(forged);expect(()=>decode(document,evidence,forged,FONT_DIGEST)).toThrow()
  }
  expect(()=>decode(document,evidence,result,RELATIONSHIPS_HASH)).toThrow()
 })
 it('binds asynchronous output to the source snapshot even when the caller edits during outlining',async()=>{
  const {renderNativeDocxTextboxPagePreviewV1:render}=await import('./nativeTextboxPageCompilerV1.js')
  const {decodeNativeDocxTextboxPagePreviewV1:decode}=await import('./nativeTextboxPagePreviewV1.js')
  const {input,document,evidence}=textboxFixture(),before=structuredClone({document,evidence})
  const result=await render(input,evidence,FONT_BYTES,{...provider,getGlyphOutline(request){
   document.body.blocks[0]!.paragraph!.runs[0]!.text='Changed during outlining'
   evidence.items[0]!.page_anchor!.x_emu+=127
   return provider.getGlyphOutline(request)
  }})
  expect(decode(before.document,before.evidence,result,FONT_DIGEST)).toEqual(result)
  expect(()=>decode(document,evidence,result,FONT_DIGEST)).toThrow()
 })
 it('follows a paragraph pushed to a later page by body flow',async()=>{
  const {renderNativeDocxTextboxPagePreviewV1:render}=await import('./nativeTextboxPageCompilerV1.js')
  const {input,document,evidence}=textboxFixture(),p=document.body.blocks[0]!.paragraph!,r=input.resolved_layout as NativeDocxResolvedLayoutInputV1
  const preceding=structuredClone(p);preceding.id='paragraph:0';preceding.anchor=anchor('/w:document[1]/w:body[1]/w:p[1]',10,90);preceding.runs[0]!.id='run:0';preceding.runs[0]!.anchor=anchor(preceding.anchor.path+'/w:r[1]',20,80)
  document.body.blocks.unshift({kind:'paragraph',id:preceding.id,paragraph:preceding});document.sections[0]!.starts_at_block_id=preceding.id
  document.sections[0]!.page.margins.bottom_twips=14000
  for(const a of [p.anchor,p.runs[0]!.anchor,document.unsupported[0]!.anchor!,evidence.items[0]!.owner.anchor,evidence.items[0]!.page_anchor!.source_anchor,evidence.items[0]!.page_anchor!.horizontal_anchor,evidence.items[0]!.page_anchor!.vertical_anchor])a.path=a.path.replace('/w:p[1]','/w:p[2]')
  r.paragraphs.unshift({...structuredClone(r.paragraphs[0]!),paragraph_id:preceding.id,properties:{}});r.runs.unshift({...structuredClone(r.runs[0]!),run_id:'run:0',paragraph_id:preceding.id})
  rewriteInventory(input,i=>{i.references[0]!.scope_ids.push('paragraph:0','run:0');i.references[0]!.scope_ids.sort()})
  const result=await render(input,evidence,FONT_BYTES,provider)
  expect(result.body_paint.pages).toHaveLength(2);expect(result.textbox.page_id).toBe(result.body_paint.pages[1]!.id)
 })
 it('refuses uncertain anchors, off-page strokes, incomplete evidence and unsupported body atomically',async()=>{
  const {renderNativeDocxTextboxPagePreviewV1:render}=await import('./nativeTextboxPageCompilerV1.js')
  const mutations:Array<(f:ReturnType<typeof textboxFixture>)=>void>=[
   f=>{f.document.body.blocks[0]!.paragraph!.runs[0]!.anchor.start_byte=105},
   f=>{f.evidence.items[0]!.page_anchor!.x_emu=0},
   f=>{f.evidence.items[0]!.page_anchor!.y_emu=127000000},
   f=>{delete f.evidence.items[0]!.page_anchor},
   f=>{f.evidence.omitted_count=1},
   f=>{f.evidence.items[0]!.geometry!.font_family='Unknown family'},
   f=>{f.document.unsupported.push({...f.document.unsupported[0]!,id:'other:1',code:'OTHER_UNSUPPORTED',capability:'unknown'})},
  ]
  for(const mutate of mutations){const f=textboxFixture();mutate(f);const before=structuredClone(f);await expect(render(f.input,f.evidence,FONT_BYTES,provider)).rejects.toThrow();expect(f).toEqual(before)}
 },20000)
 it('composes every rectangle with its own font and rejects incomplete or reordered paint',async()=>{
  const {renderNativeDocxTextboxPagesPreviewV2:render}=await import('./nativeTextboxPagesCompilerV2.js')
  const {decodeNativeDocxTextboxPagesPreviewV2:decode}=await import('./nativeTextboxPagesPreviewV2.js')
  const {nativeTextboxFontDigestV1:digest}=await import('./nativeTextboxGeometryPreviewV1.js')
  const f=multipleTextboxFixture(),mono=new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSansMono.ttf')))
  f.evidence.items[1]!.geometry!.font_family='DejaVu Sans Mono'
  const before=structuredClone(f),hashes=[FONT_DIGEST,digest(mono)],result=await render(f.input,f.evidence,[FONT_BYTES,mono],provider)
  expect(result.version).toBe(2);expect(result.textboxes.map(b=>b.paint.diagnostic_id)).toEqual(['shape:1','shape:2'])
  expect(result.textboxes.map(b=>[b.x_millipoints,b.y_millipoints])).toEqual([[72000,144000],[288000,216000]])
  expect(result.textboxes.map(b=>b.paint.font_sha256)).toEqual(hashes)
  expect(result.source_diagnostics).toEqual(f.document.unsupported);expect(f).toEqual(before)
  expect(decode(f.document,f.evidence,result,hashes)).toEqual(result)
  for(const mutate of [(v:typeof result)=>{v.textboxes.pop()},(v:typeof result)=>{v.textboxes.reverse()},(v:typeof result)=>{v.textboxes[1]=structuredClone(v.textboxes[0]!)},(v:typeof result)=>{v.textboxes[1]!.x_millipoints++},(v:typeof result)=>{v.source_diagnostics=[]}]){
   const forged=structuredClone(result);mutate(forged);expect(()=>decode(f.document,f.evidence,forged,hashes)).toThrow()
  }
  expect(()=>decode(f.document,f.evidence,result,[FONT_DIGEST,FONT_DIGEST])).toThrow()
 })
 it('uses each source anchor page and refuses reordered evidence',async()=>{
  const {renderNativeDocxTextboxPagesPreviewV2:render}=await import('./nativeTextboxPagesCompilerV2.js')
  const f=multipleTextboxFixture(true)
  const result=await render(f.input,f.evidence,[FONT_BYTES,FONT_BYTES],provider)
  expect(result.body_paint.pages).toHaveLength(2)
  expect(result.textboxes.map(b=>b.page_id)).toEqual(result.body_paint.pages.map(p=>p.id))
  expect(result.textboxes.map(b=>b.textbox_index)).toEqual([0,1])
  f.evidence.items.reverse();await expect(render(f.input,f.evidence,[FONT_BYTES,FONT_BYTES],provider)).rejects.toThrow()
 })
 it('snapshots every textbox and font before asynchronous body outlining',async()=>{
  const {renderNativeDocxTextboxPagesPreviewV2:render}=await import('./nativeTextboxPagesCompilerV2.js')
  const {decodeNativeDocxTextboxPagesPreviewV2:decode}=await import('./nativeTextboxPagesPreviewV2.js')
  const f=multipleTextboxFixture(),before=structuredClone(f),fonts=[FONT_BYTES.slice(),FONT_BYTES.slice()]
  const result=await render(f.input,f.evidence,fonts,{...provider,getGlyphOutline(request){f.evidence.items[1]!.page_anchor!.x_emu+=127;f.document.body.blocks[0]!.paragraph!.runs[0]!.text='Changed';fonts[1]!.fill(0);return provider.getGlyphOutline(request)}})
  expect(decode(before.document,before.evidence,result,[FONT_DIGEST,FONT_DIGEST])).toEqual(result)
  expect(()=>decode(f.document,f.evidence,result,[FONT_DIGEST,FONT_DIGEST])).toThrow()
 })
 it('refuses missing fonts, duplicate evidence, unknown drawings and overlapping anchors atomically',async()=>{
  const {renderNativeDocxTextboxPagesPreviewV2:render}=await import('./nativeTextboxPagesCompilerV2.js')
  const {renderNativeDocxTextboxPagePreviewV1:legacy}=await import('./nativeTextboxPageCompilerV1.js')
  for(const mutate of [(f:ReturnType<typeof multipleTextboxFixture>)=>{f.evidence.items[1]=structuredClone(f.evidence.items[0]!)},(f:ReturnType<typeof multipleTextboxFixture>)=>{f.evidence.omitted_count=1},(f:ReturnType<typeof multipleTextboxFixture>)=>{f.document.body.blocks[0]!.paragraph!.runs[0]!.anchor.start_byte=1500},(f:ReturnType<typeof multipleTextboxFixture>)=>{f.document.unsupported.push({...f.document.unsupported[0]!,id:'other-drawing'})}]){
   const f=multipleTextboxFixture();mutate(f);const before=structuredClone({document:f.document,evidence:f.evidence});await expect(render(f.input,f.evidence,[FONT_BYTES,FONT_BYTES],provider)).rejects.toThrow();expect({document:f.document,evidence:f.evidence}).toEqual(before)
  }
  const f=multipleTextboxFixture()
  await expect(render(f.input,f.evidence,[FONT_BYTES],provider)).rejects.toThrow()
  await expect(legacy(f.input,f.evidence,FONT_BYTES,provider)).rejects.toThrow()
 },15000)

 it('resolves source margins, columns, paragraphs, lines and character offsets',async()=>{
  const {renderNativeDocxTextboxPagesPreviewV2:render}=await import('./nativeTextboxPagesCompilerV2.js')
  const {decodeNativeDocxTextboxPagesPreviewV2:decode}=await import('./nativeTextboxPagesPreviewV2.js')
  for(const [h,v] of [['margin','margin'],['column','paragraph'],['character','line']] as const){
   const f=multipleTextboxFixture(true)
   for(const d of f.document.unsupported)d.code='FLOATING_DRAWING_SEMANTICS_PRESERVED'
   for(const item of f.evidence.items)item.page_anchor={...item.page_anchor!,policy:'relative-position-no-wrap-v2',horizontal_relative:h,vertical_relative:v,x_emu:-457200,y_emu:914400}
   const output=await render(f.input,f.evidence,[FONT_BYTES,FONT_BYTES],provider)
   for(const [i,box] of output.textboxes.entries()){
    const page=output.body_paint.pages[i]!,line=page.lines.find(l=>l.region==='body')!
    expect(box.x_millipoints).toBe((h==='character'?line.x_millipoints:72000)-36000)
    expect(box.y_millipoints).toBe((v==='margin'?72000:line.y_millipoints)+72000)
   }
   expect(decode(f.document,f.evidence,output,[FONT_DIGEST,FONT_DIGEST])).toEqual(output)
   const bad=structuredClone(output);bad.textboxes[1]!.y_millipoints+=72000
   expect(()=>decode(f.document,f.evidence,bad,[FONT_DIGEST,FONT_DIGEST])).toThrow()
  }
 },20000)
 it('aligns against source page and margins and clips page-edge strokes without shifting geometry',async()=>{
  const {renderNativeDocxTextboxPagesPreviewV2:render}=await import('./nativeTextboxPagesCompilerV2.js')
  const {renderNativeDocxTextboxPagePreviewV1:legacy}=await import('./nativeTextboxPageCompilerV1.js')
  for(const [base,h,v,x,y] of [['page','left','top',0,0],['page','right','bottom',396000,720000],['margin','center','center',198000,360000]] as const){
   const f=textboxFixture(),p=f.evidence.items[0]!.page_anchor!
   f.evidence.items[0]!.page_anchor={...p,policy:'relative-position-no-wrap-v2',horizontal_relative:base,vertical_relative:base,horizontal_align:h,vertical_align:v,x_emu:0,y_emu:0,horizontal_anchor:{...p.horizontal_anchor,path:p.horizontal_anchor.path.replace('posOffset','align')},vertical_anchor:{...p.vertical_anchor,path:p.vertical_anchor.path.replace('posOffset','align')}}
   const output=await render(f.input,f.evidence,[FONT_BYTES],provider)
   expect(output.textboxes[0]).toMatchObject({x_millipoints:x,y_millipoints:y})
   await expect(legacy(f.input,f.evidence,FONT_BYTES,provider)).rejects.toThrow()
  }
 },20000)
 it('refuses invalid relative axes, ambiguous alignments and off-page signed offsets',async()=>{
  const {renderNativeDocxTextboxPagesPreviewV2:render}=await import('./nativeTextboxPagesCompilerV2.js')
  for(const change of [{horizontal_relative:['page']},{vertical_align:['top']},{horizontal_relative:'paragraph'},{vertical_relative:'column'},{horizontal_align:'center'},{x_emu:-914527},{y_emu:-1828800},{horizontal_relative:'character',horizontal_align:'right',x_emu:0}]){
   const f=textboxFixture(),p={...f.evidence.items[0]!.page_anchor!,policy:'relative-position-no-wrap-v2',horizontal_relative:'margin',vertical_relative:'paragraph',...change}
   f.evidence.items[0]!.page_anchor=p as import('./nativeTextboxPageAnchorV1.js').NativeTextboxPositionAnchor
   await expect(render(f.input,f.evidence,[FONT_BYTES],provider)).rejects.toThrow()
  }
 },20000)

 it('resolves parity margin regions and inside/outside alignment using physical page order',async()=>{
  const {renderNativeDocxTextboxPagesPreviewV2:render}=await import('./nativeTextboxPagesCompilerV2.js')
  const {resolveTextboxPosition:resolve}=await import('./nativeTextboxPositionV2.js')
  const f=textboxFixture(),output=await render(f.input,f.evidence,[FONT_BYTES],provider)
  const page=output.body_paint.pages[0]!,item=f.evidence.items[0]!,paint={...output.textboxes[0]!.paint,width_millipoints:36000,height_millipoints:18000}
  Object.assign(f.document.sections[0]!.page.margins,{left_twips:1800,right_twips:1080,top_twips:720,bottom_twips:2160})
  const original=item.page_anchor!
  for(const ordinal of [0,1])for(const outside of [false,true]){
   const base=outside?'outsideMargin':'insideMargin',right=(ordinal===1)!==outside
   item.page_anchor={...original,policy:'relative-position-no-wrap-v2',horizontal_relative:base,vertical_relative:base,horizontal_align:'center',vertical_align:'center',x_emu:0,y_emu:0}
   expect(resolve(f.document,item,{...page,ordinal},paint)).toEqual({x:right?567000:27000,y:right?729000:9000})
   item.page_anchor.horizontal_relative='page';item.page_anchor.vertical_relative='page';item.page_anchor.horizontal_align=outside?'outside':'inside';item.page_anchor.vertical_align=outside?'outside':'inside'
   expect(resolve(f.document,item,{...page,ordinal},paint)).toEqual({x:right?576000:0,y:right?774000:0})
  }
  for(const [h,v,x,y] of [['leftMargin','topMargin',27000,9000],['rightMargin','bottomMargin',567000,729000]] as const){
   item.page_anchor={...original,policy:'relative-position-no-wrap-v2',horizontal_relative:h,vertical_relative:v,horizontal_align:'center',vertical_align:'center',x_emu:0,y_emu:0}
   expect(resolve(f.document,item,page,paint)).toEqual({x,y})
  }
 })
 it('binds parity alignment and coordinates to complete source evidence',async()=>{
  const {renderNativeDocxTextboxPagesPreviewV2:render}=await import('./nativeTextboxPagesCompilerV2.js')
  const {decodeNativeDocxTextboxPagesPreviewV2:decode}=await import('./nativeTextboxPagesPreviewV2.js')
  const f=multipleTextboxFixture(true)
  for(const item of f.evidence.items){const p=item.page_anchor!;item.page_anchor={...p,policy:'relative-position-no-wrap-v2',horizontal_relative:'page',vertical_relative:'page',horizontal_align:'inside',vertical_align:'outside',x_emu:0,y_emu:0,horizontal_anchor:{...p.horizontal_anchor,path:p.horizontal_anchor.path.replace('posOffset','align')},vertical_anchor:{...p.vertical_anchor,path:p.vertical_anchor.path.replace('posOffset','align')}}}
  const result=await render(f.input,f.evidence,[FONT_BYTES,FONT_BYTES],provider)
  expect(result.textboxes.map(b=>[b.x_millipoints,b.y_millipoints])).toEqual([[0,720000],[396000,0]])
  expect(decode(f.document,f.evidence,result,[FONT_DIGEST,FONT_DIGEST])).toEqual(result)
  const bad=structuredClone(result);bad.textboxes[1]!.x_millipoints=0
  expect(()=>decode(f.document,f.evidence,bad,[FONT_DIGEST,FONT_DIGEST])).toThrow()
 })

 function lateTextboxFixture(text='A '){
  const f=textboxFixture(),p=f.document.body.blocks[0]!.paragraph!,item=f.evidence.items[0]!
  f.document.body.anchor.end_byte=10000;f.document.sections[0]!.anchor=anchor(f.document.sections[0]!.anchor.path,9500,9600)
  p.anchor.end_byte=5000;p.runs[0]!.anchor.path=p.anchor.path+'/w:r[1]';p.runs[0]!.text=text
  for(const a of [item.owner.anchor,item.page_anchor!.source_anchor,item.page_anchor!.horizontal_anchor,item.page_anchor!.vertical_anchor]){a.start_byte+=2000;a.end_byte+=2000;a.path=a.path.replace('/w:r[1]','/w:r[2]')}
  item.page_anchor={...item.page_anchor!,policy:'relative-position-no-wrap-v2',horizontal_relative:'character',vertical_relative:'line',x_emu:0,y_emu:0}
  return f
 }
 it('places anchors after shaped text advances including trailing spaces and wrapped pages',async()=>{
  const {renderNativeDocxTextboxPagesPreviewV2:render}=await import('./nativeTextboxPagesCompilerV2.js')
  const {decodeNativeDocxTextboxPagesPreviewV2:decode}=await import('./nativeTextboxPagesPreviewV2.js')
  for(const text of ['A ','A '.repeat(160)]){
   const f=lateTextboxFixture(text)
   if(text.length>10)f.document.sections[0]!.page.margins.bottom_twips=13200
   const result=await render(f.input,f.evidence,[FONT_BYTES],provider)
   const last=result.body_paint.pages.at(-1)!,line=last.lines.filter(l=>l.region==='body').at(-1)!,box=result.textboxes[0]!
   expect(box.page_id).toBe(last.id);expect(box.x_millipoints).toBe(line.x_millipoints+line.width_millipoints);expect(box.y_millipoints).toBe(line.y_millipoints)
   expect(result.anchor_request).toBeDefined();expect(decode(f.document,f.evidence,result,[FONT_DIGEST])).toEqual(result)
   const bad=structuredClone(result);bad.anchor_request!.pagination_request.shaped_lines.paragraphs[0]!.lines[0]!.fragments[0]!.advance_inline_millipoints++
   expect(()=>decode(f.document,f.evidence,bad,[FONT_DIGEST])).toThrow()
   delete bad.anchor_request;expect(()=>decode(f.document,f.evidence,bad,[FONT_DIGEST])).toThrow()
   const sourceForgery=structuredClone(result),request=sourceForgery.anchor_request!
   request.pagination_request.shaped_lines.paragraphs[0]!.lines[0]!.fragments[0]!.text='Z'
   const {nativeDocxPagePaintShapedLinesSha256V1:hashLines}=await import('./nativePagePaintV1.js')
   request.integrity.shaped_lines_sha256=hashLines(request.pagination_request.shaped_lines,request.page_field_variants,request.pagination_request.column_shaped_lines)
   sourceForgery.body_paint.provenance.shaped_lines.sha256=request.integrity.shaped_lines_sha256
   expect(()=>decode(f.document,f.evidence,sourceForgery,[FONT_DIGEST])).toThrow('Textbox anchor fragment text does not match source')
  }
 },15000)

 it('locates anchors after tabs and hard breaks, before following text',async()=>{
  const {renderNativeDocxTextboxPagesPreviewV2:render}=await import('./nativeTextboxPagesCompilerV2.js')
  for(const control of ['tab','line-break'] as const){
   const f=lateTextboxFixture(),p=f.document.body.blocks[0]!.paragraph!,resolved=f.input.resolved_layout as NativeDocxResolvedLayoutInputV1
   p.runs.push({kind:'control',id:'run:control',anchor:anchor(p.anchor.path+'/w:r[1]/w:br[1]',1810,1850),control},{kind:'text',id:'run:after',anchor:anchor(p.anchor.path+'/w:r[3]/w:t[1]',3500,3800),text:'After'})
   for(const run_id of ['run:control','run:after'])resolved.runs.push({...structuredClone(resolved.runs[0]!),run_id})
   rewriteInventory(f.input,i=>{i.references[0]!.scope_ids.push('run:control','run:after');i.references[0]!.scope_ids.sort()})
   const result=await render(f.input,f.evidence,[FONT_BYTES],provider),lines=result.body_paint.pages[0]!.lines.filter(l=>l.region==='body'),line=lines[control==='tab'?0:1]!
   expect(result.textboxes[0]).toMatchObject({x_millipoints:control==='tab'?108000:72000,y_millipoints:line.y_millipoints})
  }
 })
 it('uses the logical end of a right-to-left run as the character origin',async()=>{
  const {renderNativeDocxTextboxPagesPreviewV2:render}=await import('./nativeTextboxPagesCompilerV2.js')
  const f=lateTextboxFixture('אבג'),resolved=f.input.resolved_layout as NativeDocxResolvedLayoutInputV1
  resolved.paragraphs[0]!.properties={bidi:true,alignment:'start'}
  f.evidence.items[0]!.owner.paragraphs=['Box'];Object.assign(f.evidence.items[0]!.geometry!,{width_emu:609600,height_emu:457200})
  const result=await render(f.input,f.evidence,[FONT_BYTES],provider),line=result.body_paint.pages[0]!.lines[0]!
  expect(result.textboxes[0]!.x_millipoints).toBe(line.x_millipoints)
 })

 it('binds stacking to source evidence and refuses changed or omitted layers',async()=>{
  const {renderNativeDocxTextboxPagesPreviewV2:render}=await import('./nativeTextboxPagesCompilerV2.js')
  const {decodeNativeDocxTextboxPagesPreviewV2:decode}=await import('./nativeTextboxPagesPreviewV2.js')
  const f=multipleTextboxFixture()
  for(const [i,item] of f.evidence.items.entries())item.page_anchor={...item.page_anchor!,policy:'relative-position-no-wrap-v2',horizontal_relative:'page',vertical_relative:'page',stacking:{behind_doc:i===1,relative_height:i===0?4294967295:1}}
  const result=await render(f.input,f.evidence,[FONT_BYTES,FONT_BYTES],provider)
  expect(result.textboxes.map(t=>t.stacking)).toEqual([{behind_doc:false,relative_height:4294967295},{behind_doc:true,relative_height:1}])
  for(const mutate of [(v:typeof result)=>{delete v.textboxes[1]!.stacking},(v:typeof result)=>{v.textboxes[0]!.stacking!.relative_height=1},(v:typeof result)=>{v.textboxes[1]!.stacking!.behind_doc=false}]){
   const bad=structuredClone(result);mutate(bad);expect(()=>decode(f.document,f.evidence,bad,[FONT_DIGEST,FONT_DIGEST])).toThrow()
  }
  for(const invalid of [{behind_doc:1,relative_height:1},{behind_doc:false,relative_height:-1},{behind_doc:false,relative_height:4294967296},{behind_doc:false,relative_height:1.5}]){
   const bad=structuredClone(f.evidence);(bad.items[0]!.page_anchor as unknown as {stacking:unknown}).stacking=invalid
   await expect(render(f.input,bad,[FONT_BYTES,FONT_BYTES],provider)).rejects.toThrow()
  }
 },15000)


  it('shapes an indented table-cell paragraph only in the approximate preview and declares the cell line-box policy', async () => {
    const input = tableFixture(), document = input.document as NativeDocxDocumentV1, resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1, settings = input.pagination_settings as NativeDocxPaginationSettingsV1
    const paragraph = document.body.blocks[0]!.table!.rows[0]!.cells[0]!.paragraphs[0]!
    resolved.paragraphs.find((entry) => entry.paragraph_id === paragraph.id)!.properties.indent_left_twips = 360
    const strict = await prepareNativeDocxPagePaintV1(structuredClone(input))
    expect(strict.page_paint_request.paginated_layout.status).toBe('refused')
    expect(strict.page_paint_request.paginated_layout.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'line-geometry-invalid' })]))
    settings.profile = 'unsupported'; delete settings.compatibility_mode
    settings.diagnostics = [{ code: 'COMPATIBILITY_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: SETTINGS_PART, path: '/w:settings[1]/w:compat[1]', preservation: 'preserve-verbatim', message: 'Legacy12' }]
    const eligibility = { protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: settings.document_id, revision: settings.revision, package_sha256: HASH, settings_sha256: settings.settings_sha256, status: 'eligible', legacy_compatibility_mode: 12, reasons: ['Legacy12'] }
    const outlines = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
    const provider = { providerId: input.outline_provider.provider_id, providerRevision: input.outline_provider.provider_revision, getGlyphOutline(request: import('./nativePagePaintV1.js').NativeDocxGlyphOutlineRequestV1) { const o = outlines.outline(request.glyph_id); return o.path.length ? { status: 'outlined' as const, ...request, ...o } : { status: 'empty' as const, ...request, units_per_em: o.units_per_em } } }
    const result = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, provider)
    expect(result.status).toBe('painted')
    expect(result.reasons).toContain(DOCX_APPROXIMATE_INDENTED_CELL_LINE_WARNING)
    const line = result.pages[0]!.lines.find((entry) => entry.paragraph_id === paragraph.id)!
    expect(line.x_millipoints).toBe(result.pages[0]!.body_box.x_millipoints + 100 * 50 + 360 * 50)
    expect(decodeNativeDocxApproximatePagePreviewV1(result).ok).toBe(true)
    const unindented = tableFixture(); const unindentedSettings = unindented.pagination_settings as NativeDocxPaginationSettingsV1
    unindentedSettings.profile = 'unsupported'; delete unindentedSettings.compatibility_mode; unindentedSettings.diagnostics = settings.diagnostics
    expect((await renderNativeDocxApproximatePagePreviewV1(unindented, eligibility, provider)).reasons).not.toContain(DOCX_APPROXIMATE_INDENTED_CELL_LINE_WARNING)
  }, 20000)

  it('projects evidenced Latin fallback faces only in the approximate preview, only onto attested references, and declares them', async () => {
    const legacy = (input: NativeDocxPagePaintPrepareInputV1) => {
      const settings = input.pagination_settings as NativeDocxPaginationSettingsV1
      settings.profile = 'unsupported'; delete settings.compatibility_mode
      settings.diagnostics = [{ code: 'COMPATIBILITY_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: SETTINGS_PART, path: '/w:settings[1]/w:compat[1]', preservation: 'preserve-verbatim', message: 'Legacy12' }]
      return settings
    }
    const input = fixture(), document = input.document as NativeDocxDocumentV1, resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    const run = document.body.blocks[0]!.paragraph!.runs[0]!, resolvedRun = resolved.runs.find((entry) => entry.run_id === run.id)!
    const family = resolvedRun.properties.font_family!
    // Strict resolution left this run without a face; the inventory covers only the remaining scopes.
    delete resolvedRun.properties.font_family
    rewriteInventory(input, (inventory) => { for (const reference of inventory.references) reference.scope_ids = reference.scope_ids.filter((id) => id !== run.id) })
    const fact = { scope_kind: 'run' as const, scope_id: run.id, part_name: run.anchor.part_name, path: run.anchor.path, font_family: family, package_sha256: HASH }
    const strict = await prepareNativeDocxPagePaintV1(structuredClone(input))
    expect(strict.page_paint_request.paginated_layout.status).toBe('refused')
    const settings = legacy(input)
    const eligibility = { protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: settings.document_id, revision: settings.revision, package_sha256: HASH, settings_sha256: settings.settings_sha256, status: 'eligible', legacy_compatibility_mode: 12, reasons: ['Legacy12'], latin_font_fallbacks: [fact] }
    const outlines = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
    const provider = { providerId: input.outline_provider.provider_id, providerRevision: input.outline_provider.provider_revision, getGlyphOutline(request: import('./nativePagePaintV1.js').NativeDocxGlyphOutlineRequestV1) { const o = outlines.outline(request.glyph_id); return o.path.length ? { status: 'outlined' as const, ...request, ...o } : { status: 'empty' as const, ...request, units_per_em: o.units_per_em } } }
    const before = structuredClone(input)
    const { latin_font_fallbacks: _facts, ...withoutFacts } = eligibility
    const withoutEvidence = await renderNativeDocxApproximatePagePreviewV1(input, withoutFacts, provider)
    expect(withoutEvidence.status).toBe('refused')
    const result = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, provider)
    expect(result.status).toBe('painted')
    expect(result.approximated_font_faces).toEqual([fact])
    expect(result.source_latin_font_fallbacks).toEqual([fact])
    expect(result.reasons).toContain(DOCX_LATIN_FONT_FALLBACK_WARNING)
    expect(result.pages[0]!.lines.some((line) => line.paragraph_id === document.body.blocks[0]!.paragraph!.id)).toBe(true)
    expect(input).toEqual(before)
    expect(decodeNativeDocxApproximatePagePreviewV1(result).ok).toBe(true)
    for (const mutation of [{ approximated_font_faces: [{ ...fact, font_family: 'Other' }] }, { reasons: result.reasons.filter((r) => r !== DOCX_LATIN_FONT_FALLBACK_WARNING) }, { source_latin_font_fallbacks: [] }, { approximated_font_faces: [] }]) expect(decodeNativeDocxApproximatePagePreviewV1({ ...result, ...mutation }).ok).toBe(false)
    // A face the strict inventory never attests is skipped, so the run stays unshaped and the page refuses as before.
    const unattested = await renderNativeDocxApproximatePagePreviewV1(input, { ...eligibility, latin_font_fallbacks: [{ ...fact, font_family: 'Unattested Face' }] }, provider)
    expect(unattested.status).toBe('refused')
    expect(unattested.approximated_font_faces).toBeUndefined()
    const refs = (layout: NativeDocxResolvedLayoutInputV1) => [...layout.runs.map((entry) => entry.properties), ...layout.paragraphs.map((entry) => entry.paragraph_mark_properties)].flatMap((p) => p?.font_family ? [{ family: p.font_family, weight: p.bold ? 700 : 400, style: p.italic ? 'italic' : 'normal' }] : [])
    expect(projectNativeDocxLatinFontFallbacksV1(input.document, input.resolved_layout, [fact], refs).applied).toEqual([fact])
    expect(projectNativeDocxLatinFontFallbacksV1(input.document, input.resolved_layout, [{ ...fact, font_family: 'Unattested Face' }], refs).applied).toEqual([])
    // A host manifest that attests the authored face admits it even as a new reference.
    expect(projectNativeDocxLatinFontFallbacksV1(input.document, input.resolved_layout, [{ ...fact, font_family: 'Unattested Face' }], refs, (family, weight, style) => family === 'Unattested Face' && weight === 400 && style === 'normal').applied).toEqual([{ ...fact, font_family: 'Unattested Face' }])
    expect(projectNativeDocxLatinFontFallbacksV1(input.document, input.resolved_layout, [{ ...fact, font_family: 'Unattested Face' }], refs, (_family, weight) => weight === 700).applied).toEqual([])
    expect(() => projectNativeDocxLatinFontFallbacksV1(input.document, input.resolved_layout, [{ ...fact, path: '/w:document[1]/w:body[1]/w:p[9]/w:r[1]' }], refs)).toThrow('scope anchor')
    const authored = structuredClone(resolved); authored.runs.find((entry) => entry.run_id === run.id)!.properties.font_family = family
    expect(() => projectNativeDocxLatinFontFallbacksV1(input.document, authored, [fact], refs)).toThrow('override')
    // Evidence is bound to the eligibility wire: a mismatching hash is refused at decode.
    await expect(renderNativeDocxApproximatePagePreviewV1(input, { ...eligibility, latin_font_fallbacks: [{ ...fact, package_sha256: `sha256:${'f'.repeat(64)}` }] }, provider)).rejects.toThrow()
  }, 20000)

  it('declares the indented cell line policy once per paragraph even when only a hanging indent continuation line needs it', async () => {
    const input = tableFixture(), document = input.document as NativeDocxDocumentV1, resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1, settings = input.pagination_settings as NativeDocxPaginationSettingsV1
    const paragraph = document.body.blocks[0]!.table!.rows[0]!.cells[0]!.paragraphs[0]!
    const properties = resolved.paragraphs.find((entry) => entry.paragraph_id === paragraph.id)!.properties
    properties.indent_left_twips = 360; properties.hanging_twips = 360
    settings.profile = 'unsupported'; delete settings.compatibility_mode
    settings.diagnostics = [{ code: 'COMPATIBILITY_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: SETTINGS_PART, path: '/w:settings[1]/w:compat[1]', preservation: 'preserve-verbatim', message: 'Legacy12' }]
    const eligibility = { protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: settings.document_id, revision: settings.revision, package_sha256: HASH, settings_sha256: settings.settings_sha256, status: 'eligible', legacy_compatibility_mode: 12, reasons: ['Legacy12'] }
    const outlines = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
    const provider = { providerId: input.outline_provider.provider_id, providerRevision: input.outline_provider.provider_revision, getGlyphOutline(request: import('./nativePagePaintV1.js').NativeDocxGlyphOutlineRequestV1) { const o = outlines.outline(request.glyph_id); return o.path.length ? { status: 'outlined' as const, ...request, ...o } : { status: 'empty' as const, ...request, units_per_em: o.units_per_em } } }
    // One line: the hanging first line is full width, so no policy is needed or declared.
    const single = await renderNativeDocxApproximatePagePreviewV1(structuredClone(input), eligibility, provider)
    expect(single.status).toBe('painted')
    expect(single.reasons).not.toContain(DOCX_APPROXIMATE_INDENTED_CELL_LINE_WARNING)
    // Wrapped: continuation lines are indented and take the approximate branch.
    paragraph.runs[0]!.text = Array.from({ length: 60 }, () => 'hanging').join(' ')
    const wrapped = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, provider)
    expect(wrapped.status).toBe('painted')
    const lines = wrapped.pages.flatMap((page) => page.lines.filter((line) => line.paragraph_id === paragraph.id))
    expect(lines.length).toBeGreaterThan(1)
    expect(wrapped.reasons).toContain(DOCX_APPROXIMATE_INDENTED_CELL_LINE_WARNING)
    expect(wrapped.reasons.filter((r) => r === DOCX_APPROXIMATE_INDENTED_CELL_LINE_WARNING)).toHaveLength(1)
  }, 20000)
})

describe('approximate DrawingML shapes', () => {
  const RUN_ANCHOR_START = 181, RUN_ANCHOR_END = 189
  function legacyEligibility(input: NativeDocxPagePaintPrepareInputV1) {
    const settings = input.pagination_settings as NativeDocxPaginationSettingsV1
    settings.profile = 'unsupported'; delete settings.compatibility_mode
    settings.diagnostics = [{ code: 'COMPATIBILITY_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: SETTINGS_PART, path: '/w:settings[1]/w:compat[1]', preservation: 'preserve-verbatim', message: 'Legacy mode' }]
    return { protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: settings.document_id, revision: settings.revision, package_sha256: HASH, settings_sha256: settings.settings_sha256, status: 'eligible', legacy_compatibility_mode: 12, reasons: ['Legacy mode'] }
  }
  function outlineProvider(input: NativeDocxPagePaintPrepareInputV1) {
    const outlines = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
    return { providerId: input.outline_provider.provider_id, providerRevision: input.outline_provider.provider_revision, getGlyphOutline(request: import('./nativePagePaintV1.js').NativeDocxGlyphOutlineRequestV1) { const o = outlines.outline(request.glyph_id); return o.path.length ? { status: 'outlined' as const, ...request, ...o } : { status: 'empty' as const, ...request, units_per_em: o.units_per_em } } }
  }
  /** One refused wps run after the fixture text run, with its retained source diagnostic. */
  function shapeInput(shape: Record<string, unknown>) {
    const input = fixture()
    const document = input.document as NativeDocxDocumentV1
    const paragraph = document.body.blocks[0]!.paragraph!
    document.unsupported = [{ id: 'unsupported:shape', code: 'UNMODELED_RUN_CONTENT', capability: 'runs', scope_id: paragraph.id, anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[2]/mc:AlternateContent[1]', RUN_ANCHOR_START + 1, RUN_ANCHOR_END - 1), preservation: 'refuse-mutation', message: 'Run content outside text, controls, and native references is preserved verbatim' }]
    const eligibility = legacyEligibility(input)
    const item = {
      id: 'approximate-drawing-shape:test:1', paragraph_id: paragraph.id, diagnostic_ids: ['unsupported:shape'],
      anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[2]/mc:AlternateContent[1]/mc:Choice[1]/w:drawing[1]', RUN_ANCHOR_START + 2, RUN_ANCHOR_END - 2), run_anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[2]', RUN_ANCHOR_START, RUN_ANCHOR_END),
      status: 'supported', width_emu: 914400, height_emu: 457200, rotation_degrees: 0, flip_horizontal: false, flip_vertical: false, ...shape,
    }
    const shapes = { protocol: 'injoffice.docx.approximate-drawing-shapes', version: 1, policy: 'docx.approximate-drawing-shape-preview-v1', package_sha256: HASH, part_sha256: HASH, items: [item], omitted_count: 0 }
    return { input, eligibility, shapes, item, paragraph }
  }
  const pageAnchor = (overrides: Record<string, unknown> = {}) => ({ policy: 'relative-position-no-wrap-v2', source_anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[2]/mc:AlternateContent[1]/mc:Choice[1]/w:drawing[1]/wp:anchor[1]', RUN_ANCHOR_START + 3, RUN_ANCHOR_END - 3), horizontal_anchor: anchor('/h', RUN_ANCHOR_START + 3, RUN_ANCHOR_START + 4), vertical_anchor: anchor('/v', RUN_ANCHOR_START + 4, RUN_ANCHOR_START + 5), x_emu: 914400, y_emu: 1828800, horizontal_relative: 'page', vertical_relative: 'page', stacking: { behind_doc: false, relative_height: 7 }, ...overrides })

  it('paints an anchored filled rectangle with its outline behind or in front of body text and discloses the policy', async () => {
    for (const behind of [false, true]) {
      const { input, eligibility, shapes } = shapeInput({ placement: 'anchored', preset: 'rect', fill_rgb: '4F81BD', line: { rgb: '243F60', width_emu: 25400, dash: 'solid' }, page_anchor: pageAnchor({ stacking: { behind_doc: behind, relative_height: 3 } }), wrap: 'none', notes: ['fill resolved from theme fill style'] })
      const before = structuredClone(input)
      const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input), { drawingShapes: shapes })
      expect(paint.status).toBe('painted')
      expect(input).toEqual(before)
      const commands = paint.pages[0]!.commands
      const fill = commands.find(c => c.kind === 'fill_table_cell')
      expect(fill).toMatchObject({ table_id: 'docx.approximate-drawing-shape-preview-v1', row_id: 'approximate-drawing-shape:test:1', x_millipoints: 72_000, y_millipoints: 144_000, width_millipoints: 72_000, height_millipoints: 36_000, fill_rgb: '4F81BD' })
      const strokes = commands.filter(c => c.kind === 'stroke_table_border')
      expect(strokes.map(c => c.kind === 'stroke_table_border' && c.edge)).toEqual(['top', 'right', 'bottom', 'left'])
      expect(strokes.every(c => c.kind === 'stroke_table_border' && c.width_millipoints === 2000 && c.stroke_rgb === '243F60')).toBe(true)
      const glyphIndex = commands.findIndex(c => c.kind === 'fill_glyph_path'), fillIndex = commands.indexOf(fill!)
      expect(glyphIndex).toBeGreaterThanOrEqual(0)
      expect(behind ? fillIndex < glyphIndex : fillIndex > glyphIndex).toBe(true)
      expect(paint.reasons.some(r => r.startsWith('docx.approximate-drawing-shape-preview:') && r.includes('painted 1 of 1') && r.includes('theme fill style'))).toBe(true)
      expect(decodeNativeDocxApproximatePagePreviewV1(paint).ok).toBe(true)
    }
  }, 20000)

  it('paints a group shape as its flattened children at their mapped offsets and leaves the strict lane byte-identical', async () => {
    // The sidecar flattens a wpg:wgp into one anchored item per child, each
    // already mapped out of the group's child coordinate space, so both children
    // share the group's run and drawing anchors and differ only in placement.
    const { input, eligibility, shapes, item } = shapeInput({
      placement: 'anchored', preset: 'rect', fill_rgb: '4F81BD', line: { rgb: '28415F', width_emu: 25400, dash: 'solid' },
      page_anchor: pageAnchor({ x_emu: 914400, y_emu: 1828800 }), wrap: 'none', width_emu: 1270000, height_emu: 635000,
      notes: ["group shape child placed by mapping its child coordinates into the group's declared extent"],
    })
    shapes.items.push({
      ...item, id: 'approximate-drawing-shape:test:2', width_emu: 635000, height_emu: 381000,
      page_anchor: pageAnchor({ x_emu: 1549400, y_emu: 2209800 }),
    } as never, {
      id: 'approximate-drawing-shape:test:3', paragraph_id: item.paragraph_id, diagnostic_ids: item.diagnostic_ids, anchor: item.anchor, run_anchor: item.run_anchor,
      status: 'omitted', reason: 'nested-group', width_emu: 0, height_emu: 0, rotation_degrees: 0, flip_horizontal: false, flip_vertical: false,
    } as never)
    const original = structuredClone(input)
    const { compileNativeDocxPagePaintV1 } = await import('./nativePagePaintV1.js')
    const strictBefore = await compileNativeDocxPagePaintV1((await prepareNativeDocxPagePaintV1(structuredClone(original))).page_paint_request, outlineProvider(input))
    const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input), { drawingShapes: shapes })
    expect(paint.status).toBe('painted')
    expect(input).toEqual(original)
    const fills = paint.pages[0]!.commands.flatMap(c => c.kind === 'fill_table_cell' && c.table_id === 'docx.approximate-drawing-shape-preview-v1' ? [c] : [])
    expect(fills.map(c => ({ row: c.row_id, x: c.x_millipoints, y: c.y_millipoints, w: c.width_millipoints, h: c.height_millipoints }))).toEqual([
      { row: 'approximate-drawing-shape:test:1', x: 72_000, y: 144_000, w: 100_000, h: 50_000 },
      { row: 'approximate-drawing-shape:test:2', x: 122_000, y: 174_000, w: 50_000, h: 30_000 },
    ])
    // Each child keeps its own outline, so the group reads as separate shapes.
    expect(paint.pages[0]!.commands.filter(c => c.kind === 'stroke_table_border' && c.row_id === 'approximate-drawing-shape:test:2')).toHaveLength(4)
    expect(paint.reasons).toContain(DOCX_APPROXIMATE_DRAWING_SHAPE_WARNING)
    expect(paint.reasons.some(r => r.startsWith('docx.approximate-drawing-shape-preview:') && r.includes('painted 2 of 3') && r.includes("mapping its child coordinates into the group's declared extent"))).toBe(true)
    expect(paint.reasons.some(r => r.startsWith('docx.approximate-drawing-shape-omitted:') && r.includes('approximate-drawing-shape:test:3 (nested-group)'))).toBe(true)
    expect(decodeNativeDocxApproximatePagePreviewV1(paint).ok).toBe(true)
    // Strict paint of the same bytes is unchanged by the approximate group lane.
    const strictAfter = await compileNativeDocxPagePaintV1((await prepareNativeDocxPagePaintV1(structuredClone(original))).page_paint_request, outlineProvider(input))
    expect(JSON.stringify(strictAfter)).toBe(JSON.stringify(strictBefore))
  }, 20000)

  it('reserves inline rectangles in the line, paints a line preset as one stroke, and keeps omitted shapes disclosed', async () => {
    const { input, eligibility, shapes, item } = shapeInput({ placement: 'inline', preset: 'rect', fill_rgb: '0D0D0D' })
    const { fill_rgb: _fill, ...unfilled } = item as typeof item & { fill_rgb?: string }
    shapes.items.push({ ...unfilled, id: 'approximate-drawing-shape:test:2', placement: 'anchored', preset: 'line', line: { rgb: 'FF0000', width_emu: 12700, dash: 'solid' }, page_anchor: pageAnchor({ x_emu: 0, y_emu: 0 }), flip_horizontal: true } as never, { ...item, id: 'approximate-drawing-shape:test:3', status: 'omitted', reason: 'unsupported-preset:ellipse' } as never)
    const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input), { drawingShapes: shapes })
    expect(paint.status).toBe('painted')
    const commands = paint.pages[0]!.commands
    expect(commands.filter(c => c.kind === 'fill_text_highlight')).toHaveLength(0)
    const fill = commands.find(c => c.kind === 'fill_table_cell')
    if (!fill || fill.kind !== 'fill_table_cell') throw new Error('inline fill missing')
    const line = paint.pages[0]!.lines[0]!
    // The atom follows the text run 'A' on the first line and sits on its baseline.
    expect(fill.width_millipoints).toBe(72_000); expect(fill.height_millipoints).toBe(36_000)
    expect(fill.x_millipoints).toBeGreaterThan(line.x_millipoints); expect(fill.y_millipoints + fill.height_millipoints).toBe(line.baseline_y_millipoints)
    expect(line.width_millipoints).toBeGreaterThan(72_000)
    const stroke = commands.find(c => c.kind === 'stroke_table_border')
    expect(stroke).toMatchObject({ row_id: 'approximate-drawing-shape:test:2', x1_millipoints: 72_000, y1_millipoints: 0, x2_millipoints: 0, y2_millipoints: 36_000, width_millipoints: 1000, stroke_rgb: 'FF0000' })
    expect(paint.reasons.some(r => r.startsWith('docx.approximate-drawing-shape-omitted:') && r.includes('approximate-drawing-shape:test:3 (unsupported-preset:ellipse)'))).toBe(true)
    expect(decodeNativeDocxApproximatePagePreviewV1(paint).ok).toBe(true)
  }, 20000)

  it('shapes text box paragraphs into the box, attaches glyphs to the anchor line and discloses font substitution', async () => {
    const paragraphs = [0, 1].map(index => ({
      id: `approximate-drawing-shape:test:1:p${index}`, anchor: anchor(`/w:document[1]/w:body[1]/w:p[1]/w:r[2]/mc:AlternateContent[1]/mc:Choice[1]/w:drawing[1]/wp:anchor[1]/a:graphic[1]/a:graphicData[1]/wps:wsp[1]/wps:txbx[1]/w:txbxContent[1]/w:p[${index + 1}]`, RUN_ANCHOR_START + 3 + index, RUN_ANCHOR_START + 4 + index),
      edit_policy: { mode: 'read-only' as const, allowed_operations: [], refusal: { code: 'APPROXIMATE_TEXTBOX_PREVIEW', message: 'read-only', preservation: 'refuse-mutation' as const } }, properties: {},
      runs: [{ kind: 'text' as const, id: `approximate-drawing-shape:test:1:p${index}:r0`, anchor: anchor(`/w:document[1]/w:body[1]/w:p[1]/w:r[2]/mc:AlternateContent[1]/mc:Choice[1]/w:drawing[1]/wp:anchor[1]/a:graphic[1]/a:graphicData[1]/wps:wsp[1]/wps:txbx[1]/w:txbxContent[1]/w:p[${index + 1}]/w:r[1]`, RUN_ANCHOR_START + 3 + index, RUN_ANCHOR_START + 4 + index), text: index === 0 ? 'Box heading' : 'Second line of box text' }],
    }))
    const textbox = {
      link_seq: 0, insets_emu: [91440, 45720, 91440, 45720], vertical_anchor: 't', wrap: 'square', paragraphs,
      resolved_paragraphs: paragraphs.map(p => ({ paragraph_id: p.id, applied_styles: [], properties: {}, paragraph_mark_properties: { font_family: 'DejaVu Sans', font_size_half_points: 20 } })),
      resolved_runs: paragraphs.map((p, index) => ({ run_id: p.runs[0]!.id, paragraph_id: p.id, applied_paragraph_styles: [], applied_character_styles: [], properties: { font_family: index === 0 ? 'DejaVu Sans' : 'Candara', font_size_half_points: 20, color: '1F497D', ...(index === 0 ? { underline: 'single' } : {}) } })),
      omitted_runs: 0, omitted_blocks: 0,
    }
    const { input, eligibility, shapes } = shapeInput({ placement: 'anchored', preset: 'rect', page_anchor: pageAnchor(), wrap: 'none', textbox, width_emu: 2743200, height_emu: 914400 })
    const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input), { drawingShapes: shapes })
    expect(paint.status).toBe('painted')
    const page = paint.pages[0]!
    const boxGlyphs = page.commands.flatMap(c => c.kind === 'fill_glyph_path' && c.source_id.startsWith('approximate-drawing-shape:test:1:p') ? [c] : [])
    expect(boxGlyphs.length).toBeGreaterThan(20)
    const line = page.lines[0]!
    for (const glyph of boxGlyphs) { expect(glyph.line_id).toBe(line.line_id); expect(line.command_ids).toContain(glyph.id) }
    const xs = boxGlyphs.flatMap(g => g.path.flatMap(p => 'x_millipoints' in p ? [p.x_millipoints] : []))
    const ys = boxGlyphs.flatMap(g => g.path.flatMap(p => 'y_millipoints' in p ? [p.y_millipoints] : []))
    expect(Math.min(...xs)).toBeGreaterThanOrEqual(72_000 + 7_200 - 1_000); expect(Math.max(...xs)).toBeLessThanOrEqual(72_000 + 216_000)
    expect(Math.min(...ys)).toBeGreaterThanOrEqual(144_000); expect(Math.max(...ys)).toBeLessThanOrEqual(144_000 + 72_000)
    expect(page.commands.some(c => c.kind === 'stroke_text_underline' && c.source_id === 'approximate-drawing-shape:test:1:p0:r0')).toBe(true)
    // Candara is not in the manifest: the loaded DejaVu face substitutes and the substitution is disclosed.
    expect(page.commands.some(c => c.kind === 'fill_glyph_path' && c.source_id === 'approximate-drawing-shape:test:1:p1:r0')).toBe(true)
    expect(paint.reasons.some(r => r.startsWith('docx.approximate-textbox-substituted-font: Candara / 400 / normal -> DejaVu Sans / 400 / normal'))).toBe(true)
    expect(decodeNativeDocxApproximatePagePreviewV1(paint).ok).toBe(true)
  }, 20000)

  it('paints an in-front shape fill under its own text box glyphs and still over the body line it anchors to', async () => {
    // A front shape replays after the body lines, but its own text box glyphs
    // replay with them, so an opaque fill in that layer would erase the box text.
    const paragraphs = [{
      id: 'approximate-drawing-shape:test:1:p0', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[2]/mc:AlternateContent[1]/mc:Choice[1]/w:drawing[1]/wp:anchor[1]/a:graphic[1]/a:graphicData[1]/wps:wsp[1]/wps:txbx[1]/w:txbxContent[1]/w:p[1]', RUN_ANCHOR_START + 3, RUN_ANCHOR_START + 4),
      edit_policy: { mode: 'read-only' as const, allowed_operations: [], refusal: { code: 'APPROXIMATE_TEXTBOX_PREVIEW', message: 'read-only', preservation: 'refuse-mutation' as const } }, properties: {},
      runs: [{ kind: 'text' as const, id: 'approximate-drawing-shape:test:1:p0:r0', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[2]/mc:AlternateContent[1]/mc:Choice[1]/w:drawing[1]/wp:anchor[1]/a:graphic[1]/a:graphicData[1]/wps:wsp[1]/wps:txbx[1]/w:txbxContent[1]/w:p[1]/w:r[1]', RUN_ANCHOR_START + 3, RUN_ANCHOR_START + 4), text: 'Anchored TextBox' }],
    }]
    const textbox = {
      link_seq: 0, insets_emu: [91440, 45720, 91440, 45720], vertical_anchor: 't', wrap: 'square', paragraphs,
      resolved_paragraphs: paragraphs.map(p => ({ paragraph_id: p.id, applied_styles: [], properties: {}, paragraph_mark_properties: { font_family: 'DejaVu Sans', font_size_half_points: 20 } })),
      resolved_runs: paragraphs.map(p => ({ run_id: p.runs[0]!.id, paragraph_id: p.id, applied_paragraph_styles: [], applied_character_styles: [], properties: { font_family: 'DejaVu Sans', font_size_half_points: 20, color: '000000' } })),
      omitted_runs: 0, omitted_blocks: 0,
    }
    const { input, eligibility, shapes } = shapeInput({ placement: 'anchored', preset: 'rect', fill_rgb: 'FFFFFF', page_anchor: pageAnchor({ stacking: { behind_doc: false, relative_height: 3 } }), wrap: 'none', textbox, width_emu: 2743200, height_emu: 914400 })
    const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input), { drawingShapes: shapes })
    expect(paint.status).toBe('painted')
    const commands = paint.pages[0]!.commands
    const fillIndex = commands.findIndex(c => c.kind === 'fill_table_cell' && c.row_id === 'approximate-drawing-shape:test:1')
    const boxGlyphIndexes = commands.flatMap((c, index) => c.kind === 'fill_glyph_path' && c.source_id === 'approximate-drawing-shape:test:1:p0:r0' ? [index] : [])
    const bodyGlyphIndexes = commands.flatMap((c, index) => c.kind === 'fill_glyph_path' && !c.source_id.startsWith('approximate-drawing-shape:') ? [index] : [])
    expect(fillIndex).toBeGreaterThanOrEqual(0)
    expect(boxGlyphIndexes.length).toBeGreaterThan(10)
    expect(bodyGlyphIndexes.length).toBeGreaterThan(0)
    // The opaque fill must sit under every glyph of its own box ...
    expect(Math.min(...boxGlyphIndexes)).toBeGreaterThan(fillIndex)
    // ... and still over the body text of the line it anchors to.
    expect(Math.max(...bodyGlyphIndexes)).toBeLessThan(fillIndex)
    expect(decodeNativeDocxApproximatePagePreviewV1(paint).ok).toBe(true)
  }, 20000)

  // The fixture section is 12240 twips wide with 1440 twip side margins, so the
  // single column is 9360 twips = 468_000 millipoints = 5_943_600 EMU.
  const COLUMN_MILLIPOINTS = 468_000, PAGE_MILLIPOINTS = 612_000, MARGIN_MILLIPOINTS = 72_000
  const inlineFill = (paint: Awaited<ReturnType<typeof renderNativeDocxApproximatePagePreviewV1>>) => {
    const fill = paint.pages[0]!.commands.find(c => c.kind === 'fill_table_cell')
    if (!fill || fill.kind !== 'fill_table_cell') throw new Error(`inline fill missing: ${paint.reasons.filter(r => r.includes('approximate-drawing')).join(' | ')}`)
    return fill
  }

  it('paints an over-wide inline shape at its declared extent while the reserved atom stays inside the column', async () => {
    // 540pt wide against a 468pt column: the atom must shrink to break, the object must not.
    const { input, eligibility, shapes, paragraph } = shapeInput({ placement: 'inline', preset: 'rect', fill_rgb: '0D0D0D', width_emu: 6_858_000, height_emu: 25_400 })
    const before = structuredClone(input)
    const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input), { drawingShapes: shapes })
    expect(paint.status).toBe('painted')
    expect(input).toEqual(before)
    const fill = inlineFill(paint)
    // The painted extent is the declared wp:extent, not the narrowed reservation.
    expect(fill.width_millipoints).toBe(540_000)
    expect(fill.height_millipoints).toBe(2_000)
    // The reserved atom still broke inside the column, so no line exceeds it.
    const lines = paint.pages[0]!.lines.filter(line => line.paragraph_id === paragraph.id)
    expect(lines.length).toBeGreaterThan(0)
    for (const line of lines) expect(line.width_millipoints).toBeLessThanOrEqual(COLUMN_MILLIPOINTS)
    // It starts at the atom's start edge and runs past the column into the margin.
    expect(fill.x_millipoints).toBe(MARGIN_MILLIPOINTS)
    expect(fill.x_millipoints + fill.width_millipoints).toBeGreaterThan(MARGIN_MILLIPOINTS + COLUMN_MILLIPOINTS)
    expect(fill.x_millipoints + fill.width_millipoints).toBeLessThanOrEqual(PAGE_MILLIPOINTS)
    expect(paint.reasons.some(r => r.startsWith('docx.approximate-drawing-shape-preview:') && r.includes('declared extent paints and overflows into the margin, clipped to the page'))).toBe(true)
    expect(decodeNativeDocxApproximatePagePreviewV1(paint).ok).toBe(true)
  }, 20000)

  it('clips a declared inline extent that runs past the page edge instead of painting outside the page', async () => {
    const { input, eligibility, shapes } = shapeInput({ placement: 'inline', preset: 'rect', fill_rgb: '0D0D0D', width_emu: 12_700_000, height_emu: 25_400 })
    const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input), { drawingShapes: shapes })
    expect(paint.status).toBe('painted')
    const fill = inlineFill(paint)
    // 1000pt declared from the 72pt margin would reach 1072pt; the page stops it at 612pt.
    expect(fill.x_millipoints).toBe(MARGIN_MILLIPOINTS)
    expect(fill.x_millipoints + fill.width_millipoints).toBe(PAGE_MILLIPOINTS)
    expect(fill.width_millipoints).toBeLessThan(1_000_000)
    expect(decodeNativeDocxApproximatePagePreviewV1(paint).ok).toBe(true)
  }, 20000)

  it('keeps the narrowed extent for an over-wide inline shape in a right-to-left paragraph and discloses why', async () => {
    const { input, eligibility, shapes, paragraph } = shapeInput({ placement: 'inline', preset: 'rect', fill_rgb: '0D0D0D', width_emu: 6_858_000, height_emu: 25_400 })
    const resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    resolved.paragraphs.find(entry => entry.paragraph_id === paragraph.id)!.properties.bidi = true
    const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input), { drawingShapes: shapes })
    expect(paint.status).toBe('painted')
    const fill = inlineFill(paint)
    // Overflow would have to run toward the start margin, which this preview does not model.
    expect(fill.width_millipoints).toBe(COLUMN_MILLIPOINTS)
    expect(paint.reasons.some(r => r.startsWith('docx.approximate-drawing-shape-preview:') && r.includes('overflow toward the start margin of a right-to-left paragraph is not modeled'))).toBe(true)
    expect(paint.reasons.some(r => r.includes('declared extent paints and overflows'))).toBe(false)
    expect(decodeNativeDocxApproximatePagePreviewV1(paint).ok).toBe(true)
  }, 20000)

  it('leaves an inline shape that already fits its column painted at its reserved extent with no clamp disclosure', async () => {
    const { input, eligibility, shapes } = shapeInput({ placement: 'inline', preset: 'rect', fill_rgb: '0D0D0D', width_emu: 914_400, height_emu: 457_200 })
    const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input), { drawingShapes: shapes })
    expect(paint.status).toBe('painted')
    const fill = inlineFill(paint)
    expect(fill.width_millipoints).toBe(72_000)
    expect(paint.reasons.some(r => r.includes('column width') || r.includes('declared extent paints'))).toBe(false)
    expect(decodeNativeDocxApproximatePagePreviewV1(paint).ok).toBe(true)
  }, 20000)

  it('leaves the strict page paint request and response byte-identical whether or not the inline shape evidence is supplied', async () => {
    const { input, eligibility, shapes } = shapeInput({ placement: 'inline', preset: 'rect', fill_rgb: '0D0D0D', width_emu: 6_858_000, height_emu: 25_400 })
    const strictWithout = await prepareNativeDocxPagePaintV1(structuredClone(input))
    const strictWith = await prepareNativeDocxPagePaintV1(structuredClone(input))
    expect(JSON.stringify(strictWith)).toBe(JSON.stringify(strictWithout))
    // Running the approximate tier with the evidence must not perturb the strict tier either.
    const before = structuredClone(input)
    const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input), { drawingShapes: shapes })
    expect(paint.status).toBe('painted')
    expect(inlineFill(paint).width_millipoints).toBe(540_000)
    expect(input).toEqual(before)
    const strictAfter = await prepareNativeDocxPagePaintV1(structuredClone(input))
    expect(JSON.stringify(strictAfter)).toBe(JSON.stringify(strictWithout))
    expect(strictAfter.page_paint_request.paginated_layout.status).toBe('refused')
  }, 30000)

  it('keeps painting the body when the sidecar does not exact-join, discloses the refusal, and omits nothing else', async () => {
    // A supported shape claiming a run that also carries modeled text (the Go sidecar omits these as shared-run).
    const { input, eligibility, shapes, item } = shapeInput({ placement: 'anchored', preset: 'rect', fill_rgb: '4F81BD', page_anchor: pageAnchor() })
    item.run_anchor = anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]', 105, 185); item.anchor = anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:drawing[1]', 111, 179)
    const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input), { drawingShapes: shapes })
    expect(paint.status).toBe('painted')
    expect(paint.pages[0]!.commands.some(c => c.kind === 'fill_glyph_path')).toBe(true)
    expect(paint.pages[0]!.commands.some(c => c.kind === 'fill_table_cell' || c.kind === 'stroke_table_border')).toBe(false)
    expect(paint.reasons).toContain('docx.approximate-drawing-shape-omitted: drawing-shape evidence did not exact-join the source document and was not used; refused drawings stay omitted')
    expect(paint.omitted_content.some(entry => entry.code === 'UNMODELED_RUN_CONTENT')).toBe(true)
    expect(paint.content_status).toBe('partial')
  }, 20000)

  it('centers odd-width shapes without dropping them on coordinate parity', async () => {
    for (const cx of [2990000, 2995000, 2998800]) {
      const { input, eligibility, shapes } = shapeInput({ placement: 'anchored', preset: 'rect', fill_rgb: '4F81BD', width_emu: cx, page_anchor: pageAnchor({ x_emu: 0, horizontal_align: 'center' }) })
      const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input), { drawingShapes: shapes })
      const fill = paint.pages[0]!.commands.find(c => c.kind === 'fill_table_cell')
      if (!fill || fill.kind !== 'fill_table_cell') throw new Error(`centered shape ${cx} was dropped: ${paint.reasons.filter(r => r.includes('omitted')).join(' | ')}`)
      const width = Math.round(cx / 12.7)
      expect(fill.width_millipoints).toBe(width)
      expect(Math.abs(fill.x_millipoints - (612_000 - width) / 2)).toBeLessThanOrEqual(1)
      expect(paint.content_status).toBe('complete')
    }
  }, 30000)

  it('discloses shapes the painter drops as omitted content', async () => {
    const { input, eligibility, shapes } = shapeInput({ placement: 'anchored', preset: 'rect', fill_rgb: '4F81BD', page_anchor: pageAnchor({ x_emu: 120_000_000, y_emu: 120_000_000 }) })
    const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input), { drawingShapes: shapes })
    expect(paint.status).toBe('painted')
    expect(paint.pages[0]!.commands.some(c => c.kind === 'fill_table_cell')).toBe(false)
    expect(paint.reasons.some(r => r.startsWith('docx.approximate-drawing-shape-omitted:') && r.includes('outside-page')), paint.reasons.filter(r => r.includes('approximate-drawing')).join(' / ')).toBe(true)
    expect(paint.content_status).toBe('partial')
    expect(paint.omitted_content.some(entry => entry.code === 'UNMODELED_RUN_CONTENT' && entry.scope_id === 'paragraph:1')).toBe(true)
    expect(decodeNativeDocxApproximatePagePreviewV1(paint).ok).toBe(true)
  }, 20000)

  it('refuses sidecars that do not exact-join the source document', async () => {
    const cases: Array<[string, (shapes: any, item: any) => void]> = [
      ['package', (shapes) => { shapes.package_sha256 = `sha256:${'b'.repeat(64)}` }],
      ['paragraph', (_s, item) => { item.paragraph_id = 'paragraph:missing' }],
      ['diagnostic', (_s, item) => { item.diagnostic_ids = ['unsupported:other'] }],
      ['overlap', (_s, item) => { item.run_anchor = anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]', 105, 185); item.anchor = anchor('/x', 110, 180) }],
      ['unknown field', (shapes) => { shapes.extra = true }],
      ['fill', (_s, item) => { item.fill_rgb = 'red' }],
      ['anchor', (_s, item) => { item.page_anchor = { ...item.page_anchor, policy: 'page-offset-no-wrap-v1' } }],
      ['textbox keys', (_s, item) => { item.textbox = { paragraphs: [] } }],
    ]
    for (const [name, mutate] of cases) {
      const { input, eligibility, shapes, item } = shapeInput({ placement: 'anchored', preset: 'rect', fill_rgb: '4F81BD', page_anchor: pageAnchor() })
      mutate(shapes, item)
      expect(() => decodeNativeDocxApproximateDrawingShapesV1(shapes, input.document as NativeDocxDocumentV1), name).toThrow()
      const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input), { drawingShapes: shapes })
      expect(paint.status, name).toBe('painted')
      expect(paint.pages[0]!.commands.some(c => c.kind === 'fill_table_cell'), name).toBe(false)
      expect(paint.reasons.some(r => r.includes('did not exact-join')), name).toBe(true)
    }
    const ineligible = shapeInput({ placement: 'anchored', preset: 'rect', fill_rgb: '4F81BD', page_anchor: pageAnchor() })
    await expect(renderNativeDocxApproximatePagePreviewV1(ineligible.input, { ...ineligible.eligibility, status: 'ineligible' }, outlineProvider(ineligible.input), { drawingShapes: ineligible.shapes })).rejects.toThrow()
  }, 20000)

  it('shapes an exactly qualified strict inline text box as a 10/127 EMU atom that the shaped-lines contract accepts', async () => {
    const input = fixture()
    const document = input.document as NativeDocxDocumentV1, resolved = input.resolved_layout as NativeDocxResolvedLayoutInputV1
    const paragraph = document.body.blocks[0]!.paragraph!
    paragraph.runs.push({ kind: 'drawing', id: 'run:textbox', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[2]', 181, 189), drawing: { id: 'drawing:textbox', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[2]/w:drawing[1]', 182, 188), placement: 'inline', width_emu: 914400, height_emu: 457200, textbox_text: 'Box', textbox_fill_rgb: 'FFF2CC', textbox_line_rgb: '204060', edit_policy: { mode: 'read-only', allowed_operations: [], refusal: { code: 'EXTRACT_ONLY', message: 'fixture', preservation: 'refuse-mutation' } } } })
    resolved.runs.push({ run_id: 'run:textbox', paragraph_id: paragraph.id, applied_paragraph_styles: [], applied_character_styles: [], properties: {} })
    // Strict preparation now decodes the shaper's textbox atom instead of failing the pagination request.
    const prepared = await prepareNativeDocxPagePaintV1(input)
    const shaped = prepared.page_paint_request.pagination_request.shaped_lines.paragraphs[0]!.lines[0]!
    expect(shaped.fragments.find(f => f.source_kind === 'textbox')).toMatchObject({ source_id: 'run:textbox', text: '', glyphs: [], advance_inline_millipoints: 72_000, ascent_millipoints: 36_000 })
    expect(decodeNativeDocxShapedLines(prepared.page_paint_request.pagination_request.shaped_lines).ok).toBe(true)
    const forged = structuredClone(prepared.page_paint_request.pagination_request.shaped_lines)
    const atom = forged.paragraphs[0]!.lines[0]!.fragments.find(f => f.source_kind === 'textbox')!
    atom.text = 'x'
    expect(decodeNativeDocxShapedLines(forged).ok).toBe(false)
    // Strict pagination keeps refusing drawing runs outside the exact inline-image slice; that policy is unchanged here.
    const layout = prepared.page_paint_request.paginated_layout
    expect(layout.status).toBe('refused')
    expect(layout.diagnostics.some(d => d.code === 'body-structure-unsupported' && d.scope_id === 'run:textbox')).toBe(true)
  }, 20000)
})

describe('approximate DrawingML charts', () => {
  const RUN_ANCHOR_START = 181, RUN_ANCHOR_END = 189
  /** The sidecar arrives as JSON: undefined fields are absent on the wire. */
  const wire = (value: unknown) => JSON.parse(JSON.stringify(value)) as unknown
  const CHART_ID = 'approximate-drawing-chart:test:1'
  function legacyEligibility(input: NativeDocxPagePaintPrepareInputV1) {
    const settings = input.pagination_settings as NativeDocxPaginationSettingsV1
    settings.profile = 'unsupported'; delete settings.compatibility_mode
    settings.diagnostics = [{ code: 'COMPATIBILITY_SETTING_UNSUPPORTED', severity: 'unsupported', part_name: SETTINGS_PART, path: '/w:settings[1]/w:compat[1]', preservation: 'preserve-verbatim', message: 'Legacy mode' }]
    return { protocol: 'injoffice.docx.approximation-eligibility', version: 1, document_id: settings.document_id, revision: settings.revision, package_sha256: HASH, settings_sha256: settings.settings_sha256, status: 'eligible', legacy_compatibility_mode: 12, reasons: ['Legacy mode'] }
  }
  function outlineProvider(input: NativeDocxPagePaintPrepareInputV1) {
    const outlines = createHarfBuzzOutlineProviderV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })
    return { providerId: input.outline_provider.provider_id, providerRevision: input.outline_provider.provider_revision, getGlyphOutline(request: import('./nativePagePaintV1.js').NativeDocxGlyphOutlineRequestV1) { const o = outlines.outline(request.glyph_id); return o.path.length ? { status: 'outlined' as const, ...request, ...o } : { status: 'empty' as const, ...request, units_per_em: o.units_per_em } } }
  }
  const font = (size: number) => ({ family: 'Calibri', size_hundredth_pt: size, rgb: '595959', bold: false, italic: false })
  const grey = { rgb: 'D9D9D9', width_emu: 9525, dash: 'solid' }
  /** Two series over two categories, the first with a thick dotted outline like Chart_BorderLine_Style. */
  function chartModel(overrides: Record<string, unknown> = {}) {
    return {
      kind: 'bar', bar_direction: 'column', grouping: 'clustered', gap_width_percent: 219, overlap_percent: -27, categories: ['Alpha', 'Beta'],
      series: [
        { index: 0, order: 0, title: 'First', values: ['4', '2'], fill_rgb: '4472C4', line: { rgb: '70AD47', width_emu: 76200, dash: 'sysDot' } },
        { index: 1, order: 1, title: 'Second', values: ['1', '3'], fill_rgb: 'ED7D31' },
      ],
      category_axis: { deleted: false, orientation: 'minMax', line: grey, labels: font(900), number_format: 'General' },
      value_axis: { deleted: false, orientation: 'minMax', labels: font(900), major_gridlines: grey, number_format: 'General' },
      title: { font: font(1400), overlay: false }, legend: { position: 'b', font: font(900), overlay: false },
      area_fill_rgb: 'FFFFFF', area_line: grey, ...overrides,
    }
  }
  /** One refused chart drawing after the fixture text run, with its retained PICTURE_GRAPHIC_REQUIRED diagnostic. */
  function chartInput(item: Record<string, unknown>) {
    const input = fixture()
    const document = input.document as NativeDocxDocumentV1
    const paragraph = document.body.blocks[0]!.paragraph!
    document.unsupported = [{ id: 'unsupported:chart', code: 'PICTURE_GRAPHIC_REQUIRED', capability: 'drawings', scope_id: paragraph.id, anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[2]/w:drawing[1]/wp:inline[1]/a:graphic[1]/a:graphicData[1]', RUN_ANCHOR_START + 3, RUN_ANCHOR_END - 3), preservation: 'refuse-mutation', message: 'Drawing graphic requires a picture' }]
    const eligibility = legacyEligibility(input)
    const chart = {
      id: CHART_ID, paragraph_id: paragraph.id, diagnostic_ids: ['unsupported:chart'],
      anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[2]/w:drawing[1]', RUN_ANCHOR_START + 1, RUN_ANCHOR_END - 1), run_anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[2]', RUN_ANCHOR_START, RUN_ANCHOR_END),
      status: 'supported', placement: 'inline', width_emu: 2743200, height_emu: 1828800, chart_part: 'word/charts/chart1.xml', chart_part_sha256: HASH, chart: chartModel(), ...item,
    }
    const charts = { protocol: 'injoffice.docx.approximate-drawing-charts', version: 1, policy: 'docx.approximate-drawing-chart-preview-v1', package_sha256: HASH, part_sha256: HASH, items: [chart], omitted_count: 0 }
    return { input, eligibility, charts, chart, paragraph }
  }
  const pageAnchor = () => ({ policy: 'relative-position-no-wrap-v2', source_anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[2]/w:drawing[1]/wp:anchor[1]', RUN_ANCHOR_START + 2, RUN_ANCHOR_END - 2), horizontal_anchor: anchor('/h', RUN_ANCHOR_START + 3, RUN_ANCHOR_START + 4), vertical_anchor: anchor('/v', RUN_ANCHOR_START + 4, RUN_ANCHOR_START + 5), x_emu: 914400, y_emu: 1828800, horizontal_relative: 'page', vertical_relative: 'page', stacking: { behind_doc: false, relative_height: 7 } })

  it('reserves an inline clustered column chart on the line and paints bars, axes, legend and labels inside its frame', async () => {
    const { input, eligibility, charts } = chartInput({})
    const before = structuredClone(input)
    const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input), { drawingCharts: wire(charts) })
    expect(paint.status).toBe('painted')
    expect(input).toEqual(before)
    const page = paint.pages[0]!
    const line = page.lines[0]!
    const fills = page.commands.filter((c): c is Extract<typeof c, { kind: 'fill_table_cell' }> => c.kind === 'fill_table_cell' && c.table_id === 'docx.approximate-drawing-chart-preview-v1')
    const area = fills.find(c => c.cell_id === 'area')
    if (!area) throw new Error('chart area missing')
    // The atom follows the text run 'A' on the first line and sits on its baseline.
    expect(area.width_millipoints).toBe(216_000); expect(area.height_millipoints).toBe(144_000)
    expect(area.x_millipoints).toBeGreaterThan(line.x_millipoints); expect(area.y_millipoints + area.height_millipoints).toBe(line.baseline_y_millipoints)
    expect(page.commands.filter(c => c.kind === 'fill_text_highlight')).toHaveLength(0)
    const bars = fills.filter(c => c.cell_id.startsWith('bar:'))
    expect(bars).toHaveLength(4)
    for (const bar of bars) {
      expect(bar.x_millipoints).toBeGreaterThanOrEqual(area.x_millipoints); expect(bar.x_millipoints + bar.width_millipoints).toBeLessThanOrEqual(area.x_millipoints + area.width_millipoints)
      expect(bar.y_millipoints).toBeGreaterThanOrEqual(area.y_millipoints); expect(bar.y_millipoints + bar.height_millipoints).toBeLessThanOrEqual(area.y_millipoints + area.height_millipoints)
    }
    const first = bars.filter(c => c.cell_id.startsWith('bar:0:')), second = bars.filter(c => c.cell_id.startsWith('bar:1:'))
    expect(first.map(c => c.fill_rgb)).toEqual(['4472C4', '4472C4']); expect(second.map(c => c.fill_rgb)).toEqual(['ED7D31', 'ED7D31'])
    // Cached values 4 and 2 keep their ratio on the derived 0..5 scale; bars share one baseline.
    expect(Math.abs(first[0]!.height_millipoints - 2 * first[1]!.height_millipoints)).toBeLessThanOrEqual(2)
    expect(new Set(bars.map(c => c.y_millipoints + c.height_millipoints)).size).toBe(1)
    // Series order left to right inside the first category, overlap gap between them.
    expect(first[0]!.x_millipoints + first[0]!.width_millipoints).toBeLessThan(second[0]!.x_millipoints)
    const strokes = page.commands.filter((c): c is Extract<typeof c, { kind: 'stroke_table_border' }> => c.kind === 'stroke_table_border' && c.table_id === 'docx.approximate-drawing-chart-preview-v1')
    expect(strokes.filter(c => c.cell_id === 'gridline').length).toBe(6)
    expect(strokes.filter(c => c.cell_id === 'category-axis').length).toBe(1)
    expect(strokes.filter(c => c.cell_id.startsWith('bar-outline:0:')).length).toBeGreaterThan(8)
    expect(strokes.filter(c => c.cell_id.startsWith('bar-outline:1:')).length).toBe(0)
    expect(strokes.filter(c => c.cell_id === 'area-outline').length).toBe(4)
    expect(fills.filter(c => c.cell_id.startsWith('legend:')).map(c => c.fill_rgb)).toEqual(['4472C4', 'ED7D31'])
    const glyphs = page.commands.filter((c): c is Extract<typeof c, { kind: 'fill_glyph_path' }> => c.kind === 'fill_glyph_path' && c.source_id === CHART_ID)
    // Category labels, six tick labels and two legend entries; the automatic title paints no text.
    expect(glyphs.length).toBeGreaterThan(20)
    for (const glyph of glyphs) { expect(glyph.line_id).toBe(line.line_id); expect(line.command_ids).toContain(glyph.id); expect(glyph.fill_rgb).toBe('595959') }
    const xs = glyphs.flatMap(g => g.path.flatMap(p => 'x_millipoints' in p ? [p.x_millipoints] : []))
    const ys = glyphs.flatMap(g => g.path.flatMap(p => 'y_millipoints' in p ? [p.y_millipoints] : []))
    expect(Math.min(...xs)).toBeGreaterThanOrEqual(area.x_millipoints); expect(Math.max(...xs)).toBeLessThanOrEqual(area.x_millipoints + area.width_millipoints)
    expect(Math.min(...ys)).toBeGreaterThanOrEqual(area.y_millipoints); expect(Math.max(...ys)).toBeLessThanOrEqual(area.y_millipoints + area.height_millipoints)
    // Chart text paints after the chart's own fills and strokes, so labels stay visible over the area fill.
    const chartPaintIndex = Math.max(...page.commands.map((c, index) => (c.kind === 'fill_table_cell' || c.kind === 'stroke_table_border') && c.table_id === 'docx.approximate-drawing-chart-preview-v1' ? index : -1))
    expect(Math.min(...glyphs.map(g => page.commands.indexOf(g)))).toBeGreaterThan(chartPaintIndex)
    // Body text of the same line still precedes the chart paint in replay order.
    expect(page.commands.findIndex(c => c.kind === 'fill_glyph_path' && c.source_id === 'run:1')).toBeLessThan(page.commands.indexOf(area))
    expect(paint.reasons).toContain(DOCX_APPROXIMATE_DRAWING_CHART_WARNING)
    expect(paint.reasons.some(r => r.startsWith('docx.approximate-drawing-chart-preview:') && r.includes('painted 1 of 1') && r.includes('2 categories x 2 series') && r.includes('value axis min/max/major unit derived by the host'))).toBe(true)
    // Calibri is not in the fixture manifest: the loaded DejaVu face substitutes and the substitution is disclosed.
    expect(paint.reasons.some(r => r.startsWith('docx.approximate-chart-substituted-font: Calibri / 400 / normal -> DejaVu Sans'))).toBe(true)
    // The painted chart is no longer an omitted drawing.
    expect(paint.omitted_content.some(entry => entry.code === 'PICTURE_GRAPHIC_REQUIRED')).toBe(false)
    expect(decodeNativeDocxApproximatePagePreviewV1(paint).ok).toBe(true)
  }, 20000)

  it('paints anchored charts at their resolved page position, keeps omitted charts disclosed and refuses degenerate models', async () => {
    const { input, eligibility, charts, chart } = chartInput({ placement: 'anchored', page_anchor: pageAnchor(), wrap: 'none', chart: chartModel({ legend: undefined, title: undefined, bar_direction: 'bar', category_axis: { deleted: true, orientation: 'minMax', number_format: 'General' } }) })
    charts.items.push({ ...chart, id: 'approximate-drawing-chart:test:2', status: 'omitted', reason: 'unsupported-chart-type:pieChart', chart: undefined, chart_part: undefined, chart_part_sha256: undefined, placement: undefined, page_anchor: undefined, wrap: undefined } as never)
    const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input), { drawingCharts: wire(charts) })
    expect(paint.status).toBe('painted')
    const commands = paint.pages[0]!.commands
    const area = commands.find(c => c.kind === 'fill_table_cell' && c.cell_id === 'area')
    expect(area).toMatchObject({ table_id: 'docx.approximate-drawing-chart-preview-v1', row_id: CHART_ID, x_millipoints: 72_000, y_millipoints: 144_000, width_millipoints: 216_000, height_millipoints: 144_000 })
    const bars = commands.filter((c): c is Extract<typeof c, { kind: 'fill_table_cell' }> => c.kind === 'fill_table_cell' && c.cell_id.startsWith('bar:'))
    expect(bars).toHaveLength(4)
    // Horizontal bars grow to the right from one shared value baseline.
    expect(new Set(bars.map(c => c.x_millipoints)).size).toBe(1)
    const first = bars.filter(c => c.cell_id.startsWith('bar:0:'))
    expect(Math.abs(first[0]!.width_millipoints - 2 * first[1]!.width_millipoints)).toBeLessThanOrEqual(2)
    expect(commands.some(c => c.kind === 'stroke_table_border' && c.cell_id === 'category-axis')).toBe(false)
    expect(commands.some(c => c.kind === 'fill_table_cell' && c.cell_id.startsWith('legend:'))).toBe(false)
    // The anchored chart paints in front of body text.
    expect(commands.indexOf(area!)).toBeGreaterThan(commands.findIndex(c => c.kind === 'fill_glyph_path' && c.source_id === 'run:1'))
    expect(paint.reasons.some(r => r.startsWith('docx.approximate-drawing-chart-omitted:') && r.includes('approximate-drawing-chart:test:2 (unsupported-chart-type:pieChart)'))).toBe(true)
    expect(decodeNativeDocxApproximatePagePreviewV1(paint).ok).toBe(true)
  }, 20000)

  it('discloses charts the painter drops as omitted content', async () => {
    const { input, eligibility, charts } = chartInput({ placement: 'anchored', wrap: 'none', page_anchor: { ...pageAnchor(), x_emu: 120_000_000, y_emu: 120_000_000 } })
    const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input), { drawingCharts: wire(charts) })
    expect(paint.status).toBe('painted')
    expect(paint.pages[0]!.commands.some(c => c.kind === 'fill_table_cell')).toBe(false)
    expect(paint.reasons.some(r => r.startsWith('docx.approximate-drawing-chart-omitted:') && r.includes('outside-page')), paint.reasons.filter(r => r.includes('approximate-drawing-chart')).join(' / ')).toBe(true)
    expect(paint.content_status).toBe('partial')
    expect(paint.omitted_content.some(entry => entry.code === 'PICTURE_GRAPHIC_REQUIRED' && entry.scope_id === 'paragraph:1')).toBe(true)
    expect(decodeNativeDocxApproximatePagePreviewV1(paint).ok).toBe(true)
  }, 20000)

  it('keeps a dropped shape disclosed when both sidecars are attached, and omits a chart whose authored scale is not finite', async () => {
    // A chart that paints plus an anchored shape the shape painter drops (outside the page).
    const { input, eligibility, charts, paragraph } = chartInput({})
    paragraph.anchor = anchor('/w:document[1]/w:body[1]/w:p[1]', 100, 300)
    const document = input.document as NativeDocxDocumentV1
    document.unsupported.push({ id: 'unsupported:shape', code: 'UNMODELED_RUN_CONTENT', capability: 'runs', scope_id: paragraph.id, anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[3]/mc:AlternateContent[1]', 203, 207), preservation: 'refuse-mutation', message: 'Run content outside text, controls, and native references is preserved verbatim' })
    const shapes = {
      protocol: 'injoffice.docx.approximate-drawing-shapes', version: 1, policy: 'docx.approximate-drawing-shape-preview-v1', package_sha256: HASH, part_sha256: HASH, omitted_count: 0,
      items: [{
        id: 'approximate-drawing-shape:test:1', paragraph_id: paragraph.id, diagnostic_ids: ['unsupported:shape'], anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[3]/mc:AlternateContent[1]/mc:Choice[1]/w:drawing[1]', 201, 209), run_anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[3]', 200, 210),
        status: 'supported', placement: 'anchored', preset: 'rect', fill_rgb: '4F81BD', width_emu: 914400, height_emu: 457200, rotation_degrees: 0, flip_horizontal: false, flip_vertical: false, wrap: 'none',
        page_anchor: { policy: 'relative-position-no-wrap-v2', source_anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[3]/mc:AlternateContent[1]/mc:Choice[1]/w:drawing[1]/wp:anchor[1]', 202, 208), horizontal_anchor: anchor('/h', 203, 204), vertical_anchor: anchor('/v', 204, 205), x_emu: 120_000_000, y_emu: 120_000_000, horizontal_relative: 'page', vertical_relative: 'page', stacking: { behind_doc: false, relative_height: 7 } },
      }],
    }
    const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input), { drawingShapes: wire(shapes), drawingCharts: wire(charts) })
    expect(paint.status).toBe('painted')
    const commands = paint.pages[0]!.commands
    expect(commands.some(c => c.kind === 'fill_table_cell' && c.table_id === 'docx.approximate-drawing-chart-preview-v1' && c.cell_id === 'area')).toBe(true)
    expect(commands.some(c => c.kind === 'fill_table_cell' && c.table_id === 'docx.approximate-drawing-shape-preview-v1')).toBe(false)
    expect(paint.reasons.some(r => r.startsWith('docx.approximate-drawing-shape-omitted:') && r.includes('outside-page'))).toBe(true)
    // The chart recompute must not erase the shape's restored refusal.
    expect(paint.omitted_content.some(entry => entry.code === 'UNMODELED_RUN_CONTENT' && entry.scope_id === paragraph.id)).toBe(true)
    expect(paint.omitted_content.some(entry => entry.code === 'PICTURE_GRAPHIC_REQUIRED')).toBe(false)
    expect(paint.content_status).toBe('partial')
    expect(decodeNativeDocxApproximatePagePreviewV1(paint).ok).toBe(true)
    // Authored bounds that are individually finite but span an infinite range omit only that chart.
    const infinite = chartInput({ chart: chartModel({ value_axis: { deleted: false, orientation: 'minMax', labels: font(900), number_format: 'General', min: '-1e308', max: '1e308' } }) })
    const painted = await renderNativeDocxApproximatePagePreviewV1(infinite.input, infinite.eligibility, outlineProvider(infinite.input), { drawingCharts: wire(infinite.charts) })
    expect(painted.status).toBe('painted')
    expect(painted.pages[0]!.commands.some(c => c.kind === 'fill_table_cell')).toBe(false)
    expect(painted.reasons.some(r => r.startsWith('docx.approximate-drawing-chart-omitted:') && r.includes('paint-failed: chart value scale is not finite'))).toBe(true)
    expect(painted.omitted_content.some(entry => entry.code === 'PICTURE_GRAPHIC_REQUIRED')).toBe(true)
    expect(painted.content_status).toBe('partial')
    expect(decodeNativeDocxApproximatePagePreviewV1(painted).ok).toBe(true)
  }, 30000)

  it('refuses sidecars that do not exact-join the source document or carry invalid models, and stays opt-in', async () => {
    const cases: Array<[string, (charts: any, chart: any) => void]> = [
      ['package', (charts) => { charts.package_sha256 = `sha256:${'b'.repeat(64)}` }],
      ['paragraph', (_c, chart) => { chart.paragraph_id = 'paragraph:missing' }],
      ['diagnostic', (_c, chart) => { chart.diagnostic_ids = ['unsupported:other'] }],
      ['overlap', (_c, chart) => { chart.run_anchor = anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]', 105, 185); chart.anchor = anchor('/x', 110, 180) }],
      ['unknown field', (charts) => { charts.extra = true }],
      ['value lexeme', (_c, chart) => { chart.chart.series[0].values[0] = '4,3' }],
      ['control character in category', (_c, chart) => { chart.chart.categories[0] = 'Al\u0001pha' }],
      ['carriage return in series title', (_c, chart) => { chart.chart.series[0].title = 'First\r' }],
      ['series length', (_c, chart) => { chart.chart.series[1].values = ['1'] }],
      ['fill', (_c, chart) => { chart.chart.series[0].fill_rgb = 'blue' }],
      ['grouping', (_c, chart) => { chart.chart.grouping = 'stacked' }],
      ['omitted with model', (_c, chart) => { chart.status = 'omitted'; chart.reason = 'x' }],
      ['inline anchor', (_c, chart) => { chart.page_anchor = pageAnchor() }],
    ]
    for (const [name, mutate] of cases) {
      const { input, eligibility, charts, chart } = chartInput({})
      mutate(charts, chart)
      expect(() => decodeNativeDocxApproximateDrawingChartsV1(wire(charts), input.document as NativeDocxDocumentV1), name).toThrow()
      // The body preview never depends on the sidecar: the evidence is dropped as a whole and disclosed.
      const paint = await renderNativeDocxApproximatePagePreviewV1(input, eligibility, outlineProvider(input), { drawingCharts: wire(charts) })
      expect(paint.status, name).toBe('painted')
      expect(paint.pages[0]!.commands.some(c => c.kind === 'fill_table_cell'), name).toBe(false)
      expect(paint.reasons, name).toContain(DOCX_APPROXIMATE_DRAWING_CHART_SIDECAR_REFUSED)
      expect(paint.omitted_content.some(entry => entry.code === 'PICTURE_GRAPHIC_REQUIRED'), name).toBe(true)
    }
    const ineligible = chartInput({})
    await expect(renderNativeDocxApproximatePagePreviewV1(ineligible.input, { ...ineligible.eligibility, status: 'ineligible' }, outlineProvider(ineligible.input), { drawingCharts: wire(ineligible.charts) })).rejects.toThrow()
    // Without the sidecar the same document paints no chart and keeps disclosing the refused drawing.
    const plain = chartInput({})
    const paint = await renderNativeDocxApproximatePagePreviewV1(plain.input, plain.eligibility, outlineProvider(plain.input))
    expect(paint.status).toBe('painted')
    expect(paint.pages[0]!.commands.some(c => c.kind === 'fill_table_cell' && c.table_id === 'docx.approximate-drawing-chart-preview-v1')).toBe(false)
    expect(paint.omitted_content.some(entry => entry.code === 'PICTURE_GRAPHIC_REQUIRED')).toBe(true)
    expect(paint.reasons.some(r => r.startsWith('docx.approximate-drawing-chart'))).toBe(false)
    // Strict preparation never reads the sidecar and keeps the source refusal untouched.
    const strict = chartInput({})
    const prepared = await prepareNativeDocxPagePaintV1(strict.input)
    expect(prepared.page_paint_request.pagination_request.document.unsupported.map(entry => entry.id)).toEqual(['unsupported:chart'])
    expect(JSON.stringify(prepared)).not.toContain('docx.approximate-drawing-chart')
  }, 30000)
})
