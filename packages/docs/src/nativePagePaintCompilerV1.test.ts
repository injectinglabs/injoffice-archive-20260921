import { readFileSync } from 'node:fs'
import { createHash } from 'node:crypto'
import { createRequire } from 'node:module'
import { describe, expect, it } from 'vitest'
import type { NativeFontManifest, NativeFontResolver, ResolvedFontFace } from '@injoffice/font-metrics/layout'
import { createHarfBuzzTextShaperV1, createHarfBuzzOutlineProviderV1, inspectHarfBuzzFontMetricsV1 } from '@injoffice/font-metrics/harfbuzz'
import { reorderNativeBidiLineV1 } from '@injoffice/font-metrics/bidi'
import { DOCX_NATIVE_PROTOCOL, DOCX_NATIVE_VERSION, type NativeDocxDocumentV1, type NativeDocxRunV1 } from './nativeContract.js'
import { DOCX_RESOLVED_LAYOUT_PROTOCOL, DOCX_RESOLVED_LAYOUT_VERSION, type NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'
import { DOCX_PAGINATION_SETTINGS_PROTOCOL, DOCX_PAGINATION_SETTINGS_VERSION, type NativeDocxPaginationSettingsV1 } from './nativePaginationSettings.js'
import {
  DOCX_PAGE_PAINT_COMPILER_PROTOCOL,
  DOCX_PAGE_PAINT_COMPILER_VERSION,
  completeNativeDocxPagePaintV1,
  prepareNativeDocxPagePaintV1,
  renderNativeDocxApproximatePagePreviewV1,
  decodeNativeDocxApproximatePagePreviewV1,
  type NativeDocxPagePaintPrepareInputV1,
} from './nativePagePaintCompilerV1.js'
import { encodeNativeDOCXFontInventoryV1, nativeDOCXCanonicalWireSHA256V1, type NativeDOCXFontInventoryV1 } from './nativeFontInventoryV1.js'
import { decodeNativeDocxPagePaintResourceListV1, qualifyNativeDocxInlineImageV1 } from './nativeImagePagePaintV1.js'
import { paginateNativeDocxV1, paginateNativeDocxApproximateLegacyV1 } from './nativePaginationV1.js'
import { qualifyNativeDocxTablesV1 } from './nativeTablePagePaintV1.js'
import { decodeNativeDocxShapedLines } from './nativeShapedLinesContract.js'
import { decodeNativeDocxPagePaintForRequestV1, decodeNativeDocxPagePaintRequestV1, nativeDocxPagePaintShapedLinesSha256V1, decodeNativeDocxApproximateComputedPagePaintV1, nativeDocxPagePaintPaginatedLayoutSha256V1 } from './nativePagePaintV1.js'
import { nativeDocxPageFieldDocumentV1 } from './nativePageFieldsV1.js'
import {deriveNativeSquareWrapPlanV1} from './nativeSquareWrapV1.js'

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
describe('native DOCX page-paint compiler v1', () => {
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

  it('renders a non-embedded-font document with explicit content-addressed host fonts', async () => {
    const { input, fonts } = hostFixture()
    const before = JSON.stringify(input)
    const prepared = await prepareNativeDocxPagePaintV1(input, { fonts })
    expect(prepared.outline_requests.length).toBeGreaterThan(0)
    expect(prepared.providers.resolver_id).toBe('test.host-fonts')
    expect(JSON.stringify(input)).toBe(before)
    await expect(prepareNativeDocxPagePaintV1(input)).rejects.toThrow(/configure explicit host fonts/)
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
    d.x_emu=914400;d.width_emu=5943600;expect(()=>deriveNativeSquareWrapPlanV1(p.document,p.resolved_layout,p.shaped_lines,request.paginated_layout)).toThrow('fully blocks')
  })
  it('explicitly refuses square wrapping combined with fixed or autofit table flow', async () => {
    for(const layout of ['fixed','autofit'] as const){
      const input=combinedImageTableFixture(),document=input.document as NativeDocxDocumentV1
      document.body.blocks[1]!.table!.layout=layout
      Object.assign(document.body.blocks[0]!.paragraph!.runs[0]!.drawing!,{placement:'floating',x_emu:914400,y_emu:914400,width_emu:1270000,height_emu:635000,horizontal_relative_from:'page',vertical_relative_from:'page',wrap:'square',floating_layer:'front',stacking_order:7})
      await expect(prepareNativeDocxPagePaintV1(input)).rejects.toThrow('without tables or notes')
    }
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
    for (const mutation of [ { wrap: 'tight' }, { horizontal_relative_from: 'column' }, { floating_layer: undefined }, { x_emu: -127 }, { x_emu: 1 }, { y_emu: 127000000 } ]) {
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
    await expect(prepareNativeDocxPagePaintV1(duplicate)).rejects.toThrow(/exactly cover/)

    const vector = imageFixture()
    ;(vector.document as NativeDocxDocumentV1).body.blocks[0]!.paragraph!.runs[0]!.drawing!.content_type = 'image/svg+xml'
    ;(vector.document as NativeDocxDocumentV1).passthrough_parts.at(-1)!.content_type = 'image/svg+xml'
    vector.media_assets = []
    const vectorPrepared = await prepareNativeDocxPagePaintV1(vector)
    expect(vectorPrepared.page_paint_request.paginated_layout).toMatchObject({ status: 'refused', pages: [] })
    expect(vectorPrepared.page_paint_request.media_assets).toEqual([])
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
  it('autofits unequal text columns from real font advances and replays the source-bound width policy', async () => {
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
    for(const cell of table.rows[0]!.cells)delete cell.width_twips
    const auto=await prepareNativeDocxPagePaintV1(input)
    expect(auto.page_paint_request.paginated_layout.status).toBe('paginated')
    const automatic=qualifyNativeDocxTablesV1(document,resolved,auto.page_paint_request.pagination_request.shaped_lines)
    expect(automatic.status).toBe('qualified')
    if(automatic.status==='qualified') {
      expect(automatic.tables[0]!.width_millipoints).toBeLessThanOrEqual(468_000)
      expect(automatic.tables[0]!.grid_widths_millipoints[0]).toBeLessThan(automatic.tables[0]!.grid_widths_millipoints[1]!)
    }
  })
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
    const refused = await prepareNativeDocxPagePaintV1(requiresContinuation)
    expect(refused.page_paint_request.paginated_layout).toEqual(expect.objectContaining({ status: 'refused', pages: [], sections: [], diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'note-overflow-unsupported' })]) }))
    expect(refused.outline_requests).toEqual([])
    const completed = await completeNativeDocxPagePaintV1({ prepared: refused, outline_results: [] })
    expect(completed.page_paint_output).toEqual(expect.objectContaining({ status: 'refused', pages: [] }))
  })
})
