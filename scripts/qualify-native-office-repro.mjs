import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { spawnSync } from 'node:child_process'
import {
  DOCX_DEFAULT_TAB_STOP_TWIPS,
  DOCX_NATIVE_LIMITS,
  DOCX_PAGINATION_REQUEST_PROTOCOL,
  DOCX_PAGE_PAINT_REQUEST_PROTOCOL,
  DOCX_RESOLVED_LAYOUT_PROTOCOL,
  DOCX_SHAPING_REQUEST_PROTOCOL,
  compileNativeDocxPagePaintV1,
  decodeNativeDocxDocument,
  decodeNativeDocxJson,
  encodeNativeDocxDocument,
  nativeDocxPagePaintFontManifestSha256V1,
  nativeDocxPagePaintMediaAssetsSha256V1,
  nativeDocxPagePaintPaginatedLayoutSha256V1,
  nativeDocxPagePaintShapedLinesSha256V1,
  nativeDocxTableProjectionSha256V1,
  paginateNativeDocxV1,
  prepareNativeDocxPagePaintMediaAssetsV1,
  qualifyNativeDocxTablesV1,
  shapeNativeDocxLinesV1,
} from '@injoffice/docs'
import { compileWireDeckToNativeV1 } from '@injoffice/pptx-authored'
import { PPTX_NATIVE_RESOURCE_LIMITS, stringifyNativePptx, validateNativePptx } from '@injoffice/pptx-native'
import { compileNativePptxSlide, createRecordingPaintSurface, paintSlideRenderTree, stringifySlideRenderTree } from '@injoffice/pptx-render'
import {
  NativeSheetGeometryError,
  compileNativeSheetGeometryV1,
  createNativeSheetGeometryRecordingSurfaceV1,
  emitNativeSheetGeometryCommandsV1,
  projectNativeWorkbookV1,
  validateNativeWorkbookV1,
} from '@injoffice/sheets'

const root = resolve(import.meta.dirname, '..')
const manifestPath = resolve(root, 'testdata/native-office-cross-runtime/v1/manifest.json')
const sha256 = (value) => createHash('sha256').update(value).digest('hex')
const json = (path) => JSON.parse(readFileSync(resolve(root, path), 'utf8'))
const digestJson = (value) => sha256(JSON.stringify(value))
const HASH_A = `sha256:${'a'.repeat(64)}`
const HASH_B = `sha256:${'b'.repeat(64)}`
const PNG_BYTES = new Uint8Array(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=', 'base64'))
const PNG_DIGEST = `sha256:${sha256(PNG_BYTES)}`
const FONT_DIGEST = 'sha256:054edec1d0211f624fed0cbca9d4f9400b0e491c43742af2c5b0abebf0c990d8'

function reverseObjectKeys(value) {
  if (Array.isArray(value)) return value.map(reverseObjectKeys)
  if (!value || typeof value !== 'object') return value
  return Object.fromEntries(Object.keys(value).reverse().map((key) => [key, reverseObjectKeys(value[key])]))
}

function requireOk(result, label) {
  if (!result.ok) throw new Error(`${label}: ${JSON.stringify(result)}`)
  return result.value
}

function exactError(run, label) {
  try {
    run()
  } catch (error) {
    return { name: error?.name ?? 'Error', code: error?.code ?? null }
  }
  throw new Error(`${label} did not refuse`)
}

function nativeDocxPipelineFixture() {
  const document = json('testdata/docx-native/document-v1.json')
  document.document_id = 'doc:cross-runtime'
  document.revision = 'rev:cross-runtime:1'
  document.headers = []
  document.footers = []
  document.notes = []
  document.comment_stories = []
  document.comments = []
  document.capabilities = []
  document.unsupported = []
  document.sections = [document.sections[0]]
  document.sections[0].header_refs = []
  document.sections[0].footer_refs = []
  document.sections[0].page = {
    width_twips: 1_000,
    height_twips: 1_200,
    orientation: 'portrait',
    margins: { top_twips: 100, right_twips: 100, bottom_twips: 100, left_twips: 100, header_twips: 50, footer_twips: 50, gutter_twips: 0 },
    columns: 1,
    column_spacing_twips: 100,
    column_layout: 'equal-width',
    column_definitions: [{ id: 'column:section:1:0', ordinal: 0 }],
  }

  const paragraph = document.body.blocks[0].paragraph
  const textRun = paragraph.runs[0]
  textRun.text = 'Café Ω'
  textRun.properties = { font_family: 'Fixture Sans', font_size_half_points: 20, color: '123456', language: 'en-US' }
  const drawingRun = paragraph.runs[2]
  drawingRun.id = 'run:intro:image'
  drawingRun.drawing.id = 'drawing:inline'
  drawingRun.drawing.placement = 'inline'
  drawingRun.drawing.width_emu = 127_000
  drawingRun.drawing.height_emu = 127_000
  drawingRun.drawing.media_part = 'word/media/image1.png'
  drawingRun.drawing.content_type = 'image/png'
  for (const key of ['x_emu', 'y_emu', 'horizontal_relative_from', 'vertical_relative_from', 'wrap']) delete drawingRun.drawing[key]
  paragraph.runs = [textRun, drawingRun]

  const table = document.body.blocks[1].table
  delete table.table_style_id
  table.width_twips = 800
  table.layout = 'fixed'
  table.alignment = 'left'
  table.indent_twips = 0
  table.grid_widths_twips = [800]
  table.cell_margins = { top_twips: 10, right_twips: 10, bottom_twips: 10, left_twips: 10 }
  table.borders = {
    top: { style: 'single', size_eighth_points: 8, color_rgb: '112233' },
    right: { style: 'single', size_eighth_points: 8, color_rgb: '112233' },
    bottom: { style: 'single', size_eighth_points: 8, color_rgb: '112233' },
    left: { style: 'single', size_eighth_points: 8, color_rgb: '112233' },
  }
  const row = table.rows[0]
  delete row.height_twips
  delete row.height_rule
  row.repeat_header = false
  row.cant_split = true
  const cell = row.cells[0]
  cell.width_twips = 800
  cell.shading_rgb = 'DDEEFF'
  const cellParagraph = cell.paragraphs[0]
  cellParagraph.runs[0].properties = { font_family: 'Fixture Sans', font_size_half_points: 20, color: '445566', language: 'en-US' }

  document.sections[0].starts_at_block_id = paragraph.id
  document.passthrough_parts = [
    { part_name: 'word/media/image1.png', content_type: 'image/png', byte_length: PNG_BYTES.byteLength, sha256: PNG_DIGEST, policy: 'preserve-verbatim' },
    { part_name: 'word/settings.xml', content_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml', byte_length: 1, sha256: HASH_A, policy: 'preserve-verbatim' },
    { part_name: 'word/_rels/document.xml.rels', content_type: 'application/vnd.openxmlformats-package.relationships+xml', byte_length: 1, sha256: HASH_B, policy: 'preserve-verbatim' },
  ]

  const resolvedLayout = {
    protocol: DOCX_RESOLVED_LAYOUT_PROTOCOL,
    version: 1,
    document_id: document.document_id,
    revision: document.revision,
    source_parts: { main_part: document.source.main_part },
    paragraphs: [paragraph, cellParagraph].map((entry) => ({
      paragraph_id: entry.id,
      applied_styles: [],
      properties: {},
      paragraph_mark_properties: { font_family: 'Fixture Sans', font_size_half_points: 20 },
    })),
    runs: [textRun, drawingRun, cellParagraph.runs[0]].map((entry) => ({
      run_id: entry.id,
      paragraph_id: entry === cellParagraph.runs[0] ? cellParagraph.id : paragraph.id,
      applied_paragraph_styles: [],
      applied_character_styles: [],
      properties: entry.properties ?? {},
    })),
    tables: [{ table_id: table.id }],
    fonts: [{ name: 'Fixture Sans' }],
    diagnostics: [],
  }
  const fontManifest = {
    version: 1,
    manifestId: 'fixture.cross-runtime',
    revision: '1',
    faces: [{ faceId: 'fixture.regular', family: 'Fixture Sans', weight: 400, style: 'normal', stretch: 100, source: { kind: 'bundled', resourceId: 'fixture-font', contentDigest: FONT_DIGEST } }],
    fallbackChains: [{ chainId: 'fixture.default', faceIds: ['fixture.regular'] }],
  }
  return { document, resolvedLayout, fontManifest }
}

function textProviders() {
  const face = {
    faceId: 'fixture.regular', family: 'Fixture Sans', weight: 400, style: 'normal', stretch: 100,
    sourceKind: 'bundled', resourceId: 'fixture-font', contentDigest: FONT_DIGEST,
    resolution: 'exact', matchedFamily: 'Fixture Sans',
  }
  return {
    resolver: {
      providerId: 'fixture-resolver', providerRevision: '1',
      resolve() { return { status: 'resolved', face, attemptedFaceIds: [face.faceId], decisions: [] } },
      load() { return { face, bytes: new Uint8Array([0, 1, 2, 3]), metrics: { unitsPerEm: 1_000, ascender: 800, descender: -200, lineGap: 200 } } },
    },
    shaper: {
      providerId: 'fixture-shaper', providerRevision: '1',
      shape({ run }) {
        const characters = [...run.text]
        let offset = 0
        const clusters = characters.map((character, index) => {
          const startUtf16 = offset
          offset += character.length
          return { startUtf16, endUtf16: offset, glyphStart: index, glyphEnd: index + 1, advanceInlineMilliPoints: 1_000, whitespace: /^\s$/u.test(character) }
        })
        return {
          startUtf16: 0, endUtf16: run.text.length, face,
          glyphs: clusters.map((_cluster, index) => ({ glyphId: index + 1, clusterIndex: index, advanceXMilliPoints: 1_000, advanceYMilliPoints: 0, offsetXMilliPoints: 0, offsetYMilliPoints: 0 })),
          clusters,
          metrics: { fontSizeMilliPoints: run.fontSizeMilliPoints, ascentMilliPoints: 8_000, descentMilliPoints: -2_000, lineGapMilliPoints: 2_000, lineHeightMilliPoints: 12_000 },
          advanceInlineMilliPoints: clusters.length * 1_000,
          advanceBlockMilliPoints: 0,
        }
      },
    },
  }
}

function paginationSettings(document) {
  return {
    protocol: 'injoffice.docx.pagination-settings', version: 1,
    document_id: document.document_id, revision: document.revision,
    package_sha256: document.source.package_sha256, main_part: document.source.main_part,
    relationships_part: 'word/_rels/document.xml.rels', relationships_sha256: HASH_B,
    relationship_id: 'rIdSettings', settings_part: 'word/settings.xml', settings_sha256: HASH_A,
    profile: 'word-modern-default', default_tab_stop_twips: DOCX_DEFAULT_TAB_STOP_TWIPS,
    mirror_margins: false, gutter_at_top: false, even_and_odd_headers: false, compatibility_mode: 15, diagnostics: [],
  }
}

async function qualifyDocx() {
  const source = json('testdata/docx-native/document-v1.json')
  const decoded = requireOk(decodeNativeDocxDocument(source), 'DOCX extraction fixture')
  const canonical = encodeNativeDocxDocument(decoded)
  assert.equal(encodeNativeDocxDocument(reverseObjectKeys(source)), canonical)

  const nfc = structuredClone(decoded)
  const nfd = structuredClone(decoded)
  nfc.body.blocks[0].paragraph.runs[0].text = 'Café Å'
  nfd.body.blocks[0].paragraph.runs[0].text = 'Cafe\u0301 A\u030A'
  const nfcWire = encodeNativeDocxDocument(nfc)
  const nfdWire = encodeNativeDocxDocument(nfd)
  assert.notEqual(nfcWire, nfdWire)
  assert.equal(nfdWire.includes('Cafe\u0301 A\u030A'), true)

  const { document, resolvedLayout, fontManifest } = nativeDocxPipelineFixture()
  const bodyParagraph = document.body.blocks[0].paragraph
  const tableParagraph = document.body.blocks[1].table.rows[0].cells[0].paragraphs[0]
  const shapingDocument = structuredClone(document)
  shapingDocument.body.blocks = [shapingDocument.body.blocks[0]]
  const shapingResolved = structuredClone(resolvedLayout)
  shapingResolved.paragraphs = shapingResolved.paragraphs.filter((entry) => entry.paragraph_id === bodyParagraph.id)
  shapingResolved.runs = shapingResolved.runs.filter((entry) => entry.paragraph_id === bodyParagraph.id)
  shapingResolved.tables = []
  const shapingRequest = {
    protocol: DOCX_SHAPING_REQUEST_PROTOCOL, version: 1, document: shapingDocument, resolved_layout: shapingResolved,
    font_manifest: fontManifest, available_width_millipoints: 40_000, tab_interval_millipoints: DOCX_DEFAULT_TAB_STOP_TWIPS * 50,
  }
  const shaped = requireOk(await shapeNativeDocxLinesV1(shapingRequest, textProviders()), 'DOCX shaping')
  const cellDocument = structuredClone(document)
  cellDocument.body.blocks = [{ kind: 'paragraph', id: tableParagraph.id, paragraph: structuredClone(tableParagraph) }]
  cellDocument.sections[0].starts_at_block_id = tableParagraph.id
  const cellResolved = structuredClone(resolvedLayout)
  cellResolved.paragraphs = cellResolved.paragraphs.filter((entry) => entry.paragraph_id === tableParagraph.id)
  cellResolved.runs = cellResolved.runs.filter((entry) => entry.paragraph_id === tableParagraph.id)
  cellResolved.tables = []
  const shapedCell = requireOk(await shapeNativeDocxLinesV1({ ...shapingRequest, document: cellDocument, resolved_layout: cellResolved, available_width_millipoints: 39_000 }, textProviders()), 'DOCX table-cell shaping')
  const combinedShaped = { ...shaped, paragraphs: [...shaped.paragraphs, ...shapedCell.paragraphs] }
  const invalidOrder = { ...shapingRequest, Å: true, Z: true, '😀': true, A: true }
  const invalid = await shapeNativeDocxLinesV1(invalidOrder, textProviders())
  assert.equal(invalid.ok, false)
  const unicodeIssuePaths = invalid.issues.filter((issue) => issue.code === 'UNKNOWN_FIELD').map((issue) => issue.path)
  assert.deepEqual(unicodeIssuePaths, ['/A', '/Z', '/Å', '/😀'])

  const paginationRequest = {
    protocol: DOCX_PAGINATION_REQUEST_PROTOCOL, version: 1, document, resolved_layout: resolvedLayout,
    shaped_lines: combinedShaped, pagination_settings: paginationSettings(document),
  }
  const paginated = requireOk(paginateNativeDocxV1(paginationRequest), 'DOCX pagination')
  if (paginated.status !== 'paginated') throw new Error(`DOCX pagination refused: ${JSON.stringify(paginated)}`)
  const tableProjection = qualifyNativeDocxTablesV1(document, resolvedLayout)
  assert.equal(tableProjection.status, 'qualified')
  const mediaAssets = prepareNativeDocxPagePaintMediaAssetsV1(document, [{ part_name: 'word/media/image1.png', content_type: 'image/png', content_digest: PNG_DIGEST, bytes: PNG_BYTES }])
  const pagePaintRequest = {
    protocol: DOCX_PAGE_PAINT_REQUEST_PROTOCOL, version: 1, pagination_request: paginationRequest,
    paginated_layout: paginated, font_manifest: fontManifest, media_assets: mediaAssets,
    integrity: {
      font_manifest_sha256: nativeDocxPagePaintFontManifestSha256V1(fontManifest),
      shaped_lines_sha256: nativeDocxPagePaintShapedLinesSha256V1(combinedShaped),
      paginated_layout_sha256: nativeDocxPagePaintPaginatedLayoutSha256V1(paginated),
      table_projection_sha256: tableProjection.sha256,
      media_assets_sha256: nativeDocxPagePaintMediaAssetsSha256V1(mediaAssets),
    },
    outline_provider: { provider_id: 'fixture-outline', provider_revision: '1' },
  }
  const outlineProvider = {
    providerId: 'fixture-outline', providerRevision: '1',
    getGlyphOutline(request) {
      return { status: 'outlined', face: request.face, glyph_id: request.glyph_id, units_per_em: 1_000, path: [{ kind: 'move_to', x: 0, y: 0 }, { kind: 'line_to', x: 500, y: 0 }, { kind: 'line_to', x: 500, y: 700 }, { kind: 'close_path' }] }
    },
  }
  const painted = requireOk(await compileNativeDocxPagePaintV1(pagePaintRequest, outlineProvider), 'DOCX page paint')
  assert.equal(painted.status, 'painted')
  const refusalProvider = {
    providerId: 'fixture-outline', providerRevision: '1',
    getGlyphOutline(request) { return { status: 'refused', face: request.face, glyph_id: request.glyph_id, code: 'missing-glyph', message: 'qualification refusal' } },
  }
  const refused = requireOk(await compileNativeDocxPagePaintV1(pagePaintRequest, refusalProvider), 'DOCX page-paint refusal')
  assert.deepEqual({ status: refused.status, resources: refused.resources, pages: refused.pages }, { status: 'refused', resources: [], pages: [] })
  const overLimit = decodeNativeDocxJson(' '.repeat(DOCX_NATIVE_LIMITS.maxJsonBytes + 1))
  assert.deepEqual(overLimit, { ok: false, issues: [{ code: 'LIMIT_EXCEEDED', path: '', message: `JSON payload exceeds ${DOCX_NATIVE_LIMITS.maxJsonBytes} bytes` }] })

  return {
    extraction: { canonical_sha256: sha256(canonical), reversed_keys_equal: true },
    unicode: { nfc_sha256: sha256(nfcWire), nfd_sha256: sha256(nfdWire), exact_distinct: true },
    shaping_sha256: digestJson(combinedShaped),
    unicode_issue_paths: unicodeIssuePaths,
    pagination_sha256: digestJson(paginated),
    page_paint: {
      request_sha256: digestJson(pagePaintRequest),
      output_sha256: digestJson(painted),
      command_kinds: [...new Set(painted.pages.flatMap((page) => page.commands.map((command) => command.kind)))].sort(),
      refusal: { status: refused.status, resources: refused.resources.length, pages: refused.pages.length },
    },
    resource_refusal: overLimit.issues.map(({ code, path }) => ({ code, path })),
  }
}

function pptxTextLayout() {
  const providers = textProviders()
  return {
    manifest: nativeDocxPipelineFixture().fontManifest,
    resolver: providers.resolver,
    shaper: providers.shaper,
    defaults: { fontFamilies: ['Fixture Sans'], fontSizeHundredthPt: 1_000, script: 'Latn', language: 'en-US', direction: 'ltr', fallbackChainIds: ['fixture.default'] },
  }
}

async function qualifyPptx() {
  const imported = json('go/pptxpatch/testdata/native-contract/valid/parsed-full.json')
  requireOk(validateNativePptx(imported), 'imported PPTX')
  const importedWire = stringifyNativePptx(imported)
  assert.equal(stringifyNativePptx(reverseObjectKeys(imported)), importedWire)
  const importedTree = await compileNativePptxSlide(imported, 0, { textLayout: pptxTextLayout() })
  const importedTreeWire = stringifySlideRenderTree(importedTree)

  const authoredInput = { slides: [{ background: '#FFFFFF', shapes: [
    { kind: 'rect', key: 'back', x: 10, y: 20, cx: 300, cy: 200, fill: '#112233' },
    { kind: 'textBox', key: 'title', x: 40, y: 50, cx: 1_000_000, cy: 300_000, paragraphs: [{ runs: [{ text: 'Café Å', sizePt: 12, color: '#AABBCC', font: 'Fixture Sans' }], align: 'l' }] },
  ] }] }
  const authored = compileWireDeckToNativeV1(authoredInput)
  if (!authored.ok) throw new Error(`authored PPTX: ${JSON.stringify(authored)}`)
  const authoredWire = stringifyNativePptx(authored.deck)
  const reorderedAuthored = compileWireDeckToNativeV1(reverseObjectKeys(authoredInput))
  if (!reorderedAuthored.ok) throw new Error(`reordered authored PPTX: ${JSON.stringify(reorderedAuthored)}`)
  assert.equal(stringifyNativePptx(reorderedAuthored.deck), authoredWire)

  const refusalA = compileWireDeckToNativeV1({ slides: [{ background: '#FFFFFF', shapes: [], Å: 1, Z: 1, '😀': 1, A: 1 }] })
  const refusalB = compileWireDeckToNativeV1({ slides: [{ A: 1, '😀': 1, Z: 1, Å: 1, shapes: [], background: '#FFFFFF' }] })
  assert.equal(refusalA.ok, false)
  assert.deepEqual(refusalB, refusalA)
  assert.equal(Object.hasOwn(refusalA, 'deck'), false)

  const nativeOrderA = { ...structuredClone(imported), Å: true, Z: true, '😀': true, A: true }
  const nativeOrderB = { A: true, '😀': true, Z: true, Å: true, ...structuredClone(imported) }
  const nativeRefusalA = validateNativePptx(nativeOrderA)
  const nativeRefusalB = validateNativePptx(nativeOrderB)
  assert.deepEqual(nativeRefusalB, nativeRefusalA)

  const negativeZero = validateNativePptx(json('go/pptxpatch/testdata/native-contract/invalid/negative-zero.json'))
  assert.equal(negativeZero.ok, false)
  const depthDeck = structuredClone(imported)
  let nested = { kind: 'connector', id: 'leaf', provenance: 'authored', transform: { x: 0, y: 0, cx: 1, cy: 1 }, passthrough: [], compatibility: { status: 'editable', diagnostics: [] } }
  for (let depth = 0; depth <= PPTX_NATIVE_RESOURCE_LIMITS.maxDepth; depth++) nested = { kind: 'group', id: `group-${depth}`, provenance: 'authored', transform: { x: 0, y: 0, cx: 1, cy: 1 }, children: [nested], passthrough: [], compatibility: { status: 'editable', diagnostics: [] } }
  depthDeck.slides.push({ id: 'slide-depth', provenance: 'authored', elements: [nested], passthrough: [], compatibility: { status: 'editable', diagnostics: [] } })
  const depthRefusal = validateNativePptx(depthDeck)
  assert.equal(depthRefusal.ok, false)

  const atomicSurface = createRecordingPaintSurface()
  const paintRefusal = exactError(() => paintSlideRenderTree(importedTree, atomicSurface, 1), 'PPTX paint budget')
  assert.deepEqual(atomicSurface.commands, [])

  return {
    imported: { canonical_sha256: sha256(importedWire), render_tree_sha256: sha256(importedTreeWire), reversed_keys_equal: true },
    authored: { canonical_sha256: sha256(authoredWire), refusal_sha256: digestJson(refusalA), atomic_no_deck: true },
    json_order_refusal_sha256: digestJson(nativeRefusalA),
    negative_zero_codes: negativeZero.issues.map((issue) => issue.code),
    resource_codes: [...new Set(depthRefusal.issues.map((issue) => issue.code))].sort(),
    paint_budget: { error: paintRefusal, commands: atomicSurface.commands.length },
  }
}

function xlsxMetric(workbook) {
  return {
    source_revision: workbook.revision,
    source_package_sha256: workbook.source.package_sha256,
    normal_style_xf_id: workbook.normal_style.style_xf_id,
    normal_style_font_id: workbook.normal_style.font_id,
    font_name: workbook.normal_style.font_name,
    font_size_points: workbook.normal_style.font_size_points,
    font_bold: workbook.normal_style.font_bold,
    font_italic: workbook.normal_style.font_italic,
    normal_font_record_sha256: workbook.normal_style.font_record_sha256,
    font_sha256: HASH_B,
    provider_id: 'fixture-opentype-parser',
    provider_revision: '1',
    measurement_dpi: 96,
    maximum_digit_width_pixels: 7,
  }
}

function qualifyXlsx() {
  const workbook = json('go/xlsxpatch/testdata/native-xlsx-v1/valid/lexical-render.json')
  workbook.normal_style = { style_xf_id: 0, font_id: 0, font_name: 'Calibri', font_size_points: 11, font_bold: false, font_italic: false, font_record_sha256: `sha256:${'e'.repeat(64)}` }
  requireOk(validateNativeWorkbookV1(workbook), 'XLSX workbook')
  const model = projectNativeWorkbookV1(workbook)
  const viewport = { row: 0, column: 0, end_row: 2, end_column: 2 }
  const metric = xlsxMetric(workbook)
  const geometry = compileNativeSheetGeometryV1(model, '7', viewport, metric)
  const reordered = compileNativeSheetGeometryV1(model, '7', reverseObjectKeys(viewport), reverseObjectKeys(metric))
  assert.equal(JSON.stringify(reordered), JSON.stringify(geometry))

  const unknownA = validateNativeWorkbookV1({ ...structuredClone(workbook), Å: true, Z: true, '😀': true, A: true })
  const unknownB = validateNativeWorkbookV1({ A: true, '😀': true, Z: true, Å: true, ...structuredClone(workbook) })
  assert.deepEqual(unknownB, unknownA)
  const negativeZero = exactError(() => compileNativeSheetGeometryV1(model, '7', { row: -0, column: 0, end_row: 0, end_column: 0 }, metric), 'XLSX negative zero')
  const viewportLimit = exactError(() => compileNativeSheetGeometryV1(model, '7', { row: 0, column: 0, end_row: 999, end_column: 999 }, metric), 'XLSX viewport budget')
  const recording = createNativeSheetGeometryRecordingSurfaceV1(2)
  const commandLimit = exactError(() => emitNativeSheetGeometryCommandsV1(geometry, recording), 'XLSX command budget')
  assert.deepEqual(recording.commands, [])

  return {
    geometry_sha256: geometry.geometry_sha256,
    geometry_json_sha256: digestJson(geometry),
    reversed_keys_equal: true,
    json_order_refusal_sha256: digestJson(unknownA),
    negative_zero: negativeZero,
    viewport_limit: viewportLimit,
    command_limit: { error: commandLimit, commands: recording.commands.length },
  }
}

async function worker() {
  const report = {
    protocol: 'injoffice.native-office-cross-runtime-repro/v1',
    docx: await qualifyDocx(),
    pptx: await qualifyPptx(),
    xlsx: qualifyXlsx(),
  }
  process.stdout.write(`${JSON.stringify(report)}\n`)
}

function attestFixtures(manifest) {
  for (const fixture of manifest.fixtures) {
    const bytes = readFileSync(resolve(root, fixture.path))
    assert.equal(sha256(bytes), fixture.sha256, `stale fixture digest: ${fixture.path}`)
    assert.deepEqual(Object.keys(fixture.provenance).sort(), ['license', 'producer', 'source'])
  }
}

function controller() {
  const manifest = JSON.parse(readFileSync(manifestPath, 'utf8'))
  assert.equal(manifest.protocol, 'injoffice.native-office-cross-runtime-repro/v1')
  attestFixtures(manifest)
  const outputs = manifest.environments.map((environment) => {
    const child = spawnSync(process.execPath, [import.meta.filename, '--worker'], {
      cwd: root,
      env: { ...process.env, ...environment.env },
      encoding: 'utf8',
      maxBuffer: 16 * 1024 * 1024,
    })
    assert.equal(child.status, 0, `${environment.id}: ${child.stderr || child.stdout}`)
    assert.equal(child.stderr, '', `${environment.id} wrote stderr`)
    return child.stdout.trim()
  })
  for (const output of outputs.slice(1)) assert.equal(output, outputs[0], 'semantic report changed across TZ/locale profiles')
  const semanticSha256 = sha256(outputs[0])
  assert.equal(semanticSha256, manifest.expected_semantic_sha256, 'semantic golden changed')
  process.stdout.write(`${JSON.stringify({
    protocol: manifest.protocol,
    node: process.versions.node,
    profiles: manifest.environments.map((entry) => entry.id),
    fixtures: manifest.fixtures.length,
    semantic_sha256: semanticSha256,
  })}\n`)
}

if (process.argv.includes('--worker')) await worker()
else controller()
