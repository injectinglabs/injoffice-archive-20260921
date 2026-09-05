import { createHash } from 'node:crypto'
import { mkdir, readFile, readdir, writeFile } from 'node:fs/promises'
import { dirname, relative, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { performance } from 'node:perf_hooks'

import {
  NATIVE_SHEET_DECORATION_LIMITS,
  NATIVE_SHEET_GEOMETRY_LIMITS,
  compileNativeSheetDecorationsV1,
  compileNativeSheetGeometryV1,
  createNativeSheetDecorationRecordingSurfaceV1,
  createNativeSheetGeometryRecordingSurfaceV1,
  emitNativeSheetDecorationCommandsV1,
  emitNativeSheetGeometryCommandsV1,
  nativeWorkbookStyleRawProjectionSha256V1,
  projectNativeWorkbookV1,
  replayNativeSheetDecorationCommandsV1,
} from '../packages/sheets/dist/index.js'
import {
  compileNativePptxSlide,
  createRecordingPaintSurface,
  paintSlideRenderTree,
  stringifySlideRenderTree,
} from '../packages/pptx-render/dist/index.js'
import {
  DOCX_DEFAULT_TAB_STOP_TWIPS,
  DOCX_PAGE_PAINT_REQUEST_V1_BINDING_FIELDS,
  DOCX_PAGINATION_LIMITS,
  compileNativeDocxPagePaintV1,
  nativeDocxPagePaintFontManifestSha256V1,
  nativeDocxPagePaintMediaAssetsSha256V1,
  nativeDocxPagePaintPaginatedLayoutSha256V1,
  nativeDocxPagePaintShapedLinesSha256V1,
  paginateNativeDocxV1,
} from '../packages/docs/dist/index.js'
import {
  BIDI_UNICODE_VERSION,
  NATIVE_BIDI_PROVIDER_ID,
  NATIVE_BIDI_PROVIDER_REVISION,
} from '../packages/font-metrics/dist/bidi.js'
import { UNICODE_13_CLASSIFIER_REVISION } from '../packages/font-metrics/dist/unicode13.js'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const repetitions = 7
const warmups = 2
const outputFlag = process.argv.indexOf('--output')
const outputPath = resolve(root, outputFlag === -1 ? 'artifacts/native-office-performance-report.json' : process.argv[outputFlag + 1] ?? '')
if (outputFlag !== -1 && !process.argv[outputFlag + 1]) throw new Error('--output requires a path')

const sha256 = (value) => createHash('sha256').update(value).digest('hex')
const canonical = (value) => JSON.stringify(value, (_key, child) => {
  if (child === null || Array.isArray(child) || typeof child !== 'object') return child
  return Object.fromEntries(Object.entries(child).sort(([left], [right]) => left < right ? -1 : left > right ? 1 : 0))
})

function invariant(condition, message) {
  if (!condition) throw new Error(message)
}

function budgetError(name, actual, budget) {
  if (!Number.isSafeInteger(actual) || actual < 0 || !Number.isSafeInteger(budget) || budget <= 0) return `${name}: invalid work counter/budget ${actual}/${budget}`
  return actual > budget ? `${name}: algorithmic/resource regression ${actual} exceeds ${budget}` : undefined
}

function requireAtMost(name, actual, budget) {
  const error = budgetError(name, actual, budget)
  if (error) throw new Error(error)
}

function requireStable(name, values) {
  invariant(values.length > 1, `${name}: repeatability needs at least two samples`)
  const first = values[0]
  invariant(values.every((value) => value === first), `${name}: nondeterministic output across ${values.length} runs`)
  return first
}

async function readCorpusFixture(format) {
  const packagePath = resolve(root, `go/officecompat/corpus/generated/packages/${format}-transitional-common.${format}`)
  const expectationPath = resolve(root, `go/officecompat/corpus/generated/expected/${format}-transitional-common.json`)
  const [packageBytes, expectationBytes] = await Promise.all([readFile(packagePath), readFile(expectationPath)])
  const expectation = JSON.parse(expectationBytes)
  return {
    id: `${format}-transitional-common`,
    package_sha256: sha256(packageBytes),
    package_bytes: packageBytes.length,
    native: expectation.native,
  }
}

function observeRuns(name, operation) {
  for (let index = 0; index < warmups; index++) operation()
  const outputs = []
  const elapsed = []
  const heapDeltas = []
  for (let index = 0; index < repetitions; index++) {
    const heapBefore = process.memoryUsage().heapUsed
    const started = performance.now()
    const output = operation()
    elapsed.push(Number((performance.now() - started).toFixed(3)))
    heapDeltas.push(process.memoryUsage().heapUsed - heapBefore)
    outputs.push(typeof output === 'string' ? output : canonical(output))
  }
  const stable = requireStable(name, outputs)
  return {
    output: stable,
    observation: {
      elapsed_ms: elapsed,
      heap_delta_bytes: heapDeltas,
      gating: false,
      note: 'informational only; runner scheduling, JIT, and GC are intentionally not pass/fail inputs',
    },
  }
}

async function observeAsyncRuns(name, operation) {
  for (let index = 0; index < warmups; index++) await operation()
  const outputs = []
  const elapsed = []
  const heapDeltas = []
  for (let index = 0; index < repetitions; index++) {
    const heapBefore = process.memoryUsage().heapUsed
    const started = performance.now()
    const output = await operation()
    elapsed.push(Number((performance.now() - started).toFixed(3)))
    heapDeltas.push(process.memoryUsage().heapUsed - heapBefore)
    outputs.push(typeof output === 'string' ? output : canonical(output))
  }
  const stable = requireStable(name, outputs)
  return {
    output: stable,
    observation: {
      elapsed_ms: elapsed,
      heap_delta_bytes: heapDeltas,
      gating: false,
      note: 'informational only; runner scheduling, JIT, and GC are intentionally not pass/fail inputs',
    },
  }
}

function qualifyXlsx(fixture) {
  const projected = observeRuns('xlsx/project', () => projectNativeWorkbookV1(structuredClone(fixture.native)))
  const model = JSON.parse(projected.output)
  const cells = model.sheets.reduce((sum, sheet) => sum + sheet.cells.length, 0)
  invariant(model.sheets.length === 1 && cells === 10, `xlsx projection lost objects: sheets/cells=${model.sheets.length}/${cells}`)
  invariant(fixture.native.sheets[0].cells.length === cells, 'xlsx projection cell count differs from native source')

  const synthetic = syntheticGeometryWorkbook()
  const syntheticModel = projectNativeWorkbookV1(synthetic)
  const metric = syntheticGeometryMetric(synthetic)
  const viewport = { row: 0, column: 0, end_row: 999, end_column: 99 }
  const geometry = compileNativeSheetGeometryV1(syntheticModel, synthetic.sheets[0].id, viewport, metric)
  invariant(geometry.rows.length === 1_000 && geometry.columns.length === 100, 'xlsx exact-limit geometry lost axis bands')
  const recording = createNativeSheetGeometryRecordingSurfaceV1(NATIVE_SHEET_GEOMETRY_LIMITS.maxCommands)
  emitNativeSheetGeometryCommandsV1(geometry, recording)
  const commands = recording.finish()
  const exactSurface = createNativeSheetGeometryRecordingSurfaceV1(commands.length)
  emitNativeSheetGeometryCommandsV1(geometry, exactSurface)
  invariant(exactSurface.finish().length === commands.length, 'xlsx exact command capacity did not pass')
  let commandRefusal
  try {
    emitNativeSheetGeometryCommandsV1(geometry, createNativeSheetGeometryRecordingSurfaceV1(commands.length - 1))
  } catch (error) {
    commandRefusal = error?.code
  }
  invariant(commandRefusal === 'geometry.commandBudget', `xlsx over-limit command surface did not fail closed: ${commandRefusal}`)
  let viewportRefusal
  try {
    compileNativeSheetGeometryV1(syntheticModel, synthetic.sheets[0].id, { row: 0, column: 0, end_row: 1_000, end_column: 99 }, metric)
  } catch (error) {
    viewportRefusal = error?.code
  }
  invariant(viewportRefusal === 'geometry.viewportBudget', `xlsx over-limit viewport did not fail closed: ${viewportRefusal}`)
  requireAtMost('xlsx geometry axis bands', geometry.rows.length + geometry.columns.length, budgets.node_work_ceilings.xlsx_geometry_axis_bands)
  requireAtMost('xlsx geometry commands', commands.length, budgets.node_work_ceilings.xlsx_geometry_commands)
  const decorations = compileNativeSheetDecorationsV1(syntheticModel, geometry)
  invariant(decorations.fills.length === NATIVE_SHEET_DECORATION_LIMITS.maxFills && decorations.border_segments.length === 0, 'xlsx exact-limit decorations lost resources')
  const exactBorderSegments = Array.from({ length: NATIVE_SHEET_DECORATION_LIMITS.maxBorderSegments }, (_, index) => ({
    orientation: 'horizontal', x1_emu: index, y1_emu: 0, x2_emu: index + 1, y2_emu: 0,
    border_style: 'thin', color: '#112233',
    sources: [{ cell_ref: performanceCellReference(index), edge: 'top', style_id: 0, border_id: 0, border_record_sha256: `sha256:${'e'.repeat(64)}` }],
  }))
  const { decoration_sha256: _compiledDigest, ...exactResourceUnsigned } = { ...decorations, border_segments: exactBorderSegments }
  const exactResources = { ...exactResourceUnsigned, decoration_sha256: `sha256:${sha256(JSON.stringify(exactResourceUnsigned))}` }
  const decorationSurface = createNativeSheetDecorationRecordingSurfaceV1(NATIVE_SHEET_DECORATION_LIMITS.maxCommands)
  emitNativeSheetDecorationCommandsV1(exactResources, decorationSurface)
  const decorationCommands = decorationSurface.finish()
  invariant(exactResources.border_segments.length === NATIVE_SHEET_DECORATION_LIMITS.maxBorderSegments, 'xlsx exact border-segment limit did not pass')
  invariant(decorationCommands.length === NATIVE_SHEET_DECORATION_LIMITS.maxCommands, 'xlsx exact global decoration command limit did not pass')
  const exactDecorationSurface = createNativeSheetDecorationRecordingSurfaceV1(decorationCommands.length)
  emitNativeSheetDecorationCommandsV1(exactResources, exactDecorationSurface, decorationCommands.length)
  invariant(exactDecorationSurface.finish().length === decorationCommands.length, 'xlsx exact decoration command capacity did not pass')
  let replayed = 0
  replayNativeSheetDecorationCommandsV1(null, decorationCommands, { execute() { replayed++ } })
  invariant(replayed === decorationCommands.length, 'xlsx deterministic decoration replay lost commands')
  let decorationRefusal, decorationPushes = 0
  try {
    emitNativeSheetDecorationCommandsV1({ ...exactResources, border_segments: [...exactResources.border_segments, exactResources.border_segments[0]] }, { push() { decorationPushes++ } })
  } catch (error) {
    decorationRefusal = error?.code
  }
  invariant(decorationRefusal === 'decoration.planInvalid' && decorationPushes === 0, `xlsx over-limit decoration plan did not fail atomically: ${decorationRefusal}/${decorationPushes}`)
  requireAtMost('xlsx decoration fills', decorations.fills.length, budgets.node_work_ceilings.xlsx_decoration_fills)
  requireAtMost('xlsx decoration border segments', exactResources.border_segments.length, budgets.node_work_ceilings.xlsx_decoration_border_segments)
  requireAtMost('xlsx decoration commands', decorationCommands.length, budgets.node_work_ceilings.xlsx_decoration_commands)
  return {
    counts: { sheets: model.sheets.length, cells, geometry_rows: geometry.rows.length, geometry_columns: geometry.columns.length, geometry_commands: commands.length, decoration_fills: decorations.fills.length, decoration_border_segments: exactResources.border_segments.length, decoration_commands: decorationCommands.length },
    digest: sha256(projected.output),
    observation: projected.observation,
    near_over: { viewport_cells_pass: 100_000, viewport_cells_refused: 100_100, exact_command_capacity: commands.length, refusal: commandRefusal, decoration_fills_pass: decorations.fills.length, decoration_borders_pass: exactResources.border_segments.length, decoration_commands_pass: decorationCommands.length, decoration_resources_refused: exactResources.border_segments.length + 1, decoration_refusal: decorationRefusal },
  }
}

function syntheticGeometryWorkbook() {
  const digest = `sha256:${'a'.repeat(64)}`
  const workbook = {
    protocol: 'injoffice.xlsx.native', version: 1, document_id: 'workbook:performance', revision: `rev:${'a'.repeat(64)}`,
    source: { package_sha256: digest, workbook_part: 'xl/workbook.xml', dialect: 'transitional', authority: 'exact-package-bytes' },
    normal_style: { style_xf_id: 0, font_id: 0, font_name: 'Aptos', font_size_points: 11, font_bold: false, font_italic: false, font_record_sha256: `sha256:${'c'.repeat(64)}` },
    sheets: [{ id: '1', name: 'Performance', order: 0, state: 'visible', part_name: 'xl/worksheets/sheet1.xml', sheet_format: { base_column_width: 8, default_row_height_points: 15, custom_height: false, zero_height: false }, rows: [], columns: [], cells: [], merged_ranges: [], editable: true }],
    styles: [],
    capabilities: [
      { name: 'native-ooxml-parse', level: 'read-only', detail: 'deterministic performance fixture' },
      { name: 'native-v1-mutations', level: 'partial', detail: 'deterministic performance fixture' },
      { name: 'unsupported-content', level: 'preserve-exact', detail: 'deterministic performance fixture' },
    ], passthrough_parts: [], unsupported: [],
  }
  const effective = { number_format: 'General', bold: false, italic: false, fill_color: '#AABBCC', fill: { origin: 'styles-record', fill_id: 0, record_sha256: `sha256:${'b'.repeat(64)}`, color: '#AABBCC' }, border: { origin: 'styles-record', border_id: 0, record_sha256: `sha256:${'e'.repeat(64)}` }, horizontal_alignment: 'general', vertical_alignment: 'bottom', wrap_text: false, projection: 'full', unsupported: [] }
  workbook.styles.push({ id: 0, effective, raw_projection_sha256: nativeWorkbookStyleRawProjectionSha256V1(effective) })
  return workbook
}

function performanceCellReference(index) {
  const row = Math.floor(index / 16_384) + 1
  let column = index % 16_384 + 1, letters = ''
  while (column > 0) {
    column--
    letters = String.fromCharCode(65 + column % 26) + letters
    column = Math.floor(column / 26)
  }
  return `${letters}${row}`
}

function syntheticGeometryMetric(workbook) {
  const normal = workbook.normal_style
  return {
    source_revision: workbook.revision, source_package_sha256: workbook.source.package_sha256,
    normal_style_xf_id: normal.style_xf_id, normal_style_font_id: normal.font_id,
    font_name: normal.font_name, font_size_points: normal.font_size_points, font_bold: normal.font_bold, font_italic: normal.font_italic,
    normal_font_record_sha256: normal.font_record_sha256, font_sha256: `sha256:${'d'.repeat(64)}`,
    provider_id: 'performance-metric', provider_revision: '1', measurement_dpi: 96, maximum_digit_width_pixels: 7,
  }
}

function pptxTextLayout() {
  const bytes = Uint8Array.from([0, 1, 2, 3])
  const digest = `sha256:${sha256(bytes)}`
  const face = (weight, matchedFamily = 'Aptos') => ({
    faceId: weight >= 700 ? 'performance.bold' : 'performance.regular', family: 'Performance Sans', weight, style: 'normal', stretch: 100,
    sourceKind: 'bundled', resourceId: weight >= 700 ? 'performance-bold' : 'performance-regular', contentDigest: digest,
    resolution: 'substitute', matchedFamily, fallbackChainId: 'performance.default',
  })
  return {
    manifest: {
      version: 1, manifestId: 'performance-fonts', revision: '1',
      faces: [
        { faceId: 'performance.regular', family: 'Performance Sans', aliases: ['Aptos'], weight: 400, style: 'normal', stretch: 100, source: { kind: 'bundled', resourceId: 'performance-regular', contentDigest: digest } },
        { faceId: 'performance.bold', family: 'Performance Sans', aliases: ['Aptos'], weight: 700, style: 'normal', stretch: 100, source: { kind: 'bundled', resourceId: 'performance-bold', contentDigest: digest } },
      ],
      fallbackChains: [{ chainId: 'performance.default', faceIds: ['performance.regular', 'performance.bold'] }],
    },
    resolver: {
      providerId: 'performance-resolver', providerRevision: '1',
      resolve({ run }) { return { status: 'resolved', face: face(run.font.weight, run.font.families[0]), attemptedFaceIds: ['performance.regular'], decisions: [] } },
      load(resolved) { return { face: resolved, bytes: Uint8Array.from(bytes), metrics: { unitsPerEm: 1_000, ascender: 800, descender: -200, lineGap: 200 } } },
    },
    shaper: {
      providerId: 'performance-shaper', providerRevision: '1',
      shape({ run, startUtf16, endUtf16, font }) {
        const glyphs = []
        const clusters = []
        let utf16 = startUtf16
        for (const character of run.text.slice(startUtf16, endUtf16)) {
          const end = utf16 + character.length
          glyphs.push({ glyphId: character.codePointAt(0), clusterIndex: clusters.length, advanceXMilliPoints: 1_000, advanceYMilliPoints: 0, offsetXMilliPoints: 0, offsetYMilliPoints: 0 })
          clusters.push({ startUtf16: utf16, endUtf16: end, glyphStart: glyphs.length - 1, glyphEnd: glyphs.length, advanceInlineMilliPoints: 1_000, whitespace: /^\s$/u.test(character) })
          utf16 = end
        }
        const size = run.fontSizeMilliPoints
        return {
          startUtf16, endUtf16, face: font.face, glyphs, clusters,
          metrics: { fontSizeMilliPoints: size, ascentMilliPoints: size * 0.8, descentMilliPoints: size * -0.2, lineGapMilliPoints: size * 0.2, lineHeightMilliPoints: size * 1.2 },
          advanceInlineMilliPoints: clusters.length * 1_000, advanceBlockMilliPoints: 0,
        }
      },
    },
    defaults: { fontFamilies: ['Performance Sans'], fontSizeHundredthPt: 1_000, script: 'Latn', language: 'en-US', direction: 'ltr', fallbackChainIds: ['performance.default'] },
  }
}

async function qualifyPptx(fixture) {
  const inputBefore = canonical(fixture.native)
  const layout = pptxTextLayout()
  const compiled = await observeAsyncRuns('pptx/compile', async () => stringifySlideRenderTree(await compileNativePptxSlide(structuredClone(fixture.native), 0, { textLayout: layout })))
  invariant(canonical(fixture.native) === inputBefore, 'pptx compiler mutated its source contract')
  const tree = JSON.parse(compiled.output)
  const sourceElements = countPptxSourceElements(fixture.native.slides[0].elements)
  const renderNodes = countPptxRenderNodes(tree.nodes)
  invariant(sourceElements === 5 && renderNodes === 5, `pptx compile lost objects: source/render=${sourceElements}/${renderNodes}`)
  const shapes = fixture.native.slides[0].elements.filter((element) => element.kind === 'shape').length
  invariant(shapes === 4, `pptx source shape count changed: ${shapes}`)
  const surface = createRecordingPaintSurface()
  paintSlideRenderTree(tree, surface)
  const commands = surface.finish()
  const exactSurface = createRecordingPaintSurface(commands.length)
  paintSlideRenderTree(tree, exactSurface, commands.length)
  invariant(exactSurface.finish().length === commands.length, 'pptx exact paint command capacity did not pass')
  let refusal
  try {
    paintSlideRenderTree(tree, createRecordingPaintSurface(commands.length - 1), commands.length - 1)
  } catch (error) {
    refusal = String(error?.message ?? error)
  }
  invariant(refusal?.includes('paint commands exceed'), `pptx over-limit paint surface did not fail closed: ${refusal}`)
  requireAtMost('pptx render nodes', renderNodes, budgets.node_work_ceilings.pptx_recursive_render_nodes)
  requireAtMost('pptx paint commands', commands.length, budgets.node_work_ceilings.pptx_paint_commands)
  return {
    counts: { slides: fixture.native.slides.length, source_elements: sourceElements, render_nodes: renderNodes, shapes, paint_commands: commands.length },
    digest: sha256(compiled.output), observation: compiled.observation,
    near_over: { exact_command_capacity: commands.length, refused_capacity: commands.length - 1, refusal: 'paint commands exceed' },
  }
}

function countPptxSourceElements(elements) {
  return elements.reduce((sum, element) => sum + 1 + (element.children ? countPptxSourceElements(element.children) : 0), 0)
}

function countPptxRenderNodes(nodes) {
  return nodes.reduce((sum, node) => sum + 1 + (node.children ? countPptxRenderNodes(node.children) : 0), 0)
}

function docxPaginationRequest(paragraphCount = 32, linesPerParagraph = 2) {
  const hash = `sha256:${'a'.repeat(64)}`
  const relationshipsHash = `sha256:${'b'.repeat(64)}`
  const anchor = (path, start, end) => ({ part_name: 'word/document.xml', path, start_byte: start, end_byte: end, xml_sha256: hash })
  const paragraphs = Array.from({ length: paragraphCount }, (_, index) => ({
    id: `paragraph:${index + 1}`, anchor: anchor(`/w:document[1]/w:body[1]/w:p[${index + 1}]`, 100 + index * 100, 190 + index * 100),
    edit_policy: { mode: 'read-only', allowed_operations: [], refusal: { code: 'NATIVE_READ_ONLY', message: 'Performance fixture is immutable.', preservation: 'refuse-mutation' } },
    properties: {}, runs: [{ kind: 'text', id: `run:${index + 1}`, anchor: anchor(`/w:document[1]/w:body[1]/w:p[${index + 1}]/w:r[1]`, 110 + index * 100, 180 + index * 100), text: `paragraph ${index + 1}` }],
  }))
  const document = {
    protocol: 'injoffice.docx.native', version: 1, document_id: 'document:performance', revision: 'revision:1',
    source: { package_sha256: hash, main_part: 'word/document.xml' },
    body: { id: 'story:body', kind: 'body', part_name: 'word/document.xml', anchor: anchor('/w:document[1]/w:body[1]', 1, 10_000 + paragraphCount * 100), blocks: paragraphs.map((paragraph) => ({ kind: 'paragraph', id: paragraph.id, paragraph })) },
    sections: [{
      id: 'section:1', anchor: anchor('/w:document[1]/w:body[1]/w:sectPr[1]', 8_000, 8_090), starts_at_block_id: paragraphs[0].id, break_type: 'next-page', title_page: false,
      page: { width_twips: 1_000, height_twips: 1_000, orientation: 'portrait', margins: { top_twips: 100, right_twips: 100, bottom_twips: 100, left_twips: 100, header_twips: 50, footer_twips: 50, gutter_twips: 0 }, columns: 1, column_spacing_twips: 100, column_layout: 'equal-width', column_definitions: [{ id: 'column:section:1:0', ordinal: 0 }] },
      header_refs: [], footer_refs: [],
    }],
    headers: [], footers: [], notes: [], comment_stories: [], comments: [], capabilities: [],
    passthrough_parts: [
      { part_name: 'word/settings.xml', content_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml', byte_length: 1, sha256: hash, policy: 'preserve-verbatim' },
      { part_name: 'word/_rels/document.xml.rels', content_type: 'application/vnd.openxmlformats-package.relationships+xml', byte_length: 1, sha256: relationshipsHash, policy: 'preserve-verbatim' },
    ], unsupported: [],
  }
  const resolved = {
    protocol: 'injoffice.docx.resolved-layout', version: 1, document_id: document.document_id, revision: document.revision, source_parts: { main_part: 'word/document.xml' },
    paragraphs: paragraphs.map((paragraph) => ({ paragraph_id: paragraph.id, applied_styles: [], properties: {}, paragraph_mark_properties: { font_family: 'Performance', font_size_half_points: 20 } })),
    runs: paragraphs.map((paragraph) => ({ run_id: paragraph.runs[0].id, paragraph_id: paragraph.id, applied_paragraph_styles: [], applied_character_styles: [], properties: { font_family: 'Performance', font_size_half_points: 20 } })),
    tables: [], fonts: [{ name: 'Performance' }], diagnostics: [],
  }
  const shaped = {
    protocol: 'injoffice.docx.shaped-lines', version: 1, document_id: document.document_id, revision: document.revision,
    available_width_millipoints: 40_000, tab_interval_millipoints: DOCX_DEFAULT_TAB_STOP_TWIPS * 50,
    font_manifest: { manifest_id: 'manifest:performance', revision: '1' }, providers: {
      resolver_id: 'resolver:performance', resolver_revision: '1', shaper_id: 'shaper:performance', shaper_revision: '1',
      bidi_id: NATIVE_BIDI_PROVIDER_ID, bidi_revision: NATIVE_BIDI_PROVIDER_REVISION, bidi_unicode_version: BIDI_UNICODE_VERSION,
      unicode13_revision: UNICODE_13_CLASSIFIER_REVISION,
    },
    paragraphs: paragraphs.map((paragraph) => ({
      paragraph_id: paragraph.id, story_id: document.body.id, story_kind: 'body', direction: 'ltr', alignment: 'start', spacing_before_millipoints: 0, spacing_after_millipoints: 0,
      indent_start_millipoints: 0, indent_end_millipoints: 0, first_line_delta_millipoints: 0, block_advance_millipoints: linesPerParagraph * 10_000,
      lines: Array.from({ length: linesPerParagraph }, (_, line) => ({ id: `line:${paragraph.id}:${line}`, ordinal: line, available_width_millipoints: 40_000, inline_offset_millipoints: 0, advance_inline_millipoints: 5_000, ascent_millipoints: 8_000, descent_millipoints: -2_000, line_gap_millipoints: 0, line_height_millipoints: 10_000, justified: false, logical_to_visual: [], fragments: [] })),
    })), diagnostics: [],
  }
  const settings = {
    protocol: 'injoffice.docx.pagination-settings', version: 1, document_id: document.document_id, revision: document.revision,
    package_sha256: hash, main_part: 'word/document.xml', relationships_part: 'word/_rels/document.xml.rels', relationships_sha256: relationshipsHash, relationship_id: 'rIdSettings',
    settings_part: 'word/settings.xml', settings_sha256: hash, profile: 'word-modern-default', default_tab_stop_twips: DOCX_DEFAULT_TAB_STOP_TWIPS,
    mirror_margins: false, gutter_at_top: false, even_and_odd_headers: false, compatibility_mode: 15, diagnostics: [],
  }
  return { protocol: 'injoffice.docx.pagination-request', version: 1, document, resolved_layout: resolved, shaped_lines: shaped, pagination_settings: settings }
}

async function qualifyDocx(fixture) {
  const request = docxPaginationRequest()
  const sourceBefore = canonical(request)
  const paginated = observeRuns('docx/paginate', () => {
    const result = paginateNativeDocxV1(structuredClone(request))
    invariant(result.ok && result.value.status === 'paginated', `docx pagination refused: ${canonical(result)}`)
    return result.value
  })
  invariant(canonical(request) === sourceBefore, 'docx pagination mutated its source request')
  const output = JSON.parse(paginated.output)
  const lines = output.pages.reduce((sum, page) => sum + page.lines.length, 0)
  const slices = output.pages.reduce((sum, page) => sum + page.paragraph_slices.length, 0)
  invariant(lines === 64 && slices === 32 && output.pages.length === 16, `docx pagination lost output: pages/slices/lines=${output.pages.length}/${slices}/${lines}`)
  invariant(fixture.native.body.blocks.length === 2, `pinned docx extraction object count changed: ${fixture.native.body.blocks.length}`)
  requireAtMost('docx pagination pages', output.pages.length, budgets.node_work_ceilings.docx_pages)
  requireAtMost('docx pagination placed lines', lines, budgets.node_work_ceilings.docx_placed_lines)

  const paintRequest = docxPagePaintRequest()
  const painted = await observeAsyncRuns('docx/page-paint', async () => {
    let calls = 0
    const result = await compileNativeDocxPagePaintV1(structuredClone(paintRequest), {
      providerId: 'outline:performance', providerRevision: '1',
      getGlyphOutline(request) {
        calls++
        return { status: 'outlined', face: { ...request.face }, glyph_id: request.glyph_id, units_per_em: 1_000, path: [{ kind: 'move_to', x: 0, y: 0 }, { kind: 'line_to', x: 500, y: 0 }, { kind: 'line_to', x: 500, y: 700 }, { kind: 'line_to', x: 0, y: 700 }, { kind: 'close_path' }] }
      },
    })
    invariant(result.ok && result.value.status === 'painted', `docx page paint refused: ${canonical(result)}`)
    invariant(calls === 1, `docx page paint outline cache/provider calls = ${calls}, want 1`)
    return result.value
  })
  const paint = JSON.parse(painted.output)
  const paintLines = paint.pages.reduce((sum, page) => sum + page.lines.length, 0)
  const paintCommands = paint.pages.reduce((sum, page) => sum + page.commands.length, 0)
  invariant(paint.pages.length === 1 && paintLines === 1 && paintCommands === 2, `docx page paint lost output: pages/lines/commands=${paint.pages.length}/${paintLines}/${paintCommands}`)
  requireAtMost('docx page-paint pages', paint.pages.length, budgets.node_work_ceilings.docx_paint_pages)
  requireAtMost('docx glyph commands', paintCommands, budgets.node_work_ceilings.docx_glyph_commands)

  const overLines = DOCX_PAGINATION_LIMITS.maxPages * 2 + 2
  const over = docxPaginationRequest(1, overLines)
  over.document.sections[0].page.height_twips = 600
  over.document.sections[0].page.orientation = 'landscape'
  const refused = paginateNativeDocxV1(over)
  invariant(refused.ok && refused.value.status === 'refused' && refused.value.pages.length === 0 && refused.value.diagnostics.some((entry) => entry.code === 'resource-limit'), `docx over-limit pagination did not refuse atomically: ${canonical(refused)}`)
  return {
    counts: { source_blocks: fixture.native.body.blocks.length, pages: output.pages.length, paragraph_slices: slices, placed_lines: lines, paint_pages: paint.pages.length, paint_lines: paintLines, glyph_commands: paintCommands },
    digest: { pagination: sha256(paginated.output), page_paint: sha256(painted.output) }, observation: { pagination: paginated.observation, page_paint: painted.observation },
    near_over: { qualified_pages: output.pages.length, max_pages: DOCX_PAGINATION_LIMITS.maxPages, refused_lines: overLines, refusal: 'resource-limit' },
  }
}

function docxPagePaintRequest() {
  const pagination = docxPaginationRequest(1, 1)
  pagination.document.body.blocks[0].paragraph.runs[0].text = 'AA'
  const line = pagination.shaped_lines.paragraphs[0].lines[0]
  line.advance_inline_millipoints = 10_000
  line.logical_to_visual = [0]
  line.fragments = [{
    id: 'fragment:paragraph:1:0:0', source_kind: 'run', source_id: 'run:1', start_utf16: 0, end_utf16: 2, text: 'AA',
    direction: 'ltr', bidi_level: 0, logical_order: 0, script: 'Latn', language: 'en-US', face_id: 'face:performance', whitespace: false,
    advance_inline_millipoints: 10_000, justification_expansion_millipoints: 0, ascent_millipoints: 8_000, descent_millipoints: -2_000, line_gap_millipoints: 0,
    glyphs: [
      { glyph_id: 7, advance_x_millipoints: 5_000, advance_y_millipoints: 0, offset_x_millipoints: 0, offset_y_millipoints: 0 },
      { glyph_id: 7, advance_x_millipoints: 5_000, advance_y_millipoints: 0, offset_x_millipoints: 0, offset_y_millipoints: 0 },
    ],
  }]
  const layout = paginateNativeDocxV1(pagination)
  invariant(layout.ok && layout.value.status === 'paginated', `docx page-paint pagination fixture refused: ${canonical(layout)}`)
  const manifest = {
    version: 1, manifestId: 'manifest:performance', revision: '1',
    faces: [{ faceId: 'face:performance', family: 'Performance', weight: 400, style: 'normal', stretch: 100, source: { kind: 'bundled', resourceId: 'font:performance', contentDigest: `sha256:${'a'.repeat(64)}` } }],
    fallbackChains: [],
  }
  const integrity = {
    font_manifest_sha256: nativeDocxPagePaintFontManifestSha256V1(manifest),
    shaped_lines_sha256: nativeDocxPagePaintShapedLinesSha256V1(pagination.shaped_lines),
  }
  if (DOCX_PAGE_PAINT_REQUEST_V1_BINDING_FIELDS.IntegrityV1.includes('table_projection_sha256')) {
    integrity.table_projection_sha256 = `sha256:${sha256('[]')}`
  }
  if (DOCX_PAGE_PAINT_REQUEST_V1_BINDING_FIELDS.IntegrityV1.includes('paginated_layout_sha256')) {
    integrity.paginated_layout_sha256 = nativeDocxPagePaintPaginatedLayoutSha256V1(layout.value)
  }
  const mediaAssets = []
  if (DOCX_PAGE_PAINT_REQUEST_V1_BINDING_FIELDS.IntegrityV1.includes('media_assets_sha256')) {
    integrity.media_assets_sha256 = nativeDocxPagePaintMediaAssetsSha256V1(mediaAssets)
  }
  if (DOCX_PAGE_PAINT_REQUEST_V1_BINDING_FIELDS.IntegrityV1.includes('paginated_layout_sha256')) {
    integrity.paginated_layout_sha256 = nativeDocxPagePaintPaginatedLayoutSha256V1(layout.value)
  }
  return {
    protocol: 'injoffice.docx.page-paint-request', version: 1, pagination_request: pagination, paginated_layout: layout.value, font_manifest: manifest,
    media_assets: mediaAssets,
    integrity,
    outline_provider: { provider_id: 'outline:performance', provider_revision: '1' },
  }
}

function authorityViolation(source, location) {
  const imports = [...source.matchAll(/(?:from\s+|import\s*)['"]([^'"]+)['"]/g)].map((match) => match[1].toLowerCase())
  const forbiddenImports = ['@injoffice/pdf', 'pdfjs', 'pdf-lib', 'mammoth', 'libreoffice', 'jsdom', 'playwright', 'puppeteer', 'electron', 'konva', 'react', 'canvas', 'node:http', 'node:https', 'node:net', 'node:dns']
  const dependency = imports.find((specifier) => forbiddenImports.some((token) => specifier === token || specifier.startsWith(`${token}/`)))
  if (dependency) return `${location}: forbidden runtime dependency ${dependency}`
  const authority = /\b(?:DOMParser|HTMLElement|HTMLCanvasElement|OffscreenCanvas|PDFDocument)\b|\bdocument\.createElement\b|\binnerHTML\b|\b(?:spawn|execFile|exec)\s*\(\s*['"](?:soffice|libreoffice)['"]/i.exec(source)
  return authority ? `${location}: forbidden runtime authority ${authority[0]}` : undefined
}

async function runtimeAuthorityQualification() {
  const directories = [
    ['packages/sheets/src', /^native.*\.ts$/],
    ['packages/pptx-native/src', /^(?!.*\.test\.ts$).*\.ts$/],
    ['packages/pptx-render/src', /^(?:compile|geometry|paint|types|index)\.ts$/],
    ['packages/docs/src', /^native.*\.ts$/],
  ]
  const checked = []
  for (const [directory, pattern] of directories) {
    for (const name of (await readdir(resolve(root, directory))).sort()) {
      if (!pattern.test(name) || name.endsWith('.test.ts')) continue
      const path = resolve(root, directory, name)
      const source = await readFile(path, 'utf8')
      const violation = authorityViolation(source, relative(root, path))
      if (violation) throw new Error(violation)
      checked.push(relative(root, path))
    }
  }
  return checked
}

function selfTestHarness() {
  invariant(budgetError('exact', 10, 10) === undefined, 'harness rejected exact budget')
  invariant(budgetError('over', 11, 10)?.includes('algorithmic/resource regression'), 'harness accepted budget+1')
  invariant(budgetError('invalid', 1, 0)?.includes('invalid'), 'harness accepted zero budget')
  invariant(requireStable('elapsed-independent', ['same', 'same', 'same']) === 'same', 'harness stability self-test failed')
  let nondeterminism
  try { requireStable('alternating', ['a', 'b', 'a']) } catch (error) { nondeterminism = String(error.message) }
  invariant(nondeterminism?.includes('nondeterministic'), 'harness did not reject alternating output')
  for (const source of [
    `import engine from 'electron'`, `import pdf from '@injoffice/pdf'`, `import mammoth from 'mammoth'`, `const parser = new DOMParser()`, `const pdf = new PDFDocument()`, `document.createElement('div')`, `spawn('libreoffice')`, `import https from 'node:https'`,
  ]) invariant(authorityViolation(source, 'self-test'), `authority scanner accepted ${source}`)
  invariant(!authorityViolation(`const metadata = { generator: 'example-suite@1' }`, 'metadata'), 'authority scanner rejected inert pinned metadata')
  return { exact_budget_passes: true, budget_plus_one_refused: true, nondeterminism_refused: true, timing_excluded_from_gate: true, forbidden_authority_vectors_refused: 8, inert_metadata_allowed: true }
}

const fixtures = Object.fromEntries(await Promise.all(['xlsx', 'pptx', 'docx'].map(async (format) => [format, await readCorpusFixture(format)])))
const budgetsPath = resolve(root, 'docs/qualification/native-office-performance-v1.json')
const budgets = JSON.parse(await readFile(budgetsPath, 'utf8'))
for (const [format, fixture] of Object.entries(fixtures)) {
  invariant(fixture.package_sha256 === budgets.baseline.fixtures[format].sha256, `${fixture.id}: package bytes drifted from the pinned performance budget record`)
}
const authorityFiles = await runtimeAuthorityQualification()
const selfTests = selfTestHarness()
const [xlsx, pptx, docx] = await Promise.all([
  Promise.resolve(qualifyXlsx(fixtures.xlsx)),
  qualifyPptx(fixtures.pptx),
  qualifyDocx(fixtures.docx),
])

const report = {
  protocol: 'injoffice.office.performance-report/v1',
  runtime: { node: process.version, platform: process.platform, architecture: process.arch },
  repetitions: { warmups, measured: repetitions },
  fixtures: Object.fromEntries(Object.entries(fixtures).map(([format, fixture]) => [format, { id: fixture.id, package_sha256: fixture.package_sha256, package_bytes: fixture.package_bytes }])),
  qualification: { xlsx, pptx, docx },
  authority: { checked_files: authorityFiles.length, forbidden_runtime_authority: false },
  self_tests: selfTests,
  budgets: { protocol: budgets.protocol, go_allocation_ceilings_per_op: budgets.go_allocation_ceilings_per_op, node_work_ceilings: budgets.node_work_ceilings },
  limitations: budgets.limitations,
}

await mkdir(dirname(outputPath), { recursive: true })
await writeFile(outputPath, `${JSON.stringify(report, null, 2)}\n`)
console.log(`native Office performance qualification passed (${process.version}); report: ${relative(root, outputPath)}`)
