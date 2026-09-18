/** Approximate DrawingML charts for the read-only approximate page preview.
 *
 * The strict extractor keeps refusing c:chart drawings. The Go sidecar
 * (`InspectNativeApproximateDrawingChartsV1`) joins those refusals to their
 * source drawing and describes a bounded clustered bar/column model read only
 * from the cached values in the chart part. This module validates that sidecar
 * against the current document, reserves inline charts as glyphless textbox
 * atoms in the internal body copy (the same reservation the approximate shape
 * module uses), and appends approximate paint after body pagination: bars,
 * axes, gridlines, legend swatches and the chart area reuse the rectangle and
 * border primitives; title, axis and legend text is shaped with the current
 * InjOffice shaper and painted as glyph paths attached to the anchor line.
 *
 * Layout is a host approximation (Office plot layout is not reproduced), and an
 * unauthored value axis scale is derived from the cached values and disclosed.
 * No workbook cell is read and no chart value is recalculated. Nothing here
 * touches strict paint or source bytes.
 *
 * The PPTX renderer's literal bar primitives (`createNativeLiteralBarPaths`,
 * `layoutChartAxes`) are scoped to `overlap = 0` with authored min/max axes and
 * PPTX text bodies; Word's default clustered charts author `overlap = -27` and
 * automatic scaling, so this module carries its own bounded geometry while the
 * sidecar model mirrors `NativeLiteralBar` field-for-field where they overlap. */
import type { FontResource, NativeFontManifest, NativeFontResolver, NativeTextShaper, ResolvedFontFace, ShapedSegment, TextRunInput } from '@injoffice/font-metrics/layout'
import { unicode13Script } from '@injoffice/font-metrics/unicode13'
import type { NativeDocxDocumentV1, NativeDocxRunV1, NativeDocxSourceAnchorV1, NativeDocxUnsupportedCapabilityV1 } from './nativeContract.js'
import type { NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'
import type { NativeDocxPaginationSettingsV1 } from './nativePaginationSettings.js'
import { ID, RGB, preflightWire, paintCommandID } from './nativePagePaintWireV1.js'
import { nativeDocxPageGlyphOutlineRegistryV1, nativeDocxRegisterGlyphOutlineV1, nativeDocxCaptureGlyphOutlineV1, type NativeDocxContentAddressedFaceV1, type NativeDocxFillGlyphPathCommandV1, type NativeDocxFillTableCellCommandV1, type NativeDocxFillTextHighlightCommandV1, type NativeDocxGlyphOutlineProviderV1, type NativeDocxGlyphOutlineResultV1, type NativeDocxPagePaintCommandV1, type NativeDocxPagePaintRequestV1, type NativeDocxPagePaintSuccessV1, type NativeDocxPaintLineV1, type NativeDocxPaintPageV1, type NativeDocxStrokeTableBorderCommandV1 } from './nativePagePaintV1.js'
import { textboxAnchorLine, type TextboxAnchorContext } from './nativeTextboxAnchorLineV2.js'
import { resolveTextboxPosition } from './nativeTextboxPositionV2.js'
import type { NativeDocxTextboxGeometryItemV1 } from './nativeTextboxGeometryPreviewV1.js'
import type { NativeTextboxRelativeAnchorV2 } from './nativeTextboxPageAnchorV1.js'
import { asciiLowerNative } from './nativeDeterminism.js'
import type { NativeDocxApproximateShapeLineV1 } from './nativeApproximateDrawingShapesV1.js'

export const DOCX_APPROXIMATE_DRAWING_CHARTS_PROTOCOL = 'injoffice.docx.approximate-drawing-charts' as const
export const DOCX_APPROXIMATE_DRAWING_CHART_POLICY = 'docx.approximate-drawing-chart-preview-v1' as const
export const DOCX_APPROXIMATE_DRAWING_CHART_CODE = 'docx.approximate-drawing-chart-preview' as const
export const DOCX_APPROXIMATE_DRAWING_CHART_OMITTED_CODE = 'docx.approximate-drawing-chart-omitted' as const
export const DOCX_APPROXIMATE_CHART_FONT_CODE = 'docx.approximate-chart-substituted-font' as const
export const DOCX_APPROXIMATE_DRAWING_CHART_WARNING = `${DOCX_APPROXIMATE_DRAWING_CHART_CODE}: DrawingML clustered bar/column charts are painted approximately from the values cached in the chart part with a host layout: plot margins, legend and title placement, dash segmentation and any unauthored value-axis scale are InjOffice approximations, not Office chart layout. Original drawing restrictions and source bytes are unchanged.` as const
export const DOCX_APPROXIMATE_DRAWING_CHART_SIDECAR_REFUSED = `${DOCX_APPROXIMATE_DRAWING_CHART_OMITTED_CODE}: drawing-chart evidence did not exact-join the source document and was not used; refused charts stay omitted` as const
/** Table paint primitives carry this id so consumers can tell chart paint from table paint. */
export const DOCX_APPROXIMATE_DRAWING_CHART_TABLE_ID = DOCX_APPROXIMATE_DRAWING_CHART_POLICY

const MAX_CHARTS = 16
const MAX_SERIES = 16
const MAX_CATEGORIES = 256
const MAX_TEXT = 1024
const MAX_CHART_GLYPHS = 20_000
const MAX_DASH_SEGMENTS = 6_000
const MAX_TICKS = 64
const MAX_REASONS = 24
const EMU_PER_MILLIPOINT = 12.7
/** Host layout constants in millipoints. */
const OUTER_PAD = 7_000
const TITLE_GAP = 7_000
const LEGEND_GAP = 7_000
const LABEL_GAP = 4_000
const LEGEND_ENTRY_GAP = 10_000
const LEGEND_SWATCH_GAP = 3_000
const DECIMAL = /^-?(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][-+]?[0-9]{1,3})?$/

export interface NativeDocxApproximateChartFontV1 { family: string; size_hundredth_pt: number; rgb: string; bold: boolean; italic: boolean }
export interface NativeDocxApproximateChartSeriesV1 { index: number; order: number; title?: string; values: string[]; fill_rgb: string; line?: NativeDocxApproximateShapeLineV1 }
export interface NativeDocxApproximateChartAxisV1 {
  deleted: boolean
  orientation: 'minMax' | 'maxMin'
  line?: NativeDocxApproximateShapeLineV1
  labels?: NativeDocxApproximateChartFontV1
  major_gridlines?: NativeDocxApproximateShapeLineV1
  number_format: string
  min?: string
  max?: string
  major_unit?: string
}
export interface NativeDocxApproximateChartTitleV1 { text?: string; font: NativeDocxApproximateChartFontV1; overlay: boolean }
export interface NativeDocxApproximateChartLegendV1 { position: 'b' | 't' | 'l' | 'r' | 'tr'; font: NativeDocxApproximateChartFontV1; overlay: boolean }
export interface NativeDocxApproximateChartModelV1 {
  kind: 'bar'
  bar_direction: 'column' | 'bar'
  grouping: 'clustered'
  gap_width_percent: number
  overlap_percent: number
  categories: string[]
  series: NativeDocxApproximateChartSeriesV1[]
  category_axis: NativeDocxApproximateChartAxisV1
  value_axis: NativeDocxApproximateChartAxisV1
  title?: NativeDocxApproximateChartTitleV1
  legend?: NativeDocxApproximateChartLegendV1
  area_fill_rgb?: string
  area_line?: NativeDocxApproximateShapeLineV1
  plot_fill_rgb?: string
  plot_line?: NativeDocxApproximateShapeLineV1
}
export interface NativeDocxApproximateDrawingChartV1 {
  id: string
  paragraph_id: string
  diagnostic_ids: string[]
  anchor: NativeDocxSourceAnchorV1
  run_anchor: NativeDocxSourceAnchorV1
  status: 'supported' | 'omitted'
  reason?: string
  placement?: 'inline' | 'anchored'
  width_emu: number
  height_emu: number
  page_anchor?: NativeTextboxRelativeAnchorV2
  wrap?: string
  chart_part?: string
  chart_part_sha256?: string
  chart?: NativeDocxApproximateChartModelV1
  notes?: string[]
}
export interface NativeDocxApproximateDrawingChartsV1 {
  protocol: typeof DOCX_APPROXIMATE_DRAWING_CHARTS_PROTOCOL
  version: 1
  policy: typeof DOCX_APPROXIMATE_DRAWING_CHART_POLICY
  package_sha256: string
  part_sha256: string
  items: NativeDocxApproximateDrawingChartV1[]
  omitted_count: number
}

function record(value: unknown): value is Record<string, unknown> { return value !== null && typeof value === 'object' && !Array.isArray(value) }
function exactKeys(value: Record<string, unknown>, required: readonly string[], optional: readonly string[] = []): boolean {
  const keys = Object.keys(value)
  return required.every(key => key in value) && keys.every(key => required.includes(key) || optional.includes(key))
}
function safeNonnegative(value: unknown, max = 127_000_000): value is number { return Number.isSafeInteger(value) && (value as number) >= 0 && (value as number) <= max }
function anchorValid(value: unknown, part: string): value is NativeDocxSourceAnchorV1 {
  return record(value) && exactKeys(value, ['part_name', 'path', 'start_byte', 'end_byte', 'xml_sha256']) && value.part_name === part && typeof value.path === 'string' && Number.isSafeInteger(value.start_byte) && Number.isSafeInteger(value.end_byte) && (value.start_byte as number) >= 0 && (value.end_byte as number) > (value.start_byte as number) && typeof value.xml_sha256 === 'string' && /^sha256:[0-9a-f]{64}$/.test(value.xml_sha256)
}
function within(inner: NativeDocxSourceAnchorV1, outer: NativeDocxSourceAnchorV1): boolean { return inner.part_name === outer.part_name && inner.start_byte >= outer.start_byte && inner.end_byte <= outer.end_byte }
/** C0 controls other than TAB (U+0009) and LF (U+000A) never come from a cached
 * chart string; each rejected code point is tested explicitly (no range). */
function hasForbiddenControl(value: string): boolean {
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index)
    if (code < 0x20 && code !== 0x09 && code !== 0x0a) return true
  }
  return false
}
function boundedText(value: unknown, max = MAX_TEXT): value is string { return typeof value === 'string' && value.length <= max && !hasForbiddenControl(value) }
function lineValid(value: unknown): boolean {
  return record(value) && exactKeys(value, ['rgb', 'width_emu', 'dash']) && typeof value.rgb === 'string' && RGB.test(value.rgb) && safeNonnegative(value.width_emu, 12_700_000) && (value.width_emu as number) > 0 && typeof value.dash === 'string' && value.dash.length <= 32
}
function fontValid(value: unknown): boolean {
  return record(value) && exactKeys(value, ['family', 'size_hundredth_pt', 'rgb', 'bold', 'italic']) && boundedText(value.family, 128) && safeNonnegative(value.size_hundredth_pt, 400_000) && (value.size_hundredth_pt as number) >= 100 && typeof value.rgb === 'string' && RGB.test(value.rgb) && typeof value.bold === 'boolean' && typeof value.italic === 'boolean'
}
function decimalValid(value: unknown): value is string { return typeof value === 'string' && value.length <= 64 && DECIMAL.test(value) && Number.isFinite(Number(value)) }
function axisValid(value: unknown, isValue: boolean): boolean {
  if (!record(value) || !exactKeys(value, ['deleted', 'orientation', 'number_format'], ['line', 'labels', 'major_gridlines', 'min', 'max', 'major_unit'])) return false
  if (typeof value.deleted !== 'boolean' || (value.orientation !== 'minMax' && value.orientation !== 'maxMin') || !boundedText(value.number_format, 64)) return false
  if (value.line !== undefined && !lineValid(value.line)) return false
  if (value.major_gridlines !== undefined && !lineValid(value.major_gridlines)) return false
  if (value.labels !== undefined && !fontValid(value.labels)) return false
  if (value.deleted && (value.line !== undefined || value.labels !== undefined)) return false
  for (const key of ['min', 'max', 'major_unit'] as const) if (value[key] !== undefined && (!isValue || !decimalValid(value[key]))) return false
  if (value.major_unit !== undefined && Number(value.major_unit) <= 0) return false
  if (value.min !== undefined && value.max !== undefined && Number(value.min) >= Number(value.max)) return false
  return true
}
function modelValid(value: unknown): value is NativeDocxApproximateChartModelV1 {
  if (!record(value) || !exactKeys(value, ['kind', 'bar_direction', 'grouping', 'gap_width_percent', 'overlap_percent', 'categories', 'series', 'category_axis', 'value_axis'], ['title', 'legend', 'area_fill_rgb', 'area_line', 'plot_fill_rgb', 'plot_line'])) return false
  if (value.kind !== 'bar' || (value.bar_direction !== 'column' && value.bar_direction !== 'bar') || value.grouping !== 'clustered' || !safeNonnegative(value.gap_width_percent, 500) || !Number.isSafeInteger(value.overlap_percent) || Math.abs(value.overlap_percent as number) > 100) return false
  if (!Array.isArray(value.categories) || value.categories.length < 1 || value.categories.length > MAX_CATEGORIES || value.categories.some(category => !boundedText(category))) return false
  if (!Array.isArray(value.series) || value.series.length < 1 || value.series.length > MAX_SERIES) return false
  const indices = new Set<number>(), orders = new Set<number>()
  for (const series of value.series) {
    if (!record(series) || !exactKeys(series, ['index', 'order', 'values', 'fill_rgb'], ['title', 'line'])) return false
    if (!safeNonnegative(series.index, 0xffffffff) || !safeNonnegative(series.order, 0xffffffff) || indices.has(series.index) || orders.has(series.order)) return false
    indices.add(series.index); orders.add(series.order)
    if (!Array.isArray(series.values) || series.values.length !== value.categories.length || series.values.some(entry => entry !== '' && !decimalValid(entry))) return false
    if (typeof series.fill_rgb !== 'string' || !RGB.test(series.fill_rgb) || (series.title !== undefined && !boundedText(series.title)) || (series.line !== undefined && !lineValid(series.line))) return false
  }
  if (!axisValid(value.category_axis, false) || !axisValid(value.value_axis, true)) return false
  if (value.title !== undefined && (!record(value.title) || !exactKeys(value.title, ['font', 'overlay'], ['text']) || !fontValid(value.title.font) || typeof value.title.overlay !== 'boolean' || (value.title.text !== undefined && !boundedText(value.title.text)))) return false
  if (value.legend !== undefined && (!record(value.legend) || !exactKeys(value.legend, ['position', 'font', 'overlay']) || !['b', 't', 'l', 'r', 'tr'].includes(String(value.legend.position)) || !fontValid(value.legend.font) || typeof value.legend.overlay !== 'boolean')) return false
  for (const key of ['area_fill_rgb', 'plot_fill_rgb'] as const) if (value[key] !== undefined && (typeof value[key] !== 'string' || !RGB.test(value[key] as string))) return false
  for (const key of ['area_line', 'plot_line'] as const) if (value[key] !== undefined && !lineValid(value[key])) return false
  return true
}

/** Bounded structural validation plus source joins: every chart must name a
 * body paragraph and at least one retained drawing/run diagnostic anchored
 * inside its own run, and must not overlap modeled text. */
export function decodeNativeDocxApproximateDrawingChartsV1(value: unknown, document: NativeDocxDocumentV1): NativeDocxApproximateDrawingChartsV1 {
  if (preflightWire(value, 'approximate drawing charts', 400_000, 20_000).length) throw new TypeError('Approximate drawing charts exceed their bounded wire')
  const input = structuredClone(value) as NativeDocxApproximateDrawingChartsV1
  if (!record(input) || !exactKeys(input as unknown as Record<string, unknown>, ['protocol', 'version', 'policy', 'package_sha256', 'part_sha256', 'items', 'omitted_count']) || input.protocol !== DOCX_APPROXIMATE_DRAWING_CHARTS_PROTOCOL || input.version !== 1 || input.policy !== DOCX_APPROXIMATE_DRAWING_CHART_POLICY || input.package_sha256 !== document.source.package_sha256 || typeof input.part_sha256 !== 'string' || !Array.isArray(input.items) || input.items.length > MAX_CHARTS || !safeNonnegative(input.omitted_count, 1_000_000)) throw new TypeError('Approximate drawing charts do not exact-join the source document')
  const paragraphs = new Map(document.body.blocks.flatMap(block => block.paragraph ? [[block.id, block.paragraph] as const] : []))
  const diagnostics = new Map(document.unsupported.map(entry => [entry.id, entry]))
  const ids = new Set<string>()
  const main = document.source.main_part
  for (const item of input.items) {
    if (!record(item) || !exactKeys(item as unknown as Record<string, unknown>, ['id', 'paragraph_id', 'diagnostic_ids', 'anchor', 'run_anchor', 'status', 'width_emu', 'height_emu'], ['reason', 'placement', 'page_anchor', 'wrap', 'chart_part', 'chart_part_sha256', 'chart', 'notes'])) throw new TypeError('Approximate drawing chart has unknown or missing fields')
    if (typeof item.id !== 'string' || !ID.test(item.id) || ids.has(item.id)) throw new TypeError('Approximate drawing chart id is invalid or duplicated')
    ids.add(item.id)
    const paragraph = paragraphs.get(item.paragraph_id)
    if (!paragraph || !anchorValid(item.anchor, main) || !anchorValid(item.run_anchor, main) || !within(item.run_anchor, paragraph.anchor) || !within(item.anchor, item.run_anchor)) throw new TypeError('Approximate drawing chart does not exact-join its body paragraph')
    if (!Array.isArray(item.diagnostic_ids) || item.diagnostic_ids.length === 0 || item.diagnostic_ids.length > 64 || item.diagnostic_ids.some(id => { const d = typeof id === 'string' ? diagnostics.get(id) : undefined; return !d || d.scope_id !== item.paragraph_id || !d.anchor || !within(d.anchor, item.run_anchor) })) throw new TypeError('Approximate drawing chart must join retained source drawing diagnostics')
    if (item.status !== 'supported' && item.status !== 'omitted') throw new TypeError('Approximate drawing chart status is invalid')
    if (!safeNonnegative(item.width_emu) || !safeNonnegative(item.height_emu)) throw new TypeError('Approximate drawing chart extent is out of bounds')
    if (item.notes !== undefined && (!Array.isArray(item.notes) || item.notes.length > 32 || item.notes.some(note => !boundedText(note, 512)))) throw new TypeError('Approximate drawing chart notes are unbounded')
    if (item.reason !== undefined && !boundedText(item.reason, 256)) throw new TypeError('Approximate drawing chart reason is unbounded')
    if (item.chart_part !== undefined && !boundedText(item.chart_part, 512)) throw new TypeError('Approximate drawing chart part name is unbounded')
    if (item.chart_part_sha256 !== undefined && (typeof item.chart_part_sha256 !== 'string' || !/^sha256:[0-9a-f]{64}$/.test(item.chart_part_sha256))) throw new TypeError('Approximate drawing chart part digest is invalid')
    if (item.status === 'omitted') { if (item.chart !== undefined) throw new TypeError('Omitted approximate drawing charts cannot carry a model'); continue }
    // A painted chart must own its run: a w:r that also carries modeled text is
    // omitted by the sidecar (shared-run); a supported item claiming one is forged.
    if (paragraph.runs.some(run => within(run.anchor, item.run_anchor) || within(item.run_anchor, run.anchor))) throw new TypeError('Approximate drawing chart overlaps modeled text')
    if (item.width_emu <= 0 || item.height_emu <= 0 || (item.placement !== 'inline' && item.placement !== 'anchored') || !item.chart_part || !item.chart_part_sha256) throw new TypeError('Supported approximate drawing chart requires positive extent, placement and chart part')
    if (!modelValid(item.chart)) throw new TypeError('Approximate drawing chart model is invalid')
    if (item.placement === 'anchored') {
      const anchor = item.page_anchor as unknown
      if (!record(anchor) || anchor.policy !== 'relative-position-no-wrap-v2' || !anchorValid(anchor.source_anchor, main) || !within(anchor.source_anchor as NativeDocxSourceAnchorV1, item.anchor) || !Number.isSafeInteger(anchor.x_emu) || !Number.isSafeInteger(anchor.y_emu) || Math.abs(anchor.x_emu as number) > 127_000_000 || Math.abs(anchor.y_emu as number) > 127_000_000 || typeof anchor.horizontal_relative !== 'string' || typeof anchor.vertical_relative !== 'string' || !record(anchor.stacking) || typeof anchor.stacking.behind_doc !== 'boolean' || !safeNonnegative(anchor.stacking.relative_height, 0xffffffff)) throw new TypeError('Approximate drawing chart anchor is invalid')
    } else if (item.page_anchor !== undefined) throw new TypeError('Inline approximate drawing charts cannot carry page anchors')
  }
  return input
}

export interface NativeDocxApproximateInlineChartProjectionV1 {
  document: NativeDocxDocumentV1
  resolved: NativeDocxResolvedLayoutInputV1
  /** Chart id to the synthetic glyphless drawing run id reserved in the body copy. */
  inlineRuns: Map<string, string>
  /** Source refusals removed from the body copy for supported charts; restored for charts the painter drops. */
  removedDiagnostics: NativeDocxUnsupportedCapabilityV1[]
}

/** Reserve supported inline charts as glyphless textbox atoms so surrounding
 * text reflows around their extent; drop the joined refusals from this
 * internal copy only. Anchored charts leave the body copy untouched. */
export function projectNativeDocxApproximateInlineChartsV1(document: NativeDocxDocumentV1, resolved: NativeDocxResolvedLayoutInputV1, charts: NativeDocxApproximateDrawingChartsV1): NativeDocxApproximateInlineChartProjectionV1 {
  const projected = structuredClone(document)
  const projectedResolved = structuredClone(resolved)
  const inlineRuns = new Map<string, string>()
  const removed = new Set<string>()
  const paragraphs = new Map(projected.body.blocks.flatMap(block => block.paragraph ? [[block.id, block.paragraph] as const] : []))
  const blockIndex = new Map(projected.body.blocks.map((block, index) => [block.id, index]))
  for (const chart of charts.items) {
    if (chart.status !== 'supported') continue
    for (const id of chart.diagnostic_ids) removed.add(id)
    if (chart.placement !== 'inline') continue
    const paragraph = paragraphs.get(chart.paragraph_id)
    if (!paragraph) continue
    let widthEMU = chart.width_emu
    const columnWidth = columnWidthEMU(projected, blockIndex.get(chart.paragraph_id) ?? 0)
    if (columnWidth !== undefined && widthEMU > columnWidth) { widthEMU = columnWidth; chart.notes = [...(chart.notes ?? []), 'inline chart width clamped to its column width'] }
    const runID = `${chart.id}:run`
    const run: NativeDocxRunV1 = {
      kind: 'drawing', id: runID, anchor: chart.run_anchor,
      drawing: {
        id: `${chart.id}:drawing`, anchor: chart.anchor, placement: 'inline', width_emu: widthEMU, height_emu: chart.height_emu,
        // The atom carries no painted text; the id keeps the required non-empty marker honest.
        textbox_text: chart.id, textbox_fill_rgb: 'FFFFFF', textbox_line_rgb: '000000',
        edit_policy: { mode: 'read-only', allowed_operations: [], refusal: { code: 'APPROXIMATE_DRAWING_CHART_PREVIEW', message: 'Inline chart reserved for read-only approximate preview', preservation: 'refuse-mutation' } },
      },
    }
    let index = paragraph.runs.findIndex(existing => existing.anchor.start_byte > chart.run_anchor.start_byte)
    if (index < 0) index = paragraph.runs.length
    paragraph.runs.splice(index, 0, run)
    projectedResolved.runs.push({ run_id: runID, paragraph_id: paragraph.id, applied_paragraph_styles: [], applied_character_styles: [], properties: {} })
    inlineRuns.set(chart.id, runID)
  }
  const removedDiagnostics = projected.unsupported.filter(entry => removed.has(entry.id))
  projected.unsupported = projected.unsupported.filter(entry => !removed.has(entry.id))
  return { document: projected, resolved: projectedResolved, inlineRuns, removedDiagnostics }
}

/** Equal-width column extent of the section owning a body block, in EMU. */
function columnWidthEMU(document: NativeDocxDocumentV1, blockIndex: number): number | undefined {
  const blockIndices = new Map(document.body.blocks.map((block, index) => [block.id, index]))
  let section = document.sections[0]
  for (const candidate of document.sections) {
    const start = blockIndices.get(candidate.starts_at_block_id)
    if (start !== undefined && start <= blockIndex) section = candidate
  }
  if (!section) return undefined
  const page = section.page
  const columns = Math.max(1, page.columns)
  const text = page.width_twips - page.margins.left_twips - page.margins.right_twips - page.margins.gutter_twips
  const width = Math.floor((text - page.column_spacing_twips * (columns - 1)) / columns)
  return width > 0 ? width * 635 : undefined
}

export interface NativeDocxApproximateChartPaintRuntimeV1 {
  request: NativeDocxPagePaintRequestV1
  document: NativeDocxDocumentV1
  settings: NativeDocxPaginationSettingsV1
  manifest: NativeFontManifest
  resolver: NativeFontResolver
  shaper: NativeTextShaper
  outlineProvider: NativeDocxGlyphOutlineProviderV1
}

export interface NativeDocxApproximateChartFontSubstitutionV1 { source_family: string; selected_family: string; face_id: string; weight: number; style: string }

export interface NativeDocxApproximateChartPaintResultV1 {
  painted: string[]
  omitted: Array<{ id: string; reason: string }>
  substitutions: NativeDocxApproximateChartFontSubstitutionV1[]
  reasons: string[]
}

interface PlacedChart { chart: NativeDocxApproximateDrawingChartV1; page: NativeDocxPaintPageV1; line: NativeDocxPaintLineV1; x: number; y: number; width: number; height: number; behind: boolean; order: number }

function toMillipoints(emu: number): number { return Math.round(emu / EMU_PER_MILLIPOINT) }

/** Append approximate chart paint to already painted approximate pages. The
 * pages are mutated in place; the caller re-validates the whole envelope. */
export async function paintNativeDocxApproximateDrawingChartsV1(paint: Pick<NativeDocxPagePaintSuccessV1, 'pages'>, charts: NativeDocxApproximateDrawingChartsV1, projection: NativeDocxApproximateInlineChartProjectionV1, runtime: NativeDocxApproximateChartPaintRuntimeV1): Promise<NativeDocxApproximateChartPaintResultV1> {
  const result: NativeDocxApproximateChartPaintResultV1 = { painted: [], omitted: [], substitutions: [], reasons: [] }
  const body = { pages: paint.pages } as NativeDocxPagePaintSuccessV1
  const placed: PlacedChart[] = []
  const removedIDs = new Set<string>()
  const omit = (chart: NativeDocxApproximateDrawingChartV1, reason: string) => { result.omitted.push({ id: chart.id, reason }) }
  for (const chart of charts.items) {
    if (chart.status !== 'supported' || !chart.chart) { omit(chart, chart.reason ?? 'unsupported'); continue }
    const width = toMillipoints(chart.width_emu), height = toMillipoints(chart.height_emu)
    if (width <= 0 || height <= 0) { omit(chart, 'degenerate-extent'); continue }
    if (chart.placement === 'inline') {
      const runID = projection.inlineRuns.get(chart.id)
      const found = runID ? findHighlight(paint.pages, runID) : undefined
      if (!found) { omit(chart, 'inline-chart-not-placed'); continue }
      // The reservation highlight only located the atom; the chart paints itself.
      removedIDs.add(found.command.id)
      found.line.command_ids = found.line.command_ids.filter(id => id !== found.command.id)
      placed.push({ chart, page: found.page, line: found.line, x: found.command.x_millipoints, y: found.command.y_millipoints, width: found.command.width_millipoints, height: found.command.height_millipoints, behind: false, order: 0 })
      continue
    }
    const anchor = chart.page_anchor!
    const item = { owner: { paragraph_id: chart.paragraph_id }, geometry: null, page_anchor: anchor } as unknown as NativeDocxTextboxGeometryItemV1
    let context: TextboxAnchorContext & { page: NativeDocxPaintPageV1 }
    try {
      context = textboxAnchorLine(runtime.document, item, body, runtime.request)
    } catch {
      const first = firstBodyLine(paint.pages, chart.paragraph_id)
      if (!first) { omit(chart, 'anchor-paragraph-not-placed'); continue }
      context = { page: first.page, line: first.line, character_x: first.line.x_millipoints, paragraph_y: first.line.y_millipoints }
    }
    const position = resolveChartPosition(runtime.document, item, context.page, width, height, context)
    if (!position.ok) { omit(chart, `anchor-unresolved: ${position.message}`); continue }
    placed.push({ chart, page: context.page, line: context.line, x: position.x, y: position.y, width, height, behind: anchor.stacking?.behind_doc === true, order: anchor.stacking?.relative_height ?? 0 })
  }
  placed.sort((left, right) => left.order - right.order)
  const text = new ChartTextPainter(runtime, result)
  const notes = new Set<string>()
  const layoutNotes = new Set<string>()
  const insertions: Array<{ page: NativeDocxPaintPageV1; behind: boolean; commands: NativeDocxPagePaintCommandV1[] }> = []
  for (const [ordinal, entry] of placed.entries()) {
    const maxX = entry.page.width_millipoints, maxY = entry.page.height_millipoints
    if (entry.x < 0 || entry.y < 0 || entry.x + entry.width > maxX || entry.y + entry.height > maxY) { omit(entry.chart, 'outside-page'); continue }
    try {
      const painted = await paintChart(entry, ordinal, text, layoutNotes)
      if (painted.length === 0) { omit(entry.chart, 'nothing-to-paint'); continue }
      insertions.push({ page: entry.page, behind: entry.behind, commands: painted })
      result.painted.push(entry.chart.id)
      for (const note of entry.chart.notes ?? []) notes.add(note)
    } catch (error) {
      omit(entry.chart, `paint-failed: ${(error instanceof Error ? error.message : 'unknown').slice(0, 200)}`)
    }
  }
  for (const page of paint.pages) insertCommands(page, insertions.filter(entry => entry.page === page), removedIDs)
  result.reasons = buildReasons(charts, result, [...notes, ...layoutNotes], text)
  return result
}

/** Centered/aligned axes halve (extent - chart) and refuse a .5 result; an
 * approximate chart rounds by retrying with a one-millipoint wider extent so
 * parity never drops it. Offsets that are not whole millipoints round to the
 * nearest one (at most 63 EMU). The painted extent is unchanged. */
function resolveChartPosition(document: NativeDocxDocumentV1, item: NativeDocxTextboxGeometryItemV1, page: NativeDocxPaintPageV1, width: number, height: number, context: TextboxAnchorContext): { ok: true; x: number; y: number } | { ok: false; message: string } {
  let message = 'unknown'
  const anchor = item.page_anchor!
  const rounded = { ...item, page_anchor: { ...anchor, x_emu: Math.round(anchor.x_emu / 127) * 127, y_emu: Math.round(anchor.y_emu / 127) * 127 } } as NativeDocxTextboxGeometryItemV1
  for (const [dw, dh] of [[0, 0], [1, 0], [0, 1], [1, 1]] as const) {
    try {
      const { x, y } = resolveTextboxPosition(document, rounded, page, { width_millipoints: width + dw, height_millipoints: height + dh } as never, context)
      return { ok: true, x, y }
    } catch (error) {
      message = error instanceof Error ? error.message : 'unknown'
      if (!/coordinate precision/.test(message)) break
    }
  }
  return { ok: false, message }
}

function findHighlight(pages: NativeDocxPaintPageV1[], runID: string): { page: NativeDocxPaintPageV1; line: NativeDocxPaintLineV1; command: NativeDocxFillTextHighlightCommandV1 } | undefined {
  for (const page of pages) {
    const command = page.commands.find((entry): entry is NativeDocxFillTextHighlightCommandV1 => entry.kind === 'fill_text_highlight' && entry.source_id === runID)
    if (!command) continue
    const line = page.lines.find(entry => entry.command_ids.includes(command.id))
    if (line) return { page, line, command }
  }
  return undefined
}

function firstBodyLine(pages: NativeDocxPaintPageV1[], paragraphID: string): { page: NativeDocxPaintPageV1; line: NativeDocxPaintLineV1 } | undefined {
  for (const page of pages) {
    const line = page.lines.find(entry => entry.region === 'body' && entry.paragraph_id === paragraphID && entry.source_line_ordinal === 0)
    if (line) return { page, line }
  }
  return undefined
}

const ORDINARY = new Set(['fill_glyph_path', 'fill_text_highlight', 'stroke_text_underline', 'paint_inline_image', 'stroke_note_separator'])

/** Insert chart paint without disturbing the relative order of existing
 * commands. Ordinary (glyph) commands are re-sequenced so the replay order
 * still equals the concatenated line command ids; a chart's fills and strokes
 * go immediately before its own first glyph so its labels stay visible, or
 * before the front floats when it has no text. Behind-text charts paint right
 * after the behind floats. */
function insertCommands(page: NativeDocxPaintPageV1, entries: Array<{ behind: boolean; commands: NativeDocxPagePaintCommandV1[] }>, removed: Set<string>): void {
  const kept = page.commands.filter(command => !removed.has(command.id))
  if (entries.length === 0 && kept.length === page.commands.length) return
  const ordinary = new Map<string, NativeDocxPagePaintCommandV1>()
  const others: NativeDocxPagePaintCommandV1[] = []
  let firstOrdinary = -1
  for (const command of kept) {
    if (ORDINARY.has(command.kind)) { ordinary.set(command.id, command); if (firstOrdinary < 0) firstOrdinary = others.length } else others.push(command)
  }
  for (const entry of entries) for (const command of entry.commands) if (ORDINARY.has(command.kind)) ordinary.set(command.id, command)
  const behindFloatsEnd = others.findIndex(command => !(command.kind === 'paint_floating_image' && command.layer === 'behind'))
  const frontFloatsStart = others.findIndex(command => command.kind === 'paint_floating_image' && command.layer === 'front')
  const behindAt = behindFloatsEnd < 0 ? others.length : behindFloatsEnd
  const frontAt = frontFloatsStart < 0 ? others.length : frontFloatsStart
  const sequenced: NativeDocxPagePaintCommandV1[] = []
  for (const line of page.lines) for (const id of line.command_ids) { const command = ordinary.get(id); if (command) sequenced.push(command) }
  const behind: NativeDocxPagePaintCommandV1[] = [], front: NativeDocxPagePaintCommandV1[] = []
  for (const entry of entries) {
    const paint = entry.commands.filter(command => !ORDINARY.has(command.kind))
    if (entry.behind) { behind.push(...paint); continue }
    const glyphIDs = new Set(entry.commands.filter(command => ORDINARY.has(command.kind)).map(command => command.id))
    const at = sequenced.findIndex(command => glyphIDs.has(command.id))
    if (at < 0) front.push(...paint)
    else sequenced.splice(at, 0, ...paint)
  }
  const firstBorder = others.findIndex(command => command.kind === 'stroke_table_border')
  const ordinaryAt = firstOrdinary >= 0 ? firstOrdinary : Math.max(behindAt, Math.min(frontAt, firstBorder < 0 ? frontAt : firstBorder))
  const output: NativeDocxPagePaintCommandV1[] = []
  for (const [index, command] of others.entries()) {
    if (index === behindAt) output.push(...behind)
    if (index === ordinaryAt) output.push(...sequenced)
    if (index === frontAt) output.push(...front)
    output.push(command)
  }
  if (behindAt >= others.length) output.push(...behind)
  if (ordinaryAt >= others.length) output.push(...sequenced)
  if (frontAt >= others.length) output.push(...front)
  page.commands = output
}

interface ShapedText { segment: ShapedSegment; face: NativeDocxContentAddressedFaceV1; fontSize: number; advance: number; ascent: number; descent: number }

/** Shapes chart labels with the injected resolver/shaper and paints glyph paths.
 * Unavailable families fall back to an already loaded manifest face (host faces
 * first) and the substitution is disclosed; nothing is ever invented. */
class ChartTextPainter {
  private readonly resources = new Map<string, Promise<{ resource: FontResource; face: ResolvedFontFace } | undefined>>()
  private readonly shaped = new Map<string, Promise<ShapedText | undefined>>()
  private readonly outlines = new Map<string, NativeDocxGlyphOutlineResultV1>()
  // One shared outline table per page, so repeated label glyphs are transported once.
  private readonly registries = new Map<string, ReturnType<typeof nativeDocxPageGlyphOutlineRegistryV1>>()
  private readonly substituted = new Set<string>()
  glyphBudget = MAX_CHART_GLYPHS
  dropped = 0
  failures = new Set<string>()
  constructor(private readonly runtime: NativeDocxApproximateChartPaintRuntimeV1, private readonly result: NativeDocxApproximateChartPaintResultV1) {}

  private familyLoaded(family: string, weight: number, style: string): boolean {
    return this.runtime.manifest.faces.some(face => face.weight === weight && face.style === style && face.stretch === 100 && !!face.source.contentDigest && [face.family, ...(face.aliases ?? [])].some(name => asciiLowerNative(name) === asciiLowerNative(family)))
  }

  private selectFamily(font: NativeDocxApproximateChartFontV1): { family: string; weight: number; style: 'normal' | 'italic' } | undefined {
    const weight = font.bold ? 700 : 400, style = font.italic ? 'italic' : 'normal'
    if (font.family && this.familyLoaded(font.family, weight, style)) return { family: font.family, weight, style }
    const candidates = this.runtime.manifest.faces.filter(face => !!face.source.contentDigest && face.source.kind !== 'system' && face.stretch === 100).sort((left, right) => Number(right.source.kind === 'host') - Number(left.source.kind === 'host'))
    const substitute = candidates.find(face => face.weight === weight && face.style === style) ?? candidates.find(face => face.style === style) ?? candidates.find(face => face.weight === weight) ?? candidates[0]
    if (!substitute) return undefined
    const key = `${asciiLowerNative(font.family)}\0${weight}\0${style}`
    if (!this.substituted.has(key)) { this.substituted.add(key); this.result.substitutions.push({ source_family: font.family || '(theme default unavailable)', selected_family: substitute.family, face_id: substitute.faceId, weight, style }) }
    return { family: substitute.family, weight: substitute.weight, style: substitute.style === 'italic' ? 'italic' : 'normal' }
  }

  private run(text: string, family: string, weight: number, style: 'normal' | 'italic', fontSize: number): TextRunInput {
    let script: TextRunInput['script'] = 'Latn'
    for (const character of text) {
      const value = unicode13Script(character.codePointAt(0)!)
      if (value && value !== 'Zyyy' && value !== 'Zinh' && value !== 'Other') { script = value; break }
    }
    return { version: 1, text, fontSizeMilliPoints: fontSize, font: { families: [family], weight, style, stretch: 100 }, script, language: 'und', direction: 'ltr' }
  }

  private resource(run: TextRunInput): Promise<{ resource: FontResource; face: ResolvedFontFace } | undefined> {
    const key = `${run.font.families[0]}\0${run.font.weight}\0${run.font.style}`
    let pending = this.resources.get(key)
    if (!pending) {
      pending = (async () => {
        const resolution = await this.runtime.resolver.resolve({ manifest: this.runtime.manifest, run })
        if (resolution.status !== 'resolved') return undefined
        const resource = await this.runtime.resolver.load(resolution.face)
        if ('status' in resource && resource.status === 'refused') return undefined
        return { resource: resource as FontResource, face: resolution.face }
      })().catch(() => undefined)
      this.resources.set(key, pending)
    }
    return pending
  }

  /** Shape a label; undefined when no loaded face can shape it. */
  shape(text: string, font: NativeDocxApproximateChartFontV1): Promise<ShapedText | undefined> {
    const selection = this.selectFamily(font)
    if (!selection || !text) return Promise.resolve(undefined)
    const fontSize = font.size_hundredth_pt * 10
    const key = `${selection.family}\0${selection.weight}\0${selection.style}\0${fontSize}\0${text}`
    let pending = this.shaped.get(key)
    if (!pending) {
      pending = (async () => {
        const run = this.run(text, selection.family, selection.weight, selection.style, fontSize)
        const loaded = await this.resource(run)
        if (!loaded) { this.failures.add(`${font.family || 'default'}: no loaded face`); return undefined }
        const segment = await this.runtime.shaper.shape({ run, startUtf16: 0, endUtf16: text.length, font: loaded.resource })
        if ('status' in segment && segment.status === 'refused') { this.failures.add(`${font.family || 'default'}: shaping refused`); return undefined }
        const shaped = segment as ShapedSegment
        if (!Number.isSafeInteger(shaped.advanceInlineMilliPoints) || shaped.advanceInlineMilliPoints < 0 || shaped.glyphs.length > 4096) return undefined
        const face: NativeDocxContentAddressedFaceV1 = { face_id: loaded.face.faceId, content_digest: loaded.face.contentDigest, ...(loaded.face.collectionIndex !== undefined ? { collection_index: loaded.face.collectionIndex } : {}) }
        return { segment: shaped, face, fontSize, advance: shaped.advanceInlineMilliPoints, ascent: Math.max(0, shaped.metrics.ascentMilliPoints), descent: Math.max(0, shaped.metrics.descentMilliPoints) }
      })().catch(() => undefined)
      this.shaped.set(key, pending)
    }
    return pending
  }

  /** Glyph paths for one shaped label at pen origin `x` and `baseline`. */
  async paint(shaped: ShapedText, x: number, baseline: number, rgb: string, page: NativeDocxPaintPageV1, line: NativeDocxPaintLineV1, fragmentID: string, sourceID: string): Promise<NativeDocxFillGlyphPathCommandV1[]> {
    const registry = this.registries.get(page.id) ?? nativeDocxPageGlyphOutlineRegistryV1(page)
    this.registries.set(page.id, registry)
    const commands: NativeDocxFillGlyphPathCommandV1[] = []
    let penX = x
    for (const [glyphIndex, glyph] of shaped.segment.glyphs.entries()) {
      if (this.glyphBudget <= 0) { this.dropped += 1; continue }
      const cacheKey = `${shaped.face.content_digest}\0${shaped.face.collection_index ?? ''}\0${glyph.glyphId}`
      let outline = this.outlines.get(cacheKey)
      if (!outline) {
        const live = await this.runtime.outlineProvider.getGlyphOutline(Object.freeze({ face: Object.freeze({ ...shaped.face }), glyph_id: glyph.glyphId }))
        outline = nativeDocxCaptureGlyphOutlineV1(structuredClone(live), shaped.face, glyph.glyphId)
        if (!outline) break
        this.outlines.set(cacheKey, outline)
      }
      if (outline.status === 'outlined') {
        const placement = nativeDocxRegisterGlyphOutlineV1(registry, shaped.face, glyph.glyphId, shaped.fontSize, Math.round(penX + glyph.offsetXMilliPoints), Math.round(baseline - glyph.offsetYMilliPoints), { kind: 'path', path: outline.path, units_per_em: outline.units_per_em, scale_x: shaped.fontSize, scale_y: shaped.fontSize })
        if (placement) {
          this.glyphBudget -= 1
          commands.push({ kind: 'fill_glyph_path', id: paintCommandID(line.placed_line_id, fragmentID, glyphIndex), line_id: line.line_id, fragment_id: fragmentID, source_id: sourceID, glyph_index: glyphIndex, face: shaped.face, glyph_id: glyph.glyphId, font_size_millipoints: shaped.fontSize, fill_rgb: rgb, fill_rule: 'nonzero', outline_kind: 'path', ...placement })
        }
      }
      penX += glyph.advanceXMilliPoints
    }
    return commands
  }
}

interface Rect { x: number; y: number; width: number; height: number }
interface Scale { min: number; max: number; unit: number; derived: boolean }

/** Excel-like automatic scale: zero-anchored when the data allows, 1/2/5 major
 * units aimed at roughly six intervals, five percent headroom. Disclosed as a
 * host derivation whenever the source did not author the bounds. */
function valueScale(axis: NativeDocxApproximateChartAxisV1, values: number[]): Scale {
  let dataMin = Math.min(0, ...values), dataMax = Math.max(0, ...values)
  if (!Number.isFinite(dataMin) || !Number.isFinite(dataMax)) { dataMin = 0; dataMax = 1 }
  if (dataMax === dataMin) dataMax = dataMin + 1
  const authoredMin = axis.min !== undefined ? Number(axis.min) : undefined
  const authoredMax = axis.max !== undefined ? Number(axis.max) : undefined
  const range = (authoredMax ?? dataMax) - (authoredMin ?? dataMin)
  let unit = axis.major_unit !== undefined ? Number(axis.major_unit) : niceUnit(range / 6)
  if (!(unit > 0)) unit = niceUnit(range / 6)
  let min = authoredMin ?? (dataMin < 0 ? Math.floor((dataMin - 0.05 * (dataMax - dataMin)) / unit) * unit : 0)
  let max = authoredMax ?? (dataMax > 0 ? Math.ceil((dataMax + 0.05 * (dataMax - dataMin)) / unit) * unit : 0)
  if (max <= min) max = min + unit
  // Bound the tick count against pathological authored units.
  while ((max - min) / unit > MAX_TICKS) unit *= 2
  return { min, max, unit, derived: authoredMin === undefined || authoredMax === undefined || axis.major_unit === undefined }
}

function niceUnit(raw: number): number {
  if (!(raw > 0) || !Number.isFinite(raw)) return 1
  const magnitude = 10 ** Math.floor(Math.log10(raw))
  const normalized = raw / magnitude
  return magnitude * (normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10)
}

function formatTick(value: number): string {
  const rounded = Number(value.toPrecision(12))
  return Object.is(rounded, -0) ? '0' : String(rounded)
}

const DASH_PATTERNS: Record<string, readonly number[]> = { solid: [], dot: [1, 3], dash: [4, 3], lgDash: [8, 3], dashDot: [4, 3, 1, 3], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDash: [3, 1], sysDot: [1, 1], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1] }

/** Axis-aligned stroke primitives, optionally segmented per the ECMA-376 preset
 * dash pattern in multiples of the line width. Coordinates are clamped to the page. */
function strokeCommands(idBase: string, chartID: string, cellID: string, edge: NativeDocxStrokeTableBorderCommandV1['edge'], x1: number, y1: number, x2: number, y2: number, line: NativeDocxApproximateShapeLineV1, page: NativeDocxPaintPageV1, budget: { segments: number; solidFallback: boolean }): NativeDocxStrokeTableBorderCommandV1[] {
  if (![x1, y1, x2, y2].every(Number.isFinite)) throw new RangeError('non-finite chart geometry')
  const width = Math.max(1, toMillipoints(line.width_emu))
  const clampX = (value: number) => Math.min(Math.max(Math.round(value), 0), page.width_millipoints)
  const clampY = (value: number) => Math.min(Math.max(Math.round(value), 0), page.height_millipoints)
  const make = (index: number, ax: number, ay: number, bx: number, by: number): NativeDocxStrokeTableBorderCommandV1 | undefined => {
    const command: NativeDocxStrokeTableBorderCommandV1 = { kind: 'stroke_table_border', id: `${idBase}:${index}`, table_id: DOCX_APPROXIMATE_DRAWING_CHART_TABLE_ID, row_id: chartID, cell_id: cellID, edge, x1_millipoints: clampX(ax), y1_millipoints: clampY(ay), x2_millipoints: clampX(bx), y2_millipoints: clampY(by), width_millipoints: width, stroke_rgb: line.rgb }
    return command.x1_millipoints === command.x2_millipoints && command.y1_millipoints === command.y2_millipoints ? undefined : command
  }
  const pattern = DASH_PATTERNS[line.dash] ?? []
  const horizontal = y1 === y2
  const length = horizontal ? Math.abs(x2 - x1) : Math.abs(y2 - y1)
  if (pattern.length === 0 || budget.solidFallback) { const single = make(0, x1, y1, x2, y2); return single ? [single] : [] }
  const period = pattern.reduce((sum, value) => sum + value, 0) * width
  const needed = Math.ceil(length / period) * (pattern.length / 2)
  if (budget.segments + needed > MAX_DASH_SEGMENTS) { budget.solidFallback = true; const single = make(0, x1, y1, x2, y2); return single ? [single] : [] }
  const commands: NativeDocxStrokeTableBorderCommandV1[] = []
  const startX = Math.min(x1, x2), startY = Math.min(y1, y2)
  let offset = 0, index = 0
  while (offset < length) {
    for (let step = 0; step < pattern.length && offset < length; step += 2) {
      const dash = Math.min(pattern[step]! * width, length - offset)
      const command = horizontal ? make(index, startX + offset, y1, startX + offset + dash, y1) : make(index, x1, startY + offset, x1, startY + offset + dash)
      if (command) { commands.push(command); index += 1 }
      offset += dash + pattern[step + 1]! * width
    }
  }
  budget.segments += commands.length
  return commands
}

function rectOutline(idBase: string, chartID: string, cellID: string, rect: Rect, line: NativeDocxApproximateShapeLineV1, page: NativeDocxPaintPageV1, budget: { segments: number; solidFallback: boolean }): NativeDocxStrokeTableBorderCommandV1[] {
  const right = rect.x + rect.width, bottom = rect.y + rect.height
  return [
    ...strokeCommands(`${idBase}:top`, chartID, cellID, 'top', rect.x, rect.y, right, rect.y, line, page, budget),
    ...strokeCommands(`${idBase}:right`, chartID, cellID, 'right', right, rect.y, right, bottom, line, page, budget),
    ...strokeCommands(`${idBase}:bottom`, chartID, cellID, 'bottom', rect.x, bottom, right, bottom, line, page, budget),
    ...strokeCommands(`${idBase}:left`, chartID, cellID, 'left', rect.x, rect.y, rect.x, bottom, line, page, budget),
  ]
}

function fillCommand(id: string, chartID: string, cellID: string, rect: Rect, rgb: string, page: NativeDocxPaintPageV1): NativeDocxFillTableCellCommandV1 | undefined {
  if (![rect.x, rect.y, rect.width, rect.height].every(Number.isFinite)) throw new RangeError('non-finite chart geometry')
  const left = Math.min(Math.max(Math.round(rect.x), 0), page.width_millipoints), top = Math.min(Math.max(Math.round(rect.y), 0), page.height_millipoints)
  const right = Math.min(Math.max(Math.round(rect.x + rect.width), 0), page.width_millipoints), bottom = Math.min(Math.max(Math.round(rect.y + rect.height), 0), page.height_millipoints)
  if (right <= left || bottom <= top) return undefined
  return { kind: 'fill_table_cell', id, table_id: DOCX_APPROXIMATE_DRAWING_CHART_TABLE_ID, row_id: chartID, cell_id: cellID, x_millipoints: left, y_millipoints: top, width_millipoints: right - left, height_millipoints: bottom - top, fill_rgb: rgb }
}

/** Host chart layout inside the placed frame; every command is clamped to the page. */
async function paintChart(entry: PlacedChart, ordinal: number, text: ChartTextPainter, notes: Set<string>): Promise<NativeDocxPagePaintCommandV1[]> {
  const { chart: item, page, line } = entry
  const model = item.chart!
  const chartID = item.id
  const frame: Rect = { x: entry.x, y: entry.y, width: entry.width, height: entry.height }
  const commands: NativeDocxPagePaintCommandV1[] = []
  const glyphs: NativeDocxFillGlyphPathCommandV1[] = []
  const budget = { segments: 0, solidFallback: false }
  let fragment = 0
  const label = async (shaped: ShapedText | undefined, x: number, baseline: number, rgb: string) => {
    if (!shaped) return
    const painted = await text.paint(shaped, x, baseline, rgb, page, line, `ch${ordinal}.${fragment++}`, chartID)
    glyphs.push(...painted)
  }
  const lineHeight = (shaped: ShapedText | undefined, font: NativeDocxApproximateChartFontV1) => shaped ? shaped.ascent + shaped.descent : Math.round(font.size_hundredth_pt * 12)
  // Chart area.
  if (model.area_fill_rgb) { const fill = fillCommand(`${chartID}:area`, chartID, 'area', frame, model.area_fill_rgb, page); if (fill) commands.push(fill) }
  if (model.area_line) commands.push(...rectOutline(`${chartID}:area-outline`, chartID, 'area-outline', frame, model.area_line, page, budget))
  let top = frame.y + OUTER_PAD, bottom = frame.y + frame.height - OUTER_PAD, left = frame.x + OUTER_PAD, right = frame.x + frame.width - OUTER_PAD
  // Title band.
  if (model.title && !model.title.overlay) {
    const shaped = model.title.text ? await text.shape(model.title.text, model.title.font) : undefined
    const height = lineHeight(shaped, model.title.font)
    if (shaped) await label(shaped, frame.x + (frame.width - shaped.advance) / 2, top + shaped.ascent, model.title.font.rgb)
    top += height + TITLE_GAP
  } else if (model.title?.text) {
    const shaped = await text.shape(model.title.text, model.title.font)
    if (shaped) await label(shaped, frame.x + (frame.width - shaped.advance) / 2, top + shaped.ascent, model.title.font.rgb)
  }
  const column = model.bar_direction === 'column'
  const ordered = [...model.series].sort((a, b) => a.order - b.order)
  const numeric = ordered.flatMap(series => series.values.map(value => value === '' ? Number.NaN : Number(value))).filter(value => Number.isFinite(value))
  const scale = valueScale(model.value_axis, numeric)
  if (!Number.isFinite(scale.min) || !Number.isFinite(scale.max) || !Number.isFinite(scale.unit) || !Number.isFinite(scale.max - scale.min) || scale.max <= scale.min || scale.unit <= 0) throw new RangeError('chart value scale is not finite')
  if (scale.derived) notes.add(`value axis ${(['min', 'max', 'major_unit'] as const).filter(key => model.value_axis[key] === undefined).map(key => key.replace('_', ' ')).join('/')} derived by the host from cached values (not authored)`)
  const ticks: number[] = []
  for (let value = scale.min, index = 0; value <= scale.max + scale.unit * 1e-9 && index <= MAX_TICKS; value += scale.unit, index += 1) ticks.push(Number(value.toPrecision(12)))
  if (model.value_axis.number_format !== 'General' && model.value_axis.labels && !model.value_axis.deleted) notes.add(`value axis number format ${JSON.stringify(model.value_axis.number_format)} approximated as General`)
  // Legend measurements.
  const legend = model.legend
  const legendEntries: Array<{ series: NativeDocxApproximateChartSeriesV1; shaped: ShapedText | undefined }> = []
  if (legend) for (const series of ordered) legendEntries.push({ series, shaped: series.title ? await text.shape(series.title, legend.font) : undefined })
  const swatch = legend ? Math.max(1, Math.round(legend.font.size_hundredth_pt * 6)) : 0
  const legendLine = legend ? Math.max(swatch, ...legendEntries.map(entry => lineHeight(entry.shaped, legend.font))) : 0
  const legendEntryWidth = (entry: { shaped: ShapedText | undefined }) => swatch + LEGEND_SWATCH_GAP + (entry.shaped?.advance ?? 0)
  let legendRect: Rect | undefined
  if (legend && !legend.overlay) {
    if (legend.position === 'b') { legendRect = { x: left, y: bottom - legendLine, width: right - left, height: legendLine }; bottom -= legendLine + LEGEND_GAP }
    else if (legend.position === 't') { legendRect = { x: left, y: top, width: right - left, height: legendLine }; top += legendLine + LEGEND_GAP }
    else {
      const width = Math.max(...legendEntries.map(legendEntryWidth), 1)
      const height = legendEntries.length * legendLine
      if (legend.position === 'l') { legendRect = { x: left, y: top, width, height }; left += width + LEGEND_GAP } else { legendRect = { x: right - width, y: top, width, height }; right -= width + LEGEND_GAP }
    }
  }
  // Axis label measurements.
  const categoryLabels = model.category_axis.deleted || !model.category_axis.labels ? [] : await Promise.all(model.categories.map(category => category ? text.shape(category, model.category_axis.labels!) : Promise.resolve(undefined)))
  const valueLabels = model.value_axis.deleted || !model.value_axis.labels ? [] : await Promise.all(ticks.map(tick => text.shape(formatTick(tick), model.value_axis.labels!)))
  const categoryFont = model.category_axis.labels, valueFont = model.value_axis.labels
  const categoryBand = categoryLabels.length ? Math.max(...categoryLabels.map(shaped => lineHeight(shaped, categoryFont!))) + LABEL_GAP : 0
  const valueBand = valueLabels.length ? Math.max(0, ...valueLabels.map(shaped => shaped?.advance ?? 0)) + LABEL_GAP : 0
  const categoryMaxWidth = categoryLabels.length ? Math.max(0, ...categoryLabels.map(shaped => shaped?.advance ?? 0)) : 0
  const valueBandHeight = valueLabels.length ? Math.max(...valueLabels.map(shaped => lineHeight(shaped, valueFont!))) + LABEL_GAP : 0
  if (column) { bottom -= categoryBand; left += valueBand } else { left += categoryMaxWidth ? categoryMaxWidth + LABEL_GAP : 0; bottom -= valueBandHeight }
  const plot: Rect = { x: Math.round(left), y: Math.round(top), width: Math.round(right - left), height: Math.round(bottom - top) }
  if (plot.width < 1 || plot.height < 1) throw new RangeError('chart frame leaves no plot area')
  if (model.plot_fill_rgb) { const fill = fillCommand(`${chartID}:plot`, chartID, 'plot', plot, model.plot_fill_rgb, page); if (fill) commands.push(fill) }
  // Value coordinate along the value axis; reversed orientation flips it.
  const valueReverse = model.value_axis.orientation === 'maxMin'
  const valueExtent = column ? plot.height : plot.width
  const valueCoordinate = (value: number) => {
    const clamped = Math.min(Math.max(value, scale.min), scale.max)
    const ratio = (clamped - scale.min) / (scale.max - scale.min)
    const along = ratio * valueExtent
    return column ? (valueReverse ? plot.y + along : plot.y + plot.height - along) : (valueReverse ? plot.x + plot.width - along : plot.x + along)
  }
  // Gridlines and the value axis line sit under the bars.
  if (model.value_axis.major_gridlines && !model.value_axis.deleted) for (const [index, tick] of ticks.entries()) {
    const at = valueCoordinate(tick)
    commands.push(...(column ? strokeCommands(`${chartID}:grid:${index}`, chartID, 'gridline', 'top', plot.x, at, plot.x + plot.width, at, model.value_axis.major_gridlines, page, budget) : strokeCommands(`${chartID}:grid:${index}`, chartID, 'gridline', 'left', at, plot.y, at, plot.y + plot.height, model.value_axis.major_gridlines, page, budget)))
  }
  // Bars: Office clustered geometry with gap width and (negative) overlap in
  // percent of one bar width (ECMA-376 21.2.2.75 gapWidth, 21.2.2.131 overlap).
  const count = model.categories.length, seriesCount = ordered.length
  const categoryExtent = column ? plot.width : plot.height
  const slot = categoryExtent / count
  const barWidth = slot / (model.gap_width_percent / 100 + seriesCount - (seriesCount - 1) * model.overlap_percent / 100)
  const stride = barWidth * (1 - model.overlap_percent / 100)
  const categoryReverse = column ? model.category_axis.orientation === 'maxMin' : model.category_axis.orientation === 'minMax'
  const baseline = valueCoordinate(Math.min(Math.max(0, scale.min), scale.max))
  if (barWidth * 2 < 1) throw new RangeError('chart frame cannot retain distinct bars')
  const bars: Array<{ series: NativeDocxApproximateChartSeriesV1; rect: Rect; seriesOrdinal: number; point: number }> = []
  for (let point = 0; point < count; point += 1) {
    const slotStart = categoryReverse ? categoryExtent - (point + 1) * slot : point * slot
    const clusterStart = slotStart + (model.gap_width_percent / 100) * barWidth / 2
    for (const [seriesOrdinal, series] of ordered.entries()) {
      const raw = series.values[point]
      if (raw === undefined || raw === '') continue
      const value = Number(raw)
      if (!Number.isFinite(value)) continue
      const a = clusterStart + seriesOrdinal * stride, b = a + barWidth
      const end = valueCoordinate(value)
      const near = Math.min(baseline, end), far = Math.max(baseline, end)
      const rect: Rect = column ? { x: plot.x + a, y: near, width: b - a, height: far - near } : { x: near, y: plot.y + a, width: far - near, height: b - a }
      bars.push({ series, rect, seriesOrdinal, point })
    }
  }
  for (const bar of bars) {
    const fill = fillCommand(`${chartID}:bar:${bar.series.index}:${bar.point}`, chartID, `bar:${bar.series.index}:${bar.point}`, bar.rect, bar.series.fill_rgb, page)
    if (fill) commands.push(fill)
  }
  for (const bar of bars) if (bar.series.line) commands.push(...rectOutline(`${chartID}:bar-outline:${bar.series.index}:${bar.point}`, chartID, `bar-outline:${bar.series.index}:${bar.point}`, bar.rect, bar.series.line, page, budget))
  if (bars.some(bar => bar.series.line && (DASH_PATTERNS[bar.series.line.dash] ?? []).length)) notes.add(budget.solidFallback ? 'dashed series outlines painted solid beyond the segment budget' : 'dashed series outlines segmented per the preset pattern with flat caps')
  // Axis lines.
  if (!model.category_axis.deleted && model.category_axis.line) commands.push(...(column ? strokeCommands(`${chartID}:category-axis`, chartID, 'category-axis', 'top', plot.x, baseline, plot.x + plot.width, baseline, model.category_axis.line, page, budget) : strokeCommands(`${chartID}:category-axis`, chartID, 'category-axis', 'left', baseline, plot.y, baseline, plot.y + plot.height, model.category_axis.line, page, budget)))
  if (!model.value_axis.deleted && model.value_axis.line) {
    const at = categoryReverse ? (column ? plot.x + plot.width : plot.y + plot.height) : (column ? plot.x : plot.y)
    commands.push(...(column ? strokeCommands(`${chartID}:value-axis`, chartID, 'value-axis', 'left', at, plot.y, at, plot.y + plot.height, model.value_axis.line, page, budget) : strokeCommands(`${chartID}:value-axis`, chartID, 'value-axis', 'top', plot.x, at, plot.x + plot.width, at, model.value_axis.line, page, budget)))
  }
  if (model.plot_line) commands.push(...rectOutline(`${chartID}:plot-outline`, chartID, 'plot-outline', plot, model.plot_line, page, budget))
  // Category labels centred on their slot; crowded labels are thinned like Office.
  if (categoryLabels.length && categoryFont) {
    const step = column && categoryMaxWidth > slot ? Math.ceil(categoryMaxWidth / slot) : 1
    if (step > 1) notes.add('crowded category labels thinned to fit their slots')
    for (let point = 0; point < count; point += step) {
      const shaped = categoryLabels[point]
      if (!shaped) continue
      const slotStart = categoryReverse ? categoryExtent - (point + 1) * slot : point * slot
      if (column) await label(shaped, plot.x + slotStart + (slot - shaped.advance) / 2, plot.y + plot.height + LABEL_GAP + shaped.ascent, categoryFont.rgb)
      else await label(shaped, plot.x - LABEL_GAP - shaped.advance, plot.y + slotStart + slot / 2 + (shaped.ascent - shaped.descent) / 2, categoryFont.rgb)
    }
  }
  // Value labels at each tick.
  if (valueLabels.length && valueFont) for (const [index, tick] of ticks.entries()) {
    const shaped = valueLabels[index]
    if (!shaped) continue
    const at = valueCoordinate(tick)
    if (column) await label(shaped, plot.x - LABEL_GAP - shaped.advance, at + (shaped.ascent - shaped.descent) / 2, valueFont.rgb)
    else await label(shaped, at - shaped.advance / 2, plot.y + plot.height + LABEL_GAP + shaped.ascent, valueFont.rgb)
  }
  // Legend: swatches in series fill colour followed by the cached series name.
  if (legend && legendRect) {
    const horizontalLegend = legend.position === 'b' || legend.position === 't'
    if (horizontalLegend) {
      const total = legendEntries.reduce((sum, entry) => sum + legendEntryWidth(entry), 0) + LEGEND_ENTRY_GAP * Math.max(0, legendEntries.length - 1)
      let x = legendRect.x + (legendRect.width - total) / 2
      for (const entry of legendEntries) {
        const swatchRect: Rect = { x, y: legendRect.y + (legendLine - swatch) / 2, width: swatch, height: swatch }
        const fill = fillCommand(`${chartID}:legend:${entry.series.index}`, chartID, `legend:${entry.series.index}`, swatchRect, entry.series.fill_rgb, page)
        if (fill) commands.push(fill)
        if (entry.shaped) await label(entry.shaped, x + swatch + LEGEND_SWATCH_GAP, legendRect.y + (legendLine + entry.shaped.ascent - entry.shaped.descent) / 2, legend.font.rgb)
        x += legendEntryWidth(entry) + LEGEND_ENTRY_GAP
      }
    } else {
      let y = plot.y + Math.max(0, (plot.height - legendRect.height) / 2)
      for (const entry of legendEntries) {
        const swatchRect: Rect = { x: legendRect.x, y: y + (legendLine - swatch) / 2, width: swatch, height: swatch }
        const fill = fillCommand(`${chartID}:legend:${entry.series.index}`, chartID, `legend:${entry.series.index}`, swatchRect, entry.series.fill_rgb, page)
        if (fill) commands.push(fill)
        if (entry.shaped) await label(entry.shaped, legendRect.x + swatch + LEGEND_SWATCH_GAP, y + (legendLine + entry.shaped.ascent - entry.shaped.descent) / 2, legend.font.rgb)
        y += legendLine
      }
    }
  } else if (legend?.overlay) notes.add('overlaid legend is omitted')
  for (const glyph of glyphs) line.command_ids.push(glyph.id)
  return [...commands, ...glyphs]
}

function buildReasons(charts: NativeDocxApproximateDrawingChartsV1, result: NativeDocxApproximateChartPaintResultV1, notes: string[], text: ChartTextPainter): string[] {
  const reasons: string[] = []
  if (result.painted.length) {
    reasons.push(DOCX_APPROXIMATE_DRAWING_CHART_WARNING)
    const painted = charts.items.filter(chart => result.painted.includes(chart.id))
    reasons.push(`${DOCX_APPROXIMATE_DRAWING_CHART_CODE}: painted ${result.painted.length} of ${charts.items.length} refused DrawingML charts (${painted.map(chart => `${chart.id} ${chart.placement} clustered ${chart.chart!.bar_direction} chart, ${chart.chart!.categories.length} categories x ${chart.chart!.series.length} series from ${chart.chart_part} for diagnostics ${chart.diagnostic_ids.join(',')}`).join('; ')})${notes.length ? `; approximations: ${notes.join('; ')}` : ''}`.slice(0, 8000))
  }
  if (result.omitted.length || charts.omitted_count) reasons.push(`${DOCX_APPROXIMATE_DRAWING_CHART_OMITTED_CODE}: ${result.omitted.map(entry => `${entry.id} (${entry.reason})`).join('; ')}${charts.omitted_count ? `; ${charts.omitted_count} charts beyond the ${MAX_CHARTS} chart budget` : ''}`.slice(0, 8000))
  for (const substitution of result.substitutions.slice(0, MAX_REASONS - 4)) reasons.push(`${DOCX_APPROXIMATE_CHART_FONT_CODE}: ${substitution.source_family} / ${substitution.weight} / ${substitution.style} -> ${substitution.selected_family} (loaded face ${substitution.face_id}); chart text metrics may differ`)
  if (text.dropped || text.failures.size) reasons.push(`${DOCX_APPROXIMATE_DRAWING_CHART_OMITTED_CODE}: chart text partially omitted (${text.dropped} glyphs beyond the preview budget${text.failures.size ? `, ${[...text.failures].join('; ')}` : ''})`.slice(0, 8000))
  return reasons.slice(0, MAX_REASONS)
}
