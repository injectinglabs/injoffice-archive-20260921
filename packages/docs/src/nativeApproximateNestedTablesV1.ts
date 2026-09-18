/** Approximate nested tables (one level) for the read-only approximate page
 * preview.
 *
 * The strict extractor keeps refusing tables nested inside table cells
 * (NESTED_TABLE_OR_CELL_MARKUP). The Go sidecar
 * (`InspectNativeApproximateNestedTablesV1`) joins those refusals to their
 * source nodes and describes the inner table through the ordinary extractor and
 * style resolver. This module validates that sidecar against the current
 * document, lays each supported inner table out inside the containing cell's
 * content box (authored tblGrid/tcW columns fitted to the cell, rows sized from
 * shaped content), reserves the extent as extra paragraph spacing on a
 * neighbouring cell paragraph in the internal body copy, and after body
 * pagination appends table paint: fills and borders reuse the table primitives,
 * cell glyphs attach to the anchor paragraph's paint line. Strict paint, source
 * bytes and the original diagnostics are unchanged. */
import type { NativeFontManifest, NativeFontResolver, NativeTextShaper } from '@injoffice/font-metrics/layout'
import type { NativeDocxDocumentV1, NativeDocxParagraphV1, NativeDocxSourceAnchorV1, NativeDocxTableBorderV1, NativeDocxTableBordersV1, NativeDocxTableCellV1, NativeDocxTableV1, NativeDocxUnsupportedCapabilityV1 } from './nativeContract.js'
import type { NativeDocxResolvedLayoutInputV1, NativeDocxResolvedParagraphV1, NativeDocxResolvedRunPropertiesV1, NativeDocxResolvedRunV1, NativeDocxResolvedTableGeometryV1 } from './nativeResolvedLayout.js'
import type { NativeDocxPaginationSettingsV1 } from './nativePaginationSettings.js'
import type { NativeDocxLineFragmentV1, NativeDocxShapedLinesV1, NativeDocxShapedParagraphV1 } from './nativeShapingLines.js'
import { shapeNativeDocxLinesWithParagraphWidthsV1 } from './nativeShapingLines.js'
import { ID, RGB, preflightWire, paintCommandID } from './nativePagePaintWireV1.js'
import { nativeDocxPageGlyphOutlineRegistryV1, nativeDocxRegisterGlyphOutlineV1, nativeDocxCaptureGlyphOutlineV1, type NativeDocxContentAddressedFaceV1, type NativeDocxFillGlyphPathCommandV1, type NativeDocxFillTableCellCommandV1, type NativeDocxFillTextHighlightCommandV1, type NativeDocxGlyphOutlineProviderV1, type NativeDocxGlyphOutlineResultV1, type NativeDocxPagePaintCommandV1, type NativeDocxPaintLineV1, type NativeDocxPaintPageV1, type NativeDocxStrokeTableBorderCommandV1, type NativeDocxStrokeTextUnderlineCommandV1 } from './nativePagePaintV1.js'
import { nativeTextUnderlineCommandsV1 } from './nativeTextUnderlineV1.js'
import { nativeTextHighlightCommandV1 } from './nativeTextHighlightV1.js'
import { nativeDocxTableGeometryV1 } from './nativeTablePagePaintV1.js'
import { asciiLowerNative } from './nativeDeterminism.js'
import { DOCX_APPROXIMATE_OMITTED_SOURCE_UNSUPPORTED } from './nativeApproximationV1.js'
import { collectNativeDocxApproximateOmissionsV1, DOCX_APPROXIMATE_OMITTED_CONTENT_WARNING, type NativeDocxApproximateOmissionsV1 } from './nativeApproximateOmittedContentV1.js'

export const DOCX_APPROXIMATE_NESTED_TABLES_PROTOCOL = 'injoffice.docx.approximate-nested-tables' as const
export const DOCX_APPROXIMATE_NESTED_TABLE_POLICY = 'docx.approximate-nested-table-preview-v1' as const
/** Declared layout policy string for envelope reasons. */
export const DOCX_APPROXIMATE_NESTED_TABLE_LAYOUT_POLICY = 'approximate-nested-table-v1' as const
export const DOCX_APPROXIMATE_NESTED_TABLE_CODE = 'docx.approximate-nested-table-preview' as const
export const DOCX_APPROXIMATE_NESTED_TABLE_OMITTED_CODE = 'docx.approximate-nested-table-omitted' as const
export const DOCX_APPROXIMATE_NESTED_TABLE_FONT_CODE = 'docx.approximate-nested-table-substituted-font' as const
export const DOCX_APPROXIMATE_NESTED_TABLE_WARNING = `${DOCX_APPROXIMATE_NESTED_TABLE_CODE}: tables nested one level inside a table cell are laid out inside the containing cell's content box (${DOCX_APPROXIMATE_NESTED_TABLE_LAYOUT_POLICY}): authored tblGrid/tcW columns are scaled proportionally to the cell when they exceed it, rows size from shaped content, table and cell borders and shading follow direct properties over the base table style, and the containing cell grows by the reserved extent through neighbouring paragraph spacing. Only the firstRow conditional table-style region is applied (when the table look selects it); other conditional regions are not, and deeper nesting and merged cells stay omitted. Source bytes and original diagnostics are unchanged.` as const
export const DOCX_APPROXIMATE_NESTED_TABLE_SIDECAR_REFUSED = `${DOCX_APPROXIMATE_NESTED_TABLE_OMITTED_CODE}: nested-table evidence did not exact-join the source document and was not used; refused nested tables stay omitted` as const
/** Table paint primitives carry this id so consumers can tell nested-table paint from body table paint. */
export const DOCX_APPROXIMATE_NESTED_TABLE_TABLE_ID = DOCX_APPROXIMATE_NESTED_TABLE_POLICY

const MAX_ITEMS = 64
const MAX_ROWS = 256
const MAX_COLUMNS = 64
const MAX_CELL_PARAGRAPHS = 256
const MAX_GLYPHS = 100_000
/** Interactive viewers read at most 16 MiB; keep body paint plus nested-table paint under this. */
const MAX_ENVELOPE_BYTES = 15 * 1024 * 1024
const MAX_REASONS = 24
const MAX_TWIPS = 20_000_000
const HOST_DEFAULT_SIZE_HALF_POINTS = 22

export interface NativeDocxApproximateNestedTableV1 {
  id: string
  table_id: string
  cell_id: string
  diagnostic_ids: string[]
  anchor: NativeDocxSourceAnchorV1
  preceding_paragraphs: number
  status: 'supported' | 'omitted'
  reason?: string
  table?: NativeDocxTableV1
  geometry?: NativeDocxResolvedTableGeometryV1
  style_borders?: NativeDocxTableBordersV1
  style_cell_shading_rgb?: string
  /** firstRow conditional cell fill when the inner tblLook enables the region. */
  first_row_cell_shading_rgb?: string
  /** Approximate geometry of the containing table from its own style cascade. */
  outer_geometry?: NativeDocxResolvedTableGeometryV1
  resolved_paragraphs: NativeDocxResolvedParagraphV1[]
  resolved_runs: NativeDocxResolvedRunV1[]
  omitted_runs: number
  notes?: string[]
}

export interface NativeDocxApproximateNestedTablesV1 {
  protocol: typeof DOCX_APPROXIMATE_NESTED_TABLES_PROTOCOL
  version: 1
  policy: typeof DOCX_APPROXIMATE_NESTED_TABLE_POLICY
  package_sha256: string
  part_sha256: string
  items: NativeDocxApproximateNestedTableV1[]
  omitted_count: number
}

function record(value: unknown): value is Record<string, unknown> { return value !== null && typeof value === 'object' && !Array.isArray(value) }
function exactKeys(value: Record<string, unknown>, required: readonly string[], optional: readonly string[] = []): boolean {
  const keys = Object.keys(value)
  return required.every(key => key in value) && keys.every(key => required.includes(key) || optional.includes(key))
}
function safeNonnegative(value: unknown, max = MAX_TWIPS): value is number { return Number.isSafeInteger(value) && (value as number) >= 0 && (value as number) <= max }
function anchorValid(value: unknown, part: string): value is NativeDocxSourceAnchorV1 {
  return record(value) && exactKeys(value, ['part_name', 'path', 'start_byte', 'end_byte', 'xml_sha256']) && value.part_name === part && typeof value.path === 'string' && value.path.length <= 4096 && Number.isSafeInteger(value.start_byte) && Number.isSafeInteger(value.end_byte) && (value.start_byte as number) >= 0 && (value.end_byte as number) > (value.start_byte as number) && typeof value.xml_sha256 === 'string' && /^sha256:[0-9a-f]{64}$/.test(value.xml_sha256)
}
function within(inner: NativeDocxSourceAnchorV1, outer: NativeDocxSourceAnchorV1): boolean { return inner.part_name === outer.part_name && inner.start_byte >= outer.start_byte && inner.end_byte <= outer.end_byte }
function sameAnchor(left: NativeDocxSourceAnchorV1, right: NativeDocxSourceAnchorV1): boolean { return left.part_name === right.part_name && left.path === right.path && left.start_byte === right.start_byte && left.end_byte === right.end_byte && left.xml_sha256 === right.xml_sha256 }
function borderValid(border: unknown): border is NativeDocxTableBorderV1 {
  if (border === undefined) return true
  if (!record(border) || !exactKeys(border, ['style', 'size_eighth_points'], ['color_rgb'])) return false
  if (border.style === 'none') return border.size_eighth_points === 0 && border.color_rgb === undefined
  return border.style === 'single' && safeNonnegative(border.size_eighth_points, 768) && (border.size_eighth_points as number) > 0 && typeof border.color_rgb === 'string' && RGB.test(border.color_rgb)
}
function bordersValid(borders: unknown): borders is NativeDocxTableBordersV1 | undefined {
  if (borders === undefined) return true
  return record(borders) && exactKeys(borders, [], ['top', 'right', 'bottom', 'left', 'inside_horizontal', 'inside_vertical']) && Object.values(borders).every(borderValid)
}
function geometryValid(value: unknown): value is NativeDocxResolvedTableGeometryV1 {
  if (!record(value) || !exactKeys(value, ['layout', 'alignment', 'indent_twips', 'width_type', 'width_value', 'cell_margins'])) return false
  const margins = value.cell_margins
  return (value.layout === 'fixed' || value.layout === 'autofit') && value.alignment === 'left' && safeNonnegative(value.indent_twips) && ['auto', 'dxa', 'pct'].includes(String(value.width_type)) && safeNonnegative(value.width_value) && record(margins) && exactKeys(margins, ['top_twips', 'right_twips', 'bottom_twips', 'left_twips']) && Object.values(margins).every(n => safeNonnegative(n))
}

/** Bounded structural validation plus source joins: every item must name a
 * body table cell, sit between that cell's modeled paragraphs and join the
 * retained NESTED_TABLE_OR_CELL_MARKUP refusal at the identical anchor. The
 * inner table's paragraphs and runs carry ids derived from the item id so they
 * can never collide with document ids; their resolved layout is validated by
 * the shaping request decoder when the table is shaped, and a malformed inner
 * table only omits that item. */
export function decodeNativeDocxApproximateNestedTablesV1(value: unknown, document: NativeDocxDocumentV1, mainPartSha256?: string): NativeDocxApproximateNestedTablesV1 {
  if (preflightWire(value, 'approximate nested tables', 2_000_000, 100_000).length) throw new TypeError('Approximate nested tables exceed their bounded wire')
  const input = structuredClone(value) as NativeDocxApproximateNestedTablesV1
  if (!record(input) || !exactKeys(input as unknown as Record<string, unknown>, ['protocol', 'version', 'policy', 'package_sha256', 'part_sha256', 'items', 'omitted_count']) || input.protocol !== DOCX_APPROXIMATE_NESTED_TABLES_PROTOCOL || input.version !== 1 || input.policy !== DOCX_APPROXIMATE_NESTED_TABLE_POLICY || input.package_sha256 !== document.source.package_sha256 || typeof input.part_sha256 !== 'string' || !/^sha256:[0-9a-f]{64}$/.test(input.part_sha256) || (mainPartSha256 !== undefined && input.part_sha256 !== mainPartSha256) || !Array.isArray(input.items) || input.items.length > MAX_ITEMS || !safeNonnegative(input.omitted_count, 1_000_000)) throw new TypeError('Approximate nested tables do not exact-join the source document')
  const main = document.source.main_part
  const cells = new Map<string, { table: NativeDocxTableV1; cell: NativeDocxTableCellV1 }>()
  for (const block of document.body.blocks) for (const row of block.table?.rows ?? []) for (const cell of row.cells) cells.set(cell.id, { table: block.table!, cell })
  const diagnostics = new Map(document.unsupported.map(entry => [entry.id, entry]))
  const ids = new Set<string>()
  for (const item of input.items) {
    if (!record(item) || !exactKeys(item as unknown as Record<string, unknown>, ['id', 'table_id', 'cell_id', 'diagnostic_ids', 'anchor', 'preceding_paragraphs', 'status', 'resolved_paragraphs', 'resolved_runs', 'omitted_runs'], ['reason', 'table', 'geometry', 'style_borders', 'style_cell_shading_rgb', 'first_row_cell_shading_rgb', 'outer_geometry', 'notes'])) throw new TypeError('Approximate nested table has unknown or missing fields')
    if (typeof item.id !== 'string' || !ID.test(item.id) || item.id.length > 256 || ids.has(item.id)) throw new TypeError('Approximate nested table id is invalid or duplicated')
    ids.add(item.id)
    const owner = cells.get(item.cell_id)
    if (!owner || owner.table.id !== item.table_id || !anchorValid(item.anchor, main) || !within(item.anchor, owner.cell.anchor)) throw new TypeError('Approximate nested table does not exact-join its containing cell')
    if (!Array.isArray(item.diagnostic_ids) || item.diagnostic_ids.length === 0 || item.diagnostic_ids.length > 8 || item.diagnostic_ids.some(id => { const d = typeof id === 'string' ? diagnostics.get(id) : undefined; return !d || d.code !== 'NESTED_TABLE_OR_CELL_MARKUP' || d.scope_id !== item.table_id || !d.anchor || !sameAnchor(d.anchor, item.anchor) })) throw new TypeError('Approximate nested table must join the retained nested-table refusal at its anchor')
    const paragraphs = owner.cell.paragraphs
    if (!safeNonnegative(item.preceding_paragraphs, paragraphs.length)) throw new TypeError('Approximate nested table position is outside its cell paragraphs')
    const before = paragraphs[item.preceding_paragraphs - 1], after = paragraphs[item.preceding_paragraphs]
    if ((before && before.anchor.end_byte > item.anchor.start_byte) || (after && after.anchor.start_byte < item.anchor.end_byte)) throw new TypeError('Approximate nested table position does not match its cell paragraph order')
    if (item.status !== 'supported' && item.status !== 'omitted') throw new TypeError('Approximate nested table status is invalid')
    if (item.reason !== undefined && (typeof item.reason !== 'string' || item.reason.length > 256)) throw new TypeError('Approximate nested table reason is unbounded')
    if (item.notes !== undefined && (!Array.isArray(item.notes) || item.notes.length > 32 || item.notes.some(note => typeof note !== 'string' || note.length > 512))) throw new TypeError('Approximate nested table notes are unbounded')
    if (!safeNonnegative(item.omitted_runs, 1_000_000) || !Array.isArray(item.resolved_paragraphs) || !Array.isArray(item.resolved_runs)) throw new TypeError('Approximate nested table content counts are invalid')
    if (item.status === 'omitted') continue
    if ((item.first_row_cell_shading_rgb !== undefined && (typeof item.first_row_cell_shading_rgb !== 'string' || !RGB.test(item.first_row_cell_shading_rgb))) || (item.outer_geometry !== undefined && !geometryValid(item.outer_geometry))) throw new TypeError('Approximate nested table conditional fill or outer geometry is invalid')
    const table = item.table as unknown
    if (!record(table) || table.id !== item.id || !Array.isArray(table.grid_widths_twips) || table.grid_widths_twips.length === 0 || table.grid_widths_twips.length > MAX_COLUMNS || table.grid_widths_twips.some(width => !safeNonnegative(width) || width === 0) || !Array.isArray(table.rows) || table.rows.length === 0 || table.rows.length > MAX_ROWS || !geometryValid(item.geometry) || !bordersValid(table.borders) || !bordersValid(item.style_borders) || (item.style_cell_shading_rgb !== undefined && (typeof item.style_cell_shading_rgb !== 'string' || !RGB.test(item.style_cell_shading_rgb)))) throw new TypeError('Supported approximate nested table requires a bounded grid, geometry and borders')
    if (table.width_twips !== undefined && !safeNonnegative(table.width_twips) || table.indent_twips !== undefined && !safeNonnegative(table.indent_twips) || table.width_percent_fiftieths !== undefined && !safeNonnegative(table.width_percent_fiftieths, 5000)) throw new TypeError('Approximate nested table geometry is out of bounds')
    const columns = table.grid_widths_twips.length
    const paragraphIDs = new Set<string>(), runIDs = new Set<string>()
    for (const row of table.rows as unknown[]) {
      if (!record(row) || typeof row.id !== 'string' || !row.id.startsWith(`${item.id}:`) || !Array.isArray(row.cells) || row.cells.length !== columns || !anchorValid(row.anchor, main) || !within(row.anchor, item.anchor)) throw new TypeError('Approximate nested table row does not match its grid')
      if ((row.height_twips === undefined) !== (row.height_rule === undefined) || (row.height_twips !== undefined && (!safeNonnegative(row.height_twips) || (row.height_rule !== 'atLeast' && row.height_rule !== 'exact')))) throw new TypeError('Approximate nested table row height is invalid')
      for (const cell of row.cells as unknown[]) {
        if (!record(cell) || typeof cell.id !== 'string' || !cell.id.startsWith(`${row.id}`) || cell.grid_span !== 1 || cell.vertical_merge !== 'none' || !anchorValid(cell.anchor, main) || !within(cell.anchor, row.anchor) || (cell.width_twips !== undefined && !safeNonnegative(cell.width_twips)) || !bordersValid(cell.borders) || (cell.shading_rgb !== undefined && (typeof cell.shading_rgb !== 'string' || !RGB.test(cell.shading_rgb))) || !Array.isArray(cell.paragraphs) || cell.paragraphs.length > MAX_CELL_PARAGRAPHS) throw new TypeError('Approximate nested table cell is outside the supported subset')
        for (const paragraph of cell.paragraphs as unknown[]) {
          if (!record(paragraph) || typeof paragraph.id !== 'string' || !ID.test(paragraph.id) || !paragraph.id.startsWith(`${cell.id}:`) || paragraphIDs.has(paragraph.id) || !anchorValid(paragraph.anchor, main) || !within(paragraph.anchor, cell.anchor) || !Array.isArray(paragraph.runs)) throw new TypeError('Approximate nested table paragraph is invalid')
          paragraphIDs.add(paragraph.id)
          for (const run of paragraph.runs as unknown[]) {
            if (!record(run) || typeof run.id !== 'string' || !ID.test(run.id) || !run.id.startsWith(`${paragraph.id}:`) || runIDs.has(run.id) || !(run.kind === 'text' && typeof run.text === 'string' || run.kind === 'control' && (run.control === 'tab' || run.control === 'line-break'))) throw new TypeError('Approximate nested table run is outside text, tab and line-break runs')
            runIDs.add(run.id)
          }
        }
      }
    }
    const resolvedParagraphIDs = new Set<string>()
    for (const entry of item.resolved_paragraphs as unknown[]) {
      if (!record(entry) || typeof entry.paragraph_id !== 'string' || !paragraphIDs.has(entry.paragraph_id) || resolvedParagraphIDs.has(entry.paragraph_id)) throw new TypeError('Approximate nested table resolved paragraph does not join its inner paragraph')
      resolvedParagraphIDs.add(entry.paragraph_id)
    }
    if (resolvedParagraphIDs.size !== paragraphIDs.size) throw new TypeError('Approximate nested table paragraphs are not completely resolved')
    for (const entry of item.resolved_runs as unknown[]) if (!record(entry) || typeof entry.run_id !== 'string' || !runIDs.has(entry.run_id)) throw new TypeError('Approximate nested table resolved run does not join its inner run')
  }
  return input
}

export interface NativeDocxApproximateNestedTableRuntimeV1 {
  manifest: NativeFontManifest
  resolver: NativeFontResolver
  shaper: NativeTextShaper
  settings: NativeDocxPaginationSettingsV1
  /** Declared host default size applied to inner runs whose source size is absent. */
  fontSizeHalfPoints?: number
}

export interface NativeDocxApproximateNestedTableFontSubstitutionV1 { source_family: string; selected_family: string; face_id: string; weight: number; style: string; selected_weight: number; selected_style: string }

interface PreparedCell {
  cell: NativeDocxTableCellV1
  x_millipoints: number
  width_millipoints: number
  content_x_millipoints: number
  content_width_millipoints: number
  shading_rgb?: string
}
interface PreparedRow { row_id: string; height_millipoints: number; cells: PreparedCell[] }

export interface NativeDocxApproximateNestedTablePreparedV1 {
  item: NativeDocxApproximateNestedTableV1
  /** Cell paragraph whose spacing reserves the extent, and which paint line side anchors the table. */
  anchor_paragraph_id: string
  anchor_side: 'after' | 'before'
  /** Original spacing on the anchor side, in millipoints, before the reservation was added. */
  anchor_gap_millipoints: number
  indent_millipoints: number
  width_millipoints: number
  height_millipoints: number
  rows: PreparedRow[]
  borders?: NativeDocxTableBordersV1
  shaped: NativeDocxShapedLinesV1
  resolved: NativeDocxResolvedLayoutInputV1
  fitted?: { source_grid_widths_twips: number[]; fitted_grid_widths_twips: number[]; available_width_twips: number }
  notes: string[]
}

export interface NativeDocxApproximateNestedTableStageV1 {
  sidecar: NativeDocxApproximateNestedTablesV1
  prepared: NativeDocxApproximateNestedTablePreparedV1[]
  omitted: Array<{ id: string; reason: string }>
  substitutions: NativeDocxApproximateNestedTableFontSubstitutionV1[]
  omittedContent: number
  projection: { document: NativeDocxDocumentV1; resolved: NativeDocxResolvedLayoutInputV1 }
  /** Source refusals removed from the body copy per supported item; restored for items the painter drops. */
  removedDiagnostics: Map<string, NativeDocxUnsupportedCapabilityV1[]>
}

const twips = (value: number) => value * 50

/** Largest-remainder proportional scaling to an exact integer target. */
function scaleWidths(widths: readonly number[], target: number): number[] | undefined {
  const total = widths.reduce((a, b) => a + b, 0)
  if (total <= 0 || target <= 0 || !Number.isSafeInteger(total) || !Number.isSafeInteger(target)) return undefined
  const entries = widths.map((width, index) => { const numerator = BigInt(width) * BigInt(target); return { index, whole: Number(numerator / BigInt(total)), remainder: numerator % BigInt(total) } })
  const scaled = entries.map(entry => entry.whole)
  entries.sort((a, b) => a.remainder === b.remainder ? a.index - b.index : a.remainder > b.remainder ? -1 : 1)
  const left = target - scaled.reduce((a, b) => a + b, 0)
  for (let i = 0; i < left; i += 1) scaled[entries[i]!.index]! += 1
  return scaled.some(n => n <= 0) ? undefined : scaled
}

function mergeBorders(base: NativeDocxTableBordersV1 | undefined, overlay: NativeDocxTableBordersV1 | undefined): NativeDocxTableBordersV1 | undefined {
  if (!overlay) return base ? structuredClone(base) : undefined
  if (!base) return structuredClone(overlay)
  const merged: NativeDocxTableBordersV1 = structuredClone(base)
  for (const edge of ['top', 'right', 'bottom', 'left', 'inside_horizontal', 'inside_vertical'] as const) if (overlay[edge] !== undefined) merged[edge] = structuredClone(overlay[edge])
  return merged
}

function faceMatches(manifest: NativeFontManifest, family: string, weight: number, style: string): boolean {
  return manifest.faces.some(face => face.weight === weight && face.style === style && face.stretch === 100 && [face.family, ...(face.aliases ?? [])].some(name => asciiLowerNative(name) === asciiLowerNative(family)))
}

/** Absent sizes take the declared host default; unavailable families rewrite
 * to an already loaded face of the same weight and style (host faces first).
 * Every rewrite is disclosed; nothing is invented from platform fonts. */
function prepareInnerFonts(resolved: NativeDocxResolvedLayoutInputV1, manifest: NativeFontManifest, hostSize: number, substitutions: NativeDocxApproximateNestedTableFontSubstitutionV1[], notes: Set<string>): boolean {
  const seen = new Set(substitutions.map(entry => `${asciiLowerNative(entry.source_family)}\0${entry.weight}\0${entry.style}`))
  const candidates = manifest.faces.filter(face => !!face.source.contentDigest && face.source.kind !== 'system' && face.stretch === 100).sort((left, right) => Number(right.source.kind === 'host') - Number(left.source.kind === 'host'))
  const rewrite = (properties: NativeDocxResolvedRunPropertiesV1, fallbackFamily: string | undefined): boolean => {
    if (!properties.font_size_half_points) { properties.font_size_half_points = hostSize; notes.add(`absent font sizes use the declared host default of ${hostSize} half-points`) }
    if (!properties.font_family) {
      const family = fallbackFamily ?? candidates[0]?.family
      if (!family) return false
      properties.font_family = family
      notes.add('absent font families use the paragraph mark or a loaded host face')
    }
    const weight = properties.bold ? 700 : 400, style = properties.italic ? 'italic' : 'normal'
    if (faceMatches(manifest, properties.font_family, weight, style)) return true
    const substitute = candidates.find(face => face.weight === weight && face.style === style) ?? candidates.find(face => face.style === style) ?? candidates.find(face => face.weight === weight) ?? candidates[0]
    if (!substitute) return false
    const key = `${asciiLowerNative(properties.font_family)}\0${weight}\0${style}`
    if (!seen.has(key)) { seen.add(key); substitutions.push({ source_family: properties.font_family, selected_family: substitute.family, face_id: substitute.faceId, weight, style, selected_weight: substitute.weight, selected_style: substitute.style }) }
    properties.font_family = substitute.family
    if (substitute.weight !== weight) properties.bold = substitute.weight === 700
    if (substitute.style !== style) properties.italic = substitute.style === 'italic'
    return true
  }
  const markFamily = new Map(resolved.paragraphs.map(paragraph => [paragraph.paragraph_id, paragraph.paragraph_mark_properties.font_family]))
  for (const paragraph of resolved.paragraphs) {
    if (!rewrite(paragraph.paragraph_mark_properties, undefined)) return false
    if (paragraph.numbering && !rewrite(paragraph.numbering.marker_properties, paragraph.paragraph_mark_properties.font_family)) return false
  }
  for (const run of resolved.runs) if (!rewrite(run.properties, markFamily.get(run.paragraph_id))) return false
  return true
}

function cellContentHeight(cell: NativeDocxTableCellV1, shaped: Map<string, NativeDocxShapedParagraphV1>): number {
  let height = 0, previousAfter = 0, first = true
  for (const source of cell.paragraphs) {
    const paragraph = shaped.get(source.id)
    if (!paragraph) continue
    const lines = paragraph.lines.reduce((sum, line) => sum + line.line_height_millipoints, 0)
    height += (first ? paragraph.spacing_before_millipoints : Math.max(previousAfter, paragraph.spacing_before_millipoints)) + lines
    previousAfter = paragraph.spacing_after_millipoints
    first = false
  }
  return height + previousAfter
}

function outerCellWidthTwips(table: NativeDocxTableV1, cell: NativeDocxTableCellV1): number | undefined {
  if (cell.width_twips !== undefined && cell.width_twips > 0) return cell.width_twips
  const grid = table.grid_widths_twips
  if (!grid?.length) return undefined
  for (const row of table.rows) {
    let column = 0
    for (const candidate of row.cells) {
      const span = Math.max(1, candidate.grid_span)
      if (candidate.id === cell.id) { const slice = grid.slice(column, column + span); return slice.length === span ? slice.reduce((a, b) => a + b, 0) : undefined }
      column += span
    }
  }
  return undefined
}

/** Compiler entry before body pagination: validate the sidecar, lay every
 * supported inner table out inside its cell's content box, and reserve the
 * extent in an internal body copy. Anything a single item cannot satisfy omits
 * that item with a declared reason; the body preview never depends on it. */
export async function prepareNativeDocxApproximateNestedTableStageV1(sidecarValue: unknown, document: NativeDocxDocumentV1, resolved: NativeDocxResolvedLayoutInputV1, runtime: NativeDocxApproximateNestedTableRuntimeV1, mainPartSha256?: string): Promise<NativeDocxApproximateNestedTableStageV1> {
  const sidecar = decodeNativeDocxApproximateNestedTablesV1(sidecarValue, document, mainPartSha256)
  const projected = structuredClone(document)
  const projectedResolved = structuredClone(resolved)
  const stage: NativeDocxApproximateNestedTableStageV1 = { sidecar, prepared: [], omitted: [], substitutions: [], omittedContent: 0, projection: { document: projected, resolved: projectedResolved }, removedDiagnostics: new Map() }
  const cells = new Map<string, { table: NativeDocxTableV1; cell: NativeDocxTableCellV1 }>()
  for (const block of projected.body.blocks) for (const row of block.table?.rows ?? []) for (const cell of row.cells) cells.set(cell.id, { table: block.table!, cell })
  const resolvedParagraphs = new Map(projectedResolved.paragraphs.map(paragraph => [paragraph.paragraph_id, paragraph]))
  const hostSize = runtime.fontSizeHalfPoints ?? HOST_DEFAULT_SIZE_HALF_POINTS
  const section = document.sections[0]
  const omit = (item: NativeDocxApproximateNestedTableV1, reason: string) => { stage.omitted.push({ id: item.id, reason }); stage.omittedContent += item.omitted_runs }
  for (const item of sidecar.items) {
    if (item.status !== 'supported' || !item.table || !item.geometry) { omit(item, item.reason ?? 'unsupported'); continue }
    const owner = cells.get(item.cell_id)
    if (!owner || !section) { omit(item, 'containing-cell-missing'); continue }
    const notes = new Set<string>(item.notes ?? [])
    try {
      // Containing cell content box from the outer table's authored geometry.
      // When strict resolution refused the outer style cascade (active look),
      // the sidecar's approximate outer indent and cell margins are projected
      // as direct properties of the internal table copy so the outer table and
      // this content box agree. Layout and width are left alone: the approximate
      // qualifier already sizes such tables from their authored grid, and an
      // autofit projection would re-enter the strict measurement probe.
      const outerResolved = projectedResolved.tables.find(entry => entry.table_id === owner.table.id)
      if (item.outer_geometry && !outerResolved?.geometry && (owner.table.indent_twips === undefined || owner.table.cell_margins === undefined)) {
        if (owner.table.indent_twips === undefined) owner.table.indent_twips = item.outer_geometry.indent_twips
        if (owner.table.cell_margins === undefined) owner.table.cell_margins = { ...item.outer_geometry.cell_margins }
        notes.add('containing table indent and cell margins approximated from its table style cascade')
      }
      const outer = nativeDocxTableGeometryV1(owner.table, projectedResolved)
      const outerMargins = outer.cell_margins ?? { top_twips: 0, right_twips: 115, bottom_twips: 0, left_twips: 115 }
      const cellWidth = outerCellWidthTwips(outer, owner.cell)
      if (cellWidth === undefined) { omit(item, 'containing-cell-width-unknown'); continue }
      const contentWidth = cellWidth - outerMargins.left_twips - outerMargins.right_twips
      if (contentWidth <= 0) { omit(item, 'containing-cell-content-box-empty'); continue }
      // Inner geometry: direct properties over the style cascade, like the resolver.
      const inner = item.table
      const geometry = item.geometry
      const indent = inner.indent_twips ?? geometry.indent_twips
      const margins = inner.cell_margins ?? geometry.cell_margins
      const available = contentWidth - indent
      if (available <= 0) { omit(item, 'no-room-inside-cell'); continue }
      let grid = [...inner.grid_widths_twips!]
      const preferred = inner.width_twips ?? (geometry.width_type === 'dxa' ? geometry.width_value : undefined)
      const percent = inner.width_percent_fiftieths ?? (geometry.width_type === 'pct' ? geometry.width_value : undefined)
      if (preferred !== undefined && preferred > 0 && preferred !== grid.reduce((a, b) => a + b, 0)) { const scaled = scaleWidths(grid, preferred); if (scaled) { grid = scaled; notes.add('columns scaled to the authored table width') } }
      else if (percent !== undefined && percent > 0) { const scaled = scaleWidths(grid, Math.floor(available * percent / 5000)); if (scaled) { grid = scaled; notes.add('columns scaled to the authored percentage width') } }
      let fitted: NativeDocxApproximateNestedTablePreparedV1['fitted']
      if (grid.reduce((a, b) => a + b, 0) > available) {
        const scaled = scaleWidths(grid, available)
        if (!scaled) { omit(item, 'columns-do-not-fit-cell'); continue }
        fitted = { source_grid_widths_twips: [...inner.grid_widths_twips!], fitted_grid_widths_twips: scaled, available_width_twips: available }
        grid = scaled
        notes.add(`authored columns scaled proportionally to the ${available} twip cell content box`)
      }
      const contentWidths = grid.map(width => width - margins.left_twips - margins.right_twips)
      if (contentWidths.some(width => width <= 0)) { omit(item, 'cell-margins-exceed-column'); continue }
      // Shape the inner paragraphs at their cell content widths with the body's fonts.
      const paragraphs: NativeDocxParagraphV1[] = []
      const widths = new Map<string, number>()
      for (const row of inner.rows) for (const [column, cell] of row.cells.entries()) for (const paragraph of cell.paragraphs) { paragraphs.push(structuredClone(paragraph)); widths.set(paragraph.id, twips(contentWidths[column]!)) }
      if (paragraphs.length === 0) { omit(item, 'no-inner-paragraphs'); continue }
      const innerResolved: NativeDocxResolvedLayoutInputV1 = { ...structuredClone(resolved), paragraphs: structuredClone(item.resolved_paragraphs), runs: structuredClone(item.resolved_runs), tables: [], diagnostics: [] }
      if (!prepareInnerFonts(innerResolved, runtime.manifest, hostSize, stage.substitutions, notes)) { omit(item, 'no-loaded-face-for-inner-text'); continue }
      const innerSection = structuredClone(section)
      innerSection.starts_at_block_id = paragraphs[0]!.id
      innerSection.header_refs = []
      innerSection.footer_refs = []
      const innerDocument: NativeDocxDocumentV1 = { ...structuredClone(document), body: { ...structuredClone(document.body), blocks: paragraphs.map(paragraph => ({ kind: 'paragraph' as const, id: paragraph.id, paragraph })) }, sections: [innerSection], headers: [], footers: [], notes: [], comment_stories: [], comments: [], unsupported: [] }
      const shaped = await shapeNativeDocxLinesWithParagraphWidthsV1({ protocol: 'injoffice.docx.shaping-request', version: 1, document: innerDocument, resolved_layout: innerResolved, font_manifest: runtime.manifest, available_width_millipoints: twips(Math.max(...contentWidths)), tab_interval_millipoints: Math.max(1, runtime.settings.default_tab_stop_twips * 50) }, { resolver: runtime.resolver, shaper: runtime.shaper }, widths, undefined, DOCX_APPROXIMATE_OMITTED_SOURCE_UNSUPPORTED)
      if (!shaped.ok) { omit(item, `shaping-refused: ${shaped.issues.map(issue => issue.message).slice(0, 2).join('; ')}`.slice(0, 256)); continue }
      const shapedParagraphs = new Map(shaped.value.paragraphs.map(paragraph => [paragraph.paragraph_id, paragraph]))
      const unshaped = paragraphs.filter(paragraph => !shapedParagraphs.has(paragraph.id) && paragraph.runs.some(run => run.kind === 'text' && (run.text ?? '') !== ''))
      stage.omittedContent += shaped.value.diagnostics.filter(entry => entry.severity === 'unsupported').length + unshaped.length
      if (unshaped.length) notes.add(`${unshaped.length} inner paragraphs with text were not shaped and are omitted`)
      // Rows size from shaped content plus the inner cell margins and trHeight rules.
      const top = twips(margins.top_twips), bottom = twips(margins.bottom_twips), left = twips(margins.left_twips), right = twips(margins.right_twips)
      const rows: PreparedRow[] = []
      let height = 0
      for (const row of inner.rows) {
        let rowHeight = 0
        const cellsOut: PreparedCell[] = []
        let x = 0
        for (const [column, cell] of row.cells.entries()) {
          const width = twips(grid[column]!)
          rowHeight = Math.max(rowHeight, top + cellContentHeight(cell, shapedParagraphs) + bottom)
          const shading = cell.shading_rgb ?? (rows.length === 0 ? item.first_row_cell_shading_rgb : undefined) ?? item.style_cell_shading_rgb
          cellsOut.push({ cell, x_millipoints: x, width_millipoints: width, content_x_millipoints: x + left, content_width_millipoints: width - left - right, ...(shading ? { shading_rgb: shading } : {}) })
          x += width
        }
        if (row.height_rule === 'atLeast' && row.height_twips !== undefined) rowHeight = Math.max(rowHeight, twips(row.height_twips))
        else if (row.height_rule === 'exact' && row.height_twips !== undefined) { const exact = twips(row.height_twips); if (rowHeight > exact) notes.add('exact row height smaller than its content; content is clipped by the row box in Word and overflows here'); rowHeight = exact }
        if (!Number.isSafeInteger(rowHeight) || rowHeight <= 0) { rowHeight = Math.max(1, rowHeight | 0) }
        rows.push({ row_id: row.id, height_millipoints: rowHeight, cells: cellsOut })
        height += rowHeight
      }
      if (!Number.isSafeInteger(height) || height <= 0 || height > MAX_TWIPS * 50) { omit(item, 'degenerate-extent'); continue }
      // Reserve the extent as spacing on the neighbouring cell paragraph: after
      // the preceding paragraph, else before the following one. The anchor
      // paragraph keeps its own font references, so strict inventory joins hold.
      const cellParagraphs = owner.cell.paragraphs
      const preceding = cellParagraphs[item.preceding_paragraphs - 1], following = cellParagraphs[item.preceding_paragraphs]
      const anchorParagraph = preceding ?? following
      const anchorResolved = anchorParagraph ? resolvedParagraphs.get(anchorParagraph.id) : undefined
      if (!anchorParagraph || !anchorResolved) { omit(item, 'no-anchor-paragraph-in-cell'); continue }
      const heightTwips = Math.ceil(height / 50)
      let side: 'after' | 'before', gap: number
      if (preceding) {
        const after = anchorResolved.properties.spacing_after_twips ?? 0
        const nextBefore = following ? resolvedParagraphs.get(following.id)?.properties.spacing_before_twips ?? 0 : 0
        side = 'after'; gap = twips(after)
        anchorResolved.properties = { ...anchorResolved.properties, spacing_after_twips: after + heightTwips + nextBefore }
      } else {
        const before = anchorResolved.properties.spacing_before_twips ?? 0
        side = 'before'; gap = twips(before)
        anchorResolved.properties = { ...anchorResolved.properties, spacing_before_twips: before + heightTwips }
      }
      const removed = projected.unsupported.filter(entry => item.diagnostic_ids.includes(entry.id))
      projected.unsupported = projected.unsupported.filter(entry => !item.diagnostic_ids.includes(entry.id))
      stage.removedDiagnostics.set(item.id, removed)
      stage.omittedContent += item.omitted_runs
      stage.prepared.push({ item, anchor_paragraph_id: anchorParagraph.id, anchor_side: side, anchor_gap_millipoints: gap, indent_millipoints: twips(indent), width_millipoints: twips(grid.reduce((a, b) => a + b, 0)), height_millipoints: height, rows, borders: mergeBorders(item.style_borders, inner.borders), shaped: shaped.value, resolved: innerResolved, ...(fitted ? { fitted } : {}), notes: [...notes] })
    } catch (error) {
      omit(item, `layout-failed: ${(error instanceof Error ? error.message : 'unknown').slice(0, 200)}`)
    }
  }
  return stage
}

export interface NativeDocxApproximateNestedTablePaintResultV1 {
  painted: string[]
  omitted: Array<{ id: string; reason: string }>
  reasons: string[]
}

function anchorLine(pages: readonly NativeDocxPaintPageV1[], paragraphID: string, side: 'after' | 'before'): { page: NativeDocxPaintPageV1; line: NativeDocxPaintLineV1 } | undefined {
  let best: { page: NativeDocxPaintPageV1; line: NativeDocxPaintLineV1 } | undefined
  for (const page of pages) for (const line of page.lines) {
    if (line.region !== 'body' || line.paragraph_id !== paragraphID) continue
    if (side === 'before') { if (line.source_line_ordinal === 0) return { page, line }; continue }
    if (!best || line.source_line_ordinal > best.line.source_line_ordinal || (line.source_line_ordinal === best.line.source_line_ordinal && page.ordinal > best.page.ordinal)) best = { page, line }
  }
  return best
}

function edgeBorder(cell: NativeDocxTableCellV1, edge: 'top' | 'right' | 'bottom' | 'left', fallback: NativeDocxTableBorderV1 | undefined): NativeDocxTableBorderV1 | undefined {
  const own = cell.borders?.[edge]
  return own !== undefined ? own : fallback
}

/** Rebuild page paint order: behind floats, table fills, the line-owned
 * commands in line order, table borders, front floats. Every original command
 * keeps its category; only nested-table paint is added. */
function rebuildCommands(page: NativeDocxPaintPageV1, fills: NativeDocxPagePaintCommandV1[], borders: NativeDocxPagePaintCommandV1[], extra: Map<string, NativeDocxPagePaintCommandV1>): void {
  const byID = new Map<string, NativeDocxPagePaintCommandV1>()
  const behindFloats: NativeDocxPagePaintCommandV1[] = [], frontFloats: NativeDocxPagePaintCommandV1[] = [], existingFills: NativeDocxPagePaintCommandV1[] = [], existingBorders: NativeDocxPagePaintCommandV1[] = []
  for (const command of page.commands) {
    if (command.kind === 'paint_floating_image') (command.layer === 'behind' ? behindFloats : frontFloats).push(command)
    else if (command.kind === 'fill_table_cell') existingFills.push(command)
    else if (command.kind === 'stroke_table_border') existingBorders.push(command)
    else byID.set(command.id, command)
  }
  for (const [id, command] of extra) byID.set(id, command)
  const ordinary: NativeDocxPagePaintCommandV1[] = []
  for (const line of page.lines) for (const id of line.command_ids) { const command = byID.get(id); if (command) ordinary.push(command) }
  page.commands = [...behindFloats, ...existingFills, ...fills, ...ordinary, ...existingBorders, ...borders, ...frontFloats]
}

/** Append nested-table paint to already painted approximate pages. The pages
 * are mutated in place; the caller re-validates the whole envelope. */
export async function paintNativeDocxApproximateNestedTablesV1(pages: NativeDocxPaintPageV1[], stage: NativeDocxApproximateNestedTableStageV1, source: { shaped_lines: NativeDocxShapedLinesV1 }, runtime: { manifest: NativeFontManifest; outlineProvider: NativeDocxGlyphOutlineProviderV1 }): Promise<NativeDocxApproximateNestedTablePaintResultV1> {
  const result: NativeDocxApproximateNestedTablePaintResultV1 = { painted: [], omitted: [...stage.omitted], reasons: [] }
  const faces = new Map(runtime.manifest.faces.map(face => [face.faceId, face]))
  const bodyShaped = new Map(source.shaped_lines.paragraphs.map(paragraph => [paragraph.paragraph_id, paragraph]))
  const outlineCache = new Map<string, NativeDocxGlyphOutlineResultV1>()
  // One shared outline table per page, so a nested table repeating a glyph transports it once.
  const registries = new Map<string, ReturnType<typeof nativeDocxPageGlyphOutlineRegistryV1>>()
  const registryFor = (page: NativeDocxPaintPageV1) => { const found = registries.get(page.id); if (found) return found; const created = nativeDocxPageGlyphOutlineRegistryV1(page); registries.set(page.id, created); return created }
  const fillsByPage = new Map<string, NativeDocxPagePaintCommandV1[]>(), bordersByPage = new Map<string, NativeDocxPagePaintCommandV1[]>()
  const extra = new Map<string, NativeDocxPagePaintCommandV1>()
  let glyphBudget = MAX_GLYPHS
  let byteBudget = stage.prepared.length ? Math.max(0, MAX_ENVELOPE_BYTES - JSON.stringify(pages).length) : 0
  let droppedGlyphs = 0
  const omit = (id: string, reason: string) => result.omitted.push({ id, reason })
  for (const [itemIndex, prepared] of stage.prepared.entries()) {
    const found = anchorLine(pages, prepared.anchor_paragraph_id, prepared.anchor_side)
    if (!found) { omit(prepared.item.id, 'anchor-paragraph-not-placed'); continue }
    const { page, line } = found
    const shapedAnchor = bodyShaped.get(prepared.anchor_paragraph_id)?.lines[line.source_line_ordinal]
    const contentX = line.x_millipoints - (shapedAnchor?.inline_offset_millipoints ?? 0)
    const tableX = Math.round(contentX + prepared.indent_millipoints)
    const tableY = Math.round(prepared.anchor_side === 'after' ? line.y_millipoints + line.height_millipoints + prepared.anchor_gap_millipoints : line.y_millipoints - prepared.anchor_gap_millipoints - prepared.height_millipoints)
    const maxX = page.width_millipoints, maxY = page.height_millipoints
    if (tableX < 0 || tableY < 0 || tableX + prepared.width_millipoints > maxX || tableY + prepared.height_millipoints > maxY) { omit(prepared.item.id, 'outside-page'); continue }
    const fills: NativeDocxPagePaintCommandV1[] = [], borders: NativeDocxPagePaintCommandV1[] = []
    const shaped = new Map(prepared.shaped.paragraphs.map(paragraph => [paragraph.paragraph_id, paragraph]))
    const resolvedRuns = new Map(prepared.resolved.runs.map(run => [run.run_id, run]))
    const resolvedParagraphs = new Map(prepared.resolved.paragraphs.map(paragraph => [paragraph.paragraph_id, paragraph]))
    const innerMargins = prepared.item.table!.cell_margins ?? prepared.item.geometry!.cell_margins
    const top = twips(innerMargins.top_twips)
    let rowY = tableY
    let lineCounter = 0
    const glyphCommands: NativeDocxPagePaintCommandV1[] = []
    for (const [rowIndex, row] of prepared.rows.entries()) {
      const lastRow = rowIndex === prepared.rows.length - 1
      for (const [cellIndex, cell] of row.cells.entries()) {
        const x = tableX + cell.x_millipoints, y = rowY, width = cell.width_millipoints, height = row.height_millipoints
        if (cell.shading_rgb) fills.push({ kind: 'fill_table_cell', id: `paint:nested:${prepared.item.id}:${rowIndex}:${cellIndex}:fill`, table_id: DOCX_APPROXIMATE_NESTED_TABLE_TABLE_ID, row_id: row.row_id, cell_id: cell.cell.id, x_millipoints: x, y_millipoints: y, width_millipoints: width, height_millipoints: height, fill_rgb: cell.shading_rgb } satisfies NativeDocxFillTableCellCommandV1)
        const table = prepared.borders
        const lastColumn = cellIndex === row.cells.length - 1
        const edges: Array<{ edge: NativeDocxStrokeTableBorderCommandV1['edge']; border?: NativeDocxTableBorderV1; x1: number; y1: number; x2: number; y2: number }> = [
          { edge: 'top', border: rowIndex === 0 ? edgeBorder(cell.cell, 'top', table?.top) : cell.cell.borders?.top, x1: x, y1: y, x2: x + width, y2: y },
          { edge: 'left', border: cellIndex === 0 ? edgeBorder(cell.cell, 'left', table?.left) : cell.cell.borders?.left, x1: x, y1: y, x2: x, y2: y + height },
          { edge: 'right', border: edgeBorder(cell.cell, 'right', lastColumn ? table?.right : table?.inside_vertical), x1: x + width, y1: y, x2: x + width, y2: y + height },
          { edge: 'bottom', border: edgeBorder(cell.cell, 'bottom', lastRow ? table?.bottom : table?.inside_horizontal), x1: x, y1: y + height, x2: x + width, y2: y + height },
        ]
        for (const edge of edges) if (edge.border?.style === 'single' && edge.border.color_rgb) borders.push({ kind: 'stroke_table_border', id: `paint:nested:${prepared.item.id}:${rowIndex}:${cellIndex}:${edge.edge}`, table_id: DOCX_APPROXIMATE_NESTED_TABLE_TABLE_ID, row_id: row.row_id, cell_id: cell.cell.id, edge: edge.edge, x1_millipoints: edge.x1, y1_millipoints: edge.y1, x2_millipoints: edge.x2, y2_millipoints: edge.y2, width_millipoints: Math.max(1, edge.border.size_eighth_points * 125), stroke_rgb: edge.border.color_rgb } satisfies NativeDocxStrokeTableBorderCommandV1)
        // Cell paragraphs flow like placeTableRow: top margin, collapsed gaps, line boxes.
        let localY = top, previousAfter = 0, first = true
        for (const source of cell.cell.paragraphs) {
          const paragraph = shaped.get(source.id)
          if (!paragraph) continue
          localY += first ? paragraph.spacing_before_millipoints : Math.max(previousAfter, paragraph.spacing_before_millipoints)
          first = false
          for (const shapedLine of paragraph.lines) {
            const lineTop = y + localY
            // Same seating rule as the body painter: a fixed (exact / at-least)
            // line box taller than the natural line puts its surplus leading
            // above the text, so the descent sits on the box bottom.
            const naturalHeight = shapedLine.ascent_millipoints - shapedLine.descent_millipoints + shapedLine.line_gap_millipoints
            const lineRule = resolvedParagraphs.get(paragraph.paragraph_id)?.properties?.line_rule
            const baseline = Math.round((lineRule === 'exact' || lineRule === 'atLeast') && shapedLine.line_height_millipoints > naturalHeight
              ? lineTop + shapedLine.line_height_millipoints + shapedLine.descent_millipoints
              : lineTop + shapedLine.ascent_millipoints)
            let fragmentX = x + cell.content_x_millipoints - cell.x_millipoints + shapedLine.inline_offset_millipoints
            const highlights: NativeDocxFillTextHighlightCommandV1[] = [], underlines: NativeDocxStrokeTextUnderlineCommandV1[] = [], glyphs: NativeDocxFillGlyphPathCommandV1[] = []
            for (const sourceFragment of shapedLine.fragments) {
              const fragment: NativeDocxLineFragmentV1 = { ...sourceFragment, id: `nt${itemIndex}.${lineCounter}.${sourceFragment.id}` }
              const properties = fragment.source_kind === 'list-marker' ? resolvedParagraphs.get(paragraph.paragraph_id)?.numbering?.marker_properties : resolvedRuns.get(fragment.source_id)?.properties
              const startX = Math.round(fragmentX)
              const before = glyphs.length
              if (properties && fragment.glyphs.length && properties.font_size_half_points && fragment.face_id) {
                const fontSize = properties.font_size_half_points * 500
                const manifestFace = faces.get(fragment.face_id)
                const fill = properties.color && RGB.test(properties.color) ? properties.color : '000000'
                if (manifestFace?.source.contentDigest && manifestFace.source.kind !== 'system') {
                  const face: NativeDocxContentAddressedFaceV1 = { face_id: manifestFace.faceId, content_digest: manifestFace.source.contentDigest, ...(manifestFace.source.collectionIndex !== undefined ? { collection_index: manifestFace.source.collectionIndex } : {}) }
                  let glyphX = fragmentX
                  for (const [glyphIndex, glyph] of fragment.glyphs.entries()) {
                    if (glyphBudget <= 0) { droppedGlyphs += 1; glyphX += glyph.advance_x_millipoints; continue }
                    const cacheKey = `${face.content_digest}\0${face.collection_index ?? ''}\0${glyph.glyph_id}`
                    let outline = outlineCache.get(cacheKey)
                    if (!outline) {
                      const live = await runtime.outlineProvider.getGlyphOutline(Object.freeze({ face: Object.freeze({ ...face }), glyph_id: glyph.glyph_id }))
                      outline = nativeDocxCaptureGlyphOutlineV1(structuredClone(live), face, glyph.glyph_id)
                      if (!outline) { droppedGlyphs += fragment.glyphs.length - glyphIndex; break }
                      outlineCache.set(cacheKey, outline)
                    }
                    if (outline.status === 'outlined') {
                      const registry = registryFor(page)
                      const stored = registry.outlines.length
                      const placement = nativeDocxRegisterGlyphOutlineV1(registry, face, glyph.glyph_id, fontSize, Math.round(glyphX + glyph.offset_x_millipoints), Math.round(baseline - glyph.offset_y_millipoints), { kind: 'path', path: outline.path, units_per_em: outline.units_per_em, scale_x: fontSize, scale_y: fontSize })
                      if (placement) {
                        const command: NativeDocxFillGlyphPathCommandV1 = { kind: 'fill_glyph_path', id: paintCommandID(line.placed_line_id, fragment.id, glyphIndex), line_id: line.line_id, fragment_id: fragment.id, source_id: fragment.source_id, glyph_index: glyphIndex, face, glyph_id: glyph.glyph_id, font_size_millipoints: fontSize, fill_rgb: fill, fill_rule: 'nonzero', outline_kind: 'path', ...placement }
                        // The envelope budget charges the shared outline the first time only.
                        const bytes = JSON.stringify(command).length + 1 + (registry.outlines.length > stored ? JSON.stringify(registry.outlines[placement.outline_index]).length + 1 : 0)
                        if (bytes > byteBudget) { droppedGlyphs += 1; glyphBudget = 0 } else { byteBudget -= bytes; glyphBudget -= 1; glyphs.push(command) }
                      }
                    }
                    glyphX += glyph.advance_x_millipoints
                  }
                }
              }
              if (properties && fragment.source_kind !== 'image' && fragment.source_kind !== 'textbox' && startX >= 0 && baseline >= 0 && (glyphs.length > before || fragment.glyphs.length === 0 && glyphBudget > 0)) {
                const underline = nativeTextUnderlineCommandsV1(properties.underline, properties.color, fragment, line.placed_line_id, line.line_id, startX, baseline)
                if (underline.ok) underlines.push(...underline.commands)
                const highlight = nativeTextHighlightCommandV1(properties.highlight, fragment, line.placed_line_id, line.line_id, startX, baseline)
                if (highlight.ok && highlight.command) highlights.push(highlight.command)
              }
              fragmentX += fragment.advance_inline_millipoints
            }
            glyphCommands.push(...highlights, ...glyphs, ...underlines)
            localY += shapedLine.line_height_millipoints
            lineCounter += 1
          }
          previousAfter = paragraph.spacing_after_millipoints
        }
      }
      rowY += row.height_millipoints
    }
    for (const command of glyphCommands) {
      if (extra.has(command.id) || page.commands.some(existing => existing.id === command.id)) continue
      extra.set(command.id, command)
      line.command_ids.push(command.id)
    }
    fillsByPage.set(page.id, [...(fillsByPage.get(page.id) ?? []), ...fills])
    bordersByPage.set(page.id, [...(bordersByPage.get(page.id) ?? []), ...borders])
    result.painted.push(prepared.item.id)
  }
  for (const page of pages) if (fillsByPage.has(page.id) || bordersByPage.has(page.id) || extra.size) rebuildCommands(page, fillsByPage.get(page.id) ?? [], bordersByPage.get(page.id) ?? [], extra)
  result.reasons = buildReasons(stage, result, droppedGlyphs)
  return result
}

function buildReasons(stage: NativeDocxApproximateNestedTableStageV1, result: NativeDocxApproximateNestedTablePaintResultV1, droppedGlyphs: number): string[] {
  const reasons: string[] = []
  if (result.painted.length) {
    const notes = new Set<string>()
    for (const prepared of stage.prepared) if (result.painted.includes(prepared.item.id)) for (const note of prepared.notes) notes.add(note)
    reasons.push(DOCX_APPROXIMATE_NESTED_TABLE_WARNING)
    reasons.push(`${DOCX_APPROXIMATE_NESTED_TABLE_CODE}: painted ${result.painted.length} of ${stage.sidecar.items.length} refused nested tables with ${DOCX_APPROXIMATE_NESTED_TABLE_LAYOUT_POLICY} (${stage.prepared.filter(prepared => result.painted.includes(prepared.item.id)).map(prepared => `${prepared.item.id} in cell ${prepared.item.cell_id}: ${prepared.rows.length} rows x ${prepared.rows[0]?.cells.length ?? 0} columns, ${prepared.width_millipoints} x ${prepared.height_millipoints} millipoints${prepared.fitted ? ` fitted from ${prepared.fitted.source_grid_widths_twips.join('+')} to ${prepared.fitted.fitted_grid_widths_twips.join('+')} twips` : ''} for diagnostics ${prepared.item.diagnostic_ids.join(',')}`).join('; ')})${notes.size ? `; approximations: ${[...notes].join('; ')}` : ''}`.slice(0, 8000))
  }
  if (result.omitted.length || stage.sidecar.omitted_count) reasons.push(`${DOCX_APPROXIMATE_NESTED_TABLE_OMITTED_CODE}: ${result.omitted.map(entry => `${entry.id} (${entry.reason})`).join('; ')}${stage.sidecar.omitted_count ? `; ${stage.sidecar.omitted_count} nested tables beyond the ${MAX_ITEMS} item budget` : ''}`.slice(0, 8000))
  for (const substitution of stage.substitutions.slice(0, MAX_REASONS - 4)) reasons.push(`${DOCX_APPROXIMATE_NESTED_TABLE_FONT_CODE}: ${substitution.source_family} / ${substitution.weight} / ${substitution.style} -> ${substitution.selected_family} / ${substitution.selected_weight} / ${substitution.selected_style} (loaded host face ${substitution.face_id}); nested-table metrics and layout may differ`)
  if (droppedGlyphs || stage.omittedContent) reasons.push(`${DOCX_APPROXIMATE_NESTED_TABLE_OMITTED_CODE}: nested-table content partially omitted (${droppedGlyphs} glyphs beyond the preview size budget, ${stage.omittedContent} unsupported runs/paragraphs/diagnostics)`)
  return reasons.slice(0, MAX_REASONS)
}

/** Compiler entry after body pagination: paint the prepared nested tables at
 * their reserved extents, append the declared reasons and re-derive omitted
 * content so dropped items keep their source refusals disclosed. */
export async function completeNativeDocxApproximateNestedTableStageV1(result: NativeDocxApproximateOmissionsV1 & { reasons: string[]; status: 'painted' | 'refused'; pages: NativeDocxPaintPageV1[] }, stage: NativeDocxApproximateNestedTableStageV1, source: { document: NativeDocxDocumentV1; resolved_layout: NativeDocxResolvedLayoutInputV1; shaped_lines: NativeDocxShapedLinesV1 }, runtime: { manifest: NativeFontManifest; outlineProvider: NativeDocxGlyphOutlineProviderV1 }, alsoRestored: NativeDocxDocumentV1['unsupported'] = []): Promise<NativeDocxApproximateNestedTablePaintResultV1> {
  const painted = await paintNativeDocxApproximateNestedTablesV1(result.pages, stage, source, runtime)
  for (const reason of painted.reasons) if (!result.reasons.includes(reason) && result.reasons.length < 260) result.reasons.push(reason)
  const restored = [...alsoRestored, ...painted.omitted.flatMap(entry => stage.removedDiagnostics.get(entry.id) ?? [])]
  const omissions = collectNativeDocxApproximateOmissionsV1(restored.length ? { ...source, document: { ...source.document, unsupported: [...source.document.unsupported, ...restored] } } : source, result)
  result.content_status = omissions.content_status
  result.omitted_content = omissions.omitted_content
  result.omitted_content_total = omissions.omitted_content_total
  result.unpainted_pages = omissions.unpainted_pages
  const needsWarning = omissions.omitted_content.length > 0 || omissions.unpainted_pages.length > 0
  const has = result.reasons.includes(DOCX_APPROXIMATE_OMITTED_CONTENT_WARNING)
  if (needsWarning && !has) result.reasons.push(DOCX_APPROXIMATE_OMITTED_CONTENT_WARNING)
  if (!needsWarning && has) result.reasons = result.reasons.filter(reason => reason !== DOCX_APPROXIMATE_OMITTED_CONTENT_WARNING)
  return painted
}
