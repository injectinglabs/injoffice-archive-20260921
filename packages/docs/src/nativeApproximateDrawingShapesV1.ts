/** Approximate DrawingML rectangles, lines and text boxes for the read-only
 * approximate page preview.
 *
 * The strict extractor keeps refusing wps:wsp shapes. The Go sidecar
 * (`InspectNativeApproximateDrawingShapesV1`) joins those refusals to their
 * source nodes and describes a bounded subset. This module validates that
 * sidecar against the current document, reserves inline shapes as glyphless
 * textbox atoms in the internal body copy, and appends approximate paint
 * commands after body pagination: fills and strokes reuse the existing
 * rectangle/border primitives, text-box glyphs attach to the anchor line of
 * the owning paragraph. Nothing here touches strict paint or source bytes. */
import type { NativeFontManifest, NativeFontResolver, NativeTextShaper } from '@injoffice/font-metrics/layout'
import type { NativeDocxDocumentV1, NativeDocxParagraphV1, NativeDocxRunV1, NativeDocxSourceAnchorV1 } from './nativeContract.js'
import type { NativeDocxResolvedLayoutInputV1, NativeDocxResolvedParagraphV1, NativeDocxResolvedRunPropertiesV1, NativeDocxResolvedRunV1 } from './nativeResolvedLayout.js'
import type { NativeDocxPaginationSettingsV1 } from './nativePaginationSettings.js'
import type { NativeDocxShapedLinesV1, NativeDocxLineFragmentV1 } from './nativeShapingLines.js'
import { shapeNativeDocxLinesWithParagraphWidthsV1 } from './nativeShapingLines.js'
import { ID, RGB, preflightWire, paintCommandID, DOCX_PAGE_PAINT_LIMITS } from './nativePagePaintWireV1.js'
import { nativeDocxPlaceGlyphPathV1, nativeDocxCaptureGlyphOutlineV1, type NativeDocxContentAddressedFaceV1, type NativeDocxFillGlyphPathCommandV1, type NativeDocxFillTableCellCommandV1, type NativeDocxFillTextHighlightCommandV1, type NativeDocxGlyphOutlineProviderV1, type NativeDocxGlyphOutlineResultV1, type NativeDocxPagePaintCommandV1, type NativeDocxPagePaintRequestV1, type NativeDocxPagePaintSuccessV1, type NativeDocxPaintLineV1, type NativeDocxPaintPageV1, type NativeDocxStrokeTableBorderCommandV1, type NativeDocxStrokeTextUnderlineCommandV1 } from './nativePagePaintV1.js'
import { nativeTextUnderlineCommandsV1 } from './nativeTextUnderlineV1.js'
import { nativeTextHighlightCommandV1 } from './nativeTextHighlightV1.js'
import { textboxAnchorLine, type TextboxAnchorContext } from './nativeTextboxAnchorLineV2.js'
import { resolveTextboxPosition } from './nativeTextboxPositionV2.js'
import type { NativeDocxTextboxGeometryItemV1 } from './nativeTextboxGeometryPreviewV1.js'
import type { NativeTextboxRelativeAnchorV2 } from './nativeTextboxPageAnchorV1.js'
import { asciiLowerNative } from './nativeDeterminism.js'
import { DOCX_APPROXIMATE_OMITTED_SOURCE_UNSUPPORTED } from './nativeApproximationV1.js'

export const DOCX_APPROXIMATE_DRAWING_SHAPES_PROTOCOL = 'injoffice.docx.approximate-drawing-shapes' as const
export const DOCX_APPROXIMATE_DRAWING_SHAPE_POLICY = 'docx.approximate-drawing-shape-preview-v1' as const
export const DOCX_APPROXIMATE_DRAWING_SHAPE_CODE = 'docx.approximate-drawing-shape-preview' as const
export const DOCX_APPROXIMATE_DRAWING_SHAPE_OMITTED_CODE = 'docx.approximate-drawing-shape-omitted' as const
export const DOCX_APPROXIMATE_TEXTBOX_FONT_CODE = 'docx.approximate-textbox-substituted-font' as const
export const DOCX_APPROXIMATE_DRAWING_SHAPE_WARNING = `${DOCX_APPROXIMATE_DRAWING_SHAPE_CODE}: DrawingML rectangles, lines and text boxes are painted approximately at resolved anchor positions with theme colors and outline widths approximated; body text is not wrapped around them. Original drawing restrictions and source bytes are unchanged.` as const
/** Table paint primitives carry these ids so consumers can tell shape paint from table paint. */
export const DOCX_APPROXIMATE_DRAWING_SHAPE_TABLE_ID = DOCX_APPROXIMATE_DRAWING_SHAPE_POLICY

const MAX_SHAPES = 64
const MAX_TEXTBOX_GLYPHS = 100_000
/** Interactive viewers read at most 16 MiB; keep body paint plus shape paint under this. */
const MAX_ENVELOPE_BYTES = 15 * 1024 * 1024
const MAX_REASONS = 24
const EMU_PER_MILLIPOINT = 12.7

export interface NativeDocxApproximateShapeLineV1 { rgb: string; width_emu: number; dash: string }
export interface NativeDocxApproximateTextboxV1 {
  link_id?: string
  link_seq: number
  insets_emu: [number, number, number, number]
  vertical_anchor: 't' | 'ctr' | 'b'
  wrap: 'square' | 'none'
  paragraphs: NativeDocxParagraphV1[]
  resolved_paragraphs: NativeDocxResolvedParagraphV1[]
  resolved_runs: NativeDocxResolvedRunV1[]
  omitted_runs: number
  omitted_blocks: number
}
export interface NativeDocxApproximateDrawingShapeV1 {
  id: string
  paragraph_id: string
  diagnostic_ids: string[]
  anchor: NativeDocxSourceAnchorV1
  run_anchor: NativeDocxSourceAnchorV1
  status: 'supported' | 'omitted'
  reason?: string
  placement?: 'inline' | 'anchored'
  preset?: 'rect' | 'line'
  width_emu: number
  height_emu: number
  rotation_degrees: number
  flip_horizontal: boolean
  flip_vertical: boolean
  fill_rgb?: string
  line?: NativeDocxApproximateShapeLineV1
  page_anchor?: NativeTextboxRelativeAnchorV2
  wrap?: string
  textbox?: NativeDocxApproximateTextboxV1
  notes?: string[]
}
export interface NativeDocxApproximateDrawingShapesV1 {
  protocol: typeof DOCX_APPROXIMATE_DRAWING_SHAPES_PROTOCOL
  version: 1
  policy: typeof DOCX_APPROXIMATE_DRAWING_SHAPE_POLICY
  package_sha256: string
  part_sha256: string
  items: NativeDocxApproximateDrawingShapeV1[]
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

/** Bounded structural validation plus source joins: every shape must name a
 * body paragraph and at least one retained drawing/run diagnostic anchored
 * inside its own run. Text box content is validated later by the shaping
 * request decoder; a malformed box only omits its text. */
export function decodeNativeDocxApproximateDrawingShapesV1(value: unknown, document: NativeDocxDocumentV1): NativeDocxApproximateDrawingShapesV1 {
  if (preflightWire(value, 'approximate drawing shapes', 2_000_000, 100_000).length) throw new TypeError('Approximate drawing shapes exceed their bounded wire')
  const input = structuredClone(value) as NativeDocxApproximateDrawingShapesV1
  if (!record(input) || !exactKeys(input as unknown as Record<string, unknown>, ['protocol', 'version', 'policy', 'package_sha256', 'part_sha256', 'items', 'omitted_count']) || input.protocol !== DOCX_APPROXIMATE_DRAWING_SHAPES_PROTOCOL || input.version !== 1 || input.policy !== DOCX_APPROXIMATE_DRAWING_SHAPE_POLICY || input.package_sha256 !== document.source.package_sha256 || typeof input.part_sha256 !== 'string' || !Array.isArray(input.items) || input.items.length > MAX_SHAPES || !safeNonnegative(input.omitted_count, 1_000_000)) throw new TypeError('Approximate drawing shapes do not exact-join the source document')
  const paragraphs = new Map(document.body.blocks.flatMap(block => block.paragraph ? [[block.id, block.paragraph] as const] : []))
  const diagnostics = new Map(document.unsupported.map(entry => [entry.id, entry]))
  const ids = new Set<string>()
  const main = document.source.main_part
  for (const item of input.items) {
    if (!record(item) || !exactKeys(item as unknown as Record<string, unknown>, ['id', 'paragraph_id', 'diagnostic_ids', 'anchor', 'run_anchor', 'status', 'width_emu', 'height_emu', 'rotation_degrees', 'flip_horizontal', 'flip_vertical'], ['reason', 'placement', 'preset', 'fill_rgb', 'line', 'page_anchor', 'wrap', 'textbox', 'notes'])) throw new TypeError('Approximate drawing shape has unknown or missing fields')
    if (typeof item.id !== 'string' || !ID.test(item.id) || ids.has(item.id)) throw new TypeError('Approximate drawing shape id is invalid or duplicated')
    ids.add(item.id)
    const paragraph = paragraphs.get(item.paragraph_id)
    if (!paragraph || !anchorValid(item.anchor, main) || !anchorValid(item.run_anchor, main) || !within(item.run_anchor, paragraph.anchor) || !within(item.anchor, item.run_anchor)) throw new TypeError('Approximate drawing shape does not exact-join its body paragraph')
    if (!Array.isArray(item.diagnostic_ids) || item.diagnostic_ids.length === 0 || item.diagnostic_ids.length > 64 || item.diagnostic_ids.some(id => { const d = typeof id === 'string' ? diagnostics.get(id) : undefined; return !d || d.scope_id !== item.paragraph_id || !d.anchor || !within(d.anchor, item.run_anchor) })) throw new TypeError('Approximate drawing shape must join retained source drawing diagnostics')
    if (paragraph.runs.some(run => within(run.anchor, item.run_anchor) || within(item.run_anchor, run.anchor))) throw new TypeError('Approximate drawing shape overlaps modeled text')
    if (item.status !== 'supported' && item.status !== 'omitted') throw new TypeError('Approximate drawing shape status is invalid')
    if (!safeNonnegative(item.width_emu) || !safeNonnegative(item.height_emu) || !safeNonnegative(item.rotation_degrees, 359) || typeof item.flip_horizontal !== 'boolean' || typeof item.flip_vertical !== 'boolean') throw new TypeError('Approximate drawing shape geometry is out of bounds')
    if (item.notes !== undefined && (!Array.isArray(item.notes) || item.notes.length > 32 || item.notes.some(note => typeof note !== 'string' || note.length > 512))) throw new TypeError('Approximate drawing shape notes are unbounded')
    if (item.reason !== undefined && (typeof item.reason !== 'string' || item.reason.length > 256)) throw new TypeError('Approximate drawing shape reason is unbounded')
    if (item.status === 'omitted') continue
    if (item.width_emu <= 0 || item.height_emu <= 0 || (item.placement !== 'inline' && item.placement !== 'anchored') || (item.preset !== 'rect' && item.preset !== 'line')) throw new TypeError('Supported approximate drawing shape requires positive extent, placement and preset')
    if (item.fill_rgb !== undefined && (typeof item.fill_rgb !== 'string' || !RGB.test(item.fill_rgb))) throw new TypeError('Approximate drawing shape fill is not an explicit RGB value')
    if (item.line !== undefined && (!record(item.line) || !exactKeys(item.line as unknown as Record<string, unknown>, ['rgb', 'width_emu', 'dash']) || typeof item.line.rgb !== 'string' || !RGB.test(item.line.rgb) || !safeNonnegative(item.line.width_emu, 12_700_000) || item.line.width_emu <= 0 || typeof item.line.dash !== 'string' || item.line.dash.length > 32)) throw new TypeError('Approximate drawing shape outline is invalid')
    if (item.placement === 'anchored') {
      const anchor = item.page_anchor as unknown
      if (!record(anchor) || anchor.policy !== 'relative-position-no-wrap-v2' || !anchorValid(anchor.source_anchor, main) || !within(anchor.source_anchor as NativeDocxSourceAnchorV1, item.anchor) || !Number.isSafeInteger(anchor.x_emu) || !Number.isSafeInteger(anchor.y_emu) || Math.abs(anchor.x_emu as number) > 127_000_000 || Math.abs(anchor.y_emu as number) > 127_000_000 || typeof anchor.horizontal_relative !== 'string' || typeof anchor.vertical_relative !== 'string' || !record(anchor.stacking) || typeof anchor.stacking.behind_doc !== 'boolean' || !safeNonnegative(anchor.stacking.relative_height, 0xffffffff)) throw new TypeError('Approximate drawing shape anchor is invalid')
    } else if (item.page_anchor !== undefined) throw new TypeError('Inline approximate drawing shapes cannot carry page anchors')
    if (item.textbox !== undefined) {
      const box = item.textbox as unknown
      if (!record(box) || !exactKeys(box, ['link_seq', 'insets_emu', 'vertical_anchor', 'wrap', 'paragraphs', 'resolved_paragraphs', 'resolved_runs', 'omitted_runs', 'omitted_blocks'], ['link_id']) || !safeNonnegative(box.link_seq, 1_000_000) || !Array.isArray(box.insets_emu) || box.insets_emu.length !== 4 || box.insets_emu.some(inset => !safeNonnegative(inset)) || !['t', 'ctr', 'b'].includes(String(box.vertical_anchor)) || !['square', 'none'].includes(String(box.wrap)) || !Array.isArray(box.paragraphs) || box.paragraphs.length > 256 || !Array.isArray(box.resolved_paragraphs) || !Array.isArray(box.resolved_runs) || !safeNonnegative(box.omitted_runs, 1_000_000) || !safeNonnegative(box.omitted_blocks, 1_000_000) || (box.link_id !== undefined && (typeof box.link_id !== 'string' || box.link_id.length > 64))) throw new TypeError('Approximate text box evidence is invalid')
    }
  }
  return input
}

export interface NativeDocxApproximateInlineShapeProjectionV1 {
  document: NativeDocxDocumentV1
  resolved: NativeDocxResolvedLayoutInputV1
  /** Shape id to the synthetic glyphless drawing run id reserved in the body copy. */
  inlineRuns: Map<string, string>
}

/** Reserve supported inline shapes as glyphless textbox atoms so surrounding
 * text reflows around their extent; drop the joined refusals from this
 * internal copy only. Anchored shapes leave the body copy untouched. */
export function projectNativeDocxApproximateInlineShapesV1(document: NativeDocxDocumentV1, resolved: NativeDocxResolvedLayoutInputV1, shapes: NativeDocxApproximateDrawingShapesV1): NativeDocxApproximateInlineShapeProjectionV1 {
  const projected = structuredClone(document)
  const projectedResolved = structuredClone(resolved)
  const inlineRuns = new Map<string, string>()
  const removed = new Set<string>()
  const paragraphs = new Map(projected.body.blocks.flatMap(block => block.paragraph ? [[block.id, block.paragraph] as const] : []))
  const blockIndex = new Map(projected.body.blocks.map((block, index) => [block.id, index]))
  for (const shape of shapes.items) {
    if (shape.status !== 'supported') continue
    for (const id of shape.diagnostic_ids) removed.add(id)
    if (shape.placement !== 'inline') continue
    const paragraph = paragraphs.get(shape.paragraph_id)
    if (!paragraph) continue
    // An inline atom wider than its column cannot paginate; Word lets it overflow
    // into the margin. Clamp the reserved width and disclose the clamp.
    let widthEMU = shape.width_emu
    const columnWidth = columnWidthEMU(projected, blockIndex.get(shape.paragraph_id) ?? 0)
    if (columnWidth !== undefined && widthEMU > columnWidth) { widthEMU = columnWidth; shape.notes = [...(shape.notes ?? []), 'inline shape width clamped to its column width'] }
    const runID = `${shape.id}:run`
    const run: NativeDocxRunV1 = {
      kind: 'drawing', id: runID, anchor: shape.run_anchor,
      drawing: {
        id: `${shape.id}:drawing`, anchor: shape.anchor, placement: 'inline', width_emu: widthEMU, height_emu: shape.height_emu,
        // The atom carries no painted text; the id keeps the required non-empty marker honest.
        textbox_text: shape.id, textbox_fill_rgb: shape.fill_rgb ?? 'FFFFFF', textbox_line_rgb: shape.line?.rgb ?? '000000',
        edit_policy: { mode: 'read-only', allowed_operations: [], refusal: { code: 'APPROXIMATE_DRAWING_SHAPE_PREVIEW', message: 'Inline shape reserved for read-only approximate preview', preservation: 'refuse-mutation' } },
      },
    }
    let index = paragraph.runs.findIndex(existing => existing.anchor.start_byte > shape.run_anchor.start_byte)
    if (index < 0) index = paragraph.runs.length
    paragraph.runs.splice(index, 0, run)
    projectedResolved.runs.push({ run_id: runID, paragraph_id: paragraph.id, applied_paragraph_styles: [], applied_character_styles: [], properties: {} })
    inlineRuns.set(shape.id, runID)
  }
  projected.unsupported = projected.unsupported.filter(entry => !removed.has(entry.id))
  return { document: projected, resolved: projectedResolved, inlineRuns }
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

export interface NativeDocxApproximateShapePaintRuntimeV1 {
  request: NativeDocxPagePaintRequestV1
  document: NativeDocxDocumentV1
  settings: NativeDocxPaginationSettingsV1
  manifest: NativeFontManifest
  resolver: NativeFontResolver
  shaper: NativeTextShaper
  outlineProvider: NativeDocxGlyphOutlineProviderV1
}

export interface NativeDocxApproximateTextboxFontSubstitutionV1 { source_family: string; selected_family: string; face_id: string; weight: number; style: string; selected_weight: number; selected_style: string }

export interface NativeDocxApproximateShapePaintResultV1 {
  painted: string[]
  omitted: Array<{ id: string; reason: string }>
  substitutions: NativeDocxApproximateTextboxFontSubstitutionV1[]
  reasons: string[]
}

interface PlacedShape {
  shape: NativeDocxApproximateDrawingShapeV1
  page: NativeDocxPaintPageV1
  line: NativeDocxPaintLineV1
  x: number
  y: number
  width: number
  height: number
  behind: boolean
  order: number
}

function toMillipoints(emu: number): number { return Math.round(emu / EMU_PER_MILLIPOINT) }
function clampCoordinate(value: number, max: number): number { return Math.min(Math.max(Math.round(value), 0), max) }

/** Append approximate shape paint to already painted approximate pages. The
 * pages are mutated in place; the caller re-validates the whole envelope. */
export async function paintNativeDocxApproximateDrawingShapesV1(paint: Pick<NativeDocxPagePaintSuccessV1, 'pages'>, shapes: NativeDocxApproximateDrawingShapesV1, projection: NativeDocxApproximateInlineShapeProjectionV1, runtime: NativeDocxApproximateShapePaintRuntimeV1): Promise<NativeDocxApproximateShapePaintResultV1> {
  const result: NativeDocxApproximateShapePaintResultV1 = { painted: [], omitted: [], substitutions: [], reasons: [] }
  const body = { pages: paint.pages } as NativeDocxPagePaintSuccessV1
  const extra = new Map<string, NativeDocxPagePaintCommandV1>()
  const behindByPage = new Map<string, NativeDocxPagePaintCommandV1[]>()
  const frontByPage = new Map<string, NativeDocxPagePaintCommandV1[]>()
  const removedIDs = new Set<string>()
  const placed: PlacedShape[] = []
  const omit = (shape: NativeDocxApproximateDrawingShapeV1, reason: string) => { result.omitted.push({ id: shape.id, reason }) }
  for (const shape of shapes.items) {
    if (shape.status !== 'supported') { omit(shape, shape.reason ?? 'unsupported'); continue }
    const width = toMillipoints(shape.width_emu), height = toMillipoints(shape.height_emu)
    if (width <= 0 || height <= 0) { omit(shape, 'degenerate-extent'); continue }
    if (shape.placement === 'inline') {
      const runID = projection.inlineRuns.get(shape.id)
      const found = runID ? findHighlight(paint.pages, runID) : undefined
      if (!found) { omit(shape, 'inline-shape-not-placed'); continue }
      // The reservation highlight only located the atom; the shape paints itself.
      removedIDs.add(found.command.id)
      found.line.command_ids = found.line.command_ids.filter(id => id !== found.command.id)
      placed.push({ shape, page: found.page, line: found.line, x: found.command.x_millipoints, y: found.command.y_millipoints, width: found.command.width_millipoints, height: found.command.height_millipoints, behind: false, order: 0 })
      continue
    }
    const anchor = shape.page_anchor!
    const item = { owner: { paragraph_id: shape.paragraph_id }, geometry: null, page_anchor: anchor } as unknown as NativeDocxTextboxGeometryItemV1
    let context: TextboxAnchorContext & { page: NativeDocxPaintPageV1 }
    try {
      context = textboxAnchorLine(runtime.document, item, body, runtime.request)
    } catch {
      const first = firstBodyLine(paint.pages, shape.paragraph_id)
      if (!first) { omit(shape, 'anchor-paragraph-not-placed'); continue }
      context = { page: first.page, line: first.line, character_x: first.line.x_millipoints, paragraph_y: first.line.y_millipoints }
    }
    let position: { x: number; y: number }
    try {
      position = resolveTextboxPosition(runtime.document, item, context.page, { width_millipoints: width, height_millipoints: height } as never, context)
    } catch (error) {
      omit(shape, `anchor-unresolved: ${error instanceof Error ? error.message : 'unknown'}`)
      continue
    }
    placed.push({ shape, page: context.page, line: context.line, x: position.x, y: position.y, width, height, behind: anchor.stacking?.behind_doc === true, order: anchor.stacking?.relative_height ?? 0 })
  }
  placed.sort((left, right) => left.order - right.order)
  for (const entry of placed) {
    const commands = shapeCommands(entry)
    if (commands.length === 0 && !entry.shape.textbox) { omit(entry.shape, 'outside-page'); continue }
    const bucket = entry.behind ? behindByPage : frontByPage
    bucket.set(entry.page.id, [...(bucket.get(entry.page.id) ?? []), ...commands])
    result.painted.push(entry.shape.id)
  }
  const textboxResult = await paintTextboxes(placed, runtime, extra, result)
  result.substitutions = textboxResult.substitutions
  for (const page of paint.pages) rebuildCommands(page, behindByPage.get(page.id) ?? [], frontByPage.get(page.id) ?? [], extra, removedIDs)
  result.reasons = buildReasons(shapes, result, textboxResult)
  return result
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

/** Fill and stroke primitives clipped to the page. Quarter-turn rectangles
 * swap their extent; line presets follow flips and rotation about the center. */
function shapeCommands(entry: PlacedShape): NativeDocxPagePaintCommandV1[] {
  const { shape, page } = entry
  const maxX = page.width_millipoints, maxY = page.height_millipoints
  let { x, y, width, height } = entry
  if (shape.preset === 'rect' && (shape.rotation_degrees === 90 || shape.rotation_degrees === 270)) {
    const centerX = x + width / 2, centerY = y + height / 2
    ;[width, height] = [height, width]
    x = centerX - width / 2; y = centerY - height / 2
  }
  const commands: NativeDocxPagePaintCommandV1[] = []
  const strokeWidth = shape.line ? Math.max(1, toMillipoints(shape.line.width_emu)) : 0
  if (shape.preset === 'rect') {
    const left = clampCoordinate(x, maxX), top = clampCoordinate(y, maxY), right = clampCoordinate(x + width, maxX), bottom = clampCoordinate(y + height, maxY)
    if (right <= left || bottom <= top) return []
    if (shape.fill_rgb !== undefined) {
      const fill: NativeDocxFillTableCellCommandV1 = { kind: 'fill_table_cell', id: `${shape.id}:fill`, table_id: DOCX_APPROXIMATE_DRAWING_SHAPE_TABLE_ID, row_id: shape.id, cell_id: 'fill', x_millipoints: left, y_millipoints: top, width_millipoints: right - left, height_millipoints: bottom - top, fill_rgb: shape.fill_rgb }
      commands.push(fill)
    }
    if (shape.line) {
      const edges: Array<[NativeDocxStrokeTableBorderCommandV1['edge'], number, number, number, number]> = [['top', left, top, right, top], ['right', right, top, right, bottom], ['bottom', left, bottom, right, bottom], ['left', left, top, left, bottom]]
      for (const [edge, x1, y1, x2, y2] of edges) commands.push({ kind: 'stroke_table_border', id: `${shape.id}:outline:${edge}`, table_id: DOCX_APPROXIMATE_DRAWING_SHAPE_TABLE_ID, row_id: shape.id, cell_id: 'outline', edge, x1_millipoints: x1, y1_millipoints: y1, x2_millipoints: x2, y2_millipoints: y2, width_millipoints: strokeWidth, stroke_rgb: shape.line.rgb })
    }
    return commands
  }
  if (!shape.line) return []
  // A line preset connects the top-left and bottom-right corners of its box.
  let [x1, y1, x2, y2] = [x, y, x + width, y + height]
  if (shape.flip_horizontal) [x1, x2] = [x2, x1]
  if (shape.flip_vertical) [y1, y2] = [y2, y1]
  if (shape.rotation_degrees !== 0) {
    const centerX = x + width / 2, centerY = y + height / 2, angle = shape.rotation_degrees * Math.PI / 180
    const rotate = (px: number, py: number): [number, number] => [centerX + (px - centerX) * Math.cos(angle) - (py - centerY) * Math.sin(angle), centerY + (px - centerX) * Math.sin(angle) + (py - centerY) * Math.cos(angle)]
    ;[x1, y1] = rotate(x1, y1); [x2, y2] = rotate(x2, y2)
  }
  const stroke: NativeDocxStrokeTableBorderCommandV1 = { kind: 'stroke_table_border', id: `${shape.id}:line`, table_id: DOCX_APPROXIMATE_DRAWING_SHAPE_TABLE_ID, row_id: shape.id, cell_id: 'line', edge: 'top', x1_millipoints: clampCoordinate(x1, maxX), y1_millipoints: clampCoordinate(y1, maxY), x2_millipoints: clampCoordinate(x2, maxX), y2_millipoints: clampCoordinate(y2, maxY), width_millipoints: strokeWidth, stroke_rgb: shape.line.rgb }
  if (stroke.x1_millipoints === stroke.x2_millipoints && stroke.y1_millipoints === stroke.y2_millipoints) return []
  return [stroke]
}

/** Rebuild page paint order: behind floats, table fills, behind shapes, the
 * line-owned commands in line order, table borders, front shapes, front floats.
 * Every original command keeps its category; only shape paint is added. */
function rebuildCommands(page: NativeDocxPaintPageV1, behind: NativeDocxPagePaintCommandV1[], front: NativeDocxPagePaintCommandV1[], extra: Map<string, NativeDocxPagePaintCommandV1>, removed: Set<string>): void {
  const byID = new Map<string, NativeDocxPagePaintCommandV1>()
  const behindFloats: NativeDocxPagePaintCommandV1[] = [], frontFloats: NativeDocxPagePaintCommandV1[] = [], fills: NativeDocxPagePaintCommandV1[] = [], borders: NativeDocxPagePaintCommandV1[] = []
  for (const command of page.commands) {
    if (removed.has(command.id)) continue
    if (command.kind === 'paint_floating_image') (command.layer === 'behind' ? behindFloats : frontFloats).push(command)
    else if (command.kind === 'fill_table_cell') fills.push(command)
    else if (command.kind === 'stroke_table_border') borders.push(command)
    else byID.set(command.id, command)
  }
  for (const [id, command] of extra) byID.set(id, command)
  const ordinary: NativeDocxPagePaintCommandV1[] = []
  for (const line of page.lines) for (const id of line.command_ids) { const command = byID.get(id); if (command) ordinary.push(command) }
  page.commands = [...behindFloats, ...fills, ...behind, ...ordinary, ...borders, ...front, ...frontFloats]
}

interface TextboxPaintOutcome { substitutions: NativeDocxApproximateTextboxFontSubstitutionV1[]; textboxes: number; droppedLines: number; droppedGlyphs: number; omittedContent: number; failures: string[]; byteBudget: number }

/** Shape text box content with the same shaping core as body text and place
 * the resulting lines into the linked chain of boxes in source seq order.
 * Glyph paint attaches to the anchor line so the wire decoder keeps its
 * line ownership invariants. */
async function paintTextboxes(placed: PlacedShape[], runtime: NativeDocxApproximateShapePaintRuntimeV1, extra: Map<string, NativeDocxPagePaintCommandV1>, result: NativeDocxApproximateShapePaintResultV1): Promise<TextboxPaintOutcome> {
  const outcome: TextboxPaintOutcome = { substitutions: [], textboxes: 0, droppedLines: 0, droppedGlyphs: 0, omittedContent: 0, failures: [], byteBudget: 0 }
  const chains = new Map<string, PlacedShape[]>()
  for (const entry of placed) {
    const box = entry.shape.textbox
    if (!box || !result.painted.includes(entry.shape.id)) continue
    outcome.omittedContent += box.omitted_runs + box.omitted_blocks
    const key = box.link_id ? `link:${box.link_id}` : `shape:${entry.shape.id}`
    chains.set(key, [...(chains.get(key) ?? []), entry])
  }
  const outlineCache = new Map<string, NativeDocxGlyphOutlineResultV1>()
  let glyphBudget = MAX_TEXTBOX_GLYPHS
  // Glyph paths dominate the wire; measure the body paint once and spend only
  // the remaining viewer budget on text box glyphs. Excess is disclosed.
  if (chains.size) outcome.byteBudget = Math.max(0, MAX_ENVELOPE_BYTES - JSON.stringify(placed[0]?.page ? placed.map(entry => entry.page).filter((page, index, pages) => pages.indexOf(page) === index) : []).length - extra.size * 512)
  for (const [key, boxes] of chains) {
    boxes.sort((left, right) => left.shape.textbox!.link_seq - right.shape.textbox!.link_seq)
    const head = boxes.find(entry => entry.shape.textbox!.paragraphs.length > 0)
    if (!head) continue
    const chainOrdinal = [...chains.keys()].indexOf(key)
    try {
      const shaped = await shapeTextbox(head.shape.textbox!, runtime, head, outcome)
      if (!shaped) { outcome.failures.push(`${head.shape.id}: no host face for text box fonts`); continue }
      outcome.textboxes += 1
      glyphBudget = await placeTextboxLines(shaped, boxes, chainOrdinal, runtime, extra, outlineCache, outcome, glyphBudget)
    } catch (error) {
      outcome.failures.push(`${head.shape.id}: ${(error instanceof Error ? error.message : 'shaping failed').slice(0, 200)}`)
    }
  }
  return outcome
}

interface ShapedTextbox { lines: NativeDocxShapedLinesV1; resolved: NativeDocxResolvedLayoutInputV1 }

function faceMatches(manifest: NativeFontManifest, family: string, weight: number, style: string): boolean {
  return manifest.faces.some(face => face.weight === weight && face.style === style && face.stretch === 100 && [face.family, ...(face.aliases ?? [])].some(name => asciiLowerNative(name) === asciiLowerNative(family)))
}

/** Rewrite unavailable text box families to an already loaded host face of the
 * same weight and style; disclose every substitution. Never invents a family. */
function substituteTextboxFonts(resolved: NativeDocxResolvedLayoutInputV1, manifest: NativeFontManifest, outcome: TextboxPaintOutcome): boolean {
  const seen = new Set(outcome.substitutions.map(entry => `${asciiLowerNative(entry.source_family)}\0${entry.weight}\0${entry.style}`))
  const rewrite = (properties: NativeDocxResolvedRunPropertiesV1): boolean => {
    if (!properties.font_family) return true
    const weight = properties.bold ? 700 : 400, style = properties.italic ? 'italic' : 'normal'
    if (faceMatches(manifest, properties.font_family, weight, style)) return true
    // Prefer a loaded host face with the same weight and style, then the nearest
    // host face, then a document-embedded face; bold/italic are rewritten so the
    // shaper resolves exactly. Never a system lookup, never an invented family.
    const candidates = manifest.faces.filter(face => !!face.source.contentDigest && face.source.kind !== 'system' && face.stretch === 100).sort((left, right) => Number(right.source.kind === 'host') - Number(left.source.kind === 'host'))
    const substitute = candidates.find(face => face.weight === weight && face.style === style) ?? candidates.find(face => face.style === style) ?? candidates.find(face => face.weight === weight) ?? candidates[0]
    if (!substitute) return false
    const key = `${asciiLowerNative(properties.font_family)}\0${weight}\0${style}`
    if (!seen.has(key)) { seen.add(key); outcome.substitutions.push({ source_family: properties.font_family, selected_family: substitute.family, face_id: substitute.faceId, weight, style, selected_weight: substitute.weight, selected_style: substitute.style }) }
    properties.font_family = substitute.family
    if (substitute.weight !== weight) properties.bold = substitute.weight === 700
    if (substitute.style !== style) properties.italic = substitute.style === 'italic'
    return true
  }
  for (const run of resolved.runs) if (!rewrite(run.properties)) return false
  for (const paragraph of resolved.paragraphs) {
    if (!rewrite(paragraph.paragraph_mark_properties)) return false
    if (paragraph.numbering && !rewrite(paragraph.numbering.marker_properties)) return false
  }
  return true
}

async function shapeTextbox(box: NativeDocxApproximateTextboxV1, runtime: NativeDocxApproximateShapePaintRuntimeV1, head: PlacedShape, outcome: TextboxPaintOutcome): Promise<ShapedTextbox | undefined> {
  const source = runtime.document
  const paragraphs = structuredClone(box.paragraphs)
  const first = paragraphs[0]
  if (!first || !source.sections[0]) return undefined
  const section = structuredClone(source.sections[0])
  section.starts_at_block_id = first.id
  section.header_refs = []
  section.footer_refs = []
  const document: NativeDocxDocumentV1 = { ...structuredClone(source), body: { ...structuredClone(source.body), blocks: paragraphs.map(paragraph => ({ kind: 'paragraph' as const, id: paragraph.id, paragraph })) }, sections: [section], headers: [], footers: [], notes: [], comment_stories: [], comments: [], unsupported: [] }
  const resolved: NativeDocxResolvedLayoutInputV1 = { ...structuredClone(runtime.request.pagination_request.resolved_layout), paragraphs: structuredClone(box.resolved_paragraphs), runs: structuredClone(box.resolved_runs), tables: [], diagnostics: [] }
  if (!substituteTextboxFonts(resolved, runtime.manifest, outcome)) return undefined
  const insets = box.insets_emu
  const innerWidth = Math.max(1, head.width - toMillipoints(insets[0]) - toMillipoints(insets[2]))
  const shaped = await shapeNativeDocxLinesWithParagraphWidthsV1({ protocol: 'injoffice.docx.shaping-request', version: 1, document, resolved_layout: resolved, font_manifest: runtime.manifest, available_width_millipoints: box.wrap === 'none' ? 1_000_000_000 : innerWidth, tab_interval_millipoints: Math.max(1, runtime.settings.default_tab_stop_twips * 50) }, { resolver: runtime.resolver, shaper: runtime.shaper }, new Map(), undefined, DOCX_APPROXIMATE_OMITTED_SOURCE_UNSUPPORTED)
  if (!shaped.ok) throw new TypeError(`text box shaping refused: ${shaped.issues.map(issue => issue.message).slice(0, 3).join('; ')}`)
  outcome.omittedContent += shaped.value.diagnostics.filter(entry => entry.severity === 'unsupported').length
  return { lines: shaped.value, resolved }
}

async function placeTextboxLines(shaped: ShapedTextbox, boxes: PlacedShape[], chainOrdinal: number, runtime: NativeDocxApproximateShapePaintRuntimeV1, extra: Map<string, NativeDocxPagePaintCommandV1>, outlineCache: Map<string, NativeDocxGlyphOutlineResultV1>, outcome: TextboxPaintOutcome, glyphBudget: number): Promise<number> {
  const resolvedRuns = new Map(shaped.resolved.runs.map(run => [run.run_id, run]))
  const resolvedParagraphs = new Map(shaped.resolved.paragraphs.map(paragraph => [paragraph.paragraph_id, paragraph]))
  const faces = new Map(runtime.manifest.faces.map(face => [face.faceId, face]))
  type Slot = { box: PlacedShape; top: number; bottom: number; left: number }
  const slots: Slot[] = boxes.map(box => {
    const insets = box.shape.textbox!.insets_emu
    return { box, top: box.y + toMillipoints(insets[1]), bottom: box.y + box.height - toMillipoints(insets[3]), left: box.x + toMillipoints(insets[0]) }
  })
  // First pass: assign every shaped line to a slot at a y offset.
  const placements: Array<{ slot: number; y: number; paragraphID: string; line: NativeDocxShapedLinesV1['paragraphs'][number]['lines'][number]; spacingBefore: number }> = []
  let slotIndex = 0
  let cursor = slots[0]!.top
  let exhausted = false
  for (const paragraph of shaped.lines.paragraphs) {
    if (exhausted) break
    cursor += paragraph.spacing_before_millipoints
    for (const line of paragraph.lines) {
      let height = line.line_height_millipoints
      if (height <= 0) height = 1
      while (cursor + height > slots[slotIndex]!.bottom) {
        if (slotIndex + 1 >= slots.length) { exhausted = true; break }
        slotIndex += 1
        cursor = slots[slotIndex]!.top
      }
      if (exhausted) { outcome.droppedLines += 1; continue }
      placements.push({ slot: slotIndex, y: cursor, paragraphID: paragraph.paragraph_id, line, spacingBefore: 0 })
      cursor += height
    }
    cursor += paragraph.spacing_after_millipoints
  }
  if (exhausted) for (const paragraph of shaped.lines.paragraphs) void paragraph
  // Vertical anchoring applies to a single unlinked box only.
  if (slots.length === 1 && placements.length) {
    const anchor = boxes[0]!.shape.textbox!.vertical_anchor
    const used = cursor - slots[0]!.top
    const free = slots[0]!.bottom - slots[0]!.top - used
    if (free > 0 && anchor !== 't') { const shift = anchor === 'ctr' ? Math.floor(free / 2) : free; for (const placement of placements) placement.y += shift }
  }
  let glyphIndexBudget = glyphBudget
  for (const [placementIndex, placement] of placements.entries()) {
    const slot = slots[placement.slot]!
    const paintLine = slot.box.line
    const page = slot.box.page
    const placedID = paintLine.placed_line_id
    const line = placement.line
    const baseline = Math.round(placement.y + line.ascent_millipoints)
    let x = slot.left + line.inline_offset_millipoints
    const highlights: NativeDocxFillTextHighlightCommandV1[] = []
    const underlines: NativeDocxStrokeTextUnderlineCommandV1[] = []
    const glyphs: NativeDocxFillGlyphPathCommandV1[] = []
    for (const source of line.fragments) {
      const fragment: NativeDocxLineFragmentV1 = { ...source, id: `tb${chainOrdinal}.${placementIndex}.${source.id}` }
      const properties = fragment.source_kind === 'list-marker' ? resolvedParagraphs.get(placement.paragraphID)?.numbering?.marker_properties : resolvedRuns.get(fragment.source_id)?.properties
      const fragmentX = Math.round(x)
      const glyphsBefore = glyphs.length
      if (properties && fragment.glyphs.length && properties.font_size_half_points && fragment.face_id) {
        const fontSize = properties.font_size_half_points * 500
        const manifestFace = faces.get(fragment.face_id)
        const fill = properties.color && RGB.test(properties.color) ? properties.color : '000000'
        if (manifestFace?.source.contentDigest && manifestFace.source.kind !== 'system') {
          const face: NativeDocxContentAddressedFaceV1 = { face_id: manifestFace.faceId, content_digest: manifestFace.source.contentDigest, ...(manifestFace.source.collectionIndex !== undefined ? { collection_index: manifestFace.source.collectionIndex } : {}) }
          let glyphX = x
          for (const [glyphIndex, glyph] of fragment.glyphs.entries()) {
            if (glyphIndexBudget <= 0) { outcome.droppedGlyphs += 1; continue }
            const cacheKey = `${face.content_digest}\0${face.collection_index ?? ''}\0${glyph.glyph_id}`
            let outline = outlineCache.get(cacheKey)
            if (!outline) {
              const live = await runtime.outlineProvider.getGlyphOutline(Object.freeze({ face: Object.freeze({ ...face }), glyph_id: glyph.glyph_id }))
              outline = nativeDocxCaptureGlyphOutlineV1(structuredClone(live), face, glyph.glyph_id)
              if (!outline) break
              outlineCache.set(cacheKey, outline)
            }
            if (outline.status === 'outlined') {
              const path = nativeDocxPlaceGlyphPathV1(outline.path, Math.round(glyphX + glyph.offset_x_millipoints), Math.round(baseline - glyph.offset_y_millipoints), fontSize, outline.units_per_em)
              if (path) {
                const command: NativeDocxFillGlyphPathCommandV1 = { kind: 'fill_glyph_path', id: paintCommandID(placedID, fragment.id, glyphIndex), line_id: paintLine.line_id, fragment_id: fragment.id, source_id: fragment.source_id, glyph_index: glyphIndex, face, glyph_id: glyph.glyph_id, font_size_millipoints: fontSize, fill_rgb: fill, fill_rule: 'nonzero', outline_kind: 'path', path }
                const bytes = JSON.stringify(command).length + 1
                if (bytes > outcome.byteBudget) { outcome.droppedGlyphs += 1; glyphIndexBudget = 0; glyphX += glyph.advance_x_millipoints; continue } else {
                  outcome.byteBudget -= bytes
                  glyphIndexBudget -= 1
                  glyphs.push(command)
                }
              }
            }
            glyphX += glyph.advance_x_millipoints
          }
        }
      }
      // Decorations follow painted glyphs only, so budget truncation never leaves stray rules.
      if (properties && fragment.source_kind !== 'image' && fragment.source_kind !== 'textbox' && fragmentX >= 0 && baseline >= 0 && (glyphs.length > glyphsBefore || fragment.glyphs.length === 0 && glyphIndexBudget > 0)) {
        const underline = nativeTextUnderlineCommandsV1(properties.underline, properties.color, fragment, placedID, paintLine.line_id, fragmentX, baseline)
        if (underline.ok) underlines.push(...underline.commands)
        const highlight = nativeTextHighlightCommandV1(properties.highlight, fragment, placedID, paintLine.line_id, fragmentX, baseline)
        if (highlight.ok && highlight.command) highlights.push(highlight.command)
      }
      x += fragment.advance_inline_millipoints
    }
    const commands: NativeDocxPagePaintCommandV1[] = [...highlights, ...glyphs, ...underlines]
    for (const command of commands) {
      if (extra.has(command.id) || page.commands.some(existing => existing.id === command.id)) continue
      extra.set(command.id, command)
      paintLine.command_ids.push(command.id)
    }
  }
  return glyphIndexBudget
}

function buildReasons(shapes: NativeDocxApproximateDrawingShapesV1, result: NativeDocxApproximateShapePaintResultV1, textboxes: TextboxPaintOutcome): string[] {
  const reasons: string[] = []
  if (result.painted.length) {
    const notes = new Set<string>()
    for (const shape of shapes.items) if (result.painted.includes(shape.id)) for (const note of shape.notes ?? []) notes.add(note)
    reasons.push(DOCX_APPROXIMATE_DRAWING_SHAPE_WARNING)
    reasons.push(`${DOCX_APPROXIMATE_DRAWING_SHAPE_CODE}: painted ${result.painted.length} of ${shapes.items.length} refused DrawingML shapes (${shapes.items.filter(shape => result.painted.includes(shape.id)).map(shape => `${shape.id} ${shape.placement} ${shape.preset}${shape.textbox ? ' textbox' : ''} for diagnostics ${shape.diagnostic_ids.join(',')}`).join('; ')})${textboxes.textboxes ? `; ${textboxes.textboxes} text box chains shaped with current InjOffice layout` : ''}${notes.size ? `; approximations: ${[...notes].join('; ')}` : ''}`.slice(0, 8000))
  }
  if (result.omitted.length || shapes.omitted_count) reasons.push(`${DOCX_APPROXIMATE_DRAWING_SHAPE_OMITTED_CODE}: ${result.omitted.map(entry => `${entry.id} (${entry.reason})`).join('; ')}${shapes.omitted_count ? `; ${shapes.omitted_count} shapes beyond the ${MAX_SHAPES} shape budget` : ''}`.slice(0, 8000))
  for (const substitution of result.substitutions.slice(0, MAX_REASONS - 4)) reasons.push(`${DOCX_APPROXIMATE_TEXTBOX_FONT_CODE}: ${substitution.source_family} / ${substitution.weight} / ${substitution.style} -> ${substitution.selected_family} / ${substitution.selected_weight} / ${substitution.selected_style} (loaded host face ${substitution.face_id}); text box metrics and layout may differ`)
  if (textboxes.droppedLines || textboxes.droppedGlyphs || textboxes.omittedContent || textboxes.failures.length) reasons.push(`${DOCX_APPROXIMATE_DRAWING_SHAPE_OMITTED_CODE}: text box content partially omitted (${textboxes.droppedLines} overflow lines dropped, ${textboxes.droppedGlyphs} glyphs beyond the preview size budget, ${textboxes.omittedContent} unsupported runs/blocks/diagnostics${textboxes.failures.length ? `, ${textboxes.failures.join('; ')}` : ''})`.slice(0, 8000))
  return reasons.slice(0, MAX_REASONS)
}
