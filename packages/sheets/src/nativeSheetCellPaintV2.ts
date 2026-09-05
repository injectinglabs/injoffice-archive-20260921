/**
 * Renderer-neutral native XLSX v2 cell glyph/display paint.
 *
 * Shapes producer-issued font bytes with the pinned HarfBuzz path, places
 * integer-EMU glyph origins inside geometry v2 cell rects, and emits path
 * commands a host can fill. It never evaluates formulas, invents locale or
 * CSS/DOM display, measures canvas text, or clips overflow with host heuristics.
 * Date/numeric formats paint from exact OOXML y/m/d/h/s tokens without Date, Intl,
 * or host TZ. Named months/days require an explicit OOXML locale/calendar table.
 * wrap_text uses HarfBuzz advances of the already-qualified display string, then
 * shrink-to-fit may apply integer-EMU floor(lineBox/lineAdvance) scaling.
 * BMP Unicode and HarfBuzz clusters paint when the producer-issued font has
 * outlines. Composite TrueType glyf components flatten with a bounded recursion
 * policy. Reconstructible rich runs, theme typefaces with exact TTF bytes, theme
 * RGB, 1900/1904 serials, shrink-to-fit, and quadrant rotation paint only from
 * exact projected identity.
 */

import {
  NATIVE_TEXT_LAYOUT_VERSION,
  type FontDesignMetrics,
  type FontResource,
  type ResolvedFontFace,
  type ShapedSegment,
  type TextRunInput,
} from '@injoffice/font-metrics/layout'
import { createHarfBuzzTextShaperV1, inspectHarfBuzzFontMetricsV1 } from '@injoffice/font-metrics/harfbuzz'
import { isUnicode13Control, isUnicode13DefaultIgnorable, unicode13Script } from '@injoffice/font-metrics/unicode13'
import { isProjectedNativeWorkbookV2 } from './nativeRenderModelV2.js'
import type { NativeRenderCellContentV2, NativeRenderCellV2, NativeSheetRenderModelV2, NativeWorkbookRenderModelV2 } from './nativeRenderModelV2.js'
import { EMU_PER_CSS_PIXEL, EMU_PER_POINT, isCompiledNativeSheetGeometryV2, validateNativeSheetGeometryV2 } from './nativeSheetGeometryV2.js'
import type { NativeSheetGeometryRectV2, NativeSheetGeometryV2 } from './nativeSheetGeometryV2.js'
import { sha256Hex, sha256HexBytes } from './nativeSha256.js'
import { nativeWorkbookStyleRawProjectionSha256V2 } from './nativeValidationV2.js'
import { NativePlainDataError, snapshotNativePlainData } from './nativePlainData.js'

export const NATIVE_SHEET_CELL_PAINT_V2_PROTOCOL = 'injoffice.xlsx.sheet-cell-paint'
export const NATIVE_SHEET_CELL_PAINT_V2_VERSION = 1 as const
export const NATIVE_XLSX_CELL_PAINT_SHAPER_SOURCE_REVISION = 'injoffice.xlsx.cell-paint.v2'
export const NATIVE_SHEET_CELL_GUTTER_EMU = 2 * EMU_PER_CSS_PIXEL
export const NATIVE_SHEET_CELL_PAINT_V2_LIMITS = Object.freeze({
  maxCells: 100_000,
  maxGlyphs: 100_000,
  maxUnsupported: 100_000,
  maxPathCommandsPerGlyph: 65_536,
  maxPathCommands: 4_000_000,
  maxCommands: 300_003,
  maxDesignCoordinate: 1_000_000_000,
})

const PRINTABLE_ASCII = /^[\x20-\x7E]+$/
const RGB = /^#[0-9A-F]{6}$/
const maximumFontBytes = 64 * 1024 * 1024
const appearanceAuthorityCapabilities = new Set(['conditional-formatting', 'tables', 'drawings', 'external-links', 'formula-groups'])
const appearanceAuthorityCodes = new Set([
  'FOREIGN_WORKSHEET_MARKUP', 'WORKSHEET_ATTRIBUTES', 'SHEET_VIEW_GEOMETRY', 'WORKSHEET_EXTENSIONS', 'STYLE_RECORD_ATTRIBUTES', 'STYLE_XF_OPAQUE_CONTENT',
])

export type NativeSheetCellPaintIssueCode =
  | 'paint.sheetMissing'
  | 'paint.geometryAuthority'
  | 'paint.modelAuthority'
  | 'paint.sourceUnsupported'
  | 'paint.metricAuthority'
  | 'paint.styleUnsupported'
  | 'paint.resourceBudget'
  | 'paint.commandBudget'
  | 'paint.invalidLimit'
  | 'paint.planInvalid'
  | 'paint.planDigest'
  | 'paint.commandInvalid'

export class NativeSheetCellPaintError extends Error {
  readonly code: NativeSheetCellPaintIssueCode
  readonly path: string

  constructor(code: NativeSheetCellPaintIssueCode, path: string, message: string) {
    super(message)
    this.name = 'NativeSheetCellPaintError'
    this.code = code
    this.path = path
  }
}

export type NativeSheetCellPaintDisplayKindV2 =
  | 'string-literal'
  | 'numeric-lexical'
  | 'formula-cached-lexical'
  | 'boolean-display'
  | 'error-lexical'
  | 'date-display'

/** Exact producer-issued format strings this slice can apply without locale, Intl, or host TZ. */
export const NATIVE_SHEET_CELL_PAINT_NUMBER_FORMATS_V2 = Object.freeze(['General', '0', '0.00', 'yyyy-mm-dd', 'yyyy/mm/dd'] as const)
export type NativeSheetCellPaintNumberFormatV2 = (typeof NATIVE_SHEET_CELL_PAINT_NUMBER_FORMATS_V2)[number]

export type NativeSheetCellDisplayFormatResultV2 =
  | { readonly status: 'ready'; readonly text: string }
  | { readonly status: 'refused'; readonly code: 'CELL_DATE_DISPLAY' | 'CELL_STYLE_UNSUPPORTED'; readonly message: string }

export type NativeSheetCellPaintUnsupportedCodeV2 =
  | 'CELL_RICH_TEXT'
  | 'CELL_FORMULA_NO_CACHE'
  | 'CELL_DATE_DISPLAY'
  | 'CELL_BOOLEAN_DISPLAY'
  | 'CELL_ERROR_DISPLAY'
  | 'CELL_NON_ASCII'
  | 'CELL_COMPLEX_CLUSTER'
  | 'CELL_FONT_MISMATCH'
  | 'CELL_FONT_COLOR'
  | 'CELL_WRAP_TEXT'
  | 'CELL_ALIGNMENT_EXTENDED'
  | 'CELL_STYLE_UNSUPPORTED'
  | 'CELL_MISSING_GLYPH'
  | 'CELL_COMPOSITE_OUTLINE'
  | 'CELL_OVERFLOW'
  | 'CELL_MERGE_NON_ORIGIN'
  | 'CELL_HIDDEN'

export type NativeSheetCellPaintPathCommandV2 =
  | { readonly kind: 'move_to'; readonly x_emu: number; readonly y_emu: number }
  | { readonly kind: 'line_to'; readonly x_emu: number; readonly y_emu: number }
  | { readonly kind: 'quadratic_to'; readonly control_x_emu: number; readonly control_y_emu: number; readonly x_emu: number; readonly y_emu: number }
  | { readonly kind: 'close_path' }

export interface NativeSheetPaintedCellV2 {
  readonly cell_ref: string
  readonly row: number
  readonly column: number
  readonly style_id: number
  readonly display_text: string
  readonly display_kind: NativeSheetCellPaintDisplayKindV2
  readonly horizontal_alignment: 'left' | 'center' | 'right'
  readonly vertical_alignment: 'top' | 'middle' | 'bottom'
  readonly rect: NativeSheetGeometryRectV2
  readonly content_box: NativeSheetGeometryRectV2
  readonly glyph_start: number
  readonly glyph_end: number
}

export interface NativeSheetPaintedGlyphV2 {
  readonly cell_ref: string
  readonly row: number
  readonly column: number
  readonly glyph_index: number
  readonly face_id: string
  readonly font_sha256: `sha256:${string}`
  readonly glyph_id: number
  readonly origin_x_emu: number
  readonly origin_y_emu: number
  readonly font_size_millipoints: number
  readonly fill_rgb: string
  readonly fill_rule: 'nonzero'
  readonly outline_kind: 'path'
  readonly path: ReadonlyArray<NativeSheetCellPaintPathCommandV2>
}

export interface NativeSheetCellPaintUnsupportedV2 {
  readonly cell_ref: string
  readonly row: number
  readonly column: number
  readonly code: NativeSheetCellPaintUnsupportedCodeV2
  readonly message: string
}

export interface NativeSheetCellPaintCapabilityV2 {
  readonly name: 'native-cell-glyphs'
  readonly level: 'exact'
}

export interface NativeSheetCellPaintPlanV2 {
  readonly protocol: typeof NATIVE_SHEET_CELL_PAINT_V2_PROTOCOL
  readonly version: typeof NATIVE_SHEET_CELL_PAINT_V2_VERSION
  readonly paint_sha256: `sha256:${string}`
  readonly document_id: string
  readonly sheet_id: string
  readonly source_part: string
  readonly source_revision: string
  readonly source_package_sha256: string
  readonly geometry_sha256: string
  readonly font_sha256: `sha256:${string}`
  readonly shaper_id: string
  readonly shaper_revision: string
  readonly coordinate_space: 'viewport-local'
  readonly bounds: NativeSheetGeometryRectV2
  readonly gutter_emu: typeof NATIVE_SHEET_CELL_GUTTER_EMU
  readonly capabilities: ReadonlyArray<NativeSheetCellPaintCapabilityV2>
  readonly cells: ReadonlyArray<NativeSheetPaintedCellV2>
  readonly glyphs: ReadonlyArray<NativeSheetPaintedGlyphV2>
  readonly unsupported: ReadonlyArray<NativeSheetCellPaintUnsupportedV2>
}

export type NativeSheetCellPaintCommandV2 =
  | { readonly kind: 'beginCellPaint'; readonly protocol: typeof NATIVE_SHEET_CELL_PAINT_V2_PROTOCOL; readonly version: typeof NATIVE_SHEET_CELL_PAINT_V2_VERSION; readonly document_id: string; readonly sheet_id: string; readonly source_part: string; readonly source_revision: string; readonly source_package_sha256: string; readonly geometry_sha256: string; readonly font_sha256: string; readonly shaper_id: string; readonly shaper_revision: string; readonly coordinate_space: 'viewport-local'; readonly gutter_emu: number; readonly capabilities: ReadonlyArray<NativeSheetCellPaintCapabilityV2>; readonly paint_sha256: string }
  | { readonly kind: 'clipRect'; readonly rect: NativeSheetGeometryRectV2 }
  | { readonly kind: 'paintedCell'; readonly cell: NativeSheetPaintedCellV2 }
  | { readonly kind: 'fillGlyphPath'; readonly glyph: NativeSheetPaintedGlyphV2 }
  | { readonly kind: 'unsupportedCell'; readonly unsupported: NativeSheetCellPaintUnsupportedV2 }
  | { readonly kind: 'endCellPaint' }

export interface NativeSheetCellPaintSurfaceV2 {
  readonly command_capacity?: number
  push(command: NativeSheetCellPaintCommandV2): void
}

export interface NativeSheetCellPaintRecordingSurfaceV2 extends NativeSheetCellPaintSurfaceV2 {
  readonly commands: ReadonlyArray<NativeSheetCellPaintCommandV2>
  finish(): ReadonlyArray<NativeSheetCellPaintCommandV2>
}

export interface NativeSheetCellPaintCommandAdapterV2<HostContext> {
  execute(context: HostContext, command: NativeSheetCellPaintCommandV2): void
}

interface PaintFontFace {
  digest: `sha256:${string}`
  bytes: Uint8Array
  metrics: FontDesignMetrics
  resource: FontResource
  names: ReadonlySet<string>
  outlines: Map<number, DesignOutline>
}

interface CellPaintContext {
  workbook: NativeWorkbookRenderModelV2
  sheet: NativeSheetRenderModelV2
  geometry: NativeSheetGeometryV2
  fontBytes: Uint8Array
  fontDigest: `sha256:${string}`
  metrics: FontDesignMetrics
  resource: FontResource
  fonts: PaintFontFace[]
  shaperId: string
  shaperRevision: string
  shape: (text: string, script: 'Latn' | 'Zyyy', fontSizeMilliPoints: number, resource: FontResource) => ShapedSegment | { status: 'refused'; code: string; message: string }
}

type DesignPathCommand =
  | { kind: 'move_to'; x: number; y: number }
  | { kind: 'line_to'; x: number; y: number }
  | { kind: 'quadratic_to'; control_x: number; control_y: number; x: number; y: number }
  | { kind: 'close_path' }

type DesignOutline = { status: 'empty'; unitsPerEm: number } | { status: 'outlined'; unitsPerEm: number; path: DesignPathCommand[] } | { status: 'composite' } | { status: 'invalid' }

interface QualifiedDisplay {
  text: string
  kind: NativeSheetCellPaintDisplayKindV2
  numeric: boolean
  runs?: ReadonlyArray<{ text: string; font_name?: string; bold?: boolean; italic?: boolean; font_size_points?: number; font_color?: string }>
}

export type NativeSheetCellPaintFontAuthorityV2 =
  | Uint8Array
  | { readonly font_sha256: `sha256:${string}`; readonly bytes: Uint8Array }

export function compileNativeSheetCellPaintV2(workbook: NativeWorkbookRenderModelV2, geometry: NativeSheetGeometryV2, fontBytes: Uint8Array, additionalFontBytes: ReadonlyArray<NativeSheetCellPaintFontAuthorityV2> = []): NativeSheetCellPaintPlanV2 {
  if (!isProjectedNativeWorkbookV2(workbook)) throw new NativeSheetCellPaintError('paint.modelAuthority', '$.workbook', 'workbook must be the branded frozen result of projectNativeWorkbookV2')
  validateGeometryAuthority(workbook, geometry)
  const sheet = workbook.sheets.find((candidate) => candidate.id === geometry.sheet_id)
  if (!sheet) throw new NativeSheetCellPaintError('paint.sheetMissing', '$.geometry.sheet_id', `sheet ${JSON.stringify(geometry.sheet_id)} is absent`)
  validateSourceAuthority(workbook, sheet)
  const bytes = snapshotFontBytes(fontBytes, geometry.metric_authority.font_sha256)
  const extra = snapshotAdditionalFontBytes(additionalFontBytes, bytes)
  const context = createPaintContext(workbook, sheet, geometry, bytes, extra)
  const cells = new Map<number, NativeRenderCellV2>()
  for (const cell of sheet.cells) {
    const key = cell.row * 16_384 + cell.column
    if (!Number.isSafeInteger(cell.row) || !Number.isSafeInteger(cell.column) || cell.row < 0 || cell.row >= 1_048_576 || cell.column < 0 || cell.column >= 16_384 || cell.ref !== cellReference(cell.row, cell.column) || cells.has(key)) {
      throw new NativeSheetCellPaintError('paint.sourceUnsupported', '$.sheet.cells', 'cell coordinates or identity are non-canonical')
    }
    cells.set(key, cell)
  }
  const mergeByCovered = new Map<number, NativeSheetGeometryV2['merged_ranges'][number]>()
  for (const range of geometry.merged_ranges) {
    for (let row = range.row; row <= range.end_row; row++) {
      for (let column = range.column; column <= range.end_column; column++) mergeByCovered.set(row * 16_384 + column, range)
    }
  }

  const painted: NativeSheetPaintedCellV2[] = []
  const glyphs: NativeSheetPaintedGlyphV2[] = []
  const unsupported: NativeSheetCellPaintUnsupportedV2[] = []
  let pathCommands = 0

  for (const rowBand of geometry.rows) {
    for (const columnBand of geometry.columns) {
      const key = rowBand.row * 16_384 + columnBand.column
      const cell = cells.get(key)
      const merge = mergeByCovered.get(key)
      if (merge && (merge.row !== rowBand.row || merge.column !== columnBand.column)) {
        if (cellHasDisplayContent(cell?.content)) unsupported.push(refusal(rowBand.row, columnBand.column, 'CELL_MERGE_NON_ORIGIN', 'merged-cell text is painted only from the top-left cell'))
        continue
      }
      if (!cell || cell.content.kind === 'blank') continue
      const rect = merge
        ? { ...merge.rect }
        : { x_emu: columnBand.x_emu, y_emu: rowBand.y_emu, width_emu: columnBand.width_emu, height_emu: rowBand.height_emu }
      if (rowBand.height_emu === 0 || columnBand.width_emu === 0 || rect.width_emu === 0 || rect.height_emu === 0) {
        unsupported.push(refusal(rowBand.row, columnBand.column, 'CELL_HIDDEN', 'hidden or zero-size cells cannot host exact glyph origins'))
        continue
      }
      const paintedCell = paintOneCell(context, cell, rect, glyphs.length)
      if ('code' in paintedCell) {
        unsupported.push(refusal(cell.row, cell.column, paintedCell.code, paintedCell.message))
        continue
      }
      if (paintedCell.glyphs.length + glyphs.length > NATIVE_SHEET_CELL_PAINT_V2_LIMITS.maxGlyphs) throw new NativeSheetCellPaintError('paint.resourceBudget', '$.glyphs', 'cell glyphs exceed the native resource bound')
      pathCommands += paintedCell.glyphs.reduce((sum, glyph) => sum + glyph.path.length, 0)
      if (pathCommands > NATIVE_SHEET_CELL_PAINT_V2_LIMITS.maxPathCommands) throw new NativeSheetCellPaintError('paint.resourceBudget', '$.glyphs', 'paint path commands exceed the native resource bound')
      if (painted.length >= NATIVE_SHEET_CELL_PAINT_V2_LIMITS.maxCells) throw new NativeSheetCellPaintError('paint.resourceBudget', '$.cells', 'painted cells exceed the native resource bound')
      painted.push(paintedCell.cell)
      glyphs.push(...paintedCell.glyphs)
    }
  }
  if (unsupported.length > NATIVE_SHEET_CELL_PAINT_V2_LIMITS.maxUnsupported) throw new NativeSheetCellPaintError('paint.resourceBudget', '$.unsupported', 'unsupported cell inventory exceeds the native resource bound')
  const unsigned: Omit<NativeSheetCellPaintPlanV2, 'paint_sha256'> = {
    protocol: NATIVE_SHEET_CELL_PAINT_V2_PROTOCOL,
    version: NATIVE_SHEET_CELL_PAINT_V2_VERSION,
    document_id: workbook.document_id,
    sheet_id: sheet.id,
    source_part: sheet.mutation_authority.source_part,
    source_revision: workbook.revision,
    source_package_sha256: workbook.source.package_sha256,
    geometry_sha256: geometry.geometry_sha256,
    font_sha256: context.fontDigest,
    shaper_id: context.shaperId,
    shaper_revision: context.shaperRevision,
    coordinate_space: 'viewport-local',
    bounds: { ...geometry.bounds },
    gutter_emu: NATIVE_SHEET_CELL_GUTTER_EMU,
    capabilities: [{ name: 'native-cell-glyphs', level: 'exact' }],
    cells: painted,
    glyphs,
    unsupported,
  }
  return deepFreeze({ ...unsigned, paint_sha256: paintDigest(unsigned) })
}

export function createNativeSheetCellPaintRecordingSurfaceV2(maxCommands: number = NATIVE_SHEET_CELL_PAINT_V2_LIMITS.maxCommands): NativeSheetCellPaintRecordingSurfaceV2 {
  validateCommandLimit(maxCommands, '$.maxCommands')
  const commands: NativeSheetCellPaintCommandV2[] = []
  let finished: ReadonlyArray<NativeSheetCellPaintCommandV2> | undefined
  return {
    command_capacity: maxCommands,
    get commands() { return finished ?? Object.freeze([...commands]) },
    push(command) {
      if (finished) throw new Error('native sheet cell-paint recording is already finished')
      if (commands.length >= maxCommands) throw new NativeSheetCellPaintError('paint.commandBudget', '$.commands', `cell-paint commands exceed ${maxCommands}`)
      commands.push(deepFreeze(command))
    },
    finish() {
      if (!finished) finished = Object.freeze([...commands])
      return finished
    },
  }
}

export function emitNativeSheetCellPaintCommandsV2(plan: NativeSheetCellPaintPlanV2, surface: NativeSheetCellPaintSurfaceV2, maxCommands: number = NATIVE_SHEET_CELL_PAINT_V2_LIMITS.maxCommands): void {
  const validated = validateNativeSheetCellPaintPlanV2(plan)
  validateCommandLimit(maxCommands, '$.maxCommands')
  const surfaceCapacity = optionalDataProperty(surface, 'command_capacity', '$.surface.command_capacity')
  const surfacePush = callbackDataProperty(surface, 'push', '$.surface.push')
  if (surfaceCapacity !== undefined) validateCommandLimit(surfaceCapacity, '$.surface.command_capacity')
  const capacity = Math.min(maxCommands, surfaceCapacity ?? maxCommands)
  const commands = paintCommands(validated)
  if (commands.length > capacity) throw new NativeSheetCellPaintError('paint.commandBudget', '$.commands', `cell-paint commands require ${commands.length}, exceeding ${capacity}`)
  for (const command of commands) surfacePush.call(surface, command)
}

export function replayNativeSheetCellPaintCommandsV2<HostContext>(context: HostContext, commands: ReadonlyArray<NativeSheetCellPaintCommandV2>, adapter: NativeSheetCellPaintCommandAdapterV2<HostContext>, maxCommands: number = NATIVE_SHEET_CELL_PAINT_V2_LIMITS.maxCommands): void {
  validateCommandLimit(maxCommands, '$.maxCommands')
  const validated = validateNativeSheetCellPaintCommandsV2(commands, maxCommands)
  const execute = callbackDataProperty(adapter, 'execute', '$.adapter.execute') as (context: HostContext, command: NativeSheetCellPaintCommandV2) => void
  for (const command of validated) execute.call(adapter, context, command)
}

export function validateNativeSheetCellPaintPlanV2(input: unknown): NativeSheetCellPaintPlanV2 {
  input = snapshotPaintInput(input, '$')
  const plan = exactObject(input, ['protocol', 'version', 'paint_sha256', 'document_id', 'sheet_id', 'source_part', 'source_revision', 'source_package_sha256', 'geometry_sha256', 'font_sha256', 'shaper_id', 'shaper_revision', 'coordinate_space', 'bounds', 'gutter_emu', 'capabilities', 'cells', 'glyphs', 'unsupported'], '$')
  if (plan.protocol !== NATIVE_SHEET_CELL_PAINT_V2_PROTOCOL || plan.version !== NATIVE_SHEET_CELL_PAINT_V2_VERSION) invalidPlan('$', 'cell-paint protocol or version is unsupported')
  const documentID = identifier(plan.document_id, '$.document_id', 256)
  const sheetID = canonicalSheetID(plan.sheet_id, '$.sheet_id')
  const sourcePart = canonicalPart(plan.source_part, '$.source_part')
  const sourceRevision = stringPattern(plan.source_revision, '$.source_revision', /^rev:[0-9a-f]{64}$/)
  const sourcePackage = stringPattern(plan.source_package_sha256, '$.source_package_sha256', /^sha256:[0-9a-f]{64}$/)
  if (sourceRevision.slice(4) !== sourcePackage.slice(7)) invalidPlan('$.source_revision', 'source revision does not match package digest')
  const geometryDigest = stringPattern(plan.geometry_sha256, '$.geometry_sha256', /^sha256:[0-9a-f]{64}$/)
  const fontDigest = stringPattern(plan.font_sha256, '$.font_sha256', /^sha256:[0-9a-f]{64}$/) as `sha256:${string}`
  const shaperID = boundedString(plan.shaper_id, '$.shaper_id', 256)
  const shaperRevision = boundedString(plan.shaper_revision, '$.shaper_revision', 256)
  const suppliedDigest = stringPattern(plan.paint_sha256, '$.paint_sha256', /^sha256:[0-9a-f]{64}$/)
  if (plan.coordinate_space !== 'viewport-local') invalidPlan('$.coordinate_space', 'coordinate space must be viewport-local')
  if (plan.gutter_emu !== NATIVE_SHEET_CELL_GUTTER_EMU) invalidPlan('$.gutter_emu', 'gutter must be the producer-issued 2 CSS-pixel Excel inset at the geometry 96dpi EMU basis')
  const bounds = validateRect(plan.bounds, '$.bounds', false)
  if (bounds.x_emu !== 0 || bounds.y_emu !== 0) invalidPlan('$.bounds', 'viewport-local bounds must start at the origin')
  if (!Array.isArray(plan.capabilities) || plan.capabilities.length !== 1) invalidPlan('$.capabilities', 'capabilities must be the canonical native-cell-glyphs exact singleton')
  const capability = exactObject(plan.capabilities[0], ['name', 'level'], '$.capabilities[0]')
  if (capability.name !== 'native-cell-glyphs' || capability.level !== 'exact') invalidPlan('$.capabilities[0]', 'capability name or level is not canonical')
  if (!Array.isArray(plan.cells) || plan.cells.length > NATIVE_SHEET_CELL_PAINT_V2_LIMITS.maxCells) invalidPlan('$.cells', 'painted cell inventory is missing or exceeds its resource bound')
  if (!Array.isArray(plan.glyphs) || plan.glyphs.length > NATIVE_SHEET_CELL_PAINT_V2_LIMITS.maxGlyphs) invalidPlan('$.glyphs', 'glyph inventory is missing or exceeds its resource bound')
  if (!Array.isArray(plan.unsupported) || plan.unsupported.length > NATIVE_SHEET_CELL_PAINT_V2_LIMITS.maxUnsupported) invalidPlan('$.unsupported', 'unsupported inventory is missing or exceeds its resource bound')

  const cells: NativeSheetPaintedCellV2[] = []
  const occupied = new Set<string>()
  let priorCell = -1
  let glyphCursor = 0
  let pathCommands = 0
  for (let index = 0; index < plan.cells.length; index++) {
    const path = `$.cells[${index}]`
    const value = exactObject(plan.cells[index], ['cell_ref', 'row', 'column', 'style_id', 'display_text', 'display_kind', 'horizontal_alignment', 'vertical_alignment', 'rect', 'content_box', 'glyph_start', 'glyph_end'], path)
    const row = boundedInteger(value.row, `${path}.row`, 0, 1_048_575)
    const column = boundedInteger(value.column, `${path}.column`, 0, 16_383)
    const position = row * 16_384 + column
    if (position <= priorCell) invalidPlan(path, 'painted cells must be unique and row-major ordered')
    priorCell = position
    const cellRef = stringPattern(value.cell_ref, `${path}.cell_ref`, /^[A-Z]{1,3}[1-9][0-9]{0,6}$/)
    if (cellRef !== cellReference(row, column)) invalidPlan(`${path}.cell_ref`, 'cell reference does not match coordinates')
    if (occupied.has(cellRef)) invalidPlan(path, 'cell is both painted and unsupported')
    occupied.add(cellRef)
    if (!isCanonicalDisplayText(value.display_text)) invalidPlan(`${path}.display_text`, 'display text must be a non-empty well-formed Unicode string without control characters')
    if (value.display_kind !== 'string-literal' && value.display_kind !== 'numeric-lexical' && value.display_kind !== 'formula-cached-lexical' && value.display_kind !== 'boolean-display' && value.display_kind !== 'error-lexical' && value.display_kind !== 'date-display') invalidPlan(`${path}.display_kind`, 'display kind is unsupported')
    if (value.horizontal_alignment !== 'left' && value.horizontal_alignment !== 'center' && value.horizontal_alignment !== 'right') invalidPlan(`${path}.horizontal_alignment`, 'horizontal alignment is unsupported')
    if (value.vertical_alignment !== 'top' && value.vertical_alignment !== 'middle' && value.vertical_alignment !== 'bottom') invalidPlan(`${path}.vertical_alignment`, 'vertical alignment is unsupported')
    const rect = validateRect(value.rect, `${path}.rect`, true)
    const contentBox = validateRect(value.content_box, `${path}.content_box`, true)
    if (!rectContained(rect, bounds) || !rectContained(contentBox, rect)) invalidPlan(path, 'cell rectangles must nest inside viewport bounds')
    const glyphStart = boundedInteger(value.glyph_start, `${path}.glyph_start`, 0, plan.glyphs.length)
    const glyphEnd = boundedInteger(value.glyph_end, `${path}.glyph_end`, glyphStart + 1, plan.glyphs.length)
    if (glyphStart !== glyphCursor) invalidPlan(`${path}.glyph_start`, 'cell glyph ranges must be contiguous')
    glyphCursor = glyphEnd
    cells.push({
      cell_ref: cellRef, row, column,
      style_id: boundedInteger(value.style_id, `${path}.style_id`, 0, 4_294_967_295),
      display_text: value.display_text,
      display_kind: value.display_kind,
      horizontal_alignment: value.horizontal_alignment,
      vertical_alignment: value.vertical_alignment,
      rect, content_box: contentBox, glyph_start: glyphStart, glyph_end: glyphEnd,
    })
  }
  if (glyphCursor !== plan.glyphs.length) invalidPlan('$.glyphs', 'glyphs must be partitioned exactly by painted cells')

  const glyphs: NativeSheetPaintedGlyphV2[] = []
  for (let index = 0; index < plan.glyphs.length; index++) {
    const path = `$.glyphs[${index}]`
    const value = exactObject(plan.glyphs[index], ['cell_ref', 'row', 'column', 'glyph_index', 'face_id', 'font_sha256', 'glyph_id', 'origin_x_emu', 'origin_y_emu', 'font_size_millipoints', 'fill_rgb', 'fill_rule', 'outline_kind', 'path'], path)
    const owner = cells.find((cell) => index >= cell.glyph_start && index < cell.glyph_end)
    if (!owner) invalidPlan(path, 'glyph is not owned by a painted cell')
    const row = boundedInteger(value.row, `${path}.row`, 0, 1_048_575)
    const column = boundedInteger(value.column, `${path}.column`, 0, 16_383)
    const cellRef = stringPattern(value.cell_ref, `${path}.cell_ref`, /^[A-Z]{1,3}[1-9][0-9]{0,6}$/)
    if (cellRef !== owner.cell_ref || row !== owner.row || column !== owner.column) invalidPlan(path, 'glyph cell identity must match its painted cell')
    const glyphIndex = boundedInteger(value.glyph_index, `${path}.glyph_index`, 0, owner.glyph_end - owner.glyph_start - 1)
    if (glyphIndex !== index - owner.glyph_start) invalidPlan(`${path}.glyph_index`, 'glyph indices must be dense within the cell')
    const glyphFont = stringPattern(value.font_sha256, `${path}.font_sha256`, /^sha256:[0-9a-f]{64}$/) as `sha256:${string}`
    const faceId = boundedString(value.face_id, `${path}.face_id`, 256)
    if (glyphFont === fontDigest) {
      if (faceId !== 'xlsx.normal-font') invalidPlan(`${path}.face_id`, 'face id must be the producer-issued Normal-font identity')
    } else if (faceId !== `xlsx.font:${glyphFont.slice(7)}`) {
      invalidPlan(`${path}.face_id`, 'additional face id must bind the exact font digest')
    }
    if (value.fill_rule !== 'nonzero' || value.outline_kind !== 'path') invalidPlan(path, 'glyph fill rule or outline kind is unsupported')
    if (!Array.isArray(value.path) || value.path.length < 1 || value.path.length > NATIVE_SHEET_CELL_PAINT_V2_LIMITS.maxPathCommandsPerGlyph) invalidPlan(`${path}.path`, 'glyph path is missing or exceeds its bound')
    pathCommands += value.path.length
    if (pathCommands > NATIVE_SHEET_CELL_PAINT_V2_LIMITS.maxPathCommands) invalidPlan(`${path}.path`, 'paint path commands exceed the native resource bound')
    const originX = boundedInteger(value.origin_x_emu, `${path}.origin_x_emu`, 0, Number.MAX_SAFE_INTEGER)
    const originY = boundedInteger(value.origin_y_emu, `${path}.origin_y_emu`, 0, Number.MAX_SAFE_INTEGER)
    const placedPath = validatePath(value.path, `${path}.path`, bounds)
    if (!pathContained(placedPath, owner.rect)) invalidPlan(`${path}.path`, 'glyph path escapes its cell rectangle')
    glyphs.push({
      cell_ref: cellRef, row, column, glyph_index: glyphIndex,
      face_id: faceId,
      font_sha256: glyphFont,
      glyph_id: boundedInteger(value.glyph_id, `${path}.glyph_id`, 1, 0xffff),
      origin_x_emu: originX,
      origin_y_emu: originY,
      font_size_millipoints: boundedInteger(value.font_size_millipoints, `${path}.font_size_millipoints`, 1, 10_000_000),
      fill_rgb: stringPattern(value.fill_rgb, `${path}.fill_rgb`, RGB),
      fill_rule: 'nonzero',
      outline_kind: 'path',
      path: placedPath,
    })
  }

  const unsupported: NativeSheetCellPaintUnsupportedV2[] = []
  let priorUnsupported = -1
  for (let index = 0; index < plan.unsupported.length; index++) {
    const path = `$.unsupported[${index}]`
    const value = exactObject(plan.unsupported[index], ['cell_ref', 'row', 'column', 'code', 'message'], path)
    const row = boundedInteger(value.row, `${path}.row`, 0, 1_048_575)
    const column = boundedInteger(value.column, `${path}.column`, 0, 16_383)
    const position = row * 16_384 + column
    if (position <= priorUnsupported) invalidPlan(path, 'unsupported cells must be unique and row-major ordered')
    priorUnsupported = position
    const cellRef = stringPattern(value.cell_ref, `${path}.cell_ref`, /^[A-Z]{1,3}[1-9][0-9]{0,6}$/)
    if (cellRef !== cellReference(row, column)) invalidPlan(`${path}.cell_ref`, 'cell reference does not match coordinates')
    if (occupied.has(cellRef)) invalidPlan(path, 'cell is both painted and unsupported')
    occupied.add(cellRef)
    if (typeof value.code !== 'string' || !unsupportedCodes.has(value.code as NativeSheetCellPaintUnsupportedCodeV2)) invalidPlan(`${path}.code`, 'unsupported code is not canonical')
    if (typeof value.message !== 'string' || value.message.length < 1 || value.message.length > 4_096 || /[\u0000\r\n]/.test(value.message)) invalidPlan(`${path}.message`, 'unsupported message is missing, oversized, or non-canonical')
    unsupported.push({ cell_ref: cellRef, row, column, code: value.code as NativeSheetCellPaintUnsupportedCodeV2, message: value.message })
  }
  if (3 + cells.length + glyphs.length + unsupported.length > NATIVE_SHEET_CELL_PAINT_V2_LIMITS.maxCommands) invalidPlan('$', 'plan exceeds the cell-paint command bound')
  const unsigned: Omit<NativeSheetCellPaintPlanV2, 'paint_sha256'> = {
    protocol: NATIVE_SHEET_CELL_PAINT_V2_PROTOCOL,
    version: NATIVE_SHEET_CELL_PAINT_V2_VERSION,
    document_id: documentID,
    sheet_id: sheetID,
    source_part: sourcePart,
    source_revision: sourceRevision,
    source_package_sha256: sourcePackage,
    geometry_sha256: geometryDigest,
    font_sha256: fontDigest,
    shaper_id: shaperID,
    shaper_revision: shaperRevision,
    coordinate_space: 'viewport-local',
    bounds,
    gutter_emu: NATIVE_SHEET_CELL_GUTTER_EMU,
    capabilities: [{ name: 'native-cell-glyphs', level: 'exact' }],
    cells,
    glyphs,
    unsupported,
  }
  if (paintDigest(unsigned) !== suppliedDigest) throw new NativeSheetCellPaintError('paint.planDigest', '$.paint_sha256', 'cell-paint digest does not match the canonical plan')
  return deepFreeze({ ...unsigned, paint_sha256: suppliedDigest as `sha256:${string}` })
}

function validateNativeSheetCellPaintCommandsV2(input: unknown, maxCommands: number): ReadonlyArray<NativeSheetCellPaintCommandV2> {
  input = snapshotPaintInput(input, '$.commands')
  if (!Array.isArray(input) || input.length < 3 || input.length > maxCommands) throw new NativeSheetCellPaintError('paint.commandInvalid', '$.commands', 'command stream length is outside its bound')
  const begin = exactCommand(input[0], ['kind', 'protocol', 'version', 'document_id', 'sheet_id', 'source_part', 'source_revision', 'source_package_sha256', 'geometry_sha256', 'font_sha256', 'shaper_id', 'shaper_revision', 'coordinate_space', 'gutter_emu', 'capabilities', 'paint_sha256'], '$.commands[0]')
  if (begin.kind !== 'beginCellPaint') commandInvalid('$.commands[0].kind', 'command stream must begin with beginCellPaint')
  const clip = exactCommand(input[1], ['kind', 'rect'], '$.commands[1]')
  if (clip.kind !== 'clipRect') commandInvalid('$.commands[1].kind', 'second command must be clipRect')
  const end = exactCommand(input[input.length - 1], ['kind'], `$.commands[${input.length - 1}]`)
  if (end.kind !== 'endCellPaint') commandInvalid(`$.commands[${input.length - 1}].kind`, 'command stream must end with endCellPaint')
  const cells: unknown[] = [], glyphs: unknown[] = [], unsupported: unknown[] = []
  let lane: 'cells' | 'unsupported' = 'cells'
  let expectedGlyphs = 0
  for (let index = 2; index < input.length - 1; index++) {
    const raw = input[index]
    if (typeof raw !== 'object' || raw === null || Array.isArray(raw)) commandInvalid(`$.commands[${index}]`, 'command must be an object')
    const kind = (raw as Record<string, unknown>).kind
    if (kind === 'paintedCell' && lane === 'cells' && expectedGlyphs === 0) {
      const command = exactCommand(raw, ['kind', 'cell'], `$.commands[${index}]`)
      const cell = command.cell as { glyph_end?: unknown; glyph_start?: unknown }
      if (typeof cell !== 'object' || cell === null) commandInvalid(`$.commands[${index}].cell`, 'painted cell is missing')
      expectedGlyphs = (typeof cell.glyph_end === 'number' && typeof cell.glyph_start === 'number') ? cell.glyph_end - cell.glyph_start : -1
      cells.push(command.cell)
    } else if (kind === 'fillGlyphPath' && lane === 'cells' && expectedGlyphs > 0) {
      const command = exactCommand(raw, ['kind', 'glyph'], `$.commands[${index}]`)
      glyphs.push(command.glyph)
      expectedGlyphs -= 1
    } else if (kind === 'unsupportedCell' && expectedGlyphs <= 0) {
      lane = 'unsupported'
      const command = exactCommand(raw, ['kind', 'unsupported'], `$.commands[${index}]`)
      unsupported.push(command.unsupported)
    } else commandInvalid(`$.commands[${index}].kind`, 'commands must contain painted cells and glyphs followed by unsupported cells')
  }
  if (expectedGlyphs !== 0 && expectedGlyphs !== -1) commandInvalid('$.commands', 'painted cell glyph counts must match fillGlyphPath commands')
  const plan = validateNativeSheetCellPaintPlanV2({
    protocol: begin.protocol,
    version: begin.version,
    paint_sha256: begin.paint_sha256,
    document_id: begin.document_id,
    sheet_id: begin.sheet_id,
    source_part: begin.source_part,
    source_revision: begin.source_revision,
    source_package_sha256: begin.source_package_sha256,
    geometry_sha256: begin.geometry_sha256,
    font_sha256: begin.font_sha256,
    shaper_id: begin.shaper_id,
    shaper_revision: begin.shaper_revision,
    coordinate_space: begin.coordinate_space,
    bounds: clip.rect,
    gutter_emu: begin.gutter_emu,
    capabilities: begin.capabilities,
    cells,
    glyphs,
    unsupported,
  })
  return paintCommands(plan)
}

function paintCommands(plan: NativeSheetCellPaintPlanV2): ReadonlyArray<NativeSheetCellPaintCommandV2> {
  const commands: NativeSheetCellPaintCommandV2[] = [
    {
      kind: 'beginCellPaint',
      protocol: plan.protocol,
      version: plan.version,
      document_id: plan.document_id,
      sheet_id: plan.sheet_id,
      source_part: plan.source_part,
      source_revision: plan.source_revision,
      source_package_sha256: plan.source_package_sha256,
      geometry_sha256: plan.geometry_sha256,
      font_sha256: plan.font_sha256,
      shaper_id: plan.shaper_id,
      shaper_revision: plan.shaper_revision,
      coordinate_space: plan.coordinate_space,
      gutter_emu: plan.gutter_emu,
      capabilities: plan.capabilities,
      paint_sha256: plan.paint_sha256,
    },
    { kind: 'clipRect', rect: plan.bounds },
  ]
  for (const cell of plan.cells) {
    commands.push({ kind: 'paintedCell', cell })
    for (let index = cell.glyph_start; index < cell.glyph_end; index++) commands.push({ kind: 'fillGlyphPath', glyph: plan.glyphs[index]! })
  }
  for (const item of plan.unsupported) commands.push({ kind: 'unsupportedCell', unsupported: item })
  commands.push({ kind: 'endCellPaint' })
  return deepFreeze(commands)
}

function createPaintContext(workbook: NativeWorkbookRenderModelV2, sheet: NativeSheetRenderModelV2, geometry: NativeSheetGeometryV2, fontBytes: Uint8Array, additional: Uint8Array[]): CellPaintContext {
  const digest = geometry.metric_authority.font_sha256
  const normal = createPaintFontFace(fontBytes, digest, 'xlsx.normal-font', geometry.metric_authority.font_name, geometry.metric_authority.font_bold, geometry.metric_authority.font_italic, 'xlsx.metric-authority-font')
  const fonts = [normal]
  for (const extra of additional) {
    const extraDigest = `sha256:${sha256HexBytes(extra)}` as `sha256:${string}`
    if (extraDigest === digest) continue
    const names = fontFileNames(extra)
    const family = names[0] ?? extraDigest
    const style = fontOs2Style(extra)
    fonts.push(createPaintFontFace(extra, extraDigest, `xlsx.font:${extraDigest.slice(7)}`, family, style.bold, style.italic, `xlsx.font:${extraDigest.slice(7)}`))
  }
  let shaper
  try { shaper = createHarfBuzzTextShaperV1({ sourceRevision: NATIVE_XLSX_CELL_PAINT_SHAPER_SOURCE_REVISION }) }
  catch (error) { throw new NativeSheetCellPaintError('paint.metricAuthority', '$.shaper', error instanceof Error ? error.message : 'HarfBuzz shaper construction failed') }
  return {
    workbook, sheet, geometry, fontBytes, fontDigest: digest, metrics: normal.metrics, resource: normal.resource, fonts,
    shaperId: shaper.providerId,
    shaperRevision: shaper.providerRevision,
    shape(text, script, fontSizeMilliPoints, resource) {
      const run: TextRunInput = {
        version: NATIVE_TEXT_LAYOUT_VERSION,
        text,
        fontSizeMilliPoints,
        font: { families: [resource.face.family], weight: resource.face.weight, style: resource.face.style, stretch: 100 },
        script,
        language: 'en-US',
        direction: 'ltr',
      }
      const result = shaper.shape({ run, startUtf16: 0, endUtf16: text.length, font: resource })
      if (typeof result === 'object' && result !== null && 'then' in result) return { status: 'refused', code: 'provider-failure', message: 'cell-paint v2 requires a synchronous HarfBuzz shape result' }
      if ('status' in result && result.status === 'refused') return { status: 'refused', code: result.decisions[0]?.code ?? 'provider-failure', message: result.decisions[0]?.message ?? 'HarfBuzz refused the cell run' }
      return result as ShapedSegment
    },
  }
}

function createPaintFontFace(bytes: Uint8Array, digest: `sha256:${string}`, faceId: string, family: string, bold: boolean, italic: boolean, resourceId: string): PaintFontFace {
  let metrics: FontDesignMetrics
  try { metrics = inspectHarfBuzzFontMetricsV1({ bytes, contentDigest: digest }) }
  catch (error) { throw new NativeSheetCellPaintError('paint.metricAuthority', '$.font_bytes', error instanceof Error ? error.message : 'font metric inspection failed') }
  const face: ResolvedFontFace = Object.freeze({
    faceId, family, weight: bold ? 700 : 400, style: italic ? 'italic' : 'normal', stretch: 100,
    sourceKind: 'bundled', resourceId, contentDigest: digest, resolution: 'exact', matchedFamily: family,
  })
  return { digest, bytes, metrics, resource: Object.freeze({ face, bytes, metrics }), names: new Set(fontFileNames(bytes).map(normalizePaintFontName)), outlines: new Map() }
}

function snapshotAdditionalFontBytes(additionalFontBytes: ReadonlyArray<NativeSheetCellPaintFontAuthorityV2>, primary: Uint8Array): Uint8Array[] {
  if (!Array.isArray(additionalFontBytes)) throw new NativeSheetCellPaintError('paint.metricAuthority', '$.additional_font_bytes', 'additional font bytes must be an array of bounded Uint8Array values')
  if (additionalFontBytes.length > 16) throw new NativeSheetCellPaintError('paint.resourceBudget', '$.additional_font_bytes', 'additional font faces exceed the native resource bound')
  const seen = new Set<string>([`sha256:${sha256HexBytes(primary)}`])
  const extra: Uint8Array[] = []
  for (const candidate of additionalFontBytes) {
    const { bytes, expected } = snapshotFontAuthority(candidate)
    const digest = `sha256:${sha256HexBytes(bytes)}`
    if (expected && expected !== digest) throw new NativeSheetCellPaintError('paint.metricAuthority', '$.additional_font_bytes', 'font-byte authority digest does not match the supplied TTF bytes')
    if (seen.has(digest)) continue
    seen.add(digest)
    extra.push(bytes)
  }
  return extra
}

function snapshotFontAuthority(candidate: NativeSheetCellPaintFontAuthorityV2): { bytes: Uint8Array; expected?: `sha256:${string}` } {
  if (candidate instanceof Uint8Array) {
    if (Object.getPrototypeOf(candidate) !== Uint8Array.prototype || candidate.byteLength < 12 || candidate.byteLength > maximumFontBytes) {
      throw new NativeSheetCellPaintError('paint.metricAuthority', '$.additional_font_bytes', 'additional font bytes must be a bounded direct Uint8Array')
    }
    return { bytes: candidate.slice() }
  }
  if (typeof candidate !== 'object' || candidate === null || Array.isArray(candidate)) {
    throw new NativeSheetCellPaintError('paint.metricAuthority', '$.additional_font_bytes', 'additional font authority must be bytes or a sha256-keyed byte record')
  }
  const digest = (candidate as { font_sha256?: unknown }).font_sha256
  const raw = (candidate as { bytes?: unknown }).bytes
  if (typeof digest !== 'string' || !/^sha256:[0-9a-f]{64}$/.test(digest)) {
    throw new NativeSheetCellPaintError('paint.metricAuthority', '$.additional_font_bytes', 'font-byte authority requires an exact sha256 digest')
  }
  if (!(raw instanceof Uint8Array) || Object.getPrototypeOf(raw) !== Uint8Array.prototype || raw.byteLength < 12 || raw.byteLength > maximumFontBytes) {
    throw new NativeSheetCellPaintError('paint.metricAuthority', '$.additional_font_bytes', 'additional font bytes must be a bounded direct Uint8Array')
  }
  return { bytes: raw.slice(), expected: digest as `sha256:${string}` }
}

type QualifiedStyle = {
  fontSizeMilliPoints: number
  fillRgb: string
  horizontal: 'general' | 'left' | 'center' | 'right'
  vertical: 'top' | 'middle' | 'bottom'
  wrap: boolean
  shrink: boolean
  rotation: 0 | 90 | 180 | 270
  font: PaintFontFace
  fontName: string
  bold: boolean
  italic: boolean
}

type QualifiedRun = {
  text: string
  fontSizeMilliPoints: number
  fillRgb: string
  font: PaintFontFace
}

type ShapedLine = { segment: ShapedSegment; advanceEmu: number; run: QualifiedRun }

function paintOneCell(context: CellPaintContext, cell: NativeRenderCellV2, rect: NativeSheetGeometryRectV2, glyphStart: number): { cell: NativeSheetPaintedCellV2; glyphs: NativeSheetPaintedGlyphV2[] } | { code: NativeSheetCellPaintUnsupportedCodeV2; message: string } {
  const display = qualifyDisplay(cell.content, context.workbook.styles[cell.style_id]?.effective.number_format, context.workbook.date1904)
  if ('code' in display) return display
  const style = qualifyStyle(context, cell.style_id)
  if ('code' in style) return style
  const runs = qualifyRuns(context, style, display)
  if ('code' in runs) return runs
  const contentBox = insetContentBox(rect)
  if (!contentBox) return { code: 'CELL_OVERFLOW', message: 'cell content box is smaller than the producer-issued gutter' }
  const script = paintScript(display.text)
  const shapeRun = (run: QualifiedRun, text: string): ShapedSegment | { code: NativeSheetCellPaintUnsupportedCodeV2; message: string } => {
    const shaped = context.shape(text, script, run.fontSizeMilliPoints, run.font.resource)
    if ('status' in shaped && shaped.status === 'refused') {
      return { code: shaped.code === 'missing-glyph' ? 'CELL_MISSING_GLYPH' : 'CELL_COMPLEX_CLUSTER', message: shaped.message }
    }
    return shaped as ShapedSegment
  }
  let lines: ShapedLine[]
  if (style.wrap) {
    if (runs.length !== 1) return { code: 'CELL_WRAP_TEXT', message: 'wrapped rich runs cannot split without inventing break policy' }
    const wrapped = wrapDisplayLines(display.text, contentBox.width_emu, (text) => shapeRun(runs[0]!, text), runs[0]!, style.shrink)
    if ('code' in wrapped) return wrapped
    lines = wrapped
  } else {
    const lineRuns: ShapedLine[] = []
    let advanceEmu = 0
    for (const run of runs) {
      const segment = shapeRun(run, run.text)
      if ('code' in segment) return segment
      const runAdvance = millipointsToEmu(segment.advanceInlineMilliPoints)
      if (!Number.isSafeInteger(advanceEmu + runAdvance)) return { code: 'CELL_OVERFLOW', message: 'shaped advance overflows the cell content box; neighbor overflow and shrink-to-fit are refused' }
      advanceEmu += runAdvance
      lineRuns.push({ segment, advanceEmu: runAdvance, run })
    }
    if (advanceEmu > contentBox.width_emu && !style.shrink) return { code: 'CELL_OVERFLOW', message: 'shaped advance overflows the cell content box; neighbor overflow and shrink-to-fit are refused' }
    lines = lineRuns
  }
  const wrapLines = style.wrap ? groupWrapLines(lines) : [lines]
  let scaleNum = 1, scaleDen = 1
  if (style.shrink) {
    let maxAdvance = 0
    for (const line of wrapLines) {
      const advance = line.reduce((sum, part) => sum + part.advanceEmu, 0)
      if (advance > maxAdvance) maxAdvance = advance
    }
    if (maxAdvance > contentBox.width_emu) {
      scaleNum = contentBox.width_emu
      scaleDen = maxAdvance
      if (scaleNum < 1 || scaleDen < 1) return { code: 'CELL_OVERFLOW', message: 'shrink-to-fit cannot scale this advance into integer EMU' }
    }
  }
  const unscaledAdvance = wrapLines[0]!.reduce((sum, part) => sum + part.advanceEmu, 0)
  const usedAdvance = style.wrap ? undefined : scaleEmu(unscaledAdvance, scaleNum, scaleDen)
  const metricsAscent = millipointsToEmu(lines[0]!.segment.metrics.ascentMilliPoints)
  const metricsDescent = millipointsToEmu(lines[0]!.segment.metrics.descentMilliPoints)
  const lineHeightEmu = millipointsToEmu(lines[0]!.segment.metrics.lineHeightMilliPoints)
  if (wrapLines.length > 1 && lineHeightEmu < 1) return { code: 'CELL_WRAP_TEXT', message: 'wrapped cell text has no positive HarfBuzz line height' }
  const extra = (wrapLines.length - 1) * lineHeightEmu
  if (!Number.isSafeInteger(extra) || extra < 0) return { code: 'CELL_WRAP_TEXT', message: 'wrapped cell line spacing exceeds deterministic integer EMU bounds' }
  const alignment = resolvedAlignment(style.horizontal, display.numeric)
  const firstBaseline = style.vertical === 'top' ? rect.y_emu + metricsAscent
    : style.vertical === 'middle' ? rect.y_emu + Math.floor((rect.height_emu + metricsAscent + metricsDescent - extra) / 2)
    : rect.y_emu + rect.height_emu + metricsDescent - extra
  const top = firstBaseline - metricsAscent
  const bottom = firstBaseline + extra - metricsDescent
  if (style.rotation === 0 && (top < rect.y_emu || bottom > rect.y_emu + rect.height_emu)) {
    return { code: 'CELL_OVERFLOW', message: 'glyph origins or em-box escape the cell rectangle' }
  }
  const glyphs: NativeSheetPaintedGlyphV2[] = []
  for (let lineIndex = 0; lineIndex < wrapLines.length; lineIndex++) {
    const line = wrapLines[lineIndex]!
    const lineAdvance = scaleEmu(line.reduce((sum, part) => sum + part.advanceEmu, 0), scaleNum, scaleDen)
    const originX = alignment === 'left' ? contentBox.x_emu
      : alignment === 'right' ? contentBox.x_emu + contentBox.width_emu - (style.wrap ? lineAdvance : usedAdvance!)
      : contentBox.x_emu + Math.floor((contentBox.width_emu - (style.wrap ? lineAdvance : usedAdvance!)) / 2)
    if (originX < contentBox.x_emu || originX + lineAdvance > contentBox.x_emu + contentBox.width_emu) {
      return { code: 'CELL_OVERFLOW', message: 'glyph origins or em-box escape the cell rectangle' }
    }
    const baselineY = firstBaseline + lineIndex * lineHeightEmu
    let pen = originX
    for (const part of line) {
      const painted = paintLineGlyphs(cell, part.run, rect, part.segment, pen, baselineY, part.run.fontSizeMilliPoints, glyphs.length, originX, scaleNum, scaleDen)
      if ('code' in painted) return painted
      glyphs.push(...painted)
      pen += scaleEmu(part.advanceEmu, scaleNum, scaleDen)
    }
  }
  if (glyphs.length === 0) return { code: 'CELL_MISSING_GLYPH', message: 'visible display produced no fillable glyph outlines' }
  const rotated = style.rotation === 0 ? glyphs : rotateGlyphs(glyphs, rect, style.rotation)
  if ('code' in rotated) return rotated
  return {
    cell: {
      cell_ref: cell.ref,
      row: cell.row,
      column: cell.column,
      style_id: cell.style_id,
      display_text: display.text,
      display_kind: display.kind,
      horizontal_alignment: alignment,
      vertical_alignment: style.vertical,
      rect,
      content_box: contentBox,
      glyph_start: glyphStart,
      glyph_end: glyphStart + rotated.length,
    },
    glyphs: rotated,
  }
}

function groupWrapLines(parts: ShapedLine[]): ShapedLine[][] {
  return parts.map((part) => [part])
}

/** Greedy wrap at U+0020 from HarfBuzz advances; otherwise a measured character break. */
function wrapDisplayLines(
  text: string,
  maxWidthEmu: number,
  shapeLine: (text: string) => ShapedSegment | { code: NativeSheetCellPaintUnsupportedCodeV2; message: string },
  run: QualifiedRun,
  allowOverfullAtom = false,
): ShapedLine[] | { code: NativeSheetCellPaintUnsupportedCodeV2; message: string } {
  const lines: ShapedLine[] = []
  let start = 0
  while (start < text.length) {
    const rest = text.slice(start)
    const full = shapeLine(rest)
    if ('code' in full) return full
    const fullAdvance = millipointsToEmu(full.advanceInlineMilliPoints)
    if (fullAdvance <= maxWidthEmu) {
      lines.push({ segment: full, advanceEmu: fullAdvance, run })
      break
    }
    let low = 1, high = rest.length - 1, fit = 0
    let fitSegment: ShapedSegment | undefined
    while (low <= high) {
      let mid = (low + high) >> 1
      while (mid > 0 && isTrailSurrogate(rest.charCodeAt(mid)!)) mid--
      if (mid < 1) { high = ((low + high) >> 1) - 1; continue }
      const candidate = rest.slice(0, mid)
      const shaped = shapeLine(candidate)
      if ('code' in shaped) return shaped
      if (millipointsToEmu(shaped.advanceInlineMilliPoints) <= maxWidthEmu) {
        fit = mid
        fitSegment = shaped
        low = mid + 1
      } else high = mid - 1
    }
    if (fit === 0 || !fitSegment) {
      if (!allowOverfullAtom) return { code: 'CELL_OVERFLOW', message: 'wrapped cell text cannot place the next character inside the content box' }
      let take = 1
      if (isLeadSurrogate(rest.charCodeAt(0)!) && isTrailSurrogate(rest.charCodeAt(1) ?? 0)) take = 2
      const lineText = rest.slice(0, take)
      if (lineText.length === 0) return { code: 'CELL_WRAP_TEXT', message: 'wrapped cell text cannot progress without inventing a break' }
      const segment = shapeLine(lineText)
      if ('code' in segment) return segment
      lines.push({ segment, advanceEmu: millipointsToEmu(segment.advanceInlineMilliPoints), run })
      start += take
      if (lines.length > 4_096) return { code: 'CELL_WRAP_TEXT', message: 'wrapped cell text exceeds the bounded line count' }
      continue
    }
    const prefix = rest.slice(0, fit)
    const breakAt = prefix.lastIndexOf(' ')
    const skipBreakSpace = breakAt > 0
    const take = skipBreakSpace ? breakAt : fit
    const lineText = rest.slice(0, take)
    if (lineText.length === 0) return { code: 'CELL_WRAP_TEXT', message: 'wrapped cell text cannot progress without inventing a break' }
    const segment = take === fit ? fitSegment : shapeLine(lineText)
    if ('code' in segment) return segment
    const advanceEmu = millipointsToEmu(segment.advanceInlineMilliPoints)
    if (advanceEmu > maxWidthEmu) return { code: 'CELL_WRAP_TEXT', message: 'wrapped line exceeds the content box after reshaping' }
    lines.push({ segment, advanceEmu, run })
    start += take + (skipBreakSpace ? 1 : 0)
    if (lines.length > 4_096) return { code: 'CELL_WRAP_TEXT', message: 'wrapped cell text exceeds the bounded line count' }
  }
  if (lines.length === 0) return { code: 'CELL_WRAP_TEXT', message: 'wrap produced no display lines' }
  return lines
}

function isTrailSurrogate(unit: number): boolean {
  return unit >= 0xdc00 && unit <= 0xdfff
}

function isLeadSurrogate(unit: number): boolean {
  return unit >= 0xd800 && unit <= 0xdbff
}

function paintLineGlyphs(
  cell: NativeRenderCellV2,
  run: QualifiedRun,
  rect: NativeSheetGeometryRectV2,
  segment: ShapedSegment,
  originX: number,
  baselineY: number,
  fontSizeMilliPoints: number,
  glyphIndexStart: number,
  scaleOriginX: number,
  scaleNum: number,
  scaleDen: number,
): NativeSheetPaintedGlyphV2[] | { code: NativeSheetCellPaintUnsupportedCodeV2; message: string } {
  const glyphs: NativeSheetPaintedGlyphV2[] = []
  let pen = 0
  for (const glyph of segment.glyphs) {
    if (glyph.glyphId === 0) return { code: 'CELL_MISSING_GLYPH', message: 'font produced the .notdef glyph' }
    const outline = glyphOutline(run.font, glyph.glyphId)
    if (outline.status === 'composite') return { code: 'CELL_COMPOSITE_OUTLINE', message: 'composite TrueType outlines are outside this cell-paint slice' }
    if (outline.status === 'invalid') return { code: 'CELL_MISSING_GLYPH', message: 'glyph outline could not be parsed from the producer-issued font bytes' }
    if (outline.status === 'empty') {
      pen += glyph.advanceXMilliPoints
      continue
    }
    const unscaledX = originX + millipointsToEmu(pen + glyph.offsetXMilliPoints)
    const origin_x_emu = scaleOriginX + scaleEmu(unscaledX - scaleOriginX, scaleNum, scaleDen)
    const origin_y_emu = baselineY - millipointsToEmu(glyph.offsetYMilliPoints)
    if (!Number.isSafeInteger(origin_x_emu) || !Number.isSafeInteger(origin_y_emu) || origin_x_emu < 0 || origin_y_emu < 0) return { code: 'CELL_OVERFLOW', message: 'glyph origin exceeds non-negative integer EMU bounds' }
    const unscaledPath = placePath(outline.path, unscaledX, origin_y_emu, fontSizeMilliPoints, outline.unitsPerEm)
    const path = unscaledPath ? scalePathX(unscaledPath, scaleOriginX, scaleNum, scaleDen) : undefined
    if (!path || path.length > NATIVE_SHEET_CELL_PAINT_V2_LIMITS.maxPathCommandsPerGlyph) return { code: 'CELL_MISSING_GLYPH', message: 'scaled glyph path exceeds bounded integer EMU coordinates' }
    if (!pathContained(path, rect)) return { code: 'CELL_OVERFLOW', message: 'placed glyph ink escapes the cell rectangle' }
    glyphs.push({
      cell_ref: cell.ref,
      row: cell.row,
      column: cell.column,
      glyph_index: glyphIndexStart + glyphs.length,
      face_id: run.font.resource.face.faceId,
      font_sha256: run.font.digest,
      glyph_id: glyph.glyphId,
      origin_x_emu,
      origin_y_emu,
      font_size_millipoints: fontSizeMilliPoints,
      fill_rgb: run.fillRgb,
      fill_rule: 'nonzero',
      outline_kind: 'path',
      path,
    })
    pen += glyph.advanceXMilliPoints
  }
  return glyphs
}

function qualifyDisplay(content: NativeRenderCellContentV2, numberFormat: string | undefined, date1904: boolean | undefined): QualifiedDisplay | { code: NativeSheetCellPaintUnsupportedCodeV2; message: string } {
  if (content.kind === 'blank') return { code: 'CELL_STYLE_UNSUPPORTED', message: 'blank cells have no display text' }
  const value = content.kind === 'formula' ? content.formula.cached : content.value
  if (content.kind === 'formula' && value === undefined) return { code: 'CELL_FORMULA_NO_CACHE', message: 'formula cells without a modeled cached lexical are not evaluated' }
  if (!value) return { code: 'CELL_STYLE_UNSUPPORTED', message: 'cell has no modeled display value' }
  if (value.rich) {
    if (!value.runs || value.runs.length < 1 || typeof value.text !== 'string' || value.runs.map((run) => run.text).join('') !== value.text) {
      return { code: 'CELL_RICH_TEXT', message: 'rich text runs remain source-authoritative' }
    }
    const painted = qualifyPaintedDisplay(value.text, 'string-literal', false)
    if ('code' in painted) return painted
    return { ...painted, runs: value.runs }
  }
  const cachedKind = content.kind === 'formula' ? 'formula-cached-lexical' as const : undefined
  if (value.kind === 'boolean') {
    if (value.lexical !== '0' && value.lexical !== '1' && value.lexical !== 'false' && value.lexical !== 'true') {
      return { code: 'CELL_BOOLEAN_DISPLAY', message: 'boolean lexical is not a producer-issued 0/1/true/false token' }
    }
    return { text: value.lexical === '1' || value.lexical === 'true' ? 'TRUE' : 'FALSE', kind: cachedKind ?? 'boolean-display', numeric: false }
  }
  if (value.kind === 'error') {
    if (typeof value.lexical !== 'string' || !PRINTABLE_ASCII.test(value.lexical)) {
      return { code: typeof value.lexical === 'string' ? 'CELL_NON_ASCII' : 'CELL_ERROR_DISPLAY', message: 'error lexical is missing or not printable ASCII' }
    }
    return qualifyPaintedDisplay(value.lexical, cachedKind ?? 'error-lexical', false)
  }
  if (value.kind === 'string') {
    if (typeof value.text !== 'string' || value.text.length < 1) return { code: 'CELL_STYLE_UNSUPPORTED', message: 'string display text is missing' }
    return qualifyPaintedDisplay(value.text, cachedKind ?? 'string-literal', false)
  }
  if (value.kind === 'number') {
    if (typeof value.lexical !== 'string' || value.lexical.length < 1) return { code: 'CELL_STYLE_UNSUPPORTED', message: 'numeric lexical is missing' }
    const formatted = formatNativeSheetCellDisplayV2('number', value.lexical, numberFormat, date1904)
    if (formatted.status === 'refused') return { code: formatted.code, message: formatted.message }
    const dateFormat = numberFormat !== undefined && numberFormat !== 'General' && numberFormat !== '0' && numberFormat !== '0.00' && formatUsesDateTimeTokens(numberFormat)
    return qualifyPaintedDisplay(formatted.text, cachedKind ?? (dateFormat ? 'date-display' : 'numeric-lexical'), true)
  }
  if (value.kind === 'date') {
    if (typeof value.lexical !== 'string') return { code: 'CELL_DATE_DISPLAY', message: 'date display would require locale/number-format invention' }
    const formatted = formatNativeSheetCellDisplayV2('date', value.lexical, numberFormat, date1904)
    if (formatted.status === 'refused') return { code: formatted.code, message: formatted.message }
    return qualifyPaintedDisplay(formatted.text, cachedKind ?? 'date-display', true)
  }
  return { code: 'CELL_STYLE_UNSUPPORTED', message: 'cell value kind is outside the qualified display subset' }
}

function qualifyPaintedDisplay(text: string, kind: NativeSheetCellPaintDisplayKindV2, numeric: boolean): QualifiedDisplay | { code: NativeSheetCellPaintUnsupportedCodeV2; message: string } {
  const cluster = qualifyDisplayClusterPolicy(text)
  if (cluster) return cluster
  return { text, kind, numeric }
}

/** Latin/Common/Inherited plus HarfBuzz clusters; unpaired UTF-16 and other scripts stay refused. */
function qualifyDisplayClusterPolicy(text: string): { code: NativeSheetCellPaintUnsupportedCodeV2; message: string } | undefined {
  if (text.length < 1 || text.length > 32_767) return { code: 'CELL_STYLE_UNSUPPORTED', message: 'display text is missing or exceeds the Excel 32767 bound' }
  for (let index = 0; index < text.length;) {
    const lead = text.charCodeAt(index)!
    let codePoint = lead, width = 1
    if (lead >= 0xd800 && lead <= 0xdbff) {
      const trail = text.charCodeAt(index + 1)
      if (trail === undefined || trail < 0xdc00 || trail > 0xdfff) return { code: 'CELL_COMPLEX_CLUSTER', message: 'unpaired UTF-16 surrogates require unmodeled cluster policy' }
      codePoint = 0x10000 + ((lead - 0xd800) << 10) + (trail - 0xdc00)
      width = 2
    } else if (lead >= 0xdc00 && lead <= 0xdfff) {
      return { code: 'CELL_COMPLEX_CLUSTER', message: 'unpaired UTF-16 surrogates require unmodeled cluster policy' }
    }
    if (isQualifiedClusterIgnorable(codePoint) || isCombiningMark(codePoint) || unicode13Script(codePoint) === 'Zinh') {
      index += width
      continue
    }
    if (isUnicode13DefaultIgnorable(codePoint) || isUnicode13Control(codePoint)) {
      return { code: 'CELL_COMPLEX_CLUSTER', message: 'default-ignorable and control scalars require unmodeled cluster policy' }
    }
    const script = unicode13Script(codePoint)
    if (script !== 'Latn' && script !== 'Zyyy' && script !== 'Other') {
      return { code: 'CELL_COMPLEX_CLUSTER', message: 'non-Latin script runs require unmodeled itemization and cluster policy' }
    }
    index += width
  }
}

function isQualifiedClusterIgnorable(codePoint: number): boolean {
  return codePoint === 0x200d || (codePoint >= 0xfe00 && codePoint <= 0xfe0f) || (codePoint >= 0xe0100 && codePoint <= 0xe01ef)
}

function isCombiningMark(codePoint: number): boolean {
  return (codePoint >= 0x300 && codePoint <= 0x36f)
    || (codePoint >= 0x1ab0 && codePoint <= 0x1aff)
    || (codePoint >= 0x1dc0 && codePoint <= 0x1dff)
    || (codePoint >= 0x20d0 && codePoint <= 0x20ff)
    || (codePoint >= 0xfe20 && codePoint <= 0xfe2f)
}

function paintScript(text: string): 'Latn' | 'Zyyy' {
  for (const character of text) {
    if (unicode13Script(character.codePointAt(0)!) === 'Latn') return 'Latn'
  }
  return 'Zyyy'
}

/**
 * Locale-independent display for documented OOXML date/number tokens.
 * Uses producer ISO calendar components as written; never Date, Intl, or host TZ.
 * Number cells with `General` keep their stored lexical. Date serial conversion
 * uses only a projected workbook date1904 flag. Named months/days require an
 * explicit OOXML `[$-…]` locale/calendar that this producer has a table for.
 */
export function formatNativeSheetCellDisplayV2(
  kind: 'number' | 'date',
  lexical: string,
  numberFormat: string | undefined,
  date1904?: boolean,
): NativeSheetCellDisplayFormatResultV2 {
  const format = numberFormat ?? 'General'
  if (kind === 'number') {
    if (format === 'General') return { status: 'ready', text: lexical }
    if (format === '0' || format === '0.00') {
      const text = formatFixedDecimalLexical(lexical, format === '0' ? 0 : 2)
      if (text === undefined) return { status: 'refused', code: 'CELL_STYLE_UNSUPPORTED', message: 'bounded numeric format cannot be applied to this lexical' }
      return { status: 'ready', text }
    }
  }
  const classified = classifyOoxmlDateFormat(format)
  if (classified.status === 'ready') {
    let parts: DateTimeParts | undefined
    if (kind === 'date') parts = parseIsoDateTimeParts(lexical)
    else {
      if (classified.needsDate && date1904 === undefined) return { status: 'refused', code: 'CELL_DATE_DISPLAY', message: 'date display would require locale/number-format invention' }
      parts = excelSerialToParts(lexical, date1904 === true, classified.needsDate)
    }
    const text = parts ? applyOoxmlDateTokens(classified.tokens, parts, classified.locale) : undefined
    if (text === undefined) return { status: 'refused', code: 'CELL_DATE_DISPLAY', message: 'date display would require locale/number-format invention' }
    return { status: 'ready', text }
  }
  if (classified.status === 'unusable' || kind === 'date') return { status: 'refused', code: 'CELL_DATE_DISPLAY', message: 'date display would require locale/number-format invention' }
  return { status: 'ready', text: lexical }
}

const EN_US_GREGORIAN_MONTHS_FULL = Object.freeze(['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'])
const EN_US_GREGORIAN_MONTHS_ABBR = Object.freeze(['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'])
const EN_US_GREGORIAN_MONTHS_LETTER = Object.freeze(['J', 'F', 'M', 'A', 'M', 'J', 'J', 'A', 'S', 'O', 'N', 'D'])
const EN_US_GREGORIAN_DAYS_FULL = Object.freeze(['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'])
const EN_US_GREGORIAN_DAYS_ABBR = Object.freeze(['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'])

type DateTimeParts = {
  year: number
  month: number
  day: number
  hour: number
  minute: number
  second: number
  weekday: number
  dateValid: boolean
}

type DateLocaleTable = {
  monthsFull: readonly string[]
  monthsAbbr: readonly string[]
  monthsLetter: readonly string[]
  daysFull: readonly string[]
  daysAbbr: readonly string[]
}

const EN_US_GREGORIAN: DateLocaleTable = {
  monthsFull: EN_US_GREGORIAN_MONTHS_FULL,
  monthsAbbr: EN_US_GREGORIAN_MONTHS_ABBR,
  monthsLetter: EN_US_GREGORIAN_MONTHS_LETTER,
  daysFull: EN_US_GREGORIAN_DAYS_FULL,
  daysAbbr: EN_US_GREGORIAN_DAYS_ABBR,
}

type DateFormatToken =
  | { kind: 'year'; width: number }
  | { kind: 'month'; width: number }
  | { kind: 'monthName'; width: number }
  | { kind: 'day'; width: number }
  | { kind: 'weekday'; width: number }
  | { kind: 'hour'; width: number }
  | { kind: 'minute'; width: number }
  | { kind: 'second'; width: number }
  | { kind: 'ampm'; form: 'AM/PM' | 'A/P' }
  | { kind: 'literal'; text: string }

function formatUsesDateTimeTokens(format: string): boolean {
  return classifyOoxmlDateFormat(format).status !== 'not-date'
}

function classifyOoxmlDateFormat(format: string):
  | { status: 'ready'; tokens: DateFormatToken[]; locale: DateLocaleTable | undefined; needsDate: boolean }
  | { status: 'unusable' }
  | { status: 'not-date' } {
  let body = format
  let locale: DateLocaleTable | undefined
  if (format.startsWith('[$-')) {
    const close = format.indexOf(']')
    if (close < 0) return { status: 'unusable' }
    const inner = format.slice(3, close)
    if (isSystemDateLocale(inner)) return { status: 'unusable' }
    locale = localeTableForOoxmlPrefix(inner)
    body = format.slice(close + 1)
  }
  if (body.includes(';') || body.includes('*') || body.includes('_') || body.includes('[')) {
    return /[ymdhs]/i.test(body) || format.startsWith('[$-') ? { status: 'unusable' } : { status: 'not-date' }
  }
  const tokens = tokenizeOoxmlDateFormat(body)
  if (!tokens) return /[ymdhs]/i.test(body) || format.startsWith('[$-') ? { status: 'unusable' } : { status: 'not-date' }
  const needsNames = tokens.some((token) => token.kind === 'monthName' || token.kind === 'weekday')
  if (needsNames && locale === undefined) return { status: 'unusable' }
  const needsDate = tokens.some((token) => token.kind === 'year' || token.kind === 'month' || token.kind === 'monthName' || token.kind === 'day' || token.kind === 'weekday')
  const needsTime = tokens.some((token) => token.kind === 'hour' || token.kind === 'minute' || token.kind === 'second' || token.kind === 'ampm')
  if (!needsDate && !needsTime) return { status: 'not-date' }
  return { status: 'ready', tokens, locale, needsDate }
}

function isSystemDateLocale(inner: string): boolean {
  return /f800|f400|sysdate|systime/i.test(inner)
}

function localeTableForOoxmlPrefix(inner: string): DateLocaleTable | undefined {
  const head = inner.split(',')[0]!.trim()
  if (/^en[-_]US$/i.test(head)) return EN_US_GREGORIAN
  if (!/^[0-9A-Fa-f]+$/.test(head)) return undefined
  const value = Number.parseInt(head, 16)
  if (!Number.isSafeInteger(value) || value < 0) return undefined
  const lcid = value & 0xffff
  const calendar = (value >>> 16) & 0xff
  if (lcid === 0xf800 || lcid === 0xf400) return undefined
  if (calendar !== 0 && calendar !== 1) return undefined
  if (calendar === 1 || lcid === 0x0409 || lcid === 0x0009) return EN_US_GREGORIAN
  return undefined
}

function tokenizeOoxmlDateFormat(format: string): DateFormatToken[] | undefined {
  const tokens: DateFormatToken[] = []
  let index = 0
  while (index < format.length) {
    const char = format[index]!
    if (char === '[') return undefined
    if (char === '"') {
      const close = format.indexOf('"', index + 1)
      if (close < 0) return undefined
      tokens.push({ kind: 'literal', text: format.slice(index + 1, close) })
      index = close + 1
      continue
    }
    if (char === '\\') {
      if (index + 1 >= format.length) return undefined
      tokens.push({ kind: 'literal', text: format[index + 1]! })
      index += 2
      continue
    }
    const rest = format.slice(index)
    if (/^AM\/PM/i.test(rest) || /^A\/P/i.test(rest)) {
      const ampm = /^AM\/PM/i.test(rest)
      tokens.push({ kind: 'ampm', form: ampm ? 'AM/PM' : 'A/P' })
      index += ampm ? 5 : 3
      continue
    }
    if (char === 'y' || char === 'Y') {
      const width = takeSame(format, index, 'y')
      tokens.push({ kind: 'year', width })
      index += width
      continue
    }
    if (char === 'm' || char === 'M') {
      const width = takeSame(format, index, 'm')
      if (width >= 3) tokens.push({ kind: 'monthName', width })
      else tokens.push({ kind: 'month', width })
      index += width
      continue
    }
    if (char === 'd' || char === 'D') {
      const width = takeSame(format, index, 'd')
      if (width >= 3) tokens.push({ kind: 'weekday', width })
      else tokens.push({ kind: 'day', width })
      index += width
      continue
    }
    if (char === 'h' || char === 'H') {
      const width = takeSame(format, index, 'h')
      tokens.push({ kind: 'hour', width })
      index += width
      continue
    }
    if (char === 's' || char === 'S') {
      const width = takeSame(format, index, 's')
      tokens.push({ kind: 'second', width })
      index += width
      continue
    }
    if (/[A-Za-z]/.test(char)) return undefined
    tokens.push({ kind: 'literal', text: char })
    index += 1
  }
  return resolveMinuteTokens(tokens)
}

function takeSame(format: string, start: number, letter: string): number {
  let width = 0
  while (start + width < format.length && format[start + width]!.toLowerCase() === letter) width += 1
  return width
}

function resolveMinuteTokens(tokens: DateFormatToken[]): DateFormatToken[] {
  const significant = (token: DateFormatToken): boolean => token.kind !== 'literal'
  return tokens.map((token, index) => {
    if (token.kind !== 'month' || token.width > 2) return token
    const prev = tokens.slice(0, index).reverse().find(significant)
    const next = tokens.slice(index + 1).find(significant)
    if (prev?.kind === 'hour' || next?.kind === 'second') return { kind: 'minute', width: token.width }
    return token
  })
}

function applyOoxmlDateTokens(tokens: DateFormatToken[], parts: DateTimeParts, locale: DateLocaleTable | undefined): string | undefined {
  const hour12 = tokens.some((token) => token.kind === 'ampm')
  let text = ''
  for (const token of tokens) {
    switch (token.kind) {
      case 'year':
        text += token.width <= 2 ? String(parts.year % 100).padStart(2, '0') : String(parts.year).padStart(4, '0')
        break
      case 'month':
        text += padDateNumber(parts.month, token.width)
        break
      case 'monthName': {
        if (!locale) return undefined
        const name = token.width >= 5 ? locale.monthsLetter[parts.month - 1] : token.width === 3 ? locale.monthsAbbr[parts.month - 1] : locale.monthsFull[parts.month - 1]
        if (!name) return undefined
        text += name
        break
      }
      case 'day':
        text += padDateNumber(parts.day, token.width)
        break
      case 'weekday': {
        if (!locale) return undefined
        const name = token.width === 3 ? locale.daysAbbr[parts.weekday] : locale.daysFull[parts.weekday]
        if (!name) return undefined
        text += name
        break
      }
      case 'hour': {
        const hour = hour12 ? ((parts.hour % 12) === 0 ? 12 : parts.hour % 12) : parts.hour
        text += padDateNumber(hour, token.width)
        break
      }
      case 'minute':
        text += padDateNumber(parts.minute, token.width)
        break
      case 'second':
        text += padDateNumber(parts.second, token.width)
        break
      case 'ampm': {
        const pm = parts.hour >= 12
        text += token.form === 'AM/PM' ? (pm ? 'PM' : 'AM') : (pm ? 'P' : 'A')
        break
      }
      case 'literal':
        text += token.text
        break
    }
  }
  if (text.length < 1 || text.length > 32_767 || !isCanonicalDisplayText(text)) return undefined
  return text
}

function padDateNumber(value: number, width: number): string {
  const text = String(value)
  return width >= 2 ? text.padStart(2, '0') : text
}

function parseIsoDateTimeParts(lexical: string): DateTimeParts | undefined {
  const match = /^(\d{4})-(\d{2})-(\d{2})(?:(?:T(\d{2}):(\d{2}):(\d{2})(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})?)|(?:Z|[+-]\d{2}:\d{2}))?$/.exec(lexical)
  if (!match) return undefined
  const year = Number(match[1]), month = Number(match[2]), day = Number(match[3])
  const hour = match[4] === undefined ? 0 : Number(match[4])
  const minute = match[5] === undefined ? 0 : Number(match[5])
  const second = match[6] === undefined ? 0 : Number(match[6])
  if (month < 1 || month > 12 || day < 1 || day > daysInMonth(year, month)) return undefined
  if (hour > 23 || minute > 59 || second > 59) return undefined
  const offset = /([+-])(\d{2}):(\d{2})$/.exec(lexical)
  if (offset && (Number(offset[2]) > 24 || Number(offset[3]) > 59 || (Number(offset[2]) === 24 && Number(offset[3]) !== 0))) return undefined
  const weekday = gregorianWeekday(year, month, day)
  if (weekday === undefined) return undefined
  return { year, month, day, hour, minute, second, weekday, dateValid: true }
}

function excelSerialToParts(lexical: string, date1904: boolean, needsDate: boolean): DateTimeParts | undefined {
  const parsed = parseDecimalLexical(lexical)
  if (!parsed || parsed.negative) return undefined
  const divisor = 10n ** BigInt(parsed.scale)
  const whole = parsed.scale === 0 ? parsed.coefficient : parsed.coefficient / divisor
  if (whole > BigInt(Number.MAX_SAFE_INTEGER)) return undefined
  const serial = Number(whole)
  if (!Number.isSafeInteger(serial) || serial < 0) return undefined
  const frac = parsed.scale === 0 ? 0n : parsed.coefficient % divisor
  const roundedSeconds = Number((frac * 86400n + divisor / 2n) / divisor)
  if (!Number.isSafeInteger(roundedSeconds) || roundedSeconds < 0) return undefined
  let extraDays = Math.floor(roundedSeconds / 86400)
  let seconds = roundedSeconds % 86400
  const ymd = excelSerialToYmd(serial + extraDays, date1904)
  const hour = Math.floor(seconds / 3600)
  const minute = Math.floor((seconds % 3600) / 60)
  const second = seconds % 60
  if (!ymd) {
    if (needsDate) return undefined
    return { year: 1899, month: 12, day: 31, hour, minute, second, weekday: 0, dateValid: false }
  }
  const weekday = date1904 ? (serial + extraDays + 5) % 7 : (serial + extraDays - 1) % 7
  if (weekday < 0) return undefined
  return { year: ymd.year, month: ymd.month, day: ymd.day, hour, minute, second, weekday, dateValid: true }
}

function gregorianWeekday(year: number, month: number, day: number): number | undefined {
  if (!Number.isSafeInteger(year) || !Number.isSafeInteger(month) || !Number.isSafeInteger(day)) return undefined
  const offsets = [0, 3, 2, 5, 0, 3, 5, 1, 4, 6, 2, 4]
  const y = month < 3 ? year - 1 : year
  return (y + Math.floor(y / 4) - Math.floor(y / 100) + Math.floor(y / 400) + offsets[month - 1]! + day) % 7
}

function excelSerialToYmd(serial: number, date1904: boolean): { year: number; month: number; day: number } | undefined {
  if (!Number.isSafeInteger(serial) || serial < 0) return undefined
  if (date1904) return addDaysToYmd(1904, 1, 1, serial)
  if (serial < 1 || serial === 60) return undefined
  return addDaysToYmd(1899, 12, 31, serial > 60 ? serial - 1 : serial)
}

function addDaysToYmd(year: number, month: number, day: number, days: number): { year: number; month: number; day: number } | undefined {
  let y = year, m = month, d = day, remaining = days
  if (!Number.isSafeInteger(remaining) || remaining < 0) return undefined
  while (remaining > 0) {
    const left = daysInMonth(y, m) - d + 1
    if (remaining < left) { d += remaining; remaining = 0; break }
    remaining -= left
    d = 1
    m += 1
    if (m > 12) { m = 1; y += 1 }
    if (y > 9999) return undefined
  }
  if (y < 1900 || y > 9999) return undefined
  return { year: y, month: m, day: d }
}

function daysInMonth(year: number, month: number): number {
  return [31, year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1]!
}

function formatFixedDecimalLexical(lexical: string, fractionDigits: number): string | undefined {
  const parsed = parseDecimalLexical(lexical)
  if (!parsed) return undefined
  let { coefficient, scale } = parsed
  if (scale < fractionDigits) {
    const shift = fractionDigits - scale
    if (shift > 32) return undefined
    coefficient *= 10n ** BigInt(shift)
    scale = fractionDigits
  } else if (scale > fractionDigits) {
    const shift = scale - fractionDigits
    if (shift > 32) return undefined
    const divisor = 10n ** BigInt(shift)
    const remainder = coefficient % divisor
    coefficient = coefficient / divisor
    if (remainder * 2n >= divisor) coefficient += 1n
    scale = fractionDigits
  }
  let digits = coefficient.toString()
  if (fractionDigits > 0 && digits.length <= fractionDigits) digits = digits.padStart(fractionDigits + 1, '0')
  const sign = parsed.negative && coefficient !== 0n ? '-' : ''
  const text = fractionDigits === 0 ? `${sign}${digits}` : `${sign}${digits.slice(0, digits.length - fractionDigits)}.${digits.slice(digits.length - fractionDigits)}`
  if (text.length < 1 || text.length > 32 || !PRINTABLE_ASCII.test(text)) return undefined
  return text
}

function parseDecimalLexical(lexical: string): { negative: boolean; coefficient: bigint; scale: number } | undefined {
  const match = /^([+-])?(?:([0-9]+)(?:\.([0-9]*))?|\.([0-9]+))(?:[eE]([+-]?[0-9]+))?$/.exec(lexical)
  if (!match) return undefined
  const exponent = match[5] === undefined ? 0 : Number(match[5])
  if (!Number.isSafeInteger(exponent) || exponent > 32 || exponent < -32) return undefined
  const intDigits = match[2] ?? ''
  const fracDigits = match[3] ?? match[4] ?? ''
  const raw = `${intDigits}${fracDigits}`.replace(/^0+/, '') || '0'
  if (raw.length > 32) return undefined
  let scale = fracDigits.length - exponent
  let coefficient = BigInt(raw)
  if (scale < 0) {
    const shift = -scale
    if (shift > 32 || raw.length + shift > 32) return undefined
    coefficient *= 10n ** BigInt(shift)
    scale = 0
  }
  return { negative: match[1] === '-', coefficient, scale }
}

function qualifyStyle(context: CellPaintContext, styleID: number): QualifiedStyle | { code: NativeSheetCellPaintUnsupportedCodeV2; message: string } {
  const workbook = context.workbook
  const geometry = context.geometry
  const style = workbook.styles[styleID]
  if (!style || style.id !== styleID || style.provenance.style_id !== styleID || style.provenance.source_revision !== workbook.revision || style.provenance.source_package_sha256 !== workbook.source.package_sha256 || style.provenance.projection !== style.effective.projection || style.provenance.raw_projection_sha256 !== nativeWorkbookStyleRawProjectionSha256V2(style.effective)) {
    return { code: 'CELL_STYLE_UNSUPPORTED', message: 'effective style or provenance is incomplete for exact cell paint' }
  }
  const effective = style.effective
  if (effective.unsupported.includes('alignment-extended')) return { code: 'CELL_ALIGNMENT_EXTENDED', message: 'indent, rotation, or shrink-to-fit would require unmodeled alignment heuristics' }
  if (effective.unsupported.includes('horizontal-alignment') || effective.unsupported.includes('vertical-alignment')) return { code: 'CELL_STYLE_UNSUPPORTED', message: 'alignment is not projected exactly' }
  const sizePoints = effective.font_size_points ?? geometry.metric_authority.font_size_points
  const fontSizeMilliPoints = pointsToMilliPoints(sizePoints)
  if (fontSizeMilliPoints === undefined) return { code: 'CELL_FONT_MISMATCH', message: 'font size is not an integer millipoint quantity' }
  const font = resolvePaintFont(context, effective.font_name, effective.bold, effective.italic)
  if ('code' in font) return font
  if (effective.unsupported.includes('font-color') || effective.font_color === undefined || !RGB.test(effective.font_color)) return { code: 'CELL_FONT_COLOR', message: 'cell font color is missing, theme-based, or not direct RGB' }
  const horizontal = effective.horizontal_alignment ?? 'general'
  const vertical = effective.vertical_alignment ?? 'bottom'
  if (horizontal !== 'general' && horizontal !== 'left' && horizontal !== 'center' && horizontal !== 'right') return { code: 'CELL_STYLE_UNSUPPORTED', message: 'horizontal alignment is outside the qualified set' }
  if (vertical !== 'top' && vertical !== 'middle' && vertical !== 'bottom') return { code: 'CELL_STYLE_UNSUPPORTED', message: 'vertical alignment is outside the qualified set' }
  const rotation = effective.text_rotation ?? 0
  if (rotation !== 0 && rotation !== 90 && rotation !== 180 && rotation !== 270) return { code: 'CELL_ALIGNMENT_EXTENDED', message: 'indent, rotation, or shrink-to-fit would require unmodeled alignment heuristics' }
  return {
    fontSizeMilliPoints, fillRgb: effective.font_color, horizontal, vertical, wrap: effective.wrap_text === true,
    shrink: effective.shrink_to_fit === true, rotation, font,
    fontName: effective.font_name ?? geometry.metric_authority.font_name,
    bold: effective.bold ?? geometry.metric_authority.font_bold,
    italic: effective.italic ?? geometry.metric_authority.font_italic,
  }
}

function resolvePaintFont(context: CellPaintContext, fontName: string | undefined, bold: boolean | undefined, italic: boolean | undefined): PaintFontFace | { code: NativeSheetCellPaintUnsupportedCodeV2; message: string } {
  const authority = context.geometry.metric_authority
  const wantBold = bold ?? authority.font_bold
  const wantItalic = italic ?? authority.font_italic
  const wantName = fontName ?? authority.font_name
  if (wantName === authority.font_name && wantBold === authority.font_bold && wantItalic === authority.font_italic) return context.fonts[0]!
  const normalized = normalizePaintFontName(wantName)
  const match = context.fonts.find((font, index) => {
    if (index === 0 || !font.names.has(normalized)) return false
    const bold = font.resource.face.weight >= 700
    const italic = font.resource.face.style !== 'normal'
    return bold === wantBold && italic === wantItalic
  })
  if (!match) return { code: 'CELL_FONT_MISMATCH', message: 'cell font name does not match the producer-issued Normal-font bytes' }
  return match
}

function qualifyRuns(context: CellPaintContext, style: QualifiedStyle, display: QualifiedDisplay): QualifiedRun[] | { code: NativeSheetCellPaintUnsupportedCodeV2; message: string } {
  if (!display.runs) return [{ text: display.text, fontSizeMilliPoints: style.fontSizeMilliPoints, fillRgb: style.fillRgb, font: style.font }]
  const runs: QualifiedRun[] = []
  for (const run of display.runs) {
    if (run.text.length < 1) return { code: 'CELL_RICH_TEXT', message: 'rich text runs remain source-authoritative' }
    const font = resolvePaintFont(context, run.font_name ?? style.fontName, run.bold ?? style.bold, run.italic ?? style.italic)
    if ('code' in font) return font
    const size = run.font_size_points === undefined ? style.fontSizeMilliPoints : pointsToMilliPoints(run.font_size_points)
    if (size === undefined) return { code: 'CELL_FONT_MISMATCH', message: 'font size is not an integer millipoint quantity' }
    const fillRgb = run.font_color ?? style.fillRgb
    if (!RGB.test(fillRgb)) return { code: 'CELL_FONT_COLOR', message: 'cell font color is missing, theme-based, or not direct RGB' }
    runs.push({ text: run.text, fontSizeMilliPoints: size, fillRgb, font })
  }
  return runs
}

function resolvedAlignment(horizontal: 'general' | 'left' | 'center' | 'right', numeric: boolean): 'left' | 'center' | 'right' {
  if (horizontal === 'general') return numeric ? 'right' : 'left'
  return horizontal
}

function insetContentBox(rect: NativeSheetGeometryRectV2): NativeSheetGeometryRectV2 | undefined {
  const gutter = NATIVE_SHEET_CELL_GUTTER_EMU
  if (rect.width_emu <= gutter * 2 || rect.height_emu < 1) return undefined
  return { x_emu: rect.x_emu + gutter, y_emu: rect.y_emu, width_emu: rect.width_emu - gutter * 2, height_emu: rect.height_emu }
}

function glyphOutline(font: PaintFontFace, glyphId: number): DesignOutline {
  const cached = font.outlines.get(glyphId)
  if (cached) return cached
  const parsed = parseGlyfOutline(font.bytes, glyphId)
  font.outlines.set(glyphId, parsed)
  return parsed
}

function scaleEmu(value: number, scaleNum: number, scaleDen: number): number {
  if (scaleNum === 1 && scaleDen === 1) return value
  const product = value * scaleNum
  if (!Number.isSafeInteger(product) || scaleDen < 1) throw new NativeSheetCellPaintError('paint.sourceUnsupported', '$.coordinates', 'EMU scaling exceeds deterministic integer precision')
  return Math.floor(product / scaleDen)
}

function scalePathX(path: NativeSheetCellPaintPathCommandV2[], originX: number, scaleNum: number, scaleDen: number): NativeSheetCellPaintPathCommandV2[] | undefined {
  if (scaleNum === 1 && scaleDen === 1) return path
  const output: NativeSheetCellPaintPathCommandV2[] = []
  for (const command of path) {
    if (command.kind === 'close_path') { output.push(command); continue }
    const x = originX + scaleEmu(command.x_emu - originX, scaleNum, scaleDen)
    if (!Number.isSafeInteger(x) || x < 0) return undefined
    if (command.kind === 'move_to' || command.kind === 'line_to') output.push({ kind: command.kind, x_emu: x, y_emu: command.y_emu })
    else {
      const controlX = originX + scaleEmu(command.control_x_emu - originX, scaleNum, scaleDen)
      if (!Number.isSafeInteger(controlX) || controlX < 0) return undefined
      output.push({ kind: 'quadratic_to', control_x_emu: controlX, control_y_emu: command.control_y_emu, x_emu: x, y_emu: command.y_emu })
    }
  }
  return output
}

function rotateGlyphs(glyphs: NativeSheetPaintedGlyphV2[], rect: NativeSheetGeometryRectV2, rotation: 90 | 180 | 270): NativeSheetPaintedGlyphV2[] | { code: NativeSheetCellPaintUnsupportedCodeV2; message: string } {
  const cx = rect.x_emu + Math.floor(rect.width_emu / 2)
  const cy = rect.y_emu + Math.floor(rect.height_emu / 2)
  const rotated: NativeSheetPaintedGlyphV2[] = []
  for (const glyph of glyphs) {
    const origin = rotatePoint(glyph.origin_x_emu, glyph.origin_y_emu, cx, cy, rotation)
    if (!origin || origin.x < 0 || origin.y < 0) return { code: 'CELL_OVERFLOW', message: 'rotated glyph origin exceeds non-negative integer EMU bounds' }
    const path: NativeSheetCellPaintPathCommandV2[] = []
    for (const command of glyph.path) {
      if (command.kind === 'close_path') { path.push(command); continue }
      const point = rotatePoint(command.x_emu, command.y_emu, cx, cy, rotation)
      if (!point || point.x < 0 || point.y < 0) return { code: 'CELL_OVERFLOW', message: 'rotated glyph path exceeds non-negative integer EMU bounds' }
      if (command.kind === 'move_to' || command.kind === 'line_to') path.push({ kind: command.kind, x_emu: point.x, y_emu: point.y })
      else {
        const control = rotatePoint(command.control_x_emu, command.control_y_emu, cx, cy, rotation)
        if (!control || control.x < 0 || control.y < 0) return { code: 'CELL_OVERFLOW', message: 'rotated glyph path exceeds non-negative integer EMU bounds' }
        path.push({ kind: 'quadratic_to', control_x_emu: control.x, control_y_emu: control.y, x_emu: point.x, y_emu: point.y })
      }
    }
    if (!pathContained(path, rect)) return { code: 'CELL_OVERFLOW', message: 'placed glyph ink escapes the cell rectangle' }
    rotated.push({ ...glyph, origin_x_emu: origin.x, origin_y_emu: origin.y, path })
  }
  return rotated
}

function rotatePoint(x: number, y: number, cx: number, cy: number, rotation: 90 | 180 | 270): { x: number; y: number } | undefined {
  const dx = x - cx, dy = y - cy
  const point = rotation === 90 ? { x: cx - dy, y: cy + dx } : rotation === 180 ? { x: cx - dx, y: cy - dy } : { x: cx + dy, y: cy - dx }
  return Number.isSafeInteger(point.x) && Number.isSafeInteger(point.y) ? point : undefined
}

function placePath(path: DesignPathCommand[], originX: number, originY: number, fontSizeMilliPoints: number, unitsPerEm: number): NativeSheetCellPaintPathCommandV2[] | undefined {
  const output: NativeSheetCellPaintPathCommandV2[] = []
  for (const command of path) {
    if (command.kind === 'close_path') { output.push(command); continue }
    const x = designToEmu(originX, command.x, unitsPerEm, fontSizeMilliPoints)
    const y = designToEmu(originY, command.y, unitsPerEm, fontSizeMilliPoints, true)
    if (x === undefined || y === undefined) return undefined
    if (command.kind === 'move_to' || command.kind === 'line_to') output.push({ kind: command.kind, x_emu: x, y_emu: y })
    else {
      const controlX = designToEmu(originX, command.control_x, unitsPerEm, fontSizeMilliPoints)
      const controlY = designToEmu(originY, command.control_y, unitsPerEm, fontSizeMilliPoints, true)
      if (controlX === undefined || controlY === undefined) return undefined
      output.push({ kind: 'quadratic_to', control_x_emu: controlX, control_y_emu: controlY, x_emu: x, y_emu: y })
    }
  }
  return pathContained(output, { x_emu: 0, y_emu: 0, width_emu: Number.MAX_SAFE_INTEGER, height_emu: Number.MAX_SAFE_INTEGER }) ? output : undefined
}

function designToEmu(origin: number, design: number, unitsPerEm: number, fontSizeMilliPoints: number, invertY = false): number | undefined {
  let scaled: number
  try { scaled = scaleDesignToMilliPoints(design, unitsPerEm, fontSizeMilliPoints) } catch { return undefined }
  const emu = millipointsToEmu(scaled)
  const result = origin + (invertY ? -emu : emu)
  return Number.isSafeInteger(result) && !Object.is(result, -0) && result >= 0 ? result : undefined
}

/** Integer or half-integer TrueType design units to millipoints. Implied on-curve midpoints are halves. */
function scaleDesignToMilliPoints(design: number, unitsPerEm: number, fontSizeMilliPoints: number): number {
  const twice = design * 2
  if (!Number.isSafeInteger(twice) || !Number.isSafeInteger(unitsPerEm) || unitsPerEm <= 0 || !Number.isSafeInteger(fontSizeMilliPoints) || fontSizeMilliPoints <= 0) {
    throw new RangeError('font metrics and size must be safe integers, and unitsPerEm/size must be positive')
  }
  const product = twice * fontSizeMilliPoints
  if (!Number.isSafeInteger(product)) throw new RangeError('font metric scaling exceeds deterministic integer precision')
  const scaled = product / (2 * unitsPerEm)
  if (!Number.isSafeInteger(Math.round(scaled))) throw new RangeError('scaled font metric exceeds the safe integer range')
  return Math.round(scaled)
}

function millipointsToEmu(millipoints: number): number {
  if (!Number.isSafeInteger(millipoints) || Object.is(millipoints, -0)) throw new NativeSheetCellPaintError('paint.sourceUnsupported', '$.coordinates', 'millipoint quantity is not a canonical integer')
  const product = millipoints * EMU_PER_POINT
  if (!Number.isSafeInteger(product)) throw new NativeSheetCellPaintError('paint.sourceUnsupported', '$.coordinates', 'EMU scaling exceeds deterministic integer precision')
  const rounded = Math.round(product / 1_000)
  if (!Number.isSafeInteger(rounded)) throw new NativeSheetCellPaintError('paint.sourceUnsupported', '$.coordinates', 'EMU scaling exceeds the safe integer range')
  return Object.is(rounded, -0) ? 0 : rounded
}

function pointsToMilliPoints(points: number): number | undefined {
  if (!Number.isFinite(points) || points <= 0) return undefined
  const milli = points * 1_000
  const rounded = Math.round(milli)
  if (!Number.isSafeInteger(rounded) || rounded < 1 || Math.abs(milli - rounded) > 1e-9) return undefined
  return rounded
}

function parseGlyfOutline(bytes: Uint8Array, glyphId: number): DesignOutline {
  return parseGlyfOutlineAt(bytes, glyphId, new Set(), 0, { xx: 16384, xy: 0, yx: 0, yy: 16384, dx: 0, dy: 0 })
}

type GlyfTransform = { xx: number; xy: number; yx: number; yy: number; dx: number; dy: number }

function parseGlyfOutlineAt(bytes: Uint8Array, glyphId: number, stack: Set<number>, depth: number, transform: GlyfTransform): DesignOutline {
  if (depth > 16 || stack.has(glyphId)) return { status: 'composite' }
  const tables = sfntTables(bytes)
  if (!tables) return { status: 'invalid' }
  const head = tables.get('head'), loca = tables.get('loca'), glyf = tables.get('glyf'), maxp = tables.get('maxp')
  if (!head || !loca || !glyf || !maxp || head.length < 54 || maxp.length < 6) return { status: 'invalid' }
  const glyphCount = u16(bytes, maxp.offset + 4)
  if (glyphId < 0 || glyphId >= glyphCount) return { status: 'invalid' }
  const longLoca = i16(bytes, head.offset + 50) === 1
  const locaOffset = (index: number): number | undefined => {
    if (longLoca) {
      if (loca.offset + (index + 1) * 4 > loca.offset + loca.length) return undefined
      return u32(bytes, loca.offset + index * 4)
    }
    if (loca.offset + (index + 1) * 2 > loca.offset + loca.length) return undefined
    return u16(bytes, loca.offset + index * 2) * 2
  }
  const start = locaOffset(glyphId), end = locaOffset(glyphId + 1)
  if (start === undefined || end === undefined || end < start || start > glyf.length || end > glyf.length) return { status: 'invalid' }
  const unitsPerEm = u16(bytes, head.offset + 18)
  if (end === start) return { status: 'empty', unitsPerEm }
  const at = glyf.offset + start
  const limit = glyf.offset + end
  if (!range(bytes, at, 10) || at + 10 > limit) return { status: 'invalid' }
  const contourCount = i16(bytes, at)
  if (contourCount < 0) {
    stack.add(glyphId)
    const flattened = flattenCompositeGlyf(bytes, at + 10, limit, stack, depth, transform, unitsPerEm)
    stack.delete(glyphId)
    return flattened
  }
  if (contourCount === 0) return { status: 'empty', unitsPerEm }
  const endPts = at + 10
  if (!range(bytes, endPts, contourCount * 2) || endPts + contourCount * 2 > limit) return { status: 'invalid' }
  const lastPoint = u16(bytes, endPts + (contourCount - 1) * 2)
  const pointCount = lastPoint + 1
  if (pointCount < 1 || pointCount > 10_000) return { status: 'invalid' }
  for (let index = 1; index < contourCount; index++) if (u16(bytes, endPts + index * 2) <= u16(bytes, endPts + (index - 1) * 2)) return { status: 'invalid' }
  const instructionLength = u16(bytes, endPts + contourCount * 2)
  let cursor = endPts + contourCount * 2 + 2 + instructionLength
  if (instructionLength > 65_536 || cursor > limit) return { status: 'invalid' }
  const flags: number[] = []
  while (flags.length < pointCount) {
    if (!range(bytes, cursor, 1)) return { status: 'invalid' }
    const flag = bytes[cursor++]!
    flags.push(flag)
    if (flag & 0x08) {
      if (!range(bytes, cursor, 1)) return { status: 'invalid' }
      const repeat = bytes[cursor++]!
      if (flags.length + repeat > pointCount) return { status: 'invalid' }
      for (let count = 0; count < repeat; count++) flags.push(flag)
    }
  }
  const xs = new Array<number>(pointCount), ys = new Array<number>(pointCount)
  let x = 0, y = 0
  for (let index = 0; index < pointCount; index++) {
    const flag = flags[index]!
    if (flag & 0x02) {
      if (!range(bytes, cursor, 1)) return { status: 'invalid' }
      const value = bytes[cursor++]!
      x += (flag & 0x10) ? value : -value
    } else if (!(flag & 0x10)) {
      if (!range(bytes, cursor, 2)) return { status: 'invalid' }
      x += i16(bytes, cursor); cursor += 2
    }
    xs[index] = x
  }
  for (let index = 0; index < pointCount; index++) {
    const flag = flags[index]!
    if (flag & 0x04) {
      if (!range(bytes, cursor, 1)) return { status: 'invalid' }
      const value = bytes[cursor++]!
      y += (flag & 0x20) ? value : -value
    } else if (!(flag & 0x20)) {
      if (!range(bytes, cursor, 2)) return { status: 'invalid' }
      y += i16(bytes, cursor); cursor += 2
    }
    ys[index] = y
  }
  if (cursor > limit) return { status: 'invalid' }
  const path: DesignPathCommand[] = []
  let firstPoint = 0
  for (let contour = 0; contour < contourCount; contour++) {
    const last = u16(bytes, endPts + contour * 2)
    const points: Array<{ x: number; y: number; on: boolean }> = []
    for (let index = firstPoint; index <= last; index++) points.push({ x: xs[index]!, y: ys[index]!, on: (flags[index]! & 0x01) !== 0 })
    const contourPath = contourToPath(points)
    if (!contourPath) return { status: 'invalid' }
    path.push(...contourPath)
    firstPoint = last + 1
  }
  if (path.length === 0) return { status: 'empty', unitsPerEm }
  const transformed = transformDesignPath(path, transform)
  if (!transformed) return { status: 'composite' }
  return { status: 'outlined', unitsPerEm, path: transformed }
}

const GLYF_ARG_1_AND_2_ARE_WORDS = 0x0001
const GLYF_ARGS_ARE_XY_VALUES = 0x0002
const GLYF_ROUND_XY_TO_GRID = 0x0004
const GLYF_WE_HAVE_A_SCALE = 0x0008
const GLYF_MORE_COMPONENTS = 0x0020
const GLYF_WE_HAVE_AN_X_AND_Y_SCALE = 0x0040
const GLYF_WE_HAVE_A_TWO_BY_TWO = 0x0080
const GLYF_WE_HAVE_INSTRUCTIONS = 0x0100
const GLYF_USE_MY_METRICS = 0x0200
const GLYF_OVERLAP_COMPOUND = 0x0400
const GLYF_SCALED_COMPONENT_OFFSET = 0x0800
const GLYF_UNSCALED_COMPONENT_OFFSET = 0x1000
const GLYF_MODELED_FLAGS = GLYF_ARG_1_AND_2_ARE_WORDS | GLYF_ARGS_ARE_XY_VALUES | GLYF_ROUND_XY_TO_GRID | GLYF_WE_HAVE_A_SCALE | GLYF_MORE_COMPONENTS | GLYF_WE_HAVE_AN_X_AND_Y_SCALE | GLYF_WE_HAVE_A_TWO_BY_TWO | GLYF_WE_HAVE_INSTRUCTIONS | GLYF_USE_MY_METRICS | GLYF_OVERLAP_COMPOUND | GLYF_SCALED_COMPONENT_OFFSET | GLYF_UNSCALED_COMPONENT_OFFSET

function flattenCompositeGlyf(bytes: Uint8Array, start: number, limit: number, stack: Set<number>, depth: number, parent: GlyfTransform, unitsPerEm: number): DesignOutline {
  let cursor = start
  const path: DesignPathCommand[] = []
  let more = true
  while (more) {
    if (!range(bytes, cursor, 4)) return { status: 'composite' }
    const flags = u16(bytes, cursor); cursor += 2
    const glyphIndex = u16(bytes, cursor); cursor += 2
    if ((flags & ~GLYF_MODELED_FLAGS) !== 0 || (flags & GLYF_ARGS_ARE_XY_VALUES) === 0) return { status: 'composite' }
    if ((flags & GLYF_SCALED_COMPONENT_OFFSET) !== 0 && (flags & GLYF_UNSCALED_COMPONENT_OFFSET) !== 0) return { status: 'composite' }
    let arg1: number, arg2: number
    if (flags & GLYF_ARG_1_AND_2_ARE_WORDS) {
      if (!range(bytes, cursor, 4)) return { status: 'composite' }
      arg1 = i16(bytes, cursor); arg2 = i16(bytes, cursor + 2); cursor += 4
    } else {
      if (!range(bytes, cursor, 2)) return { status: 'composite' }
      arg1 = bytes[cursor]! >= 0x80 ? bytes[cursor]! - 256 : bytes[cursor]!
      arg2 = bytes[cursor + 1]! >= 0x80 ? bytes[cursor + 1]! - 256 : bytes[cursor + 1]!
      cursor += 2
    }
    let xx = 16384, xy = 0, yx = 0, yy = 16384
    if (flags & GLYF_WE_HAVE_A_SCALE) {
      if (!range(bytes, cursor, 2)) return { status: 'composite' }
      xx = yy = i16(bytes, cursor); cursor += 2
    } else if (flags & GLYF_WE_HAVE_AN_X_AND_Y_SCALE) {
      if (!range(bytes, cursor, 4)) return { status: 'composite' }
      xx = i16(bytes, cursor); yy = i16(bytes, cursor + 2); cursor += 4
    } else if (flags & GLYF_WE_HAVE_A_TWO_BY_TWO) {
      if (!range(bytes, cursor, 8)) return { status: 'composite' }
      xx = i16(bytes, cursor); xy = i16(bytes, cursor + 2); yx = i16(bytes, cursor + 4); yy = i16(bytes, cursor + 6); cursor += 8
    }
    let dx = arg1, dy = arg2
    if (flags & GLYF_SCALED_COMPONENT_OFFSET) {
      const scaledX = applyF2Dot14(arg1, xx)
      const scaledY = applyF2Dot14(arg2, yy)
      if (scaledX === undefined || scaledY === undefined) return { status: 'composite' }
      dx = scaledX
      dy = scaledY
    }
    const child = composeGlyfTransform(parent, { xx, xy, yx, yy, dx, dy })
    if (!child) return { status: 'composite' }
    const outlined = parseGlyfOutlineAt(bytes, glyphIndex, stack, depth + 1, child)
    if (outlined.status === 'composite' || outlined.status === 'invalid') return outlined
    if (outlined.status === 'outlined') path.push(...outlined.path)
    more = (flags & GLYF_MORE_COMPONENTS) !== 0
    if (!more && (flags & GLYF_WE_HAVE_INSTRUCTIONS)) {
      if (!range(bytes, cursor, 2)) return { status: 'composite' }
      const instructionLength = u16(bytes, cursor)
      cursor += 2 + instructionLength
      if (cursor > limit) return { status: 'composite' }
    }
  }
  if (path.length === 0) return { status: 'empty', unitsPerEm }
  return { status: 'outlined', unitsPerEm, path }
}

function composeGlyfTransform(parent: GlyfTransform, child: GlyfTransform): GlyfTransform | undefined {
  const xx = mulF2Dot14(child.xx, parent.xx)
  const xy = mulF2Dot14(child.xx, parent.xy)
  const yx = mulF2Dot14(child.yx, parent.xx)
  const yy = mulF2Dot14(child.yx, parent.xy)
  const xx2 = mulF2Dot14(child.xy, parent.yx)
  const xy2 = mulF2Dot14(child.xy, parent.yy)
  const yx2 = mulF2Dot14(child.yy, parent.yx)
  const yy2 = mulF2Dot14(child.yy, parent.yy)
  if ([xx, xy, yx, yy, xx2, xy2, yx2, yy2].some((value) => value === undefined)) return undefined
  const dx = applyF2Dot14(child.dx, parent.xx)
  const dx2 = applyF2Dot14(child.dy, parent.yx)
  const dy = applyF2Dot14(child.dx, parent.xy)
  const dy2 = applyF2Dot14(child.dy, parent.yy)
  if ([dx, dx2, dy, dy2].some((value) => value === undefined)) return undefined
  return { xx: xx! + xx2!, xy: xy! + xy2!, yx: yx! + yx2!, yy: yy! + yy2!, dx: dx! + dx2! + parent.dx, dy: dy! + dy2! + parent.dy }
}

function mulF2Dot14(left: number, right: number): number | undefined {
  const product = left * right
  if (!Number.isSafeInteger(product) || product % 16384 !== 0) return undefined
  return product / 16384
}

function applyF2Dot14(design: number, scale: number): number | undefined {
  const twice = design * 2
  if (!Number.isSafeInteger(twice)) return undefined
  const product = twice * scale
  if (!Number.isSafeInteger(product) || product % 16384 !== 0) return undefined
  return product / 32768
}

function transformDesignPath(path: DesignPathCommand[], transform: GlyfTransform): DesignPathCommand[] | undefined {
  if (transform.xx === 16384 && transform.xy === 0 && transform.yx === 0 && transform.yy === 16384 && transform.dx === 0 && transform.dy === 0) return path
  const output: DesignPathCommand[] = []
  for (const command of path) {
    if (command.kind === 'close_path') { output.push(command); continue }
    const point = transformDesignPoint(command.x, command.y, transform)
    if (!point) return undefined
    if (command.kind === 'move_to' || command.kind === 'line_to') output.push({ kind: command.kind, x: point.x, y: point.y })
    else {
      const control = transformDesignPoint(command.control_x, command.control_y, transform)
      if (!control) return undefined
      output.push({ kind: 'quadratic_to', control_x: control.x, control_y: control.y, x: point.x, y: point.y })
    }
  }
  return output
}

function transformDesignPoint(x: number, y: number, transform: GlyfTransform): { x: number; y: number } | undefined {
  const x1 = applyF2Dot14(x, transform.xx)
  const x2 = applyF2Dot14(y, transform.yx)
  const y1 = applyF2Dot14(x, transform.xy)
  const y2 = applyF2Dot14(y, transform.yy)
  if ([x1, x2, y1, y2].some((value) => value === undefined)) return undefined
  return { x: x1! + x2! + transform.dx, y: y1! + y2! + transform.dy }
}

function contourToPath(points: Array<{ x: number; y: number; on: boolean }>): DesignPathCommand[] | undefined {
  if (points.length < 2) return undefined
  const expanded: Array<{ x: number; y: number; on: boolean }> = []
  for (let index = 0; index < points.length; index++) {
    const current = points[index]!, next = points[(index + 1) % points.length]!
    expanded.push(current)
    if (!current.on && !next.on) expanded.push({ x: (current.x + next.x) / 2, y: (current.y + next.y) / 2, on: true })
  }
  const start = expanded.findIndex((point) => point.on)
  if (start < 0) return undefined
  const path: DesignPathCommand[] = [{ kind: 'move_to', x: expanded[start]!.x, y: expanded[start]!.y }]
  let index = start
  do {
    const next = (index + 1) % expanded.length
    const nextPoint = expanded[next]!
    if (nextPoint.on) {
      path.push({ kind: 'line_to', x: nextPoint.x, y: nextPoint.y })
      index = next
    } else {
      const end = (next + 1) % expanded.length
      const endPoint = expanded[end]!
      if (!endPoint.on) return undefined
      path.push({ kind: 'quadratic_to', control_x: nextPoint.x, control_y: nextPoint.y, x: endPoint.x, y: endPoint.y })
      index = end
    }
  } while (index !== start)
  path.push({ kind: 'close_path' })
  return path
}

function sfntTables(bytes: Uint8Array): Map<string, { offset: number; length: number }> | undefined {
  if (bytes.byteLength < 12) return undefined
  const flavor = u32(bytes, 0)
  if (flavor !== 0x00010000 && flavor !== 0x74727565) return undefined
  const count = u16(bytes, 4)
  if (count < 1 || count > 256 || !range(bytes, 12, count * 16)) return undefined
  const tables = new Map<string, { offset: number; length: number }>()
  for (let index = 0; index < count; index++) {
    const at = 12 + index * 16
    const name = ascii(bytes, at, 4), offset = u32(bytes, at + 8), length = u32(bytes, at + 12)
    if (tables.has(name) || length < 1 || !range(bytes, offset, length)) return undefined
    tables.set(name, { offset, length })
  }
  return tables
}

function normalizePaintFontName(value: string): string {
  return value.normalize('NFKC').toLowerCase().replace(/[\s_-]/g, '')
}

function fontOs2Style(bytes: Uint8Array): { bold: boolean; italic: boolean } {
  const tables = sfntTables(bytes)
  const os2 = tables?.get('OS/2')
  const head = tables?.get('head')
  const selection = os2 && os2.length >= 64 ? u16(bytes, os2.offset + 62) : undefined
  const macStyle = head && head.length >= 46 ? u16(bytes, head.offset + 44) : 0
  return {
    bold: selection === undefined ? (macStyle & 1) !== 0 : (selection & 0x20) !== 0,
    italic: selection === undefined ? (macStyle & 2) !== 0 : (selection & 1) !== 0,
  }
}

function fontFileNames(bytes: Uint8Array): string[] {
  const tables = sfntTables(bytes)
  const table = tables?.get('name')
  if (!table || table.length < 6) return []
  const count = u16(bytes, table.offset + 2), strings = u16(bytes, table.offset + 4), names = new Set<string>()
  if (count > 16_384 || table.length < 6 + count * 12 || strings > table.length) return []
  for (let index = 0; index < count; index++) {
    const at = table.offset + 6 + index * 12, platform = u16(bytes, at), nameID = u16(bytes, at + 6), length = u16(bytes, at + 8), offset = u16(bytes, at + 10)
    if (![1, 4, 6, 16].includes(nameID) || offset + length > table.length - strings) continue
    const valueAt = table.offset + strings + offset
    let value = ''
    if (platform === 0 || platform === 3) {
      if (length % 2 !== 0) continue
      for (let unit = 0; unit < length; unit += 2) value += String.fromCharCode(u16(bytes, valueAt + unit))
    } else if (platform === 1) value = ascii(bytes, valueAt, length)
    if (value.length > 0 && value.length <= 255) names.add(value)
  }
  return [...names]
}

function snapshotFontBytes(fontBytes: Uint8Array, expected: string): Uint8Array {
  if (!(fontBytes instanceof Uint8Array) || Object.getPrototypeOf(fontBytes) !== Uint8Array.prototype || fontBytes.byteLength < 12 || fontBytes.byteLength > maximumFontBytes) {
    throw new NativeSheetCellPaintError('paint.metricAuthority', '$.font_bytes', 'font bytes must be a bounded direct Uint8Array')
  }
  let bytes: Uint8Array
  try { bytes = fontBytes.slice() } catch { throw new NativeSheetCellPaintError('paint.metricAuthority', '$.font_bytes', 'Proxy font bytes are refused') }
  if (`sha256:${sha256HexBytes(bytes)}` !== expected) throw new NativeSheetCellPaintError('paint.metricAuthority', '$.font_bytes', 'font bytes do not match the geometry metric-authority digest')
  return bytes
}

function validateSourceAuthority(workbook: NativeWorkbookRenderModelV2, sheet: NativeSheetRenderModelV2): void {
  if (sheet.mutation_authority.source_revision !== workbook.revision || typeof sheet.mutation_authority.source_part !== 'string' || sheet.mutation_authority.source_part.length === 0) {
    throw new NativeSheetCellPaintError('paint.sourceUnsupported', '$.sheet.mutation_authority', 'sheet source authority is not bound to this workbook revision')
  }
  if (sheet.rows.some((row) => row.style_id !== undefined) || sheet.columns.some((column) => column.style_id !== undefined)) {
    throw new NativeSheetCellPaintError('paint.styleUnsupported', '$.sheet', 'row and column style precedence is not projected by native v2')
  }
  const issue = workbook.unsupported.find((item) => {
    if (item.scope_id !== 'workbook' && item.scope_id !== `sheet:${sheet.id}`) return false
    return appearanceAuthorityCapabilities.has(item.capability) || appearanceAuthorityCodes.has(item.code)
  })
  if (issue) throw new NativeSheetCellPaintError('paint.sourceUnsupported', '$.unsupported', `${issue.code} can change source-authoritative sheet appearance`)
}

function validateGeometryAuthority(workbook: NativeWorkbookRenderModelV2, geometry: NativeSheetGeometryV2): void {
  if (geometry.document_id !== workbook.document_id || geometry.source_revision !== workbook.revision || geometry.source_package_sha256 !== workbook.source.package_sha256 || geometry.coordinate_space !== 'viewport-local') {
    throw new NativeSheetCellPaintError('paint.geometryAuthority', '$.geometry', 'geometry is not bound to this exact workbook revision')
  }
  try {
    if (!isCompiledNativeSheetGeometryV2(geometry)) throw new Error('geometry is not branded')
    const canonical = validateNativeSheetGeometryV2(geometry)
    if (JSON.stringify(canonical) !== JSON.stringify(geometry)) throw new Error('geometry is not canonical')
  } catch {
    throw new NativeSheetCellPaintError('paint.geometryAuthority', '$.geometry', 'geometry is not the deterministic projection of this workbook and metric authority')
  }
}

function cellHasDisplayContent(content: NativeRenderCellContentV2 | undefined): boolean {
  if (!content || content.kind === 'blank') return false
  if (content.kind === 'literal') return true
  return content.formula.cached !== undefined
}

function refusal(row: number, column: number, code: NativeSheetCellPaintUnsupportedCodeV2, message: string): NativeSheetCellPaintUnsupportedV2 {
  return { cell_ref: cellReference(row, column), row, column, code, message }
}

const unsupportedCodes = new Set<NativeSheetCellPaintUnsupportedCodeV2>([
  'CELL_RICH_TEXT', 'CELL_FORMULA_NO_CACHE', 'CELL_DATE_DISPLAY', 'CELL_BOOLEAN_DISPLAY', 'CELL_ERROR_DISPLAY', 'CELL_NON_ASCII',
  'CELL_COMPLEX_CLUSTER', 'CELL_FONT_MISMATCH', 'CELL_FONT_COLOR', 'CELL_WRAP_TEXT', 'CELL_ALIGNMENT_EXTENDED', 'CELL_STYLE_UNSUPPORTED', 'CELL_MISSING_GLYPH',
  'CELL_COMPOSITE_OUTLINE', 'CELL_OVERFLOW', 'CELL_MERGE_NON_ORIGIN', 'CELL_HIDDEN',
])

type UnknownRecord = Record<string, unknown>

function validatePath(value: unknown, path: string, bounds: NativeSheetGeometryRectV2): NativeSheetCellPaintPathCommandV2[] {
  if (!Array.isArray(value) || value.length < 1 || value.length > NATIVE_SHEET_CELL_PAINT_V2_LIMITS.maxPathCommandsPerGlyph) invalidPlan(path, 'glyph path is missing or exceeds its bound')
  const result: NativeSheetCellPaintPathCommandV2[] = []
  let contourOpen = false
  let contourDrawn = false
  for (let index = 0; index < value.length; index++) {
    const commandPath = `${path}[${index}]`
    if (typeof value[index] !== 'object' || value[index] === null || Array.isArray(value[index])) invalidPlan(commandPath, 'path command must be an object')
    const command = value[index] as UnknownRecord
    if (command.kind === 'move_to') {
      const record = exactObject(command, ['kind', 'x_emu', 'y_emu'], commandPath)
      if (contourOpen) invalidPlan(commandPath, 'move_to cannot start a nested contour')
      contourOpen = true
      contourDrawn = false
      result.push({ kind: 'move_to', x_emu: boundedInteger(record.x_emu, `${commandPath}.x_emu`, 0, Number.MAX_SAFE_INTEGER), y_emu: boundedInteger(record.y_emu, `${commandPath}.y_emu`, 0, Number.MAX_SAFE_INTEGER) })
    } else if (command.kind === 'line_to') {
      const record = exactObject(command, ['kind', 'x_emu', 'y_emu'], commandPath)
      if (!contourOpen) invalidPlan(commandPath, 'line_to must follow move_to')
      contourDrawn = true
      result.push({ kind: 'line_to', x_emu: boundedInteger(record.x_emu, `${commandPath}.x_emu`, 0, Number.MAX_SAFE_INTEGER), y_emu: boundedInteger(record.y_emu, `${commandPath}.y_emu`, 0, Number.MAX_SAFE_INTEGER) })
    } else if (command.kind === 'quadratic_to') {
      const record = exactObject(command, ['kind', 'control_x_emu', 'control_y_emu', 'x_emu', 'y_emu'], commandPath)
      if (!contourOpen) invalidPlan(commandPath, 'quadratic_to must follow move_to')
      contourDrawn = true
      result.push({
        kind: 'quadratic_to',
        control_x_emu: boundedInteger(record.control_x_emu, `${commandPath}.control_x_emu`, 0, Number.MAX_SAFE_INTEGER),
        control_y_emu: boundedInteger(record.control_y_emu, `${commandPath}.control_y_emu`, 0, Number.MAX_SAFE_INTEGER),
        x_emu: boundedInteger(record.x_emu, `${commandPath}.x_emu`, 0, Number.MAX_SAFE_INTEGER),
        y_emu: boundedInteger(record.y_emu, `${commandPath}.y_emu`, 0, Number.MAX_SAFE_INTEGER),
      })
    } else if (command.kind === 'close_path') {
      const record = exactObject(command, ['kind'], commandPath)
      if (!contourOpen || !contourDrawn || record.kind !== 'close_path') invalidPlan(commandPath, 'close_path must close a drawn contour')
      contourOpen = false
      contourDrawn = false
      result.push({ kind: 'close_path' })
    } else invalidPlan(commandPath, 'path command kind is unsupported')
  }
  if (contourOpen) invalidPlan(path, 'glyph path left an open contour')
  if (!pathContained(result, bounds)) invalidPlan(path, 'glyph path escapes decoration bounds')
  return result
}

function pathContained(path: ReadonlyArray<NativeSheetCellPaintPathCommandV2>, bounds: NativeSheetGeometryRectV2): boolean {
  let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity
  const visit = (x: number, y: number): void => { minX = Math.min(minX, x); maxX = Math.max(maxX, x); minY = Math.min(minY, y); maxY = Math.max(maxY, y) }
  for (const command of path) {
    if (command.kind === 'close_path') continue
    if (command.kind === 'quadratic_to') visit(command.control_x_emu, command.control_y_emu)
    visit(command.x_emu, command.y_emu)
  }
  if (!(maxX > minX && maxY > minY)) return false
  return minX >= bounds.x_emu && minY >= bounds.y_emu && maxX <= bounds.x_emu + bounds.width_emu && maxY <= bounds.y_emu + bounds.height_emu
}

function snapshotPaintInput(value: unknown, path: string): unknown {
  try { return snapshotNativePlainData(value, { maxDepth: 32, maxNodes: NATIVE_SHEET_CELL_PAINT_V2_LIMITS.maxCommands * 16 }) }
  catch (error) {
    if (error instanceof NativePlainDataError) throw new NativeSheetCellPaintError('paint.planInvalid', error.path || path, error.message)
    throw error
  }
}

function exactObject(value: unknown, keys: ReadonlyArray<string>, path: string): UnknownRecord {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) invalidPlan(path, 'expected an object')
  const record = value as UnknownRecord
  const actual = Object.keys(record).sort()
  const expected = [...keys].sort()
  if (actual.length !== expected.length || actual.some((key, index) => key !== expected[index])) invalidPlan(path, 'object fields are missing or unknown')
  return record
}

function exactCommand(value: unknown, keys: ReadonlyArray<string>, path: string): UnknownRecord {
  try { return exactObject(value, keys, path) }
  catch (error) {
    if (error instanceof NativeSheetCellPaintError) throw new NativeSheetCellPaintError('paint.commandInvalid', error.path, error.message)
    throw error
  }
}

function invalidPlan(path: string, message: string): never { throw new NativeSheetCellPaintError('paint.planInvalid', path, message) }
function commandInvalid(path: string, message: string): never { throw new NativeSheetCellPaintError('paint.commandInvalid', path, message) }

function stringPattern(value: unknown, path: string, pattern: RegExp): string {
  if (typeof value !== 'string' || !pattern.test(value)) invalidPlan(path, 'string is missing or non-canonical')
  return value
}

function boundedString(value: unknown, path: string, maximum: number): string {
  if (typeof value !== 'string' || value.length < 1 || value.length > maximum || /[\u0000-\u001f\u007f]/.test(value)) invalidPlan(path, 'string is missing, oversized, or contains control characters')
  return value
}

function isCanonicalDisplayText(value: unknown): value is string {
  if (typeof value !== 'string' || value.length < 1 || value.length > 32_767 || /[\u0000-\u001f\u007f]/.test(value)) return false
  for (let index = 0; index < value.length; index++) {
    const unit = value.charCodeAt(index)!
    if (unit >= 0xd800 && unit <= 0xdbff) {
      const trail = value.charCodeAt(index + 1)
      if (trail === undefined || trail < 0xdc00 || trail > 0xdfff) return false
      index += 1
      continue
    }
    if (unit >= 0xdc00 && unit <= 0xdfff) return false
  }
  return true
}

function identifier(value: unknown, path: string, maximum: number): string {
  if (typeof value !== 'string' || value.length < 1 || value.length > maximum || !/^[A-Za-z0-9][A-Za-z0-9._:-]*$/.test(value)) invalidPlan(path, 'identifier is missing, oversized, or non-canonical')
  return value
}

function canonicalSheetID(value: unknown, path: string): string {
  const result = stringPattern(value, path, /^[1-9][0-9]{0,9}$/)
  if (Number(result) > 4_294_967_295) invalidPlan(path, 'sheet id is outside uint32 bounds')
  return result
}

function canonicalPart(value: unknown, path: string): string {
  if (typeof value !== 'string' || value.length < 1 || value.length > 8_192 || value.startsWith('/') || /[?#\\\u0000-\u001f\u007f]/.test(value) || /^[A-Za-z][A-Za-z0-9+.-]*:/.test(value) || /%(?:2f|5c)/i.test(value) || /%(?![0-9A-F]{2})/.test(value)) invalidPlan(path, 'source part is not a canonical OPC part name')
  for (const match of value.matchAll(/%([0-9A-F]{2})/g)) {
    const decodedUnit = String.fromCharCode(Number.parseInt(match[1]!, 16))
    if (/^[A-Za-z0-9._~-]$/.test(decodedUnit)) invalidPlan(path, 'unreserved OPC part characters must not be percent-encoded')
  }
  let decoded: string
  try { decoded = decodeURIComponent(value) } catch { invalidPlan(path, 'source part has a malformed percent escape') }
  const segments = decoded.split('/')
  if (!decoded || decoded.startsWith('/') || decoded.endsWith('/') || decoded.includes('//') || segments.some((segment) => !segment || segment === '.' || segment === '..' || segment.endsWith('.') || /[\\\u0000-\u001f\u007f]/.test(segment))) invalidPlan(path, 'source part is not a canonical OPC part name')
  return value
}

function boundedInteger(value: unknown, path: string, minimum: number, maximum: number): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || Object.is(value, -0) || value < minimum || value > maximum) invalidPlan(path, `integer must be within ${minimum}..${maximum} and not negative zero`)
  return value
}

function validateRect(value: unknown, path: string, requireArea: boolean): NativeSheetGeometryRectV2 {
  const rect = exactObject(value, ['x_emu', 'y_emu', 'width_emu', 'height_emu'], path)
  const result = {
    x_emu: boundedInteger(rect.x_emu, `${path}.x_emu`, 0, Number.MAX_SAFE_INTEGER),
    y_emu: boundedInteger(rect.y_emu, `${path}.y_emu`, 0, Number.MAX_SAFE_INTEGER),
    width_emu: boundedInteger(rect.width_emu, `${path}.width_emu`, requireArea ? 1 : 0, Number.MAX_SAFE_INTEGER),
    height_emu: boundedInteger(rect.height_emu, `${path}.height_emu`, requireArea ? 1 : 0, Number.MAX_SAFE_INTEGER),
  }
  if (!Number.isSafeInteger(result.x_emu + result.width_emu) || !Number.isSafeInteger(result.y_emu + result.height_emu)) invalidPlan(path, 'rectangle endpoint exceeds safe integer bounds')
  return result
}

function rectContained(rect: NativeSheetGeometryRectV2, bounds: NativeSheetGeometryRectV2): boolean {
  return rect.x_emu >= bounds.x_emu && rect.y_emu >= bounds.y_emu && rect.x_emu + rect.width_emu <= bounds.x_emu + bounds.width_emu && rect.y_emu + rect.height_emu <= bounds.y_emu + bounds.height_emu
}

function paintDigest(unsigned: Omit<NativeSheetCellPaintPlanV2, 'paint_sha256'>): `sha256:${string}` {
  return `sha256:${sha256Hex(`injoffice.xlsx.sheet-cell-paint.v1\0${JSON.stringify(unsigned)}`)}`
}

function validateCommandLimit(value: unknown, path: string): asserts value is number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || Object.is(value, -0) || value < 1 || value > NATIVE_SHEET_CELL_PAINT_V2_LIMITS.maxCommands) throw new NativeSheetCellPaintError('paint.invalidLimit', path, `command limit must be from 1 through ${NATIVE_SHEET_CELL_PAINT_V2_LIMITS.maxCommands}`)
}

function optionalDataProperty(value: unknown, name: string, path: string): unknown {
  if ((typeof value !== 'object' && typeof value !== 'function') || value === null) commandInvalid(path, 'callback owner must be an object')
  const descriptor = Object.getOwnPropertyDescriptor(value, name)
  if (!descriptor) return undefined
  if (!('value' in descriptor)) commandInvalid(path, 'callback preflight refuses accessors')
  return descriptor.value
}

function callbackDataProperty(value: unknown, name: string, path: string): Function {
  const callback = optionalDataProperty(value, name, path)
  if (typeof callback !== 'function') commandInvalid(path, 'callback must be an own data-property function')
  return callback
}

function cellReference(row: number, column: number): string {
  let value = column + 1, name = ''
  while (value > 0) { value--; name = String.fromCharCode(65 + value % 26) + name; value = Math.floor(value / 26) }
  return `${name}${row + 1}`
}

function range(bytes: Uint8Array, offset: number, length: number): boolean {
  return Number.isSafeInteger(offset) && Number.isSafeInteger(length) && offset >= 0 && length >= 0 && offset <= bytes.length && length <= bytes.length - offset
}

function u16(bytes: Uint8Array, offset: number): number { return bytes[offset]! * 256 + bytes[offset + 1]! }
function i16(bytes: Uint8Array, offset: number): number { const value = u16(bytes, offset); return value >= 0x8000 ? value - 0x10000 : value }
function u32(bytes: Uint8Array, offset: number): number { return ((bytes[offset]! * 0x1000000) + (bytes[offset + 1]! << 16) + (bytes[offset + 2]! << 8) + bytes[offset + 3]!) >>> 0 }
function ascii(bytes: Uint8Array, offset: number, length: number): string { let result = ''; for (let index = 0; index < length; index++) result += String.fromCharCode(bytes[offset + index]!); return result }

function deepFreeze<T>(value: T): T {
  if (typeof value !== 'object' || value === null || Object.isFrozen(value)) return value
  for (const child of Object.values(value)) deepFreeze(child)
  return Object.freeze(value)
}
