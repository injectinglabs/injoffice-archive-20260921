import type { NativeBorderStyle, NativeWorkbookBorderSideV2 } from './nativeContractV2.generated.js'
import { isProjectedNativeWorkbookV2 } from './nativeRenderModelV2.js'
import type { NativeRenderCellV2, NativeRenderStyleV2, NativeSheetRenderModelV2, NativeWorkbookRenderModelV2 } from './nativeRenderModelV2.js'
import { isCompiledNativeSheetGeometryV2, validateNativeSheetGeometryV2 } from './nativeSheetGeometryV2.js'
import type { NativeSheetGeometryRectV2, NativeSheetGeometryV2 } from './nativeSheetGeometryV2.js'
import { sha256Hex } from './nativeSha256.js'
import { nativeWorkbookStyleRawProjectionSha256V2 } from './nativeValidationV2.js'
import { NativePlainDataError, snapshotNativePlainData } from './nativePlainData.js'

export const NATIVE_SHEET_DECORATION_V2_PROTOCOL = 'injoffice.xlsx.sheet-decoration'
export const NATIVE_SHEET_DECORATION_V2_VERSION = 1 as const
export const NATIVE_SHEET_DECORATION_V2_LIMITS = Object.freeze({
  maxFills: 100_000,
  maxBorderSegments: 200_000,
  maxCommands: 300_003,
})

export type NativeSheetDecorationIssueCode =
  | 'decoration.sheetMissing'
  | 'decoration.geometryAuthority'
  | 'decoration.modelAuthority'
  | 'decoration.sourceUnsupported'
  | 'decoration.stylePrecedence'
  | 'decoration.styleUnsupported'
  | 'decoration.mergeUnsupported'
  | 'decoration.borderConflict'
  | 'decoration.resourceBudget'
  | 'decoration.commandBudget'
  | 'decoration.invalidLimit'
  | 'decoration.hiddenDimension'
  | 'decoration.planInvalid'
  | 'decoration.planDigest'
  | 'decoration.commandInvalid'

export class NativeSheetDecorationError extends Error {
  readonly code: NativeSheetDecorationIssueCode
  readonly path: string

  constructor(code: NativeSheetDecorationIssueCode, path: string, message: string) {
    super(message)
    this.name = 'NativeSheetDecorationError'
    this.code = code
    this.path = path
  }
}

export interface NativeSheetFillDecorationV2 {
  readonly cell_ref: string
  readonly row: number
  readonly column: number
  readonly style_id: number
  readonly fill_id: number
  readonly fill_record_sha256: string
  readonly rect: NativeSheetGeometryRectV2
  readonly color: string
}

export type NativeBorderEdgeV2 = 'left' | 'right' | 'top' | 'bottom'

export interface NativeSheetBorderSourceV2 {
  readonly cell_ref: string
  readonly edge: NativeBorderEdgeV2
  readonly style_id: number
  readonly border_id: number
  readonly border_record_sha256: string
}

export interface NativeSheetBorderSegmentV2 {
  readonly orientation: 'horizontal' | 'vertical'
  readonly x1_emu: number
  readonly y1_emu: number
  readonly x2_emu: number
  readonly y2_emu: number
  /** Exact OOXML token. It deliberately does not invent a physical stroke width. */
  readonly border_style: NativeBorderStyle
  readonly color: string
  readonly sources: ReadonlyArray<NativeSheetBorderSourceV2>
}

export interface NativeSheetDecorationPlanV2 {
  readonly protocol: typeof NATIVE_SHEET_DECORATION_V2_PROTOCOL
  readonly version: typeof NATIVE_SHEET_DECORATION_V2_VERSION
  readonly decoration_sha256: `sha256:${string}`
  readonly document_id: string
  readonly sheet_id: string
  readonly source_part: string
  readonly source_revision: string
  readonly source_package_sha256: string
  readonly geometry_sha256: string
  readonly coordinate_space: 'viewport-local'
  readonly bounds: NativeSheetGeometryRectV2
  readonly fills: ReadonlyArray<NativeSheetFillDecorationV2>
  readonly border_segments: ReadonlyArray<NativeSheetBorderSegmentV2>
}

export type NativeSheetDecorationCommandV2 =
  | { readonly kind: 'beginDecorations'; readonly protocol: typeof NATIVE_SHEET_DECORATION_V2_PROTOCOL; readonly version: typeof NATIVE_SHEET_DECORATION_V2_VERSION; readonly document_id: string; readonly sheet_id: string; readonly source_part: string; readonly source_revision: string; readonly source_package_sha256: string; readonly geometry_sha256: string; readonly coordinate_space: 'viewport-local'; readonly decoration_sha256: string }
  | { readonly kind: 'clipRect'; readonly rect: NativeSheetGeometryRectV2 }
  | { readonly kind: 'fillRect'; readonly decoration: NativeSheetFillDecorationV2 }
  | { readonly kind: 'borderSegment'; readonly decoration: NativeSheetBorderSegmentV2 }
  | { readonly kind: 'endDecorations' }

export interface NativeSheetDecorationSurfaceV2 {
  readonly command_capacity?: number
  push(command: NativeSheetDecorationCommandV2): void
}

export interface NativeSheetDecorationRecordingSurfaceV2 extends NativeSheetDecorationSurfaceV2 {
  readonly commands: ReadonlyArray<NativeSheetDecorationCommandV2>
  finish(): ReadonlyArray<NativeSheetDecorationCommandV2>
}

export interface NativeSheetDecorationCommandAdapterV2<HostContext> {
  execute(context: HostContext, command: NativeSheetDecorationCommandV2): void
}

type BorderCandidate = Omit<NativeSheetBorderSegmentV2, 'sources'> & { sources: NativeSheetBorderSourceV2[] }

export function compileNativeSheetDecorationsV2(workbook: NativeWorkbookRenderModelV2, geometry: NativeSheetGeometryV2): NativeSheetDecorationPlanV2 {
  if (!isProjectedNativeWorkbookV2(workbook)) throw new NativeSheetDecorationError('decoration.modelAuthority', '$.workbook', 'workbook must be the branded frozen result of projectNativeWorkbookV2')
  validateGeometryAuthority(workbook, geometry)
  const sheet = workbook.sheets.find((candidate) => candidate.id === geometry.sheet_id)
  if (!sheet) throw new NativeSheetDecorationError('decoration.sheetMissing', '$.geometry.sheet_id', `sheet ${JSON.stringify(geometry.sheet_id)} is absent`)
  validateSourceAuthority(workbook, sheet, geometry)

  const cells = new Map<number, NativeRenderCellV2>()
  for (const cell of sheet.cells) {
    const key = cell.row * 16_384 + cell.column
    if (!Number.isSafeInteger(cell.row) || !Number.isSafeInteger(cell.column) || cell.row < 0 || cell.row >= 1_048_576 || cell.column < 0 || cell.column >= 16_384 || cell.ref !== cellReference(cell.row, cell.column) || cells.has(key)) {
      throw new NativeSheetDecorationError('decoration.sourceUnsupported', '$.sheet.cells', 'cell coordinates or identity are non-canonical')
    }
    cells.set(key, cell)
  }
  const fills: NativeSheetFillDecorationV2[] = []
  const segments = new Map<string, BorderCandidate>()

  for (let rowOffset = 0; rowOffset < geometry.rows.length; rowOffset++) {
    const rowBand = geometry.rows[rowOffset]!
    for (let columnOffset = 0; columnOffset < geometry.columns.length; columnOffset++) {
      const columnBand = geometry.columns[columnOffset]!
      const cell = cells.get(rowBand.row * 16_384 + columnBand.column)
      const styleID = cell?.style_id ?? 0
      const style = requireDecorationStyle(workbook, styleID)
      if (rowBand.height_emu === 0 || columnBand.width_emu === 0) {
        if (styleHasDecoration(style)) throw new NativeSheetDecorationError('decoration.hiddenDimension', '$.geometry', `styled hidden or zero-size cell ${cell?.ref ?? cellReference(rowBand.row, columnBand.column)} has unresolved edge semantics`)
        continue
      }
      const rect = {
        x_emu: columnBand.x_emu,
        y_emu: rowBand.y_emu,
        width_emu: columnBand.width_emu,
        height_emu: rowBand.height_emu,
      }
      const cellRef = cell?.ref ?? cellReference(rowBand.row, columnBand.column)
      if (style.effective.fill_color !== undefined) {
        if (fills.length >= NATIVE_SHEET_DECORATION_V2_LIMITS.maxFills) throw new NativeSheetDecorationError('decoration.resourceBudget', '$.fills', 'fill decorations exceed the native resource bound')
        const fill = style.effective.fill!
        if (fill.origin !== 'styles-record' || fill.fill_id === undefined || fill.record_sha256 === undefined) throw new NativeSheetDecorationError('decoration.styleUnsupported', `$.styles[${styleID}].effective.fill`, 'visible fill lacks exact styles-table provenance')
        fills.push({ cell_ref: cellRef, row: rowBand.row, column: columnBand.column, style_id: styleID, fill_id: fill.fill_id, fill_record_sha256: fill.record_sha256, rect, color: style.effective.fill_color })
      }
      const border = style.effective.border!
      if (border.origin === 'styles-record') {
        addBorder(segments, border.left, 'left', rect.x_emu, rect.y_emu, rect.x_emu, rect.y_emu + rect.height_emu, style, cellRef)
        addBorder(segments, border.top, 'top', rect.x_emu, rect.y_emu, rect.x_emu + rect.width_emu, rect.y_emu, style, cellRef)
        addBorder(segments, border.right, 'right', rect.x_emu + rect.width_emu, rect.y_emu, rect.x_emu + rect.width_emu, rect.y_emu + rect.height_emu, style, cellRef)
        addBorder(segments, border.bottom, 'bottom', rect.x_emu, rect.y_emu + rect.height_emu, rect.x_emu + rect.width_emu, rect.y_emu + rect.height_emu, style, cellRef)
      }
    }
  }
  // A one-cell source halo is sufficient to resolve every shared edge that is
  // clipped to the viewport. Only the halo edge facing inward is compiled.
  const haloBorder = (row: number, column: number, edge: NativeBorderEdgeV2, x1: number, y1: number, x2: number, y2: number, visible: boolean): void => {
    const cell = cells.get(row * 16_384 + column)
    const style = requireDecorationStyle(workbook, cell?.style_id ?? 0)
    const border = style.effective.border!
    const side = edge === 'left' ? border.left : edge === 'right' ? border.right : edge === 'top' ? border.top : border.bottom
    if (!visible && side !== undefined) throw new NativeSheetDecorationError('decoration.hiddenDimension', '$.geometry', `halo border ${cell?.ref ?? cellReference(row, column)}:${edge} meets a hidden viewport dimension`)
    if (visible && border.origin === 'styles-record') addBorder(segments, side, edge, x1, y1, x2, y2, style, cell?.ref ?? cellReference(row, column))
  }
  if (geometry.viewport.column > 0) for (const rowBand of geometry.rows) haloBorder(rowBand.row, geometry.viewport.column - 1, 'right', 0, rowBand.y_emu, 0, rowBand.y_emu + rowBand.height_emu, rowBand.height_emu > 0)
  if (geometry.viewport.end_column < 16_383) for (const rowBand of geometry.rows) haloBorder(rowBand.row, geometry.viewport.end_column + 1, 'left', geometry.bounds.width_emu, rowBand.y_emu, geometry.bounds.width_emu, rowBand.y_emu + rowBand.height_emu, rowBand.height_emu > 0)
  if (geometry.viewport.row > 0) for (const columnBand of geometry.columns) haloBorder(geometry.viewport.row - 1, columnBand.column, 'bottom', columnBand.x_emu, 0, columnBand.x_emu + columnBand.width_emu, 0, columnBand.width_emu > 0)
  if (geometry.viewport.end_row < 1_048_575) for (const columnBand of geometry.columns) haloBorder(geometry.viewport.end_row + 1, columnBand.column, 'top', columnBand.x_emu, geometry.bounds.height_emu, columnBand.x_emu + columnBand.width_emu, geometry.bounds.height_emu, columnBand.width_emu > 0)
  if (segments.size > NATIVE_SHEET_DECORATION_V2_LIMITS.maxBorderSegments) throw new NativeSheetDecorationError('decoration.resourceBudget', '$.border_segments', 'border segments exceed the native resource bound')
  const borderSegments = [...segments.values()]
    .map((segment) => ({ ...segment, sources: [...segment.sources].sort(compareBorderSources) }))
    .sort(compareBorderSegments)
  const unsigned: Omit<NativeSheetDecorationPlanV2, 'decoration_sha256'> = {
    protocol: NATIVE_SHEET_DECORATION_V2_PROTOCOL,
    version: NATIVE_SHEET_DECORATION_V2_VERSION,
    document_id: workbook.document_id,
    sheet_id: sheet.id,
    source_part: sheet.mutation_authority.source_part,
    source_revision: workbook.revision,
    source_package_sha256: workbook.source.package_sha256,
    geometry_sha256: geometry.geometry_sha256,
    coordinate_space: 'viewport-local' as const,
    bounds: { ...geometry.bounds },
    fills,
    border_segments: borderSegments,
  } as const
  return deepFreeze({ ...unsigned, decoration_sha256: decorationDigest(unsigned) })
}

export function createNativeSheetDecorationRecordingSurfaceV2(maxCommands: number = NATIVE_SHEET_DECORATION_V2_LIMITS.maxCommands): NativeSheetDecorationRecordingSurfaceV2 {
  validateCommandLimit(maxCommands, '$.maxCommands')
  const commands: NativeSheetDecorationCommandV2[] = []
  let finished: ReadonlyArray<NativeSheetDecorationCommandV2> | undefined
  return {
    command_capacity: maxCommands,
    get commands() { return finished ?? Object.freeze([...commands]) },
    push(command) {
      if (finished) throw new Error('native sheet decoration recording is already finished')
      if (commands.length >= maxCommands) throw new NativeSheetDecorationError('decoration.commandBudget', '$.commands', `decoration commands exceed ${maxCommands}`)
      commands.push(deepFreeze(command))
    },
    finish() {
      if (!finished) finished = Object.freeze([...commands])
      return finished
    },
  }
}

export function emitNativeSheetDecorationCommandsV2(plan: NativeSheetDecorationPlanV2, surface: NativeSheetDecorationSurfaceV2, maxCommands: number = NATIVE_SHEET_DECORATION_V2_LIMITS.maxCommands): void {
  const validated = validateNativeSheetDecorationPlanV2(plan)
  validateCommandLimit(maxCommands, '$.maxCommands')
  const surfaceCapacity = optionalDataProperty(surface, 'command_capacity', '$.surface.command_capacity')
  const surfacePush = callbackDataProperty(surface, 'push', '$.surface.push')
  if (surfaceCapacity !== undefined) validateCommandLimit(surfaceCapacity, '$.surface.command_capacity')
  const capacity = Math.min(maxCommands, surfaceCapacity ?? maxCommands)
  const commands = decorationCommands(validated)
  const required = commands.length
  if (required > capacity) throw new NativeSheetDecorationError('decoration.commandBudget', '$.commands', `decoration commands require ${required}, exceeding ${capacity}`)
  for (const command of commands) surfacePush.call(surface, command)
}

export function replayNativeSheetDecorationCommandsV2<HostContext>(context: HostContext, commands: ReadonlyArray<NativeSheetDecorationCommandV2>, adapter: NativeSheetDecorationCommandAdapterV2<HostContext>, maxCommands: number = NATIVE_SHEET_DECORATION_V2_LIMITS.maxCommands): void {
  validateCommandLimit(maxCommands, '$.maxCommands')
  const validated = validateNativeSheetDecorationCommandsV2(commands, maxCommands)
  const execute = callbackDataProperty(adapter, 'execute', '$.adapter.execute') as (context: HostContext, command: NativeSheetDecorationCommandV2) => void
  for (const command of validated) execute.call(adapter, context, command)
}

/** Rebuilds an immutable canonical plan and verifies every untrusted field. */
export function validateNativeSheetDecorationPlanV2(input: unknown): NativeSheetDecorationPlanV2 {
  input = snapshotDecorationInput(input, '$')
  const plan = exactObject(input, ['protocol', 'version', 'decoration_sha256', 'document_id', 'sheet_id', 'source_part', 'source_revision', 'source_package_sha256', 'geometry_sha256', 'coordinate_space', 'bounds', 'fills', 'border_segments'], '$')
  if (plan.protocol !== NATIVE_SHEET_DECORATION_V2_PROTOCOL || plan.version !== NATIVE_SHEET_DECORATION_V2_VERSION) invalidPlan('$', 'decoration protocol or version is unsupported')
  const documentID = identifier(plan.document_id, '$.document_id', 256)
  const sheetID = canonicalSheetID(plan.sheet_id, '$.sheet_id')
  const sourcePart = canonicalPart(plan.source_part, '$.source_part')
  const sourceRevision = stringPattern(plan.source_revision, '$.source_revision', /^rev:[0-9a-f]{64}$/)
  const sourcePackage = stringPattern(plan.source_package_sha256, '$.source_package_sha256', /^sha256:[0-9a-f]{64}$/)
  if (sourceRevision.slice(4) !== sourcePackage.slice(7)) invalidPlan('$.source_revision', 'source revision does not match package digest')
  const geometryDigest = stringPattern(plan.geometry_sha256, '$.geometry_sha256', /^sha256:[0-9a-f]{64}$/)
  const suppliedDigest = stringPattern(plan.decoration_sha256, '$.decoration_sha256', /^sha256:[0-9a-f]{64}$/)
  if (plan.coordinate_space !== 'viewport-local') invalidPlan('$.coordinate_space', 'coordinate space must be viewport-local')
  const bounds = validateRect(plan.bounds, '$.bounds', false)
  if (bounds.x_emu !== 0 || bounds.y_emu !== 0) invalidPlan('$.bounds', 'viewport-local bounds must start at the origin')
  if (!Array.isArray(plan.fills) || plan.fills.length > NATIVE_SHEET_DECORATION_V2_LIMITS.maxFills) invalidPlan('$.fills', 'fill inventory is missing or exceeds its resource bound')
  if (!Array.isArray(plan.border_segments) || plan.border_segments.length > NATIVE_SHEET_DECORATION_V2_LIMITS.maxBorderSegments) invalidPlan('$.border_segments', 'border inventory is missing or exceeds its resource bound')
  const fills: NativeSheetFillDecorationV2[] = []
  let priorCell = -1
  for (let index = 0; index < plan.fills.length; index++) {
    const value = exactObject(plan.fills[index], ['cell_ref', 'row', 'column', 'style_id', 'fill_id', 'fill_record_sha256', 'rect', 'color'], `$.fills[${index}]`)
    const row = boundedInteger(value.row, `$.fills[${index}].row`, 0, 1_048_575)
    const column = boundedInteger(value.column, `$.fills[${index}].column`, 0, 16_383)
    const position = row * 16_384 + column
    if (position <= priorCell) invalidPlan(`$.fills[${index}]`, 'fills must be unique and row-major ordered')
    priorCell = position
    const cellRef = stringPattern(value.cell_ref, `$.fills[${index}].cell_ref`, /^[A-Z]{1,3}[1-9][0-9]{0,6}$/)
    if (cellRef !== cellReference(row, column)) invalidPlan(`$.fills[${index}].cell_ref`, 'cell reference does not match coordinates')
    const rect = validateRect(value.rect, `$.fills[${index}].rect`, true)
    if (!rectContained(rect, bounds)) invalidPlan(`$.fills[${index}].rect`, 'fill rectangle escapes decoration bounds')
    fills.push({
      cell_ref: cellRef, row, column,
      style_id: boundedInteger(value.style_id, `$.fills[${index}].style_id`, 0, 4_294_967_295),
      fill_id: boundedInteger(value.fill_id, `$.fills[${index}].fill_id`, 0, 4_294_967_295),
      fill_record_sha256: stringPattern(value.fill_record_sha256, `$.fills[${index}].fill_record_sha256`, /^sha256:[0-9a-f]{64}$/),
      rect,
      color: stringPattern(value.color, `$.fills[${index}].color`, /^#[0-9A-F]{6}$/),
    })
  }
  const borderSegments: NativeSheetBorderSegmentV2[] = []
  const seenSegmentCoordinates = new Set<string>()
  const seenGlobalSources = new Set<string>()
  let priorSegment: NativeSheetBorderSegmentV2 | undefined
  let totalSources = 0
  for (let index = 0; index < plan.border_segments.length; index++) {
    const path = `$.border_segments[${index}]`
    const value = exactObject(plan.border_segments[index], ['orientation', 'x1_emu', 'y1_emu', 'x2_emu', 'y2_emu', 'border_style', 'color', 'sources'], path)
    if (value.orientation !== 'horizontal' && value.orientation !== 'vertical') invalidPlan(`${path}.orientation`, 'border orientation is invalid')
    const x1 = boundedInteger(value.x1_emu, `${path}.x1_emu`, 0, Number.MAX_SAFE_INTEGER)
    const y1 = boundedInteger(value.y1_emu, `${path}.y1_emu`, 0, Number.MAX_SAFE_INTEGER)
    const x2 = boundedInteger(value.x2_emu, `${path}.x2_emu`, 0, Number.MAX_SAFE_INTEGER)
    const y2 = boundedInteger(value.y2_emu, `${path}.y2_emu`, 0, Number.MAX_SAFE_INTEGER)
    if (value.orientation === 'horizontal' ? (y1 !== y2 || x2 <= x1) : (x1 !== x2 || y2 <= y1)) invalidPlan(path, 'border segment must be forward, nonzero, and axis-aligned')
    if (!pointContained(x1, y1, bounds) || !pointContained(x2, y2, bounds)) invalidPlan(path, 'border segment escapes decoration bounds')
    const coordinateKey = `${String(value.orientation)}:${x1}:${y1}:${x2}:${y2}`
    if (seenSegmentCoordinates.has(coordinateKey)) invalidPlan(path, 'border segment coordinates must be unique')
    seenSegmentCoordinates.add(coordinateKey)
    if (!supportedBorderStyles.has(value.border_style as NativeBorderStyle)) invalidPlan(`${path}.border_style`, 'border style token is unsupported')
    if (!Array.isArray(value.sources) || value.sources.length < 1 || value.sources.length > 2) invalidPlan(`${path}.sources`, 'border segment requires one or two bounded sources')
    totalSources += value.sources.length
    if (totalSources > NATIVE_SHEET_DECORATION_V2_LIMITS.maxBorderSegments * 2) invalidPlan(`${path}.sources`, 'border source inventory exceeds its resource bound')
    const sources: NativeSheetBorderSourceV2[] = []
    const seenSources = new Set<string>()
    for (let sourceIndex = 0; sourceIndex < value.sources.length; sourceIndex++) {
      const sourcePath = `${path}.sources[${sourceIndex}]`
      const source = exactObject(value.sources[sourceIndex], ['cell_ref', 'edge', 'style_id', 'border_id', 'border_record_sha256'], sourcePath)
      const cellRef = canonicalCellReference(source.cell_ref, `${sourcePath}.cell_ref`)
      if (!['left', 'right', 'top', 'bottom'].includes(source.edge as string)) invalidPlan(`${sourcePath}.edge`, 'border source edge is invalid')
      if (value.orientation === 'horizontal' ? !['top', 'bottom'].includes(source.edge as string) : !['left', 'right'].includes(source.edge as string)) invalidPlan(`${sourcePath}.edge`, 'border source edge contradicts segment orientation')
      const key = `${cellRef}:${String(source.edge)}`
      if (seenSources.has(key)) invalidPlan(sourcePath, 'border sources must be unique')
      seenSources.add(key)
      if (seenGlobalSources.has(key)) invalidPlan(sourcePath, 'cell edge source must occur in exactly one segment')
      seenGlobalSources.add(key)
      sources.push({
        cell_ref: cellRef,
        edge: source.edge as NativeBorderEdgeV2,
        style_id: boundedInteger(source.style_id, `${sourcePath}.style_id`, 0, 4_294_967_295),
        border_id: boundedInteger(source.border_id, `${sourcePath}.border_id`, 0, 4_294_967_295),
        border_record_sha256: stringPattern(source.border_record_sha256, `${sourcePath}.border_record_sha256`, /^sha256:[0-9a-f]{64}$/),
      })
    }
    if (sources.length === 2) {
      const edges = new Set(sources.map((source) => source.edge))
      const complementary = value.orientation === 'horizontal' ? edges.has('top') && edges.has('bottom') : edges.has('left') && edges.has('right')
      if (!complementary) invalidPlan(`${path}.sources`, 'two-source border joins require complementary cell edges')
      if (!areActuallyAdjacentOppositeSources(sources, value.orientation)) invalidPlan(`${path}.sources`, 'two-source border joins require actually adjacent cells on opposite edges')
    }
    const canonicalSources = [...sources].sort(compareBorderSources)
    if (sources.some((source, sourceIndex) => compareBorderSources(source, canonicalSources[sourceIndex]!) !== 0)) invalidPlan(`${path}.sources`, 'border sources must use canonical order')
    const segment: NativeSheetBorderSegmentV2 = {
      orientation: value.orientation,
      x1_emu: x1, y1_emu: y1, x2_emu: x2, y2_emu: y2,
      border_style: value.border_style as NativeBorderStyle,
      color: stringPattern(value.color, `${path}.color`, /^#[0-9A-F]{6}$/),
      sources,
    }
    if (priorSegment !== undefined && compareBorderSegments(priorSegment, segment) >= 0) invalidPlan(path, 'border segments must be unique and canonically ordered')
    priorSegment = segment
    borderSegments.push(segment)
  }
  if (3 + fills.length + borderSegments.length > NATIVE_SHEET_DECORATION_V2_LIMITS.maxCommands) invalidPlan('$', 'plan exceeds the decoration command bound')
  const unsigned: Omit<NativeSheetDecorationPlanV2, 'decoration_sha256'> = {
    protocol: NATIVE_SHEET_DECORATION_V2_PROTOCOL,
    version: NATIVE_SHEET_DECORATION_V2_VERSION,
    document_id: documentID,
    sheet_id: sheetID,
    source_part: sourcePart,
    source_revision: sourceRevision,
    source_package_sha256: sourcePackage,
    geometry_sha256: geometryDigest,
    coordinate_space: 'viewport-local' as const,
    bounds,
    fills,
    border_segments: borderSegments,
  }
  if (decorationDigest(unsigned) !== suppliedDigest) throw new NativeSheetDecorationError('decoration.planDigest', '$.decoration_sha256', 'decoration digest does not match the canonical plan')
  return deepFreeze({ ...unsigned, decoration_sha256: suppliedDigest as `sha256:${string}` })
}

function validateNativeSheetDecorationCommandsV2(input: unknown, maxCommands: number): ReadonlyArray<NativeSheetDecorationCommandV2> {
  input = snapshotDecorationInput(input, '$.commands')
  if (!Array.isArray(input) || input.length < 3 || input.length > maxCommands) throw new NativeSheetDecorationError('decoration.commandInvalid', '$.commands', 'command stream length is outside its bound')
  const begin = exactCommand(input[0], ['kind', 'protocol', 'version', 'document_id', 'sheet_id', 'source_part', 'source_revision', 'source_package_sha256', 'geometry_sha256', 'coordinate_space', 'decoration_sha256'], '$.commands[0]')
  if (begin.kind !== 'beginDecorations') commandInvalid('$.commands[0].kind', 'command stream must begin with beginDecorations')
  const clip = exactCommand(input[1], ['kind', 'rect'], '$.commands[1]')
  if (clip.kind !== 'clipRect') commandInvalid('$.commands[1].kind', 'second command must be clipRect')
  const end = exactCommand(input[input.length - 1], ['kind'], `$.commands[${input.length - 1}]`)
  if (end.kind !== 'endDecorations') commandInvalid(`$.commands[${input.length - 1}].kind`, 'command stream must end with endDecorations')
  const fills: unknown[] = [], borders: unknown[] = []
  let borderLane = false
  for (let index = 2; index < input.length - 1; index++) {
    const command = exactCommand(input[index], ['kind', 'decoration'], `$.commands[${index}]`)
    if (command.kind === 'fillRect' && !borderLane) fills.push(command.decoration)
    else if (command.kind === 'borderSegment') { borderLane = true; borders.push(command.decoration) }
    else commandInvalid(`$.commands[${index}].kind`, 'commands must contain fills followed by border segments')
  }
  const plan = validateNativeSheetDecorationPlanV2({
    protocol: begin.protocol,
    version: begin.version,
    decoration_sha256: begin.decoration_sha256,
    document_id: begin.document_id,
    sheet_id: begin.sheet_id,
    source_part: begin.source_part,
    source_revision: begin.source_revision,
    source_package_sha256: begin.source_package_sha256,
    geometry_sha256: begin.geometry_sha256,
    coordinate_space: begin.coordinate_space,
    bounds: clip.rect,
    fills,
    border_segments: borders,
  })
  return decorationCommands(plan)
}

function decorationCommands(plan: NativeSheetDecorationPlanV2): ReadonlyArray<NativeSheetDecorationCommandV2> {
  return deepFreeze([
    {
      kind: 'beginDecorations' as const,
      protocol: plan.protocol,
      version: plan.version,
      document_id: plan.document_id,
      sheet_id: plan.sheet_id,
      source_part: plan.source_part,
      source_revision: plan.source_revision,
      source_package_sha256: plan.source_package_sha256,
      geometry_sha256: plan.geometry_sha256,
      coordinate_space: plan.coordinate_space,
      decoration_sha256: plan.decoration_sha256,
    },
    { kind: 'clipRect' as const, rect: plan.bounds },
    ...plan.fills.map((decoration) => ({ kind: 'fillRect' as const, decoration })),
    ...plan.border_segments.map((decoration) => ({ kind: 'borderSegment' as const, decoration })),
    { kind: 'endDecorations' as const },
  ])
}

function validateSourceAuthority(workbook: NativeWorkbookRenderModelV2, sheet: NativeSheetRenderModelV2, geometry: NativeSheetGeometryV2): void {
  if (sheet.mutation_authority.source_revision !== workbook.revision || typeof sheet.mutation_authority.source_part !== 'string' || sheet.mutation_authority.source_part.length === 0) throw new NativeSheetDecorationError('decoration.sourceUnsupported', '$.sheet.mutation_authority', 'sheet source authority is not bound to this workbook revision')
  if (sheet.merged_ranges.length !== 0 || geometry.merged_ranges.length !== 0) throw new NativeSheetDecorationError('decoration.mergeUnsupported', '$.sheet.merged_ranges', 'border joins across merged cells are not projected by native v2')
  if (sheet.rows.some((row) => row.style_id !== undefined) || sheet.columns.some((column) => column.style_id !== undefined)) {
    throw new NativeSheetDecorationError('decoration.stylePrecedence', '$.sheet', 'row and column style precedence is not projected by native v2')
  }
  const issue = workbook.unsupported.find((item) => {
    if (item.scope_id !== 'workbook' && item.scope_id !== `sheet:${sheet.id}`) return false
    return appearanceAuthorityCapabilities.has(item.capability) || appearanceAuthorityCodes.has(item.code)
  })
  if (issue) throw new NativeSheetDecorationError('decoration.sourceUnsupported', '$.unsupported', `${issue.code} can change source-authoritative sheet appearance`)
}

function validateGeometryAuthority(workbook: NativeWorkbookRenderModelV2, geometry: NativeSheetGeometryV2): void {
  if (geometry.document_id !== workbook.document_id || geometry.source_revision !== workbook.revision || geometry.source_package_sha256 !== workbook.source.package_sha256 || geometry.coordinate_space !== 'viewport-local') {
    throw new NativeSheetDecorationError('decoration.geometryAuthority', '$.geometry', 'geometry is not bound to this exact workbook revision')
  }
  try {
    if (!isCompiledNativeSheetGeometryV2(geometry)) throw new Error('geometry is not branded')
    const canonical = validateNativeSheetGeometryV2(geometry)
    if (JSON.stringify(canonical) !== JSON.stringify(geometry)) throw new Error('geometry is not canonical')
  } catch {
    throw new NativeSheetDecorationError('decoration.geometryAuthority', '$.geometry', 'geometry is not the deterministic projection of this workbook and metric authority')
  }
}

function requireDecorationStyle(workbook: NativeWorkbookRenderModelV2, styleID: number): NativeRenderStyleV2 {
  const style = workbook.styles[styleID]
  const contradictoryDiagnostic = workbook.unsupported.some((item) => item.scope_id === `style:${styleID}` && (item.code === 'STYLE_FILL' || item.code === 'STYLE_BORDER'))
  if (!style || style.id !== styleID || style.provenance.style_id !== styleID || style.provenance.source_revision !== workbook.revision || style.provenance.source_package_sha256 !== workbook.source.package_sha256 || style.provenance.projection !== style.effective.projection || style.provenance.raw_projection_sha256 !== nativeWorkbookStyleRawProjectionSha256V2(style.effective) || style.effective.unsupported.includes('fill') || style.effective.unsupported.includes('border') || style.effective.fill === undefined || style.effective.border === undefined || contradictoryDiagnostic) {
    throw new NativeSheetDecorationError('decoration.styleUnsupported', `$.styles[${styleID}]`, 'effective style or provenance is incomplete for exact decorations')
  }
  const fill = style.effective.fill
  if (fill.origin === 'implicit-default') {
    if (fill.fill_id !== undefined || fill.record_sha256 !== undefined || fill.color !== undefined || style.effective.fill_color !== undefined) throw new NativeSheetDecorationError('decoration.styleUnsupported', `$.styles[${styleID}].effective.fill`, 'implicit default fill carries forged styles-table authority')
  } else if (fill.origin !== 'styles-record' || !Number.isSafeInteger(fill.fill_id) || fill.fill_id! < 0 || !/^sha256:[0-9a-f]{64}$/.test(fill.record_sha256 ?? '') || fill.color !== style.effective.fill_color || (fill.color !== undefined && !/^#[0-9A-F]{6}$/.test(fill.color))) {
    throw new NativeSheetDecorationError('decoration.styleUnsupported', `$.styles[${styleID}].effective.fill`, 'fill projection is not canonical direct-RGB styles-table authority')
  }
  const border = style.effective.border
  const sides = [border.left, border.right, border.top, border.bottom]
  if (border.origin === 'implicit-default') {
    if (border.border_id !== undefined || border.record_sha256 !== undefined || sides.some((side) => side !== undefined)) throw new NativeSheetDecorationError('decoration.styleUnsupported', `$.styles[${styleID}].effective.border`, 'implicit default border carries forged styles-table authority')
  } else if (border.origin !== 'styles-record' || !Number.isSafeInteger(border.border_id) || border.border_id! < 0 || !/^sha256:[0-9a-f]{64}$/.test(border.record_sha256 ?? '') || sides.some((side) => side !== undefined && (!supportedBorderStyles.has(side.style) || !/^#[0-9A-F]{6}$/.test(side.color)))) {
    throw new NativeSheetDecorationError('decoration.styleUnsupported', `$.styles[${styleID}].effective.border`, 'border projection is not canonical direct-RGB styles-table authority')
  }
  return style
}

const appearanceAuthorityCapabilities = new Set(['conditional-formatting', 'tables', 'drawings', 'external-links', 'formula-groups'])
const appearanceAuthorityCodes = new Set([
  'FOREIGN_WORKSHEET_MARKUP', 'WORKSHEET_ATTRIBUTES', 'SHEET_VIEW_GEOMETRY', 'WORKSHEET_EXTENSIONS', 'STYLE_RECORD_ATTRIBUTES', 'STYLE_XF_OPAQUE_CONTENT',
])

const supportedBorderStyles = new Set<NativeBorderStyle>(['dashDot', 'dashDotDot', 'dashed', 'dotted', 'double', 'hair', 'medium', 'mediumDashDot', 'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin'])

function styleHasDecoration(style: NativeRenderStyleV2): boolean {
  const border = style.effective.border
  return style.effective.fill_color !== undefined || border?.left !== undefined || border?.right !== undefined || border?.top !== undefined || border?.bottom !== undefined
}

function addBorder(segments: Map<string, BorderCandidate>, side: NativeWorkbookBorderSideV2 | undefined, edge: NativeBorderEdgeV2, x1: number, y1: number, x2: number, y2: number, style: NativeRenderStyleV2, cellRef: string): void {
  if (!side) return
  const border = style.effective.border!
  if (border.origin !== 'styles-record' || border.border_id === undefined || border.record_sha256 === undefined) throw new NativeSheetDecorationError('decoration.styleUnsupported', `$.styles[${style.id}].effective.border`, 'border side lacks exact styles-table provenance')
  const orientation = y1 === y2 ? 'horizontal' : 'vertical'
  const key = `${orientation}:${x1}:${y1}:${x2}:${y2}`
  const source = { cell_ref: cellRef, edge, style_id: style.id, border_id: border.border_id, border_record_sha256: border.record_sha256 }
  const existing = segments.get(key)
  if (existing) {
    if (existing.border_style !== side.style || existing.color !== side.color) throw new NativeSheetDecorationError('decoration.borderConflict', '$.border_segments', `conflicting OOXML borders meet at ${key}`)
    if (existing.sources.some((candidate) => candidate.cell_ref === source.cell_ref && candidate.edge === source.edge) || existing.sources.length >= 2) throw new NativeSheetDecorationError('decoration.borderConflict', '$.border_segments', `non-canonical border topology meets at ${key}`)
    existing.sources.push(source)
    return
  }
  segments.set(key, { orientation, x1_emu: x1, y1_emu: y1, x2_emu: x2, y2_emu: y2, border_style: side.style, color: side.color, sources: [source] })
}

function compareBorderSources(left: NativeSheetBorderSourceV2, right: NativeSheetBorderSourceV2): number {
  return compareASCII(left.cell_ref, right.cell_ref)
    || compareASCII(left.edge, right.edge)
    || compareNumber(left.style_id, right.style_id)
    || compareNumber(left.border_id, right.border_id)
    || compareASCII(left.border_record_sha256, right.border_record_sha256)
}

function compareBorderSegments(left: NativeSheetBorderSegmentV2, right: NativeSheetBorderSegmentV2): number {
  return compareASCII(left.orientation, right.orientation)
    || compareNumber(left.x1_emu, right.x1_emu)
    || compareNumber(left.y1_emu, right.y1_emu)
    || compareNumber(left.x2_emu, right.x2_emu)
    || compareNumber(left.y2_emu, right.y2_emu)
    || compareASCII(left.border_style, right.border_style)
    || compareASCII(left.color, right.color)
}

function compareASCII(left: string, right: string): number { return left < right ? -1 : left > right ? 1 : 0 }
function compareNumber(left: number, right: number): number { return left < right ? -1 : left > right ? 1 : 0 }

type UnknownRecord = Record<string, unknown>

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
    if (error instanceof NativeSheetDecorationError) throw new NativeSheetDecorationError('decoration.commandInvalid', error.path, error.message)
    throw error
  }
}

function invalidPlan(path: string, message: string): never { throw new NativeSheetDecorationError('decoration.planInvalid', path, message) }
function commandInvalid(path: string, message: string): never { throw new NativeSheetDecorationError('decoration.commandInvalid', path, message) }

function stringPattern(value: unknown, path: string, pattern: RegExp): string {
  if (typeof value !== 'string' || !pattern.test(value)) invalidPlan(path, 'string is missing or non-canonical')
  return value
}

function identifier(value: unknown, path: string, maximum: number): string {
  if (typeof value !== 'string' || value.length < 1 || value.length > maximum || !/^[A-Za-z0-9][A-Za-z0-9._:-]*$/.test(value)) invalidPlan(path, 'identifier is missing, oversized, or non-canonical')
  return value
}

function canonicalSheetID(value: unknown, path: string): string {
  const result = stringPattern(value, path, /^[1-9][0-9]{0,9}$/)
  const numeric = Number(result)
  if (!Number.isSafeInteger(numeric) || numeric > 4_294_967_295) invalidPlan(path, 'sheet id is outside uint32 bounds')
  return result
}

function canonicalCellReference(value: unknown, path: string): string {
  const result = stringPattern(value, path, /^([A-Z]{1,3})([1-9][0-9]{0,6})$/)
  const match = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(result)!
  let column = 0
  for (const character of match[1]!) column = column * 26 + character.charCodeAt(0) - 64
  const row = Number(match[2])
  if (column < 1 || column > 16_384 || row < 1 || row > 1_048_576) invalidPlan(path, 'cell reference is outside Excel bounds')
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

function pointContained(x: number, y: number, bounds: NativeSheetGeometryRectV2): boolean {
  return x >= bounds.x_emu && y >= bounds.y_emu && x <= bounds.x_emu + bounds.width_emu && y <= bounds.y_emu + bounds.height_emu
}

function decorationDigest(unsigned: Omit<NativeSheetDecorationPlanV2, 'decoration_sha256'>): `sha256:${string}` {
  return `sha256:${sha256Hex(`injoffice.xlsx.sheet-decoration.v1\0${JSON.stringify(unsigned)}`)}`
}

function validateCommandLimit(value: unknown, path: string): asserts value is number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || Object.is(value, -0) || value < 1 || value > NATIVE_SHEET_DECORATION_V2_LIMITS.maxCommands) throw new NativeSheetDecorationError('decoration.invalidLimit', path, `command limit must be from 1 through ${NATIVE_SHEET_DECORATION_V2_LIMITS.maxCommands}`)
}

function snapshotDecorationInput(value: unknown, path: string): unknown {
  try { return snapshotNativePlainData(value, { maxDepth: 32, maxNodes: NATIVE_SHEET_DECORATION_V2_LIMITS.maxCommands * 16 }) }
  catch (error) {
    if (error instanceof NativePlainDataError) throw new NativeSheetDecorationError('decoration.planInvalid', error.path || path, error.message)
    throw error
  }
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

function areActuallyAdjacentOppositeSources(sources: ReadonlyArray<NativeSheetBorderSourceV2>, orientation: 'horizontal' | 'vertical'): boolean {
  const parsed = sources.map((source) => ({ ...parseCanonicalCellReference(source.cell_ref), edge: source.edge }))
  const first = parsed[0]!, second = parsed[1]!
  if (orientation === 'vertical') {
    const right = first.edge === 'right' ? first : second.edge === 'right' ? second : undefined
    const left = first.edge === 'left' ? first : second.edge === 'left' ? second : undefined
    return right !== undefined && left !== undefined && right.row === left.row && right.column + 1 === left.column
  }
  const bottom = first.edge === 'bottom' ? first : second.edge === 'bottom' ? second : undefined
  const top = first.edge === 'top' ? first : second.edge === 'top' ? second : undefined
  return bottom !== undefined && top !== undefined && bottom.column === top.column && bottom.row + 1 === top.row
}

function parseCanonicalCellReference(value: string): { row: number; column: number } {
  const match = /^([A-Z]+)([1-9][0-9]*)$/.exec(value)!
  let column = 0
  for (const character of match[1]!) column = column * 26 + character.charCodeAt(0) - 64
  return { row: Number(match[2]) - 1, column: column - 1 }
}

function cellReference(row: number, column: number): string {
  let value = column + 1, name = ''
  while (value > 0) { value--; name = String.fromCharCode(65 + value % 26) + name; value = Math.floor(value / 26) }
  return `${name}${row + 1}`
}

function deepFreeze<T>(value: T): T {
  if (typeof value !== 'object' || value === null || Object.isFrozen(value)) return value
  for (const child of Object.values(value)) deepFreeze(child)
  return Object.freeze(value)
}
