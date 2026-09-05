import type { NativeBorderStyle, NativeWorkbookBorderSideV1 } from './nativeContract.generated.js'
import { isProjectedNativeWorkbookV1 } from './nativeRenderModel.js'
import type { NativeRenderCellV1, NativeRenderStyleV1, NativeSheetRenderModelV1, NativeWorkbookRenderModelV1 } from './nativeRenderModel.js'
import { compileNativeSheetGeometryV1 } from './nativeSheetGeometry.js'
import type { NativeSheetGeometryRectV1, NativeSheetGeometryV1 } from './nativeSheetGeometry.js'
import { sha256Hex } from './nativeSha256.js'
import { nativeWorkbookStyleRawProjectionSha256V1 } from './nativeValidation.js'

export const NATIVE_SHEET_DECORATION_PROTOCOL = 'injoffice.xlsx.sheet-decoration'
export const NATIVE_SHEET_DECORATION_VERSION = 1 as const
export const NATIVE_SHEET_DECORATION_LIMITS = Object.freeze({
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

export interface NativeSheetFillDecorationV1 {
  readonly cell_ref: string
  readonly row: number
  readonly column: number
  readonly style_id: number
  readonly fill_id: number
  readonly fill_record_sha256: string
  readonly rect: NativeSheetGeometryRectV1
  readonly color: string
}

export type NativeBorderEdgeV1 = 'left' | 'right' | 'top' | 'bottom'

export interface NativeSheetBorderSourceV1 {
  readonly cell_ref: string
  readonly edge: NativeBorderEdgeV1
  readonly style_id: number
  readonly border_id: number
  readonly border_record_sha256: string
}

export interface NativeSheetBorderSegmentV1 {
  readonly orientation: 'horizontal' | 'vertical'
  readonly x1_emu: number
  readonly y1_emu: number
  readonly x2_emu: number
  readonly y2_emu: number
  /** Exact OOXML token. It deliberately does not invent a physical stroke width. */
  readonly border_style: NativeBorderStyle
  readonly color: string
  readonly sources: ReadonlyArray<NativeSheetBorderSourceV1>
}

export interface NativeSheetDecorationPlanV1 {
  readonly protocol: typeof NATIVE_SHEET_DECORATION_PROTOCOL
  readonly version: typeof NATIVE_SHEET_DECORATION_VERSION
  readonly decoration_sha256: `sha256:${string}`
  readonly document_id: string
  readonly sheet_id: string
  readonly source_part: string
  readonly source_revision: string
  readonly source_package_sha256: string
  readonly geometry_sha256: string
  readonly coordinate_space: 'viewport-local'
  readonly bounds: NativeSheetGeometryRectV1
  readonly fills: ReadonlyArray<NativeSheetFillDecorationV1>
  readonly border_segments: ReadonlyArray<NativeSheetBorderSegmentV1>
}

export type NativeSheetDecorationCommandV1 =
  | { readonly kind: 'beginDecorations'; readonly protocol: typeof NATIVE_SHEET_DECORATION_PROTOCOL; readonly version: typeof NATIVE_SHEET_DECORATION_VERSION; readonly document_id: string; readonly sheet_id: string; readonly source_part: string; readonly source_revision: string; readonly source_package_sha256: string; readonly geometry_sha256: string; readonly coordinate_space: 'viewport-local'; readonly decoration_sha256: string }
  | { readonly kind: 'clipRect'; readonly rect: NativeSheetGeometryRectV1 }
  | { readonly kind: 'fillRect'; readonly decoration: NativeSheetFillDecorationV1 }
  | { readonly kind: 'borderSegment'; readonly decoration: NativeSheetBorderSegmentV1 }
  | { readonly kind: 'endDecorations' }

export interface NativeSheetDecorationSurfaceV1 {
  readonly command_capacity?: number
  push(command: NativeSheetDecorationCommandV1): void
}

export interface NativeSheetDecorationRecordingSurfaceV1 extends NativeSheetDecorationSurfaceV1 {
  readonly commands: ReadonlyArray<NativeSheetDecorationCommandV1>
  finish(): ReadonlyArray<NativeSheetDecorationCommandV1>
}

export interface NativeSheetDecorationCommandAdapterV1<HostContext> {
  execute(context: HostContext, command: NativeSheetDecorationCommandV1): void
}

type BorderCandidate = Omit<NativeSheetBorderSegmentV1, 'sources'> & { sources: NativeSheetBorderSourceV1[] }

export function compileNativeSheetDecorationsV1(workbook: NativeWorkbookRenderModelV1, geometry: NativeSheetGeometryV1): NativeSheetDecorationPlanV1 {
  if (!isProjectedNativeWorkbookV1(workbook)) throw new NativeSheetDecorationError('decoration.modelAuthority', '$.workbook', 'workbook must be the branded frozen result of projectNativeWorkbookV1')
  validateGeometryAuthority(workbook, geometry)
  const sheet = workbook.sheets.find((candidate) => candidate.id === geometry.sheet_id)
  if (!sheet) throw new NativeSheetDecorationError('decoration.sheetMissing', '$.geometry.sheet_id', `sheet ${JSON.stringify(geometry.sheet_id)} is absent`)
  validateSourceAuthority(workbook, sheet, geometry)

  const cells = new Map<number, NativeRenderCellV1>()
  for (const cell of sheet.cells) {
    const key = cell.row * 16_384 + cell.column
    if (!Number.isSafeInteger(cell.row) || !Number.isSafeInteger(cell.column) || cell.row < 0 || cell.row >= 1_048_576 || cell.column < 0 || cell.column >= 16_384 || cell.ref !== cellReference(cell.row, cell.column) || cells.has(key)) {
      throw new NativeSheetDecorationError('decoration.sourceUnsupported', '$.sheet.cells', 'cell coordinates or identity are non-canonical')
    }
    cells.set(key, cell)
  }
  const fills: NativeSheetFillDecorationV1[] = []
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
        if (fills.length >= NATIVE_SHEET_DECORATION_LIMITS.maxFills) throw new NativeSheetDecorationError('decoration.resourceBudget', '$.fills', 'fill decorations exceed the native resource bound')
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
  if (segments.size > NATIVE_SHEET_DECORATION_LIMITS.maxBorderSegments) throw new NativeSheetDecorationError('decoration.resourceBudget', '$.border_segments', 'border segments exceed the native resource bound')
  const borderSegments = [...segments.values()]
    .map((segment) => ({ ...segment, sources: [...segment.sources].sort(compareBorderSources) }))
    .sort(compareBorderSegments)
  const unsigned: Omit<NativeSheetDecorationPlanV1, 'decoration_sha256'> = {
    protocol: NATIVE_SHEET_DECORATION_PROTOCOL,
    version: NATIVE_SHEET_DECORATION_VERSION,
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

export function createNativeSheetDecorationRecordingSurfaceV1(maxCommands: number = NATIVE_SHEET_DECORATION_LIMITS.maxCommands): NativeSheetDecorationRecordingSurfaceV1 {
  validateCommandLimit(maxCommands, '$.maxCommands')
  const commands: NativeSheetDecorationCommandV1[] = []
  let finished: ReadonlyArray<NativeSheetDecorationCommandV1> | undefined
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

export function emitNativeSheetDecorationCommandsV1(plan: NativeSheetDecorationPlanV1, surface: NativeSheetDecorationSurfaceV1, maxCommands: number = NATIVE_SHEET_DECORATION_LIMITS.maxCommands): void {
  const validated = validateNativeSheetDecorationPlanV1(plan)
  validateCommandLimit(maxCommands, '$.maxCommands')
  if ((typeof surface !== 'object' && typeof surface !== 'function') || surface === null) throw new NativeSheetDecorationError('decoration.commandInvalid', '$.surface', 'surface must be an object')
  const surfaceCapacity = surface.command_capacity
  const surfacePush = surface.push
  if (typeof surfacePush !== 'function') throw new NativeSheetDecorationError('decoration.commandInvalid', '$.surface.push', 'surface push must be a function')
  if (surfaceCapacity !== undefined) validateCommandLimit(surfaceCapacity, '$.surface.command_capacity')
  const capacity = Math.min(maxCommands, surfaceCapacity ?? maxCommands)
  const commands = decorationCommands(validated)
  const required = commands.length
  if (required > capacity) throw new NativeSheetDecorationError('decoration.commandBudget', '$.commands', `decoration commands require ${required}, exceeding ${capacity}`)
  for (const command of commands) surfacePush.call(surface, command)
}

export function replayNativeSheetDecorationCommandsV1<HostContext>(context: HostContext, commands: ReadonlyArray<NativeSheetDecorationCommandV1>, adapter: NativeSheetDecorationCommandAdapterV1<HostContext>, maxCommands: number = NATIVE_SHEET_DECORATION_LIMITS.maxCommands): void {
  validateCommandLimit(maxCommands, '$.maxCommands')
  const validated = validateNativeSheetDecorationCommandsV1(commands, maxCommands)
  if ((typeof adapter !== 'object' && typeof adapter !== 'function') || adapter === null) throw new NativeSheetDecorationError('decoration.commandInvalid', '$.adapter', 'adapter must be an object')
  const execute = adapter.execute
  if (typeof execute !== 'function') throw new NativeSheetDecorationError('decoration.commandInvalid', '$.adapter.execute', 'adapter execute must be a function')
  for (const command of validated) execute.call(adapter, context, command)
}

/** Rebuilds an immutable canonical plan and verifies every untrusted field. */
export function validateNativeSheetDecorationPlanV1(input: unknown): NativeSheetDecorationPlanV1 {
  const plan = exactObject(input, ['protocol', 'version', 'decoration_sha256', 'document_id', 'sheet_id', 'source_part', 'source_revision', 'source_package_sha256', 'geometry_sha256', 'coordinate_space', 'bounds', 'fills', 'border_segments'], '$')
  if (plan.protocol !== NATIVE_SHEET_DECORATION_PROTOCOL || plan.version !== NATIVE_SHEET_DECORATION_VERSION) invalidPlan('$', 'decoration protocol or version is unsupported')
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
  if (!Array.isArray(plan.fills) || plan.fills.length > NATIVE_SHEET_DECORATION_LIMITS.maxFills) invalidPlan('$.fills', 'fill inventory is missing or exceeds its resource bound')
  if (!Array.isArray(plan.border_segments) || plan.border_segments.length > NATIVE_SHEET_DECORATION_LIMITS.maxBorderSegments) invalidPlan('$.border_segments', 'border inventory is missing or exceeds its resource bound')
  const fills: NativeSheetFillDecorationV1[] = []
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
  const borderSegments: NativeSheetBorderSegmentV1[] = []
  const seenSegmentCoordinates = new Set<string>()
  const seenBorderSources = new Set<string>()
  let priorSegment: NativeSheetBorderSegmentV1 | undefined
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
    if (totalSources > NATIVE_SHEET_DECORATION_LIMITS.maxBorderSegments * 2) invalidPlan(`${path}.sources`, 'border source inventory exceeds its resource bound')
    const sources: NativeSheetBorderSourceV1[] = []
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
      if (seenBorderSources.has(key)) invalidPlan(sourcePath, 'a cell edge cannot source more than one border segment')
      seenBorderSources.add(key)
      sources.push({
        cell_ref: cellRef,
        edge: source.edge as NativeBorderEdgeV1,
        style_id: boundedInteger(source.style_id, `${sourcePath}.style_id`, 0, 4_294_967_295),
        border_id: boundedInteger(source.border_id, `${sourcePath}.border_id`, 0, 4_294_967_295),
        border_record_sha256: stringPattern(source.border_record_sha256, `${sourcePath}.border_record_sha256`, /^sha256:[0-9a-f]{64}$/),
      })
    }
    if (sources.length === 2) {
      const edges = new Set(sources.map((source) => source.edge))
      const complementary = value.orientation === 'horizontal' ? edges.has('top') && edges.has('bottom') : edges.has('left') && edges.has('right')
      if (!complementary) invalidPlan(`${path}.sources`, 'two-source border joins require complementary cell edges')
      if (value.orientation === 'horizontal') {
        const upper = sources.find((source) => source.edge === 'bottom')!
        const lower = sources.find((source) => source.edge === 'top')!
        const upperCell = cellReferenceCoordinates(upper.cell_ref), lowerCell = cellReferenceCoordinates(lower.cell_ref)
        if (upperCell.column !== lowerCell.column || upperCell.row + 1 !== lowerCell.row) invalidPlan(`${path}.sources`, 'horizontal joins require vertically adjacent cells with bottom/top edges')
      } else {
        const left = sources.find((source) => source.edge === 'right')!
        const right = sources.find((source) => source.edge === 'left')!
        const leftCell = cellReferenceCoordinates(left.cell_ref), rightCell = cellReferenceCoordinates(right.cell_ref)
        if (leftCell.row !== rightCell.row || leftCell.column + 1 !== rightCell.column) invalidPlan(`${path}.sources`, 'vertical joins require horizontally adjacent cells with right/left edges')
      }
    }
    const canonicalSources = [...sources].sort(compareBorderSources)
    if (sources.some((source, sourceIndex) => compareBorderSources(source, canonicalSources[sourceIndex]!) !== 0)) invalidPlan(`${path}.sources`, 'border sources must use canonical order')
    const segment: NativeSheetBorderSegmentV1 = {
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
  if (3 + fills.length + borderSegments.length > NATIVE_SHEET_DECORATION_LIMITS.maxCommands) invalidPlan('$', 'plan exceeds the decoration command bound')
  const unsigned: Omit<NativeSheetDecorationPlanV1, 'decoration_sha256'> = {
    protocol: NATIVE_SHEET_DECORATION_PROTOCOL,
    version: NATIVE_SHEET_DECORATION_VERSION,
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

function validateNativeSheetDecorationCommandsV1(input: unknown, maxCommands: number): ReadonlyArray<NativeSheetDecorationCommandV1> {
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
  const plan = validateNativeSheetDecorationPlanV1({
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

function decorationCommands(plan: NativeSheetDecorationPlanV1): ReadonlyArray<NativeSheetDecorationCommandV1> {
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

function validateSourceAuthority(workbook: NativeWorkbookRenderModelV1, sheet: NativeSheetRenderModelV1, geometry: NativeSheetGeometryV1): void {
  if (sheet.mutation_authority.source_revision !== workbook.revision || typeof sheet.mutation_authority.source_part !== 'string' || sheet.mutation_authority.source_part.length === 0) throw new NativeSheetDecorationError('decoration.sourceUnsupported', '$.sheet.mutation_authority', 'sheet source authority is not bound to this workbook revision')
  if (sheet.merged_ranges.length !== 0 || geometry.merged_ranges.length !== 0) throw new NativeSheetDecorationError('decoration.mergeUnsupported', '$.sheet.merged_ranges', 'border joins across merged cells are not projected by native v1')
  if (sheet.rows.some((row) => row.style_id !== undefined) || sheet.columns.some((column) => column.style_id !== undefined)) {
    throw new NativeSheetDecorationError('decoration.stylePrecedence', '$.sheet', 'row and column style precedence is not projected by native v1')
  }
  const issue = workbook.unsupported.find((item) => {
    if (item.scope_id !== 'workbook' && item.scope_id !== `sheet:${sheet.id}`) return false
    return appearanceAuthorityCapabilities.has(item.capability) || appearanceAuthorityCodes.has(item.code)
  })
  if (issue) throw new NativeSheetDecorationError('decoration.sourceUnsupported', '$.unsupported', `${issue.code} can change source-authoritative sheet appearance`)
}

function validateGeometryAuthority(workbook: NativeWorkbookRenderModelV1, geometry: NativeSheetGeometryV1): void {
  if (geometry.document_id !== workbook.document_id || geometry.source_revision !== workbook.revision || geometry.source_package_sha256 !== workbook.source.package_sha256 || geometry.coordinate_space !== 'viewport-local') {
    throw new NativeSheetDecorationError('decoration.geometryAuthority', '$.geometry', 'geometry is not bound to this exact workbook revision')
  }
  try {
    const rebuilt = compileNativeSheetGeometryV1(workbook, geometry.sheet_id, geometry.viewport, geometry.metric_authority)
    if (JSON.stringify(rebuilt) !== JSON.stringify(geometry)) throw new Error('geometry does not match deterministic recompilation')
  } catch {
    throw new NativeSheetDecorationError('decoration.geometryAuthority', '$.geometry', 'geometry is not the deterministic projection of this workbook and metric authority')
  }
}

function requireDecorationStyle(workbook: NativeWorkbookRenderModelV1, styleID: number): NativeRenderStyleV1 {
  const style = workbook.styles[styleID]
  const contradictoryDiagnostic = workbook.unsupported.some((item) => item.scope_id === `style:${styleID}` && (item.code === 'STYLE_FILL' || item.code === 'STYLE_BORDER'))
  if (!style || style.id !== styleID || style.provenance.style_id !== styleID || style.provenance.source_revision !== workbook.revision || style.provenance.source_package_sha256 !== workbook.source.package_sha256 || style.provenance.projection !== style.effective.projection || style.provenance.raw_projection_sha256 !== nativeWorkbookStyleRawProjectionSha256V1(style.effective) || style.effective.unsupported.includes('fill') || style.effective.unsupported.includes('border') || style.effective.fill === undefined || style.effective.border === undefined || contradictoryDiagnostic) {
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

function styleHasDecoration(style: NativeRenderStyleV1): boolean {
  const border = style.effective.border
  return style.effective.fill_color !== undefined || border?.left !== undefined || border?.right !== undefined || border?.top !== undefined || border?.bottom !== undefined
}

function addBorder(segments: Map<string, BorderCandidate>, side: NativeWorkbookBorderSideV1 | undefined, edge: NativeBorderEdgeV1, x1: number, y1: number, x2: number, y2: number, style: NativeRenderStyleV1, cellRef: string): void {
  if (!side) return
  const border = style.effective.border!
  if (border.origin !== 'styles-record' || border.border_id === undefined || border.record_sha256 === undefined) throw new NativeSheetDecorationError('decoration.styleUnsupported', `$.styles[${style.id}].effective.border`, 'border side lacks exact styles-table provenance')
  const orientation = y1 === y2 ? 'horizontal' : 'vertical'
  const key = `${orientation}:${x1}:${y1}:${x2}:${y2}`
  const source = { cell_ref: cellRef, edge, style_id: style.id, border_id: border.border_id, border_record_sha256: border.record_sha256 }
  const existing = segments.get(key)
  if (existing) {
    if (existing.border_style !== side.style || existing.color !== side.color) throw new NativeSheetDecorationError('decoration.borderConflict', '$.border_segments', `conflicting OOXML borders meet at ${key}`)
    existing.sources.push(source)
    return
  }
  segments.set(key, { orientation, x1_emu: x1, y1_emu: y1, x2_emu: x2, y2_emu: y2, border_style: side.style, color: side.color, sources: [source] })
}

function compareBorderSources(left: NativeSheetBorderSourceV1, right: NativeSheetBorderSourceV1): number {
  return compareASCII(left.cell_ref, right.cell_ref)
    || compareASCII(left.edge, right.edge)
    || left.style_id - right.style_id
    || left.border_id - right.border_id
    || compareASCII(left.border_record_sha256, right.border_record_sha256)
}

function compareBorderSegments(left: NativeSheetBorderSegmentV1, right: NativeSheetBorderSegmentV1): number {
  return compareASCII(left.orientation, right.orientation)
    || left.x1_emu - right.x1_emu
    || left.y1_emu - right.y1_emu
    || left.x2_emu - right.x2_emu
    || left.y2_emu - right.y2_emu
    || compareASCII(left.border_style, right.border_style)
    || compareASCII(left.color, right.color)
}

function compareASCII(left: string, right: string): number { return left < right ? -1 : left > right ? 1 : 0 }

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

function cellReferenceCoordinates(ref: string): { row: number; column: number } {
  const match = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(ref)!
  let column = 0
  for (const character of match[1]!) column = column * 26 + character.charCodeAt(0) - 64
  return { row: Number(match[2]) - 1, column: column - 1 }
}

function canonicalPart(value: unknown, path: string): string {
  if (typeof value !== 'string' || value.length < 1 || value.length > 8_192 || /[?#\\\u0000-\u001f\u007f]/.test(value) || /^[A-Za-z][A-Za-z0-9+.-]*:/.test(value) || /%(?:2f|5c)/i.test(value)) invalidPlan(path, 'source part is not a canonical OPC part name')
  let decoded: string
  try { decoded = decodeURIComponent(value) } catch { invalidPlan(path, 'source part has a malformed percent escape') }
  decoded = decoded.startsWith('/') ? decoded.slice(1) : decoded
  const segments = decoded.split('/')
  if (!decoded || decoded.startsWith('/') || decoded.endsWith('/') || decoded.includes('//') || segments.some((segment) => !segment || segment === '.' || segment === '..' || segment.endsWith('.') || /[\\\u0000-\u001f\u007f]/.test(segment))) invalidPlan(path, 'source part is not a canonical OPC part name')
  return value
}

function boundedInteger(value: unknown, path: string, minimum: number, maximum: number): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < minimum || value > maximum) invalidPlan(path, `integer must be within ${minimum}..${maximum}`)
  return value
}

function validateRect(value: unknown, path: string, requireArea: boolean): NativeSheetGeometryRectV1 {
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

function rectContained(rect: NativeSheetGeometryRectV1, bounds: NativeSheetGeometryRectV1): boolean {
  return rect.x_emu >= bounds.x_emu && rect.y_emu >= bounds.y_emu && rect.x_emu + rect.width_emu <= bounds.x_emu + bounds.width_emu && rect.y_emu + rect.height_emu <= bounds.y_emu + bounds.height_emu
}

function pointContained(x: number, y: number, bounds: NativeSheetGeometryRectV1): boolean {
  return x >= bounds.x_emu && y >= bounds.y_emu && x <= bounds.x_emu + bounds.width_emu && y <= bounds.y_emu + bounds.height_emu
}

function decorationDigest(unsigned: Omit<NativeSheetDecorationPlanV1, 'decoration_sha256'>): `sha256:${string}` {
  return `sha256:${sha256Hex(JSON.stringify(unsigned))}`
}

function validateCommandLimit(value: number, path: string): void {
  if (!Number.isSafeInteger(value) || value < 1 || value > NATIVE_SHEET_DECORATION_LIMITS.maxCommands) throw new NativeSheetDecorationError('decoration.invalidLimit', path, `command limit must be from 1 through ${NATIVE_SHEET_DECORATION_LIMITS.maxCommands}`)
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
