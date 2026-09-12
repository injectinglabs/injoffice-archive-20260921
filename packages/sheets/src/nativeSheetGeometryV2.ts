import { EXCEL_MAX_COLUMNS, EXCEL_MAX_ROWS } from './mutationProtocol.js'
import { isProjectedNativeWorkbookV2 } from './nativeRenderModelV2.js'
import type { NativeSheetRenderModelV2, NativeWorkbookRenderModelV2 } from './nativeRenderModelV2.js'
import { sha256Hex } from './nativeSha256.js'
import { NativePlainDataError, snapshotNativePlainData } from './nativePlainData.js'
import { isNativeMaximumDigitWidthAuthorityV2 } from './nativeMaximumDigitWidthV2.js'
import {decodeNativeWorkbookObjectsV1,type NativeWorkbookObjectsV1} from './nativeObjectsPreviewV1.js'
import type {NativeStoredRowGeometryV1} from './nativeStoredRowsPreviewV1.js'

export const NATIVE_SHEET_GEOMETRY_V2_PROTOCOL = 'injoffice.xlsx.sheet-geometry'
export const NATIVE_SHEET_GEOMETRY_V2_VERSION = 1 as const
export const EMU_PER_POINT = 12_700
export const EMU_PER_CSS_PIXEL = 9_525
export const NATIVE_SHEET_GEOMETRY_V2_LIMITS = Object.freeze({
  maxViewportRows: 4_096,
  maxViewportColumns: 1_024,
  maxViewportCells: 100_000,
  maxCommands: 110_000,
})
const compiledNativeSheetGeometries = new WeakSet<object>()
const compiledStoredRowGeometries = new WeakSet<object>()
export interface NativeStoredRowSheetGeometryV1 extends NativeSheetGeometryV2 {
 readonly approximation:{readonly policy:'source-stored-rows-v1';readonly read_only:true;readonly source:NativeStoredRowGeometryV1}
}
export function isCompiledNativeStoredRowSheetGeometryV1(value:unknown):value is NativeStoredRowSheetGeometryV1 {
 return typeof value==='object'&&value!==null&&compiledStoredRowGeometries.has(value)&&Object.isFrozen(value)
}
/** Explicit read-only row-height policy. Descender metadata does not alter the
 * stored row boxes; text fitting/baselines are not qualified by this projection.
 * This is deliberately NOT branded or accepted as strict sheet geometry.
 */
export function compileNativeStoredRowSheetGeometryV1(workbook:NativeWorkbookRenderModelV2,sheetId:string,viewport:NativeSheetViewportV2,metricAuthority:NativeMaximumDigitWidthAuthorityV2,objects:NativeWorkbookObjectsV1):NativeStoredRowSheetGeometryV1 {
 if(!isProjectedNativeWorkbookV2(workbook))throw new TypeError('Stored-row geometry requires a projected source workbook')
 const evidence=decodeNativeWorkbookObjectsV1(objects,workbook.source.package_sha256)
 const sheet=workbook.sheets.find(s=>s.id===sheetId)
 const rows=evidence.row_geometry?.filter(r=>r.sheet_part===sheet?.mutation_authority.source_part)??[]
 if(rows.length!==1||rows[0]!.rows.length!==32)throw new TypeError('Qualified stored row dimensions unavailable')
 const view=snapshotViewport(viewport)
 if(view.end_row>=32)throw new RangeError('Stored row approximation covers only the first 32 rows')
 return compileGeometry(workbook,sheetId,view,metricAuthority,rows[0]) as NativeStoredRowSheetGeometryV1
}

export type NativeSheetGeometryIssueCode =
  | 'geometry.invalidViewport'
  | 'geometry.viewportBudget'
  | 'geometry.sheetMissing'
  | 'geometry.sheetFormatUnavailable'
  | 'geometry.sourceUnsupported'
  | 'geometry.zeroHeightUnavailable'
  | 'geometry.metricAuthority'
  | 'geometry.metricUnavailable'
  | 'geometry.mergeClipped'
  | 'geometry.integerOverflow'
  | 'geometry.commandBudget'
  | 'geometry.invalidLimit'
  | 'geometry.planInvalid'
  | 'geometry.planDigest'
  | 'geometry.commandInvalid'

export class NativeSheetGeometryV2Error extends Error {
  readonly code: NativeSheetGeometryIssueCode
  readonly path: string

  constructor(code: NativeSheetGeometryIssueCode, path: string, message: string) {
    super(message)
    this.name = 'NativeSheetGeometryV2Error'
    this.code = code
    this.path = path
  }
}

export interface NativeSheetViewportV2 {
  readonly row: number
  readonly column: number
  readonly end_row: number
  readonly end_column: number
}

/**
 * Source-bound Normal-font maximum digit width. Hosts may obtain this from a
 * deterministic font parser or shaper; DOM/canvas measurement is deliberately
 * outside this contract and an unbound metric is refused.
 */
export interface NativeMaximumDigitWidthAuthorityV2 {
  readonly source_revision: string
  readonly source_package_sha256: string
  readonly normal_style_xf_id: number
  readonly normal_style_font_id: number
  readonly font_name: string
  readonly font_size_points: number
  readonly font_bold: boolean
  readonly font_italic: boolean
  readonly normal_font_record_sha256: `sha256:${string}`
  readonly font_sha256: `sha256:${string}`
  readonly provider_id: string
  readonly provider_revision: string
  readonly measurement_dpi: 96
  readonly maximum_digit_width_pixels: number
}

export interface NativeSheetGeometryRectV2 {
  readonly x_emu: number
  readonly y_emu: number
  readonly width_emu: number
  readonly height_emu: number
}

export interface NativeSheetRowBandV2 {
  readonly row: number
  readonly y_emu: number
  readonly height_emu: number
  readonly hidden: boolean
  readonly source: 'sheet-default' | 'row-override'
}

export interface NativeSheetColumnBandV2 {
  readonly column: number
  readonly x_emu: number
  readonly width_emu: number
  readonly hidden: boolean
  readonly width_characters: number
  readonly source: 'sheet-default' | 'column-override'
}

export interface NativeSheetMergedGeometryV2 {
  readonly ref: string
  readonly row: number
  readonly column: number
  readonly end_row: number
  readonly end_column: number
  readonly rect: NativeSheetGeometryRectV2
}

export interface NativeSheetGeometryV2 {
  readonly protocol: typeof NATIVE_SHEET_GEOMETRY_V2_PROTOCOL
  readonly version: typeof NATIVE_SHEET_GEOMETRY_V2_VERSION
  readonly geometry_sha256: `sha256:${string}`
  readonly document_id: string
  readonly sheet_id: string
  readonly source_revision: string
  readonly source_package_sha256: string
  /** Axis coordinates and bounds are relative to the requested viewport. */
  readonly coordinate_space: 'viewport-local'
  readonly origin_cell: { readonly row: number; readonly column: number }
  readonly viewport: NativeSheetViewportV2
  readonly metric_authority: NativeMaximumDigitWidthAuthorityV2
  readonly bounds: NativeSheetGeometryRectV2
  readonly rows: ReadonlyArray<NativeSheetRowBandV2>
  readonly columns: ReadonlyArray<NativeSheetColumnBandV2>
  readonly merged_ranges: ReadonlyArray<NativeSheetMergedGeometryV2>
}

export type NativeSheetGeometryCommandV2 =
  | { readonly kind: 'beginSheet'; readonly protocol: typeof NATIVE_SHEET_GEOMETRY_V2_PROTOCOL; readonly version: typeof NATIVE_SHEET_GEOMETRY_V2_VERSION; readonly document_id: string; readonly sheet_id: string; readonly source_revision: string; readonly source_package_sha256: string; readonly coordinate_space: 'viewport-local'; readonly origin_cell: { readonly row: number; readonly column: number }; readonly viewport: NativeSheetViewportV2; readonly metric_authority: NativeMaximumDigitWidthAuthorityV2; readonly geometry_sha256: string }
  | { readonly kind: 'clipRect'; readonly rect: NativeSheetGeometryRectV2 }
  | { readonly kind: 'rowBand'; readonly band: NativeSheetRowBandV2 }
  | { readonly kind: 'columnBand'; readonly band: NativeSheetColumnBandV2 }
  | { readonly kind: 'mergedRect'; readonly range: NativeSheetMergedGeometryV2 }
  | { readonly kind: 'endSheet' }

export interface NativeSheetGeometrySurfaceV2 {
  /** Optional atomic preflight capacity advertised by bounded surfaces. */
  readonly command_capacity?: number
  push(command: NativeSheetGeometryCommandV2): void
}

export interface NativeSheetGeometryRecordingSurfaceV2 extends NativeSheetGeometrySurfaceV2 {
  readonly commands: ReadonlyArray<NativeSheetGeometryCommandV2>
  finish(): ReadonlyArray<NativeSheetGeometryCommandV2>
}

export interface NativeSheetGeometryCommandAdapterV2<HostContext> {
  execute(context: HostContext, command: NativeSheetGeometryCommandV2): void
}

export function compileNativeSheetGeometryV2(
  workbook: NativeWorkbookRenderModelV2,
  sheetId: string,
  viewport: NativeSheetViewportV2,
  metricAuthority: NativeMaximumDigitWidthAuthorityV2,
): NativeSheetGeometryV2 {
 return compileGeometry(workbook,sheetId,viewport,metricAuthority)
}
function compileGeometry(workbook:NativeWorkbookRenderModelV2,sheetId:string,viewport:NativeSheetViewportV2,metricAuthority:NativeMaximumDigitWidthAuthorityV2,storedRows?:NativeStoredRowGeometryV1):NativeSheetGeometryV2 {
  if (!isProjectedNativeWorkbookV2(workbook)) throw new NativeSheetGeometryV2Error('geometry.sourceUnsupported', '$.workbook', 'workbook must be the branded frozen result of projectNativeWorkbookV2')
  const safeViewport = snapshotViewport(viewport)
  if (!isNativeMaximumDigitWidthAuthorityV2(metricAuthority)) throw new NativeSheetGeometryV2Error('geometry.metricAuthority', '$.metric_authority', 'maximum digit width must come from the pinned native sfnt provider over exact font bytes')
  const safeMetricAuthority = snapshotMetricAuthority(metricAuthority)
  if (typeof sheetId !== 'string' || !/^[1-9][0-9]{0,9}$/.test(sheetId) || Number(sheetId) > 0xffff_ffff) throw new NativeSheetGeometryV2Error('geometry.sheetMissing', '$.sheetId', 'sheet id is not canonical uint32 text')
  const sheet = workbook.sheets.find((candidate) => candidate.id === sheetId)
  if (!sheet) throw new NativeSheetGeometryV2Error('geometry.sheetMissing', '$.sheetId', `sheet ${JSON.stringify(sheetId)} is absent`)
  const format = sheet.sheet_format
  if (!format) throw new NativeSheetGeometryV2Error('geometry.sheetFormatUnavailable', '$.sheet.sheet_format', 'source worksheet has no authoritative sheetFormatPr geometry')
  if (format.zero_height) throw new NativeSheetGeometryV2Error('geometry.zeroHeightUnavailable', '$.sheet.sheet_format.zero_height', 'zeroHeight needs explicit-row visibility provenance not available in native v2')
  const dimensionIssue = workbook.unsupported.find((item) => item.scope_id === `sheet:${sheet.id}` && !(storedRows&&(item.code==='SHEET_FORMAT_EXTRAS'||item.code==='ROW_DIMENSION_EXTRAS'||(item.code==='WORKSHEET_ATTRIBUTES'&&storedRows.root_policy==='x14ac-descent-only-v1'))) && (
    item.code === 'SHEET_FORMAT_EXTRAS' || item.code === 'ROW_DIMENSION_EXTRAS' || item.code === 'COLUMN_DIMENSION_EXTRAS' || item.code === 'COLS_ATTRIBUTES'
    || item.code === 'SHEET_VIEW_GEOMETRY' || item.code === 'WORKSHEET_ATTRIBUTES' || item.code === 'FOREIGN_WORKSHEET_MARKUP'
  ))
  if (dimensionIssue) throw new NativeSheetGeometryV2Error('geometry.sourceUnsupported', '$.sheet', `source dimension semantics ${dimensionIssue.code} are not projected exactly`)
  validateMetricAuthority(workbook, safeMetricAuthority)

  const defaultColumnWidth = format.default_column_width ?? paddedBaseColumnWidth(format.base_column_width ?? 8, safeMetricAuthority.maximum_digit_width_pixels)
  const rows: NativeSheetRowBandV2[] = []
  const columns: NativeSheetColumnBandV2[] = []
  let y = 0
  let rowCursor = 0
  for (let row = safeViewport.row; row <= safeViewport.end_row; row++) {
    while (rowCursor < sheet.rows.length && sheet.rows[rowCursor]!.row < row) rowCursor++
    const override = sheet.rows[rowCursor]?.row === row ? sheet.rows[rowCursor] : undefined
    const points = storedRows?.rows[row]?.height_points ?? override?.height_points ?? format.default_row_height_points
    const hidden = (storedRows?.rows[row]?.hidden ?? override?.hidden ?? false) || points === 0
    const height = hidden ? 0 : checkedInteger(Math.round(points * EMU_PER_POINT), `$.rows[${rows.length}].height_emu`)
    rows.push({ row, y_emu: y, height_emu: height, hidden, source: override ? 'row-override' : 'sheet-default' })
    y = checkedSum(y, height, `$.rows[${rows.length - 1}].y_emu`)
  }
  let x = 0
  let columnCursor = 0
  for (let column = safeViewport.column; column <= safeViewport.end_column; column++) {
    while (columnCursor < sheet.columns.length && sheet.columns[columnCursor]!.end_column < column) columnCursor++
    const candidate = sheet.columns[columnCursor]
    const override = candidate && candidate.column <= column && column <= candidate.end_column ? candidate : undefined
    if (override?.best_fit && override.width === undefined) {
      throw new NativeSheetGeometryV2Error('geometry.sheetFormatUnavailable', `$.sheet.columns[${columnCursor}].width`, 'bestFit column has no stored width; content-derived geometry is unavailable')
    }
    const widthCharacters = override?.width ?? defaultColumnWidth
    const hidden = (override?.hidden ?? false) || widthCharacters === 0
    const pixels = hidden || widthCharacters === 0 ? 0 : characterWidthToPixels(widthCharacters, safeMetricAuthority.maximum_digit_width_pixels)
    const width = checkedInteger(pixels * EMU_PER_CSS_PIXEL, `$.columns[${columns.length}].width_emu`)
    columns.push({ column, x_emu: x, width_emu: width, hidden, width_characters: widthCharacters, source: override ? 'column-override' : 'sheet-default' })
    x = checkedSum(x, width, `$.columns[${columns.length - 1}].x_emu`)
  }
  const merged = compileMergedGeometry(sheet, safeViewport, rows, columns)
  const canonicalViewport: NativeSheetViewportV2 = {
    row: safeViewport.row,
    column: safeViewport.column,
    end_row: safeViewport.end_row,
    end_column: safeViewport.end_column,
  }
  const canonicalMetricAuthority: NativeMaximumDigitWidthAuthorityV2 = {
    ...safeMetricAuthority,
  }
  const unsigned = {
    ...(storedRows?{approximation:{policy:'source-stored-rows-v1' as const,read_only:true as const,source:storedRows}}:{}),
    protocol: NATIVE_SHEET_GEOMETRY_V2_PROTOCOL,
    version: NATIVE_SHEET_GEOMETRY_V2_VERSION,
    document_id: workbook.document_id,
    sheet_id: sheet.id,
    source_revision: workbook.revision,
    source_package_sha256: workbook.source.package_sha256,
    coordinate_space: 'viewport-local' as const,
    origin_cell: { row: safeViewport.row, column: safeViewport.column },
    viewport: canonicalViewport,
    metric_authority: canonicalMetricAuthority,
    bounds: { x_emu: 0, y_emu: 0, width_emu: x, height_emu: y },
    rows,
    columns,
    merged_ranges: merged,
  } as const
  const result: NativeSheetGeometryV2 = {
    ...unsigned,
    geometry_sha256: geometryDigest(unsigned),
  }
  const frozen = deepFreeze(result)
  if(storedRows)compiledStoredRowGeometries.add(frozen)
  else compiledNativeSheetGeometries.add(frozen)
  return frozen
}

export function isCompiledNativeSheetGeometryV2(value: unknown): value is NativeSheetGeometryV2 {
  return typeof value === 'object' && value !== null && compiledNativeSheetGeometries.has(value) && Object.isFrozen(value)
}

export function createNativeSheetGeometryRecordingSurfaceV2(maxCommands: number = NATIVE_SHEET_GEOMETRY_V2_LIMITS.maxCommands): NativeSheetGeometryRecordingSurfaceV2 {
  validateCommandLimit(maxCommands, '$.maxCommands')
  const commands: NativeSheetGeometryCommandV2[] = []
  let finished: ReadonlyArray<NativeSheetGeometryCommandV2> | undefined
  return {
    command_capacity: maxCommands,
    get commands() { return finished ?? Object.freeze([...commands]) },
    push(command) {
      if (finished) throw new Error('native sheet geometry recording is already finished')
      if (commands.length >= maxCommands) throw new NativeSheetGeometryV2Error('geometry.commandBudget', '$.commands', `geometry commands exceed ${maxCommands}`)
      commands.push(deepFreeze(command))
    },
    finish() {
      if (!finished) finished = Object.freeze([...commands])
      return finished
    },
  }
}

export function emitNativeSheetGeometryCommandsV2(geometry: NativeSheetGeometryV2, surface: NativeSheetGeometrySurfaceV2, maxCommands: number = NATIVE_SHEET_GEOMETRY_V2_LIMITS.maxCommands): void {
  const validated = validateNativeSheetGeometryV2(geometry)
  validateCommandLimit(maxCommands, '$.maxCommands')
  const callback = callbackDataProperty(surface, 'push', '$.surface.push')
  const capacityValue = optionalDataProperty(surface, 'command_capacity', '$.surface.command_capacity')
  if (capacityValue !== undefined) validateCommandLimit(capacityValue, '$.surface.command_capacity')
  const commands = geometryCommands(validated)
  const capacity = Math.min(maxCommands, capacityValue ?? maxCommands)
  if (commands.length > capacity) throw new NativeSheetGeometryV2Error('geometry.commandBudget', '$.commands', `geometry commands require ${commands.length}, exceeding ${capacity}`)
  for (const command of commands) callback.call(surface, command)
}

export function replayNativeSheetGeometryCommandsV2<HostContext>(context: HostContext, commands: ReadonlyArray<NativeSheetGeometryCommandV2>, adapter: NativeSheetGeometryCommandAdapterV2<HostContext>, maxCommands: number = NATIVE_SHEET_GEOMETRY_V2_LIMITS.maxCommands): void {
  validateCommandLimit(maxCommands, '$.maxCommands')
  const validated = validateNativeSheetGeometryCommandsV2(commands, maxCommands)
  const execute = callbackDataProperty(adapter, 'execute', '$.adapter.execute') as (context: HostContext, command: NativeSheetGeometryCommandV2) => void
  for (const command of validated) execute.call(adapter, context, command)
}

/** Rebuilds and freezes the canonical geometry while verifying its digest. */
export function validateNativeSheetGeometryV2(input: unknown): NativeSheetGeometryV2 {
  const value = snapshotGeometryInput(input, '$')
  const geometry = exactObject(value, ['protocol', 'version', 'geometry_sha256', 'document_id', 'sheet_id', 'source_revision', 'source_package_sha256', 'coordinate_space', 'origin_cell', 'viewport', 'metric_authority', 'bounds', 'rows', 'columns', 'merged_ranges'], '$')
  if (geometry.protocol !== NATIVE_SHEET_GEOMETRY_V2_PROTOCOL || geometry.version !== NATIVE_SHEET_GEOMETRY_V2_VERSION) invalidGeometry('$', 'geometry protocol or version is unsupported')
  const documentID = identifier(geometry.document_id, '$.document_id')
  const sheetID = canonicalSheetID(geometry.sheet_id, '$.sheet_id')
  const sourceRevision = stringPattern(geometry.source_revision, '$.source_revision', /^rev:[0-9a-f]{64}$/)
  const sourcePackage = stringPattern(geometry.source_package_sha256, '$.source_package_sha256', /^sha256:[0-9a-f]{64}$/)
  if (sourceRevision.slice(4) !== sourcePackage.slice(7)) invalidGeometry('$.source_revision', 'source revision does not match package digest')
  if (geometry.coordinate_space !== 'viewport-local') invalidGeometry('$.coordinate_space', 'coordinate space must be viewport-local')
  const viewport = snapshotViewport(geometry.viewport)
  const origin = exactObject(geometry.origin_cell, ['row', 'column'], '$.origin_cell')
  const originCell = { row: boundedInteger(origin.row, '$.origin_cell.row', 0, EXCEL_MAX_ROWS - 1), column: boundedInteger(origin.column, '$.origin_cell.column', 0, EXCEL_MAX_COLUMNS - 1) }
  if (originCell.row !== viewport.row || originCell.column !== viewport.column) invalidGeometry('$.origin_cell', 'origin cell must equal viewport origin')
  const metricAuthority = snapshotMetricAuthority(geometry.metric_authority)
  const bounds = geometryRect(geometry.bounds, '$.bounds', false)
  if (bounds.x_emu !== 0 || bounds.y_emu !== 0) invalidGeometry('$.bounds', 'viewport-local bounds must start at zero')
  if (!Array.isArray(geometry.rows) || geometry.rows.length !== viewport.end_row - viewport.row + 1) invalidGeometry('$.rows', 'row bands must exactly cover the viewport')
  if (!Array.isArray(geometry.columns) || geometry.columns.length !== viewport.end_column - viewport.column + 1) invalidGeometry('$.columns', 'column bands must exactly cover the viewport')
  if (!Array.isArray(geometry.merged_ranges) || geometry.merged_ranges.length > NATIVE_SHEET_GEOMETRY_V2_LIMITS.maxViewportCells) invalidGeometry('$.merged_ranges', 'merged range inventory is invalid or oversized')
  const rows: NativeSheetRowBandV2[] = []
  let y = 0
  for (let index = 0; index < geometry.rows.length; index++) {
    const path = `$.rows[${index}]`, row = exactObject(geometry.rows[index], ['row', 'y_emu', 'height_emu', 'hidden', 'source'], path)
    const item: NativeSheetRowBandV2 = {
      row: boundedInteger(row.row, `${path}.row`, viewport.row, viewport.end_row),
      y_emu: boundedInteger(row.y_emu, `${path}.y_emu`, 0, Number.MAX_SAFE_INTEGER),
      height_emu: boundedInteger(row.height_emu, `${path}.height_emu`, 0, Number.MAX_SAFE_INTEGER),
      hidden: booleanValue(row.hidden, `${path}.hidden`),
      source: enumValue(row.source, `${path}.source`, ['sheet-default', 'row-override']),
    }
    if (item.row !== viewport.row + index || item.y_emu !== y || item.hidden !== (item.height_emu === 0)) invalidGeometry(path, 'row bands must be consecutive, cumulative, and canonically hidden')
    y = checkedGeometrySum(y, item.height_emu, path)
    rows.push(item)
  }
  const columns: NativeSheetColumnBandV2[] = []
  let x = 0
  for (let index = 0; index < geometry.columns.length; index++) {
    const path = `$.columns[${index}]`, column = exactObject(geometry.columns[index], ['column', 'x_emu', 'width_emu', 'hidden', 'width_characters', 'source'], path)
    const widthCharacters = finiteNumber(column.width_characters, `${path}.width_characters`, 0, 255)
    const item: NativeSheetColumnBandV2 = {
      column: boundedInteger(column.column, `${path}.column`, viewport.column, viewport.end_column),
      x_emu: boundedInteger(column.x_emu, `${path}.x_emu`, 0, Number.MAX_SAFE_INTEGER),
      width_emu: boundedInteger(column.width_emu, `${path}.width_emu`, 0, Number.MAX_SAFE_INTEGER),
      hidden: booleanValue(column.hidden, `${path}.hidden`), width_characters: widthCharacters,
      source: enumValue(column.source, `${path}.source`, ['sheet-default', 'column-override']),
    }
    if (item.column !== viewport.column + index || item.x_emu !== x || item.hidden !== (item.width_emu === 0)) invalidGeometry(path, 'column bands must be consecutive, cumulative, and canonically hidden')
    x = checkedGeometrySum(x, item.width_emu, path)
    columns.push(item)
  }
  if (bounds.width_emu !== x || bounds.height_emu !== y) invalidGeometry('$.bounds', 'bounds must equal cumulative axis extents')
  const mergedRanges: NativeSheetMergedGeometryV2[] = []
  const mergedOccupancy = new Uint8Array(rows.length * columns.length)
  let priorMerged = -1
  for (let index = 0; index < geometry.merged_ranges.length; index++) {
    const path = `$.merged_ranges[${index}]`, merged = exactObject(geometry.merged_ranges[index], ['ref', 'row', 'column', 'end_row', 'end_column', 'rect'], path)
    const row = boundedInteger(merged.row, `${path}.row`, viewport.row, viewport.end_row), column = boundedInteger(merged.column, `${path}.column`, viewport.column, viewport.end_column)
    const endRow = boundedInteger(merged.end_row, `${path}.end_row`, row, viewport.end_row), endColumn = boundedInteger(merged.end_column, `${path}.end_column`, column, viewport.end_column)
    if (row === endRow && column === endColumn) invalidGeometry(path, 'merged range must span multiple cells')
    const position = row * EXCEL_MAX_COLUMNS + column
    if (position <= priorMerged) invalidGeometry(path, 'merged ranges must be uniquely row-major ordered')
    priorMerged = position
    const ref = stringPattern(merged.ref, `${path}.ref`, /^[A-Z]{1,3}[1-9][0-9]{0,6}:[A-Z]{1,3}[1-9][0-9]{0,6}$/)
    if (ref !== `${cellReference(row, column)}:${cellReference(endRow, endColumn)}`) invalidGeometry(`${path}.ref`, 'merged reference does not match coordinates')
    const rect = geometryRect(merged.rect, `${path}.rect`, false)
    const firstRow = rows[row - viewport.row]!, lastRow = rows[endRow - viewport.row]!, firstColumn = columns[column - viewport.column]!, lastColumn = columns[endColumn - viewport.column]!
    const expected = { x_emu: firstColumn.x_emu, y_emu: firstRow.y_emu, width_emu: lastColumn.x_emu + lastColumn.width_emu - firstColumn.x_emu, height_emu: lastRow.y_emu + lastRow.height_emu - firstRow.y_emu }
    if (JSON.stringify(rect) !== JSON.stringify(expected)) invalidGeometry(`${path}.rect`, 'merged rectangle does not match axis bands')
    for (let occupiedRow = row; occupiedRow <= endRow; occupiedRow++) for (let occupiedColumn = column; occupiedColumn <= endColumn; occupiedColumn++) {
      const occupied = (occupiedRow - viewport.row) * columns.length + occupiedColumn - viewport.column
      if (mergedOccupancy[occupied] !== 0) invalidGeometry(path, 'merged ranges must not overlap')
      mergedOccupancy[occupied] = 1
    }
    mergedRanges.push({ ref, row, column, end_row: endRow, end_column: endColumn, rect })
  }
  const unsigned: Omit<NativeSheetGeometryV2, 'geometry_sha256'> = { protocol: NATIVE_SHEET_GEOMETRY_V2_PROTOCOL, version: NATIVE_SHEET_GEOMETRY_V2_VERSION, document_id: documentID, sheet_id: sheetID, source_revision: sourceRevision, source_package_sha256: sourcePackage, coordinate_space: 'viewport-local', origin_cell: originCell, viewport, metric_authority: metricAuthority, bounds, rows, columns, merged_ranges: mergedRanges }
  const supplied = stringPattern(geometry.geometry_sha256, '$.geometry_sha256', /^sha256:[0-9a-f]{64}$/)
  if (geometryDigest(unsigned) !== supplied) throw new NativeSheetGeometryV2Error('geometry.planDigest', '$.geometry_sha256', 'geometry digest does not match canonical geometry')
  return deepFreeze({ ...unsigned, geometry_sha256: supplied as `sha256:${string}` })
}

function validateNativeSheetGeometryCommandsV2(input: unknown, maxCommands: number): ReadonlyArray<NativeSheetGeometryCommandV2> {
  const value = snapshotGeometryInput(input, '$.commands')
  if (!Array.isArray(value) || value.length < 3 || value.length > maxCommands) commandInvalid('$.commands', 'command stream length is outside its bound')
  const begin = exactObject(value[0], ['kind', 'protocol', 'version', 'document_id', 'sheet_id', 'source_revision', 'source_package_sha256', 'coordinate_space', 'origin_cell', 'viewport', 'metric_authority', 'geometry_sha256'], '$.commands[0]')
  if (begin.kind !== 'beginSheet') commandInvalid('$.commands[0].kind', 'command stream must begin with beginSheet')
  const clip = exactObject(value[1], ['kind', 'rect'], '$.commands[1]')
  if (clip.kind !== 'clipRect') commandInvalid('$.commands[1].kind', 'second command must be clipRect')
  const end = exactObject(value[value.length - 1], ['kind'], `$.commands[${value.length - 1}]`)
  if (end.kind !== 'endSheet') commandInvalid(`$.commands[${value.length - 1}].kind`, 'command stream must end with endSheet')
  const rows: unknown[] = [], columns: unknown[] = [], merged: unknown[] = []
  let lane: 'rows' | 'columns' | 'merged' = 'rows'
  for (let index = 2; index < value.length - 1; index++) {
    const raw = value[index]
    if (typeof raw !== 'object' || raw === null || Array.isArray(raw)) commandInvalid(`$.commands[${index}]`, 'command must be an object')
    const kind = (raw as Record<string, unknown>).kind
    const command = exactObject(raw, ['kind', kind === 'mergedRect' ? 'range' : 'band'], `$.commands[${index}]`)
    if (command.kind === 'rowBand' && lane === 'rows') rows.push(command.band)
    else if (command.kind === 'columnBand' && lane !== 'merged') { lane = 'columns'; columns.push(command.band) }
    else if (command.kind === 'mergedRect') { lane = 'merged'; merged.push(command.range) }
    else commandInvalid(`$.commands[${index}].kind`, 'commands must contain row bands, column bands, then merged rectangles')
  }
  const geometry = validateNativeSheetGeometryV2({
    protocol: begin.protocol, version: begin.version, geometry_sha256: begin.geometry_sha256,
    document_id: begin.document_id, sheet_id: begin.sheet_id, source_revision: begin.source_revision,
    source_package_sha256: begin.source_package_sha256, coordinate_space: begin.coordinate_space,
    origin_cell: begin.origin_cell, viewport: begin.viewport, metric_authority: begin.metric_authority,
    bounds: clip.rect, rows, columns, merged_ranges: merged,
  })
  return geometryCommands(geometry)
}

function geometryCommands(geometry: NativeSheetGeometryV2): ReadonlyArray<NativeSheetGeometryCommandV2> {
  return deepFreeze([
    {
      kind: 'beginSheet' as const, protocol: geometry.protocol, version: geometry.version,
      document_id: geometry.document_id, sheet_id: geometry.sheet_id, source_revision: geometry.source_revision,
      source_package_sha256: geometry.source_package_sha256, coordinate_space: geometry.coordinate_space,
      origin_cell: geometry.origin_cell, viewport: geometry.viewport, metric_authority: geometry.metric_authority,
      geometry_sha256: geometry.geometry_sha256,
    },
    { kind: 'clipRect' as const, rect: geometry.bounds },
    ...geometry.rows.map((band) => ({ kind: 'rowBand' as const, band })),
    ...geometry.columns.map((band) => ({ kind: 'columnBand' as const, band })),
    ...geometry.merged_ranges.map((range) => ({ kind: 'mergedRect' as const, range })),
    { kind: 'endSheet' as const },
  ])
}

/** ISO/IEC 29500 stored character width to runtime grid pixels. */
export function characterWidthToPixels(width: number, maximumDigitWidthPixels: number): number {
  if (!Number.isFinite(width) || Object.is(width, -0) || width < 0 || width > 255) throw new NativeSheetGeometryV2Error('geometry.metricUnavailable', '$.width', 'column width must be finite, non-negative-zero, and within 0..255')
  validateMdw(maximumDigitWidthPixels, '$.maximumDigitWidthPixels')
  return width === 0 ? 0 : Math.floor(((256 * width + Math.floor(128 / maximumDigitWidthPixels)) / 256) * maximumDigitWidthPixels)
}

/** Default base character count plus Excel's five-pixel cell padding, snapped down to 1/256. */
export function paddedBaseColumnWidth(baseColumnWidth: number, maximumDigitWidthPixels: number): number {
  if (!Number.isSafeInteger(baseColumnWidth) || Object.is(baseColumnWidth, -0) || baseColumnWidth < 0 || baseColumnWidth > 255) throw new NativeSheetGeometryV2Error('geometry.sheetFormatUnavailable', '$.sheet.sheet_format.base_column_width', 'base column width must be a non-negative-zero integer within 0..255')
  validateMdw(maximumDigitWidthPixels, '$.metric_authority.maximum_digit_width_pixels')
  return Math.floor((baseColumnWidth + 5 / maximumDigitWidthPixels) * 256) / 256
}

function validateViewport(viewport: NativeSheetViewportV2): void {
  const values = [viewport.row, viewport.column, viewport.end_row, viewport.end_column]
  if (values.some((value) => !Number.isSafeInteger(value) || Object.is(value, -0)) || viewport.row < 0 || viewport.column < 0 || viewport.end_row < viewport.row || viewport.end_column < viewport.column || viewport.end_row >= EXCEL_MAX_ROWS || viewport.end_column >= EXCEL_MAX_COLUMNS) {
    throw new NativeSheetGeometryV2Error('geometry.invalidViewport', '$.viewport', 'viewport must be a forward bounded zero-based Excel rectangle')
  }
  const rowCount = viewport.end_row - viewport.row + 1
  const columnCount = viewport.end_column - viewport.column + 1
  if (rowCount > NATIVE_SHEET_GEOMETRY_V2_LIMITS.maxViewportRows || columnCount > NATIVE_SHEET_GEOMETRY_V2_LIMITS.maxViewportColumns || rowCount * columnCount > NATIVE_SHEET_GEOMETRY_V2_LIMITS.maxViewportCells) {
    throw new NativeSheetGeometryV2Error('geometry.viewportBudget', '$.viewport', 'viewport exceeds native geometry resource limits')
  }
}

function snapshotViewport(input: unknown): NativeSheetViewportV2 {
  const value = exactObject(snapshotGeometryInput(input, '$.viewport'), ['row', 'column', 'end_row', 'end_column'], '$.viewport')
  const result = {
    row: boundedInteger(value.row, '$.viewport.row', 0, EXCEL_MAX_ROWS - 1),
    column: boundedInteger(value.column, '$.viewport.column', 0, EXCEL_MAX_COLUMNS - 1),
    end_row: boundedInteger(value.end_row, '$.viewport.end_row', 0, EXCEL_MAX_ROWS - 1),
    end_column: boundedInteger(value.end_column, '$.viewport.end_column', 0, EXCEL_MAX_COLUMNS - 1),
  }
  validateViewport(result)
  return result
}

function snapshotMetricAuthority(input: unknown): NativeMaximumDigitWidthAuthorityV2 {
  const value = exactObject(snapshotGeometryInput(input, '$.metric_authority'), ['source_revision', 'source_package_sha256', 'normal_style_xf_id', 'normal_style_font_id', 'font_name', 'font_size_points', 'font_bold', 'font_italic', 'normal_font_record_sha256', 'font_sha256', 'provider_id', 'provider_revision', 'measurement_dpi', 'maximum_digit_width_pixels'], '$.metric_authority')
  const result: NativeMaximumDigitWidthAuthorityV2 = {
    source_revision: stringPattern(value.source_revision, '$.metric_authority.source_revision', /^rev:[0-9a-f]{64}$/),
    source_package_sha256: stringPattern(value.source_package_sha256, '$.metric_authority.source_package_sha256', /^sha256:[0-9a-f]{64}$/),
    normal_style_xf_id: boundedInteger(value.normal_style_xf_id, '$.metric_authority.normal_style_xf_id', 0, 0xffff_ffff),
    normal_style_font_id: boundedInteger(value.normal_style_font_id, '$.metric_authority.normal_style_font_id', 0, 0xffff_ffff),
    font_name: boundedString(value.font_name, '$.metric_authority.font_name', 255),
    font_size_points: finiteNumber(value.font_size_points, '$.metric_authority.font_size_points', Number.MIN_VALUE, 409.5),
    font_bold: booleanValue(value.font_bold, '$.metric_authority.font_bold'),
    font_italic: booleanValue(value.font_italic, '$.metric_authority.font_italic'),
    normal_font_record_sha256: stringPattern(value.normal_font_record_sha256, '$.metric_authority.normal_font_record_sha256', /^sha256:[0-9a-f]{64}$/) as `sha256:${string}`,
    font_sha256: stringPattern(value.font_sha256, '$.metric_authority.font_sha256', /^sha256:[0-9a-f]{64}$/) as `sha256:${string}`,
    provider_id: boundedString(value.provider_id, '$.metric_authority.provider_id', 256),
    provider_revision: boundedString(value.provider_revision, '$.metric_authority.provider_revision', 256),
    measurement_dpi: boundedInteger(value.measurement_dpi, '$.metric_authority.measurement_dpi', 96, 96) as 96,
    maximum_digit_width_pixels: boundedInteger(value.maximum_digit_width_pixels, '$.metric_authority.maximum_digit_width_pixels', 1, 512),
  }
  if (result.source_revision.slice(4) !== result.source_package_sha256.slice(7)) invalidGeometry('$.metric_authority.source_revision', 'metric source revision does not match package digest')
  return result
}

function validateMetricAuthority(workbook: NativeWorkbookRenderModelV2, authority: NativeMaximumDigitWidthAuthorityV2): void {
  if (authority.source_revision !== workbook.revision || authority.source_package_sha256 !== workbook.source.package_sha256) {
    throw new NativeSheetGeometryV2Error('geometry.metricAuthority', '$.metric_authority', 'maximum digit width is not bound to this exact workbook revision and Normal style')
  }
  const normal = workbook.normal_style
  if (!normal) {
    throw new NativeSheetGeometryV2Error('geometry.metricUnavailable', '$.normal_style', 'the built-in Normal-style font identity is unavailable')
  }
  if (authority.normal_style_xf_id !== normal.style_xf_id || authority.normal_style_font_id !== normal.font_id || authority.font_name !== normal.font_name || authority.font_size_points !== normal.font_size_points || authority.font_bold !== normal.font_bold || authority.font_italic !== normal.font_italic || authority.normal_font_record_sha256 !== normal.font_record_sha256) {
    throw new NativeSheetGeometryV2Error('geometry.metricAuthority', '$.metric_authority', 'maximum digit width font identity disagrees with the projected Normal style')
  }
  if (!/^sha256:[0-9a-f]{64}$/.test(authority.normal_font_record_sha256) || !/^sha256:[0-9a-f]{64}$/.test(authority.font_sha256) || !boundedIdentifier(authority.provider_id) || !boundedIdentifier(authority.provider_revision) || authority.measurement_dpi !== 96) {
    throw new NativeSheetGeometryV2Error('geometry.metricAuthority', '$.metric_authority', 'font digest and provider provenance are required')
  }
  validateMdw(authority.maximum_digit_width_pixels, '$.metric_authority.maximum_digit_width_pixels')
}

function validateMdw(value: number, path: string): void {
  if (!Number.isSafeInteger(value) || Object.is(value, -0) || value < 1 || value > 512) throw new NativeSheetGeometryV2Error('geometry.metricUnavailable', path, 'maximum digit width must be a non-negative-zero integer from 1 through 512 pixels')
}

function boundedIdentifier(value: string): boolean { return value.length >= 1 && value.length <= 256 && !/[\u0000-\u001f\u007f]/.test(value) }

function compileMergedGeometry(sheet: NativeSheetRenderModelV2, viewport: NativeSheetViewportV2, rows: NativeSheetRowBandV2[], columns: NativeSheetColumnBandV2[]): NativeSheetMergedGeometryV2[] {
  const result: NativeSheetMergedGeometryV2[] = []
  for (const range of sheet.merged_ranges) {
    const intersects = range.row <= viewport.end_row && range.end_row >= viewport.row && range.column <= viewport.end_column && range.end_column >= viewport.column
    if (!intersects) continue
    if (range.row < viewport.row || range.end_row > viewport.end_row || range.column < viewport.column || range.end_column > viewport.end_column) {
      throw new NativeSheetGeometryV2Error('geometry.mergeClipped', `$.sheet.merged_ranges[${result.length}]`, `viewport clips merged range ${range.ref}`)
    }
    const firstRow = rows[range.row - viewport.row]!
    const lastRow = rows[range.end_row - viewport.row]!
    const firstColumn = columns[range.column - viewport.column]!
    const lastColumn = columns[range.end_column - viewport.column]!
    result.push({
      ref: range.ref, row: range.row, column: range.column, end_row: range.end_row, end_column: range.end_column,
      rect: {
        x_emu: firstColumn.x_emu,
        y_emu: firstRow.y_emu,
        width_emu: checkedSum(lastColumn.x_emu, lastColumn.width_emu, '$.merged.width_emu') - firstColumn.x_emu,
        height_emu: checkedSum(lastRow.y_emu, lastRow.height_emu, '$.merged.height_emu') - firstRow.y_emu,
      },
    })
  }
  return result
}

function checkedInteger(value: number, path: string): number {
  if (!Number.isSafeInteger(value) || Object.is(value, -0) || value < 0) throw new NativeSheetGeometryV2Error('geometry.integerOverflow', path, 'geometry exceeds non-negative safe integer EMU bounds')
  return value
}

function checkedSum(left: number, right: number, path: string): number { return checkedInteger(left + right, path) }

type UnknownRecord = Record<string, unknown>

function snapshotGeometryInput(value: unknown, path: string): unknown {
  try { return snapshotNativePlainData(value, { maxDepth: 32, maxNodes: NATIVE_SHEET_GEOMETRY_V2_LIMITS.maxCommands * 16 }) }
  catch (error) {
    if (error instanceof NativePlainDataError) throw new NativeSheetGeometryV2Error('geometry.planInvalid', error.path || path, error.message)
    throw error
  }
}

function exactObject(value: unknown, keys: ReadonlyArray<string>, path: string): UnknownRecord {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) invalidGeometry(path, 'expected an object')
  const record = value as UnknownRecord
  const actual = Object.keys(record).sort(), expected = [...keys].sort()
  if (actual.length !== expected.length || actual.some((key, index) => key !== expected[index])) invalidGeometry(path, 'object fields are missing or unknown')
  return record
}

function boundedInteger(value: unknown, path: string, minimum: number, maximum: number): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || Object.is(value, -0) || value < minimum || value > maximum) invalidGeometry(path, `integer must be within ${minimum}..${maximum} and not negative zero`)
  return value
}

function finiteNumber(value: unknown, path: string, minimum: number, maximum: number): number {
  if (typeof value !== 'number' || !Number.isFinite(value) || Object.is(value, -0) || value < minimum || value > maximum) invalidGeometry(path, `number must be finite within ${minimum}..${maximum} and not negative zero`)
  return value
}

function booleanValue(value: unknown, path: string): boolean {
  if (typeof value !== 'boolean') invalidGeometry(path, 'expected a boolean')
  return value
}

function enumValue<T extends string>(value: unknown, path: string, options: readonly T[]): T {
  if (typeof value !== 'string' || !options.includes(value as T)) invalidGeometry(path, 'value is outside the canonical enumeration')
  return value as T
}

function stringPattern(value: unknown, path: string, pattern: RegExp): string {
  if (typeof value !== 'string' || !pattern.test(value)) invalidGeometry(path, 'string is missing or non-canonical')
  return value
}

function boundedString(value: unknown, path: string, maximum: number): string {
  if (typeof value !== 'string' || value.length < 1 || value.length > maximum || /[\u0000-\u001f\u007f]/.test(value)) invalidGeometry(path, 'string is missing, oversized, or contains control characters')
  return value
}

function identifier(value: unknown, path: string): string {
  const result = boundedString(value, path, 256)
  if (!/^[A-Za-z0-9][A-Za-z0-9._:-]*$/.test(result)) invalidGeometry(path, 'identifier is non-canonical')
  return result
}

function canonicalSheetID(value: unknown, path: string): string {
  const result = stringPattern(value, path, /^[1-9][0-9]{0,9}$/)
  if (Number(result) > 0xffff_ffff) invalidGeometry(path, 'sheet id exceeds uint32')
  return result
}

function cellReference(row: number, column: number): string {
  let value = column + 1, name = ''
  while (value > 0) { value--; name = String.fromCharCode(65 + value % 26) + name; value = Math.floor(value / 26) }
  return `${name}${row + 1}`
}

function geometryRect(value: unknown, path: string, requireArea: boolean): NativeSheetGeometryRectV2 {
  const record = exactObject(value, ['x_emu', 'y_emu', 'width_emu', 'height_emu'], path)
  const result = {
    x_emu: boundedInteger(record.x_emu, `${path}.x_emu`, 0, Number.MAX_SAFE_INTEGER),
    y_emu: boundedInteger(record.y_emu, `${path}.y_emu`, 0, Number.MAX_SAFE_INTEGER),
    width_emu: boundedInteger(record.width_emu, `${path}.width_emu`, requireArea ? 1 : 0, Number.MAX_SAFE_INTEGER),
    height_emu: boundedInteger(record.height_emu, `${path}.height_emu`, requireArea ? 1 : 0, Number.MAX_SAFE_INTEGER),
  }
  checkedGeometrySum(result.x_emu, result.width_emu, path); checkedGeometrySum(result.y_emu, result.height_emu, path)
  return result
}

function checkedGeometrySum(left: number, right: number, path: string): number {
  const result = left + right
  if (!Number.isSafeInteger(result)) invalidGeometry(path, 'geometry endpoint exceeds safe integer bounds')
  return result
}

function geometryDigest(unsigned: Omit<NativeSheetGeometryV2, 'geometry_sha256'>): `sha256:${string}` {
  return `sha256:${sha256Hex(`injoffice.xlsx.sheet-geometry.v1\0${JSON.stringify(unsigned)}`)}`
}

function invalidGeometry(path: string, message: string): never { throw new NativeSheetGeometryV2Error('geometry.planInvalid', path, message) }
function commandInvalid(path: string, message: string): never { throw new NativeSheetGeometryV2Error('geometry.commandInvalid', path, message) }
function validateCommandLimit(value: unknown, path: string): asserts value is number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || Object.is(value, -0) || value < 1 || value > NATIVE_SHEET_GEOMETRY_V2_LIMITS.maxCommands) throw new NativeSheetGeometryV2Error('geometry.invalidLimit', path, `command limit must be from 1 through ${NATIVE_SHEET_GEOMETRY_V2_LIMITS.maxCommands}`)
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

function deepFreeze<T>(value: T): T {
  if (typeof value !== 'object' || value === null || Object.isFrozen(value)) return value
  for (const child of Object.values(value)) deepFreeze(child)
  return Object.freeze(value)
}
