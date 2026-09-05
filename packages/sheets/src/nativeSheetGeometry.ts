import { EXCEL_MAX_COLUMNS, EXCEL_MAX_ROWS } from './mutationProtocol.js'
import { isProjectedNativeWorkbookV1 } from './nativeRenderModel.js'
import type { NativeSheetRenderModelV1, NativeWorkbookRenderModelV1 } from './nativeRenderModel.js'
import { sha256Hex } from './nativeSha256.js'

export const NATIVE_SHEET_GEOMETRY_PROTOCOL = 'injoffice.xlsx.sheet-geometry'
export const NATIVE_SHEET_GEOMETRY_VERSION = 1 as const
export const EMU_PER_POINT = 12_700
export const EMU_PER_CSS_PIXEL = 9_525
export const NATIVE_SHEET_GEOMETRY_LIMITS = Object.freeze({
  maxViewportRows: 4_096,
  maxViewportColumns: 1_024,
  maxViewportCells: 100_000,
  maxCommands: 110_000,
})

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

export class NativeSheetGeometryError extends Error {
  readonly code: NativeSheetGeometryIssueCode
  readonly path: string

  constructor(code: NativeSheetGeometryIssueCode, path: string, message: string) {
    super(message)
    this.name = 'NativeSheetGeometryError'
    this.code = code
    this.path = path
  }
}

export interface NativeSheetViewportV1 {
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
export interface NativeMaximumDigitWidthAuthorityV1 {
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

export interface NativeSheetGeometryRectV1 {
  readonly x_emu: number
  readonly y_emu: number
  readonly width_emu: number
  readonly height_emu: number
}

export interface NativeSheetRowBandV1 {
  readonly row: number
  readonly y_emu: number
  readonly height_emu: number
  readonly hidden: boolean
  readonly source: 'sheet-default' | 'row-override'
}

export interface NativeSheetColumnBandV1 {
  readonly column: number
  readonly x_emu: number
  readonly width_emu: number
  readonly hidden: boolean
  readonly width_characters: number
  readonly source: 'sheet-default' | 'column-override'
}

export interface NativeSheetMergedGeometryV1 {
  readonly ref: string
  readonly row: number
  readonly column: number
  readonly end_row: number
  readonly end_column: number
  readonly rect: NativeSheetGeometryRectV1
}

export interface NativeSheetGeometryV1 {
  readonly protocol: typeof NATIVE_SHEET_GEOMETRY_PROTOCOL
  readonly version: typeof NATIVE_SHEET_GEOMETRY_VERSION
  readonly geometry_sha256: `sha256:${string}`
  readonly document_id: string
  readonly sheet_id: string
  readonly source_revision: string
  readonly source_package_sha256: string
  /** Axis coordinates and bounds are relative to the requested viewport. */
  readonly coordinate_space: 'viewport-local'
  readonly origin_cell: { readonly row: number; readonly column: number }
  readonly viewport: NativeSheetViewportV1
  readonly metric_authority: NativeMaximumDigitWidthAuthorityV1
  readonly bounds: NativeSheetGeometryRectV1
  readonly rows: ReadonlyArray<NativeSheetRowBandV1>
  readonly columns: ReadonlyArray<NativeSheetColumnBandV1>
  readonly merged_ranges: ReadonlyArray<NativeSheetMergedGeometryV1>
}

export type NativeSheetGeometryCommandV1 =
  | { readonly kind: 'beginSheet'; readonly document_id: string; readonly sheet_id: string; readonly geometry_sha256: string; readonly bounds: NativeSheetGeometryRectV1 }
  | { readonly kind: 'clipRect'; readonly rect: NativeSheetGeometryRectV1 }
  | { readonly kind: 'rowBand'; readonly band: NativeSheetRowBandV1 }
  | { readonly kind: 'columnBand'; readonly band: NativeSheetColumnBandV1 }
  | { readonly kind: 'mergedRect'; readonly range: NativeSheetMergedGeometryV1 }
  | { readonly kind: 'endSheet' }

export interface NativeSheetGeometrySurfaceV1 {
  /** Optional atomic preflight capacity advertised by bounded surfaces. */
  readonly command_capacity?: number
  push(command: NativeSheetGeometryCommandV1): void
}

export interface NativeSheetGeometryRecordingSurfaceV1 extends NativeSheetGeometrySurfaceV1 {
  readonly commands: ReadonlyArray<NativeSheetGeometryCommandV1>
  finish(): ReadonlyArray<NativeSheetGeometryCommandV1>
}

export interface NativeSheetGeometryCommandAdapterV1<HostContext> {
  execute(context: HostContext, command: NativeSheetGeometryCommandV1): void
}

export function compileNativeSheetGeometryV1(
  workbook: NativeWorkbookRenderModelV1,
  sheetId: string,
  viewport: NativeSheetViewportV1,
  metricAuthority: NativeMaximumDigitWidthAuthorityV1,
): NativeSheetGeometryV1 {
  if (!isProjectedNativeWorkbookV1(workbook)) throw new NativeSheetGeometryError('geometry.sourceUnsupported', '$.workbook', 'workbook must be the branded frozen result of projectNativeWorkbookV1')
  validateViewport(viewport)
  const sheet = workbook.sheets.find((candidate) => candidate.id === sheetId)
  if (!sheet) throw new NativeSheetGeometryError('geometry.sheetMissing', '$.sheetId', `sheet ${JSON.stringify(sheetId)} is absent`)
  const format = sheet.sheet_format
  if (!format) throw new NativeSheetGeometryError('geometry.sheetFormatUnavailable', '$.sheet.sheet_format', 'source worksheet has no authoritative sheetFormatPr geometry')
  if (format.zero_height) throw new NativeSheetGeometryError('geometry.zeroHeightUnavailable', '$.sheet.sheet_format.zero_height', 'zeroHeight needs explicit-row visibility provenance not available in native v1')
  const dimensionIssue = workbook.unsupported.find((item) => item.scope_id === `sheet:${sheet.id}` && (
    item.code === 'SHEET_FORMAT_EXTRAS' || item.code === 'ROW_DIMENSION_EXTRAS' || item.code === 'COLUMN_DIMENSION_EXTRAS' || item.code === 'COLS_ATTRIBUTES'
    || item.code === 'SHEET_VIEW_GEOMETRY' || item.code === 'WORKSHEET_ATTRIBUTES' || item.code === 'FOREIGN_WORKSHEET_MARKUP'
  ))
  if (dimensionIssue) throw new NativeSheetGeometryError('geometry.sourceUnsupported', '$.sheet', `source dimension semantics ${dimensionIssue.code} are not projected exactly`)
  validateMetricAuthority(workbook, metricAuthority)

  const defaultColumnWidth = format.default_column_width ?? paddedBaseColumnWidth(format.base_column_width ?? 8, metricAuthority.maximum_digit_width_pixels)
  const rows: NativeSheetRowBandV1[] = []
  const columns: NativeSheetColumnBandV1[] = []
  let y = 0
  let rowCursor = 0
  for (let row = viewport.row; row <= viewport.end_row; row++) {
    while (rowCursor < sheet.rows.length && sheet.rows[rowCursor]!.row < row) rowCursor++
    const override = sheet.rows[rowCursor]?.row === row ? sheet.rows[rowCursor] : undefined
    const points = override?.height_points ?? format.default_row_height_points
    const hidden = (override?.hidden ?? false) || points === 0
    const height = hidden ? 0 : checkedInteger(Math.round(points * EMU_PER_POINT), `$.rows[${rows.length}].height_emu`)
    rows.push({ row, y_emu: y, height_emu: height, hidden, source: override ? 'row-override' : 'sheet-default' })
    y = checkedSum(y, height, `$.rows[${rows.length - 1}].y_emu`)
  }
  let x = 0
  let columnCursor = 0
  for (let column = viewport.column; column <= viewport.end_column; column++) {
    while (columnCursor < sheet.columns.length && sheet.columns[columnCursor]!.end_column < column) columnCursor++
    const candidate = sheet.columns[columnCursor]
    const override = candidate && candidate.column <= column && column <= candidate.end_column ? candidate : undefined
    if (override?.best_fit && override.width === undefined) {
      throw new NativeSheetGeometryError('geometry.sheetFormatUnavailable', `$.sheet.columns[${columnCursor}].width`, 'bestFit column has no stored width; content-derived geometry is unavailable')
    }
    const widthCharacters = override?.width ?? defaultColumnWidth
    const hidden = (override?.hidden ?? false) || widthCharacters === 0
    const pixels = hidden || widthCharacters === 0 ? 0 : characterWidthToPixels(widthCharacters, metricAuthority.maximum_digit_width_pixels)
    const width = checkedInteger(pixels * EMU_PER_CSS_PIXEL, `$.columns[${columns.length}].width_emu`)
    columns.push({ column, x_emu: x, width_emu: width, hidden, width_characters: widthCharacters, source: override ? 'column-override' : 'sheet-default' })
    x = checkedSum(x, width, `$.columns[${columns.length - 1}].x_emu`)
  }
  const merged = compileMergedGeometry(sheet, viewport, rows, columns)
  const canonicalViewport: NativeSheetViewportV1 = {
    row: viewport.row,
    column: viewport.column,
    end_row: viewport.end_row,
    end_column: viewport.end_column,
  }
  const canonicalMetricAuthority: NativeMaximumDigitWidthAuthorityV1 = {
    source_revision: metricAuthority.source_revision,
    source_package_sha256: metricAuthority.source_package_sha256,
    normal_style_xf_id: metricAuthority.normal_style_xf_id,
    normal_style_font_id: metricAuthority.normal_style_font_id,
    font_name: metricAuthority.font_name,
    font_size_points: metricAuthority.font_size_points,
    font_bold: metricAuthority.font_bold,
    font_italic: metricAuthority.font_italic,
    normal_font_record_sha256: metricAuthority.normal_font_record_sha256,
    font_sha256: metricAuthority.font_sha256,
    provider_id: metricAuthority.provider_id,
    provider_revision: metricAuthority.provider_revision,
    measurement_dpi: metricAuthority.measurement_dpi,
    maximum_digit_width_pixels: metricAuthority.maximum_digit_width_pixels,
  }
  const unsigned = {
    protocol: NATIVE_SHEET_GEOMETRY_PROTOCOL,
    version: NATIVE_SHEET_GEOMETRY_VERSION,
    document_id: workbook.document_id,
    sheet_id: sheet.id,
    source_revision: workbook.revision,
    source_package_sha256: workbook.source.package_sha256,
    coordinate_space: 'viewport-local' as const,
    origin_cell: { row: viewport.row, column: viewport.column },
    viewport: canonicalViewport,
    metric_authority: canonicalMetricAuthority,
    bounds: { x_emu: 0, y_emu: 0, width_emu: x, height_emu: y },
    rows,
    columns,
    merged_ranges: merged,
  } as const
  const result: NativeSheetGeometryV1 = {
    ...unsigned,
    geometry_sha256: `sha256:${sha256Hex(JSON.stringify(unsigned))}`,
  }
  return deepFreeze(result)
}

export function createNativeSheetGeometryRecordingSurfaceV1(maxCommands: number = NATIVE_SHEET_GEOMETRY_LIMITS.maxCommands): NativeSheetGeometryRecordingSurfaceV1 {
  if (!Number.isSafeInteger(maxCommands) || maxCommands < 1 || maxCommands > NATIVE_SHEET_GEOMETRY_LIMITS.maxCommands) {
    throw new NativeSheetGeometryError('geometry.invalidLimit', '$.maxCommands', `maxCommands must be from 1 through ${NATIVE_SHEET_GEOMETRY_LIMITS.maxCommands}`)
  }
  const commands: NativeSheetGeometryCommandV1[] = []
  let finished: ReadonlyArray<NativeSheetGeometryCommandV1> | undefined
  return {
    command_capacity: maxCommands,
    get commands() { return finished ?? Object.freeze([...commands]) },
    push(command) {
      if (finished) throw new Error('native sheet geometry recording is already finished')
      if (commands.length >= maxCommands) throw new NativeSheetGeometryError('geometry.commandBudget', '$.commands', `geometry commands exceed ${maxCommands}`)
      commands.push(deepFreeze(command))
    },
    finish() {
      if (!finished) finished = Object.freeze([...commands])
      return finished
    },
  }
}

export function emitNativeSheetGeometryCommandsV1(geometry: NativeSheetGeometryV1, surface: NativeSheetGeometrySurfaceV1, maxCommands: number = NATIVE_SHEET_GEOMETRY_LIMITS.maxCommands): void {
  if (!Number.isSafeInteger(maxCommands) || maxCommands < 1 || maxCommands > NATIVE_SHEET_GEOMETRY_LIMITS.maxCommands) {
    throw new NativeSheetGeometryError('geometry.invalidLimit', '$.maxCommands', `maxCommands must be from 1 through ${NATIVE_SHEET_GEOMETRY_LIMITS.maxCommands}`)
  }
  if (surface.command_capacity !== undefined && (!Number.isSafeInteger(surface.command_capacity) || surface.command_capacity < 1 || surface.command_capacity > NATIVE_SHEET_GEOMETRY_LIMITS.maxCommands)) {
    throw new NativeSheetGeometryError('geometry.invalidLimit', '$.surface.command_capacity', 'surface command capacity is outside native geometry limits')
  }
  const capacity = Math.min(maxCommands, surface.command_capacity ?? maxCommands)
  const requiredCommands = 3 + geometry.rows.length + geometry.columns.length + geometry.merged_ranges.length
  if (requiredCommands > capacity) throw new NativeSheetGeometryError('geometry.commandBudget', '$.commands', `geometry commands require ${requiredCommands}, exceeding ${capacity}`)
  const push = (command: NativeSheetGeometryCommandV1): void => surface.push(command)
  push({ kind: 'beginSheet', document_id: geometry.document_id, sheet_id: geometry.sheet_id, geometry_sha256: geometry.geometry_sha256, bounds: geometry.bounds })
  push({ kind: 'clipRect', rect: geometry.bounds })
  for (const band of geometry.rows) push({ kind: 'rowBand', band })
  for (const band of geometry.columns) push({ kind: 'columnBand', band })
  for (const range of geometry.merged_ranges) push({ kind: 'mergedRect', range })
  push({ kind: 'endSheet' })
}

export function replayNativeSheetGeometryCommandsV1<HostContext>(context: HostContext, commands: ReadonlyArray<NativeSheetGeometryCommandV1>, adapter: NativeSheetGeometryCommandAdapterV1<HostContext>): void {
  for (const command of commands) adapter.execute(context, command)
}

/** ISO/IEC 29500 stored character width to runtime grid pixels. */
export function characterWidthToPixels(width: number, maximumDigitWidthPixels: number): number {
  if (!Number.isFinite(width) || width < 0 || width > 255) throw new NativeSheetGeometryError('geometry.metricUnavailable', '$.width', 'column width must be finite and within 0..255')
  validateMdw(maximumDigitWidthPixels, '$.maximumDigitWidthPixels')
  return width === 0 ? 0 : Math.floor(((256 * width + Math.floor(128 / maximumDigitWidthPixels)) / 256) * maximumDigitWidthPixels)
}

/** Default base character count plus Excel's five-pixel cell padding, snapped down to 1/256. */
export function paddedBaseColumnWidth(baseColumnWidth: number, maximumDigitWidthPixels: number): number {
  if (!Number.isSafeInteger(baseColumnWidth) || baseColumnWidth < 0 || baseColumnWidth > 255) throw new NativeSheetGeometryError('geometry.sheetFormatUnavailable', '$.sheet.sheet_format.base_column_width', 'base column width must be an integer within 0..255')
  validateMdw(maximumDigitWidthPixels, '$.metric_authority.maximum_digit_width_pixels')
  return Math.floor((baseColumnWidth + 5 / maximumDigitWidthPixels) * 256) / 256
}

function validateViewport(viewport: NativeSheetViewportV1): void {
  const values = [viewport.row, viewport.column, viewport.end_row, viewport.end_column]
  if (values.some((value) => !Number.isSafeInteger(value) || Object.is(value, -0)) || viewport.row < 0 || viewport.column < 0 || viewport.end_row < viewport.row || viewport.end_column < viewport.column || viewport.end_row >= EXCEL_MAX_ROWS || viewport.end_column >= EXCEL_MAX_COLUMNS) {
    throw new NativeSheetGeometryError('geometry.invalidViewport', '$.viewport', 'viewport must be a forward bounded zero-based Excel rectangle')
  }
  const rowCount = viewport.end_row - viewport.row + 1
  const columnCount = viewport.end_column - viewport.column + 1
  if (rowCount > NATIVE_SHEET_GEOMETRY_LIMITS.maxViewportRows || columnCount > NATIVE_SHEET_GEOMETRY_LIMITS.maxViewportColumns || rowCount * columnCount > NATIVE_SHEET_GEOMETRY_LIMITS.maxViewportCells) {
    throw new NativeSheetGeometryError('geometry.viewportBudget', '$.viewport', 'viewport exceeds native geometry resource limits')
  }
}

function validateMetricAuthority(workbook: NativeWorkbookRenderModelV1, authority: NativeMaximumDigitWidthAuthorityV1): void {
  if (authority.source_revision !== workbook.revision || authority.source_package_sha256 !== workbook.source.package_sha256) {
    throw new NativeSheetGeometryError('geometry.metricAuthority', '$.metric_authority', 'maximum digit width is not bound to this exact workbook revision and Normal style')
  }
  const normal = workbook.normal_style
  if (!normal) {
    throw new NativeSheetGeometryError('geometry.metricUnavailable', '$.normal_style', 'the built-in Normal-style font identity is unavailable')
  }
  if (authority.normal_style_xf_id !== normal.style_xf_id || authority.normal_style_font_id !== normal.font_id || authority.font_name !== normal.font_name || authority.font_size_points !== normal.font_size_points || authority.font_bold !== normal.font_bold || authority.font_italic !== normal.font_italic || authority.normal_font_record_sha256 !== normal.font_record_sha256) {
    throw new NativeSheetGeometryError('geometry.metricAuthority', '$.metric_authority', 'maximum digit width font identity disagrees with the projected Normal style')
  }
  if (!/^sha256:[0-9a-f]{64}$/.test(authority.normal_font_record_sha256) || !/^sha256:[0-9a-f]{64}$/.test(authority.font_sha256) || !boundedIdentifier(authority.provider_id) || !boundedIdentifier(authority.provider_revision) || authority.measurement_dpi !== 96) {
    throw new NativeSheetGeometryError('geometry.metricAuthority', '$.metric_authority', 'font digest and provider provenance are required')
  }
  validateMdw(authority.maximum_digit_width_pixels, '$.metric_authority.maximum_digit_width_pixels')
}

function validateMdw(value: number, path: string): void {
  if (!Number.isSafeInteger(value) || value < 1 || value > 512) throw new NativeSheetGeometryError('geometry.metricUnavailable', path, 'maximum digit width must be an integer from 1 through 512 pixels')
}

function boundedIdentifier(value: string): boolean { return value.length >= 1 && value.length <= 256 && !/[\u0000-\u001f\u007f]/.test(value) }

function compileMergedGeometry(sheet: NativeSheetRenderModelV1, viewport: NativeSheetViewportV1, rows: NativeSheetRowBandV1[], columns: NativeSheetColumnBandV1[]): NativeSheetMergedGeometryV1[] {
  const result: NativeSheetMergedGeometryV1[] = []
  for (const range of sheet.merged_ranges) {
    const intersects = range.row <= viewport.end_row && range.end_row >= viewport.row && range.column <= viewport.end_column && range.end_column >= viewport.column
    if (!intersects) continue
    if (range.row < viewport.row || range.end_row > viewport.end_row || range.column < viewport.column || range.end_column > viewport.end_column) {
      throw new NativeSheetGeometryError('geometry.mergeClipped', `$.sheet.merged_ranges[${result.length}]`, `viewport clips merged range ${range.ref}`)
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
  if (!Number.isSafeInteger(value) || value < 0) throw new NativeSheetGeometryError('geometry.integerOverflow', path, 'geometry exceeds non-negative safe integer EMU bounds')
  return value
}

function checkedSum(left: number, right: number, path: string): number { return checkedInteger(left + right, path) }

function deepFreeze<T>(value: T): T {
  if (typeof value !== 'object' || value === null || Object.isFrozen(value)) return value
  for (const child of Object.values(value)) deepFreeze(child)
  return Object.freeze(value)
}
