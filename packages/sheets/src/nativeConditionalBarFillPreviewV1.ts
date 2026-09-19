import { isNativePreviewPartPathV1 } from './nativePreviewPartPathV1.js'
import { snapshotNativePlainData } from './nativePlainData.js'
import type { NativeWorkbookObjectsV1 } from './nativeObjectsPreviewV1.js'

/** One source range a data bar was read from. Disclosure only. */
export interface NativeConditionalBarFillRangeV1 {
  ref: string
  priority: number
}
/** One painted bar. Coordinates are zero-based; the span and the axis are
 * thousandths of the cell's width, measured from its left edge. A negative
 * axis position means the rule draws no axis. */
export interface NativeConditionalBarFillCellV1 {
  row: number
  column: number
  start_permille: number
  end_permille: number
  axis_permille: number
  color: string
  border_color?: string
  axis_color?: string
}
export type NativeConditionalBarFillPreviewV1 = {
  sheet_id: string
  sheet_part: string
  warnings: string[]
} & (
  | { status: 'available'; cells: NativeConditionalBarFillCellV1[]; ranges: NativeConditionalBarFillRangeV1[] }
  | { status: 'unavailable'; cells?: never; ranges?: never }
)

const RANGE = /^([A-Z]{1,3})([1-9][0-9]{0,6}):([A-Z]{1,3})([1-9][0-9]{0,6})$/
const COLOR = /^#[0-9A-F]{6}$/
const MAX_CELLS = 16384

/** Validate the additive source projection. Lengths are read as written; no
 * worksheet formula is evaluated and no source style is changed. */
export function decodeNativeConditionalBarFillPreviewsV1(input: unknown): NativeConditionalBarFillPreviewV1[] {
  const value = snapshotNativePlainData(input, { maxDepth: 8, maxNodes: 200000 })
  const fail = (): never => { throw new TypeError('Invalid data-bar fill preview') }
  const exact = (candidate: unknown, keys: string[]) => {
    if (!candidate || typeof candidate !== 'object' || Array.isArray(candidate)) return fail()
    const object = candidate as Record<string, unknown>
    if (Object.keys(object).length !== keys.length || keys.some(key => !Object.hasOwn(object, key))) return fail()
    return object
  }
  const permille = (candidate: unknown) => (Number.isInteger(candidate) && !Object.is(candidate, -0) && Number(candidate) >= 0 && Number(candidate) <= 1000 ? Number(candidate) : fail())
  if (!Array.isArray(value) || value.length > 64) return fail()
  const ids = new Set<string>(), parts = new Set<string>()
  let total = 0
  return value.map(raw => {
    const status = (raw as Record<string, unknown> | undefined)?.status
    const entry = exact(raw, ['sheet_id', 'sheet_part', 'status', 'warnings', ...(status === 'available' ? ['cells', 'ranges'] : [])])
    if (typeof entry.sheet_id !== 'string' || !/^[1-9][0-9]{0,9}$/.test(entry.sheet_id) || Number(entry.sheet_id) > 0xffffffff || ids.has(entry.sheet_id) || !isNativePreviewPartPathV1(entry.sheet_part) || parts.has(entry.sheet_part)) return fail()
    ids.add(entry.sheet_id); parts.add(entry.sheet_part as string)
    if (!Array.isArray(entry.warnings) || entry.warnings.length < 1 || entry.warnings.length > 8 || entry.warnings.some(warning => typeof warning !== 'string' || !warning || warning.length > 4096 || /[\u0000-\u001f\u007f]/.test(warning))) return fail()
    const base = { sheet_id: entry.sheet_id, sheet_part: entry.sheet_part as string, warnings: entry.warnings as string[] }
    if (status === 'unavailable') return { ...base, status }
    if (status !== 'available') return fail()
    if (!Array.isArray(entry.ranges) || entry.ranges.length < 1 || entry.ranges.length > 64) return fail()
    const ranges = entry.ranges.map(candidate => {
      const source = exact(candidate, ['ref', 'priority'])
      if (typeof source.ref !== 'string' || !RANGE.test(source.ref) || !Number.isInteger(source.priority) || Number(source.priority) < 1 || Number(source.priority) > 2147483647) return fail()
      return { ref: source.ref, priority: Number(source.priority) }
    })
    if (!Array.isArray(entry.cells) || entry.cells.length < 1 || (total += entry.cells.length) > MAX_CELLS) return fail()
    let previous = -1
    const cells = entry.cells.map(candidate => {
      const object = candidate as Record<string, unknown>
      const hasBorder = !!object && typeof object === 'object' && Object.hasOwn(object, 'border_color')
      const hasAxis = !!object && typeof object === 'object' && Object.hasOwn(object, 'axis_color')
      const cell = exact(candidate, ['row', 'column', 'start_permille', 'end_permille', 'axis_permille', 'color', ...(hasBorder ? ['border_color'] : []), ...(hasAxis ? ['axis_color'] : [])])
      if (!Number.isInteger(cell.row) || !Number.isInteger(cell.column) || Object.is(cell.row, -0) || Object.is(cell.column, -0) || Number(cell.row) < 0 || Number(cell.row) > 1048575 || Number(cell.column) < 0 || Number(cell.column) > 16383) return fail()
      const start = permille(cell.start_permille), end = permille(cell.end_permille)
      // A reversed span would paint a negative-width rectangle, which some
      // renderers silently flip and others drop.
      if (end < start) return fail()
      const axis = cell.axis_permille === -1 ? -1 : permille(cell.axis_permille)
      if (axis === -1 ? hasAxis : !hasAxis) return fail()
      if (typeof cell.color !== 'string' || !COLOR.test(cell.color)) return fail()
      if (hasBorder && (typeof cell.border_color !== 'string' || !COLOR.test(cell.border_color))) return fail()
      if (hasAxis && (typeof cell.axis_color !== 'string' || !COLOR.test(cell.axis_color))) return fail()
      const position = Number(cell.row) * 16384 + Number(cell.column)
      if (position <= previous) return fail()
      previous = position
      return {
        row: Number(cell.row), column: Number(cell.column), start_permille: start, end_permille: end, axis_permille: axis, color: cell.color,
        ...(hasBorder ? { border_color: cell.border_color as string } : {}),
        ...(hasAxis ? { axis_color: cell.axis_color as string } : {}),
      }
    })
    return { ...base, status, cells, ranges }
  })
}

const indexes = new WeakMap<NativeWorkbookObjectsV1, Map<string, Map<number, NativeConditionalBarFillCellV1>>>()

function barIndex(objects: NativeWorkbookObjectsV1): Map<string, Map<number, NativeConditionalBarFillCellV1>> {
  const cached = indexes.get(objects)
  if (cached) return cached
  const index = new Map<string, Map<number, NativeConditionalBarFillCellV1>>()
  for (const entry of decodeNativeConditionalBarFillPreviewsV1(objects.conditional_bar_fills ?? [])) {
    if (entry.status !== 'available') continue
    const sheet = new Map<number, NativeConditionalBarFillCellV1>()
    for (const cell of entry.cells) sheet.set(cell.row * 16384 + cell.column, cell)
    index.set(entry.sheet_part, sheet)
  }
  indexes.set(objects, index)
  return index
}

/** Read-only data-bar decoration. Coordinates are zero-based. The bar is drawn
 * over the cell's fill and under its text, which is the order Excel paints. */
export function nativeConditionalBarFillPreview(objects: NativeWorkbookObjectsV1, revision: string, sheetPart: string, row: number, column: number): NativeConditionalBarFillCellV1 | undefined {
  if (objects.package_sha256 !== revision || !Number.isSafeInteger(row) || !Number.isSafeInteger(column) || row < 0 || column < 0 || row > 1048575 || column > 16383) return undefined
  return barIndex(objects).get(sheetPart)?.get(row * 16384 + column)
}
