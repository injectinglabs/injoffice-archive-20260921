import { isNativePreviewPartPathV1 } from './nativePreviewPartPathV1.js'
import { snapshotNativePlainData } from './nativePlainData.js'
import type { NativeWorkbookObjectsV1 } from './nativeObjectsPreviewV1.js'

/** One source range a colour scale was read from. Disclosure only. */
export interface NativeConditionalScaleFillRangeV1 {
  ref: string
  priority: number
  stops: 2 | 3
}
/** One painted cell. Coordinates are zero-based; the colour is opaque. */
export interface NativeConditionalScaleFillCellV1 {
  row: number
  column: number
  color: string
}
export type NativeConditionalScaleFillPreviewV1 = {
  sheet_id: string
  sheet_part: string
  warnings: string[]
} & (
  | { status: 'available'; cells: NativeConditionalScaleFillCellV1[]; ranges: NativeConditionalScaleFillRangeV1[] }
  | { status: 'unavailable'; cells?: never; ranges?: never }
)

const RANGE = /^([A-Z]{1,3})([1-9][0-9]{0,6}):([A-Z]{1,3})([1-9][0-9]{0,6})$/
const MAX_CELLS = 16384

/** Validate the additive source projection. Colours are read as written; no
 * worksheet formula is evaluated and no source style is changed. */
export function decodeNativeConditionalScaleFillPreviewsV1(input: unknown): NativeConditionalScaleFillPreviewV1[] {
  const value = snapshotNativePlainData(input, { maxDepth: 8, maxNodes: 140000 })
  const fail = (): never => { throw new TypeError('Invalid colour-scale fill preview') }
  const exact = (candidate: unknown, keys: string[]) => {
    if (!candidate || typeof candidate !== 'object' || Array.isArray(candidate)) return fail()
    const object = candidate as Record<string, unknown>
    if (Object.keys(object).length !== keys.length || keys.some(key => !Object.hasOwn(object, key))) return fail()
    return object
  }
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
      const source = exact(candidate, ['ref', 'priority', 'stops'])
      if (typeof source.ref !== 'string' || !RANGE.test(source.ref) || !Number.isInteger(source.priority) || Number(source.priority) < 1 || Number(source.priority) > 2147483647 || (source.stops !== 2 && source.stops !== 3)) return fail()
      return { ref: source.ref, priority: Number(source.priority), stops: source.stops as 2 | 3 }
    })
    if (!Array.isArray(entry.cells) || entry.cells.length < 1 || (total += entry.cells.length) > MAX_CELLS) return fail()
    let previous = -1
    const cells = entry.cells.map(candidate => {
      const cell = exact(candidate, ['row', 'column', 'color'])
      if (!Number.isInteger(cell.row) || !Number.isInteger(cell.column) || Object.is(cell.row, -0) || Object.is(cell.column, -0) || Number(cell.row) < 0 || Number(cell.row) > 1048575 || Number(cell.column) < 0 || Number(cell.column) > 16383 || typeof cell.color !== 'string' || !/^#[0-9A-F]{6}$/.test(cell.color)) return fail()
      const position = Number(cell.row) * 16384 + Number(cell.column)
      if (position <= previous) return fail()
      previous = position
      return { row: Number(cell.row), column: Number(cell.column), color: cell.color }
    })
    return { ...base, status, cells, ranges }
  })
}

const indexes = new WeakMap<NativeWorkbookObjectsV1, Map<string, Map<number, string>>>()

function scaleIndex(objects: NativeWorkbookObjectsV1): Map<string, Map<number, string>> {
  const cached = indexes.get(objects)
  if (cached) return cached
  const index = new Map<string, Map<number, string>>()
  for (const entry of decodeNativeConditionalScaleFillPreviewsV1(objects.conditional_scale_fills ?? [])) {
    if (entry.status !== 'available') continue
    const sheet = new Map<number, string>()
    for (const cell of entry.cells) sheet.set(cell.row * 16384 + cell.column, cell.color)
    index.set(entry.sheet_part, sheet)
  }
  indexes.set(objects, index)
  return index
}

/** Read-only colour-scale cell decoration. Coordinates are zero-based. An
 * explicit source cell fill always wins: Excel paints a colour scale behind a
 * cell that has no fill of its own, so a styled cell keeps its own colour. */
export function nativeConditionalScaleFillPreview(objects: NativeWorkbookObjectsV1, revision: string, sheetPart: string, row: number, column: number): string | undefined {
  if (objects.package_sha256 !== revision || !Number.isSafeInteger(row) || !Number.isSafeInteger(column) || row < 0 || column < 0 || row > 1048575 || column > 16383) return undefined
  return scaleIndex(objects).get(sheetPart)?.get(row * 16384 + column)
}
