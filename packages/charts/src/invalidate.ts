import type { CellRangeRef } from './types'

// Keep in sync with packages/pivots/src/invalidate.ts — packages stay independently consumable.

// Range-aware invalidation for ChartManager.
//
// Univer's command bus fires on every selection, scroll, and style tweak.
// Rebuilding every chart on that stream is wasted work: a chart only needs a
// new option when the cells it reads actually changed (or the sheet's
// geometry shifted underneath it). Structural mutations (insert/remove/move
// rows or columns) can shift values inside a source block without the
// mutation's own range intersecting it, so those refresh every chart on the
// affected sheet.

export interface CommandInfoLike {
  id?: string
  type?: number
  params?: unknown
}

export type InvalidateMode =
  | { kind: 'skip' }
  | { kind: 'all' }
  | { kind: 'sheet'; sheetId: string }
  | { kind: 'ranges'; ranges: CellRangeRef[] }

const MUTATION = 2

const STRUCTURAL = new Set([
  'sheet.mutation.insert-row',
  'sheet.mutation.insert-col',
  'sheet.mutation.remove-rows',
  'sheet.mutation.remove-col',
  'sheet.mutation.move-rows',
  'sheet.mutation.move-columns',
  'sheet.mutation.remove-sheet',
  'sheet.mutation.insert-sheet',
])

const VALUE = new Set([
  'sheet.mutation.set-range-values',
  'sheet.mutation.move-range',
  'sheet.mutation.reorder-range',
])

export function rangesIntersect(a: CellRangeRef, b: CellRangeRef): boolean {
  if (a.sheetId !== b.sheetId) return false
  return a.startRow <= b.endRow && a.endRow >= b.startRow && a.startColumn <= b.endColumn && a.endColumn >= b.startColumn
}

export function sourceNeedsRefresh(source: CellRangeRef, mode: InvalidateMode): boolean {
  if (mode.kind === 'skip') return false
  if (mode.kind === 'all') return true
  if (mode.kind === 'sheet') return source.sheetId === mode.sheetId
  return mode.ranges.some((range) => rangesIntersect(source, range))
}

export function invalidateMode(info: CommandInfoLike): InvalidateMode {
  const id = typeof info.id === 'string' ? info.id : ''
  if (!isMutation(id, info.type)) return { kind: 'skip' }
  if (STRUCTURAL.has(id)) {
    const sheetId = sheetIdOf(info.params)
    return sheetId ? { kind: 'sheet', sheetId } : { kind: 'all' }
  }
  if (!VALUE.has(id)) return { kind: 'skip' }
  const ranges = editedRanges(info.params)
  return ranges.length === 0 ? { kind: 'skip' } : { kind: 'ranges', ranges }
}

function isMutation(id: string, type: number | undefined): boolean {
  if (type === MUTATION) return true
  // Tests (and some facade listeners) omit CommandType; id is enough.
  return id.startsWith('sheet.mutation.')
}

function sheetIdOf(params: unknown): string | undefined {
  if (!params || typeof params !== 'object') return undefined
  const p = params as Record<string, unknown>
  if (typeof p.subUnitId === 'string' && p.subUnitId) return p.subUnitId
  if (typeof p.sheetId === 'string' && p.sheetId) return p.sheetId
  const from = p.from
  if (from && typeof from === 'object' && typeof (from as { subUnitId?: unknown }).subUnitId === 'string') {
    return (from as { subUnitId: string }).subUnitId
  }
  return undefined
}

function editedRanges(params: unknown): CellRangeRef[] {
  if (!params || typeof params !== 'object') return []
  const p = params as Record<string, unknown>
  const sheetId = sheetIdOf(p)
  const out: CellRangeRef[] = []

  const fromCellValue = boundingBox(p.cellValue)
  if (fromCellValue && sheetId) out.push({ sheetId, ...fromCellValue })

  pushRange(out, sheetId, p.range)
  if (Array.isArray(p.ranges)) for (const r of p.ranges) pushRange(out, sheetId, r)
  if (Array.isArray(p.rangeList)) for (const r of p.rangeList) pushRange(out, sheetId, r)

  const from = asEndpoint(p.from)
  if (from) {
    const box = boundingBox(from.value)
    if (box) out.push({ sheetId: from.subUnitId, ...box })
    pushRange(out, from.subUnitId, p.fromRange)
  } else {
    pushRange(out, sheetId, p.fromRange)
  }
  const to = asEndpoint(p.to)
  if (to) {
    const box = boundingBox(to.value)
    if (box) out.push({ sheetId: to.subUnitId, ...box })
    pushRange(out, to.subUnitId, p.toRange)
  } else {
    pushRange(out, sheetId, p.toRange)
  }

  return out
}

function asEndpoint(value: unknown): { subUnitId: string; value?: unknown } | null {
  if (!value || typeof value !== 'object') return null
  const p = value as { subUnitId?: unknown; value?: unknown }
  if (typeof p.subUnitId !== 'string' || !p.subUnitId) return null
  return { subUnitId: p.subUnitId, value: p.value }
}

function pushRange(out: CellRangeRef[], sheetId: string | undefined, value: unknown): void {
  const r = asRange(value)
  if (!r) return
  const id = (typeof (value as { sheetId?: unknown }).sheetId === 'string' && (value as { sheetId: string }).sheetId) || sheetId
  if (!id) return
  out.push({ sheetId: id, ...r })
}

function asRange(value: unknown): Omit<CellRangeRef, 'sheetId'> | null {
  if (!value || typeof value !== 'object') return null
  const r = value as Record<string, unknown>
  if (
    typeof r.startRow !== 'number' ||
    typeof r.startColumn !== 'number' ||
    typeof r.endRow !== 'number' ||
    typeof r.endColumn !== 'number'
  ) {
    return null
  }
  return { startRow: r.startRow, startColumn: r.startColumn, endRow: r.endRow, endColumn: r.endColumn }
}

function boundingBox(cellValue: unknown): Omit<CellRangeRef, 'sheetId'> | null {
  if (!cellValue || typeof cellValue !== 'object') return null
  let startRow = Infinity
  let endRow = -Infinity
  let startColumn = Infinity
  let endColumn = -Infinity
  for (const [rowKey, cols] of Object.entries(cellValue as Record<string, unknown>)) {
    const row = Number(rowKey)
    if (!Number.isFinite(row) || !cols || typeof cols !== 'object') continue
    for (const colKey of Object.keys(cols as object)) {
      const col = Number(colKey)
      if (!Number.isFinite(col)) continue
      if (row < startRow) startRow = row
      if (row > endRow) endRow = row
      if (col < startColumn) startColumn = col
      if (col > endColumn) endColumn = col
    }
  }
  if (!Number.isFinite(startRow) || !Number.isFinite(startColumn)) return null
  return { startRow, startColumn, endRow, endColumn }
}
