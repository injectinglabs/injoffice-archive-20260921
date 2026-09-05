import { describe, expect, it } from 'vitest'
import { invalidateMode, rangesIntersect, sourceNeedsRefresh } from './invalidate'
import type { CellRangeRef } from './types'

const source = (over: Partial<CellRangeRef> = {}): CellRangeRef => ({
  sheetId: 's1',
  startRow: 0,
  startColumn: 0,
  endRow: 3,
  endColumn: 2,
  ...over,
})

const setRange = (cellValue: Record<string, Record<string, unknown>>, subUnitId = 's1') => ({
  id: 'sheet.mutation.set-range-values',
  type: 2,
  params: { unitId: 'wb1', subUnitId, cellValue },
})

describe('rangesIntersect', () => {
  it('requires the same sheet', () => {
    expect(rangesIntersect(source(), source({ sheetId: 's2' }))).toBe(false)
  })
  it('detects overlap and disjoint blocks', () => {
    expect(rangesIntersect(source(), source({ startRow: 3, startColumn: 2, endRow: 5, endColumn: 4 }))).toBe(true)
    expect(rangesIntersect(source(), source({ startRow: 4, startColumn: 0, endRow: 5, endColumn: 2 }))).toBe(false)
    expect(rangesIntersect(source(), source({ startRow: 0, startColumn: 3, endRow: 3, endColumn: 4 }))).toBe(false)
  })
})

describe('invalidateMode', () => {
  it('skips selection/UI commands', () => {
    expect(invalidateMode({ id: 'sheet.operation.set-selections', params: { subUnitId: 's1' } })).toEqual({ kind: 'skip' })
    expect(invalidateMode({ id: 'sheet.command.set-range-values', type: 0, params: { range: source() } })).toEqual({
      kind: 'skip',
    })
  })

  it('extracts the bounding box of a set-range-values cell matrix', () => {
    expect(invalidateMode(setRange({ 1: { 2: { v: 9 }, 4: { v: 1 } }, 5: { 2: { v: 0 } } }))).toEqual({
      kind: 'ranges',
      ranges: [{ sheetId: 's1', startRow: 1, startColumn: 2, endRow: 5, endColumn: 4 }],
    })
  })

  it('treats insert/remove row-col as a whole-sheet invalidation', () => {
    expect(
      invalidateMode({
        id: 'sheet.mutation.insert-row',
        type: 2,
        params: { subUnitId: 's1', range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 } },
      }),
    ).toEqual({ kind: 'sheet', sheetId: 's1' })
  })

  it('uses from/to endpoints on move-range', () => {
    const mode = invalidateMode({
      id: 'sheet.mutation.move-range',
      params: {
        unitId: 'wb1',
        fromRange: { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 },
        toRange: { startRow: 10, startColumn: 10, endRow: 11, endColumn: 11 },
        from: { subUnitId: 's1', value: { 0: { 0: { v: 1 } } } },
        to: { subUnitId: 's2', value: { 10: { 10: { v: 1 } } } },
      },
    })
    expect(mode.kind).toBe('ranges')
    if (mode.kind !== 'ranges') return
    expect(mode.ranges.some((r) => r.sheetId === 's1' && r.startRow === 0)).toBe(true)
    expect(mode.ranges.some((r) => r.sheetId === 's2' && r.startRow === 10)).toBe(true)
    expect(mode.ranges.some((r) => r.sheetId === 's1' && r.startRow === 10)).toBe(false)
  })
})

describe('sourceNeedsRefresh', () => {
  it('refreshes only when the edited cells intersect the source', () => {
    const src = source()
    expect(sourceNeedsRefresh(src, invalidateMode(setRange({ 1: { 1: { v: 4 } } })))).toBe(true)
    expect(sourceNeedsRefresh(src, invalidateMode(setRange({ 10: { 10: { v: 4 } } })))).toBe(false)
    expect(sourceNeedsRefresh(src, invalidateMode(setRange({ 1: { 1: { v: 4 } } }, 's2')))).toBe(false)
    expect(sourceNeedsRefresh(src, { kind: 'skip' })).toBe(false)
    expect(sourceNeedsRefresh(src, { kind: 'sheet', sheetId: 's1' })).toBe(true)
    expect(sourceNeedsRefresh(src, { kind: 'sheet', sheetId: 's2' })).toBe(false)
    expect(sourceNeedsRefresh(src, { kind: 'all' })).toBe(true)
  })
})
