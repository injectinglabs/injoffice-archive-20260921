import { describe, expect, it } from 'vitest'
import { invalidateMode, rangesIntersect, sourceNeedsRefresh } from './invalidate'
import type { CellRangeRef } from './types'

const source = (over: Partial<CellRangeRef> = {}): CellRangeRef => ({
  sheetId: 's1',
  startRow: 0,
  startColumn: 0,
  endRow: 6,
  endColumn: 3,
  ...over,
})

const setRange = (cellValue: Record<string, Record<string, unknown>>, subUnitId = 's1') => ({
  id: 'sheet.mutation.set-range-values',
  type: 2,
  params: { unitId: 'wb1', subUnitId, cellValue },
})

describe('rangesIntersect', () => {
  it('does not treat a baked target block as overlapping a disjoint source', () => {
    const src = source()
    const target = source({ startColumn: 5, endColumn: 8 })
    expect(rangesIntersect(src, target)).toBe(false)
  })
})

describe('invalidateMode', () => {
  it('skips non-mutation traffic', () => {
    expect(invalidateMode({ id: 'sheet.operation.set-selections' })).toEqual({ kind: 'skip' })
  })

  it('extracts edited cells from set-range-values payloads', () => {
    expect(invalidateMode(setRange({ 2: { 1: { v: 50 } } }))).toEqual({
      kind: 'ranges',
      ranges: [{ sheetId: 's1', startRow: 2, startColumn: 1, endRow: 2, endColumn: 1 }],
    })
  })

  it('invalidates the whole sheet on row insert', () => {
    expect(
      invalidateMode({
        id: 'sheet.mutation.insert-row',
        params: { subUnitId: 's1', range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 } },
      }),
    ).toEqual({ kind: 'sheet', sheetId: 's1' })
  })

  it('uses from/to endpoints on move-range without attributing toRange to the from sheet', () => {
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
  it('ignores edits that only hit a pivot target, not its source', () => {
    const src = source()
    expect(sourceNeedsRefresh(src, invalidateMode(setRange({ 0: { 1: { v: 9 } } })))).toBe(true)
    expect(sourceNeedsRefresh(src, invalidateMode(setRange({ 0: { 6: { v: 9 } } })))).toBe(false)
  })
})
