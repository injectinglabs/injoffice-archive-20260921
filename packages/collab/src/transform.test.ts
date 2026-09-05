import { describe, expect, it } from 'vitest'
import { MOVE_RANGE, type OpRecord, SET_RANGE_VALUES } from './sync'
import { mapIndex, mapIndexThroughMove, mapSpan, mapSpanThroughMove, parseMove, parseStruct, structEditsOf, transformOp, transformOps, transformSelection, transformSelectionThroughMove, type MoveEdit, type StructEdit } from './transform'

const range = (r0: number, r1: number, c0 = 0, c1 = 0) => ({ startRow: r0, endRow: r1, startColumn: c0, endColumn: c1 })
const insRows = (s: number, n: number): OpRecord => ({ id: 'sheet.mutation.insert-row', params: { unitId: 'wb', subUnitId: 's1', range: range(s, s + n - 1, 0, 25) } })
const remRows = (s: number, n: number): OpRecord => ({ id: 'sheet.mutation.remove-rows', params: { unitId: 'wb', subUnitId: 's1', range: range(s, s + n - 1, 0, 25) } })
const insCols = (s: number, n: number): OpRecord => ({ id: 'sheet.mutation.insert-col', params: { unitId: 'wb', subUnitId: 's1', range: { startRow: 0, endRow: 99, startColumn: s, endColumn: s + n - 1 } } })
const cells = (m: Record<string, Record<string, unknown>>, sub = 's1'): OpRecord => ({ id: SET_RANGE_VALUES, params: { unitId: 'wb', subUnitId: sub, cellValue: m } })
const moveRows = (from: number, to: number, count = 1): OpRecord => ({ id: 'sheet.mutation.move-rows', params: { unitId: 'wb', subUnitId: 's1', sourceRange: range(from, from + count - 1, 0, 25), targetRange: range(to, to + count - 1, 0, 25) } })
const moveCols = (from: number, to: number, count = 1): OpRecord => ({ id: 'sheet.mutation.move-columns', params: { unitId: 'wb', subUnitId: 's1', sourceRange: range(0, 99, from, from + count - 1), targetRange: range(0, 99, to, to + count - 1) } })
const moveRange = (): OpRecord => ({
  id: MOVE_RANGE,
  params: {
    unitId: 'wb', fromRange: range(1, 1), toRange: range(5, 5),
    from: { subUnitId: 's1', value: { 1: { 0: null } } },
    to: { subUnitId: 's1', value: { 5: { 0: { v: 'moved' } } } },
  },
})

const edit = (op: OpRecord): StructEdit => parseStruct(op)!
const move = (op: OpRecord): MoveEdit => parseMove(op)!

describe('parseStruct / mapIndex / mapSpan', () => {
  it('normalizes the four structural mutations', () => {
    expect(edit(insRows(3, 2))).toEqual({ kind: 'insert', axis: 'row', subUnitId: 's1', start: 3, count: 2 })
    expect(edit(remRows(5, 3))).toEqual({ kind: 'remove', axis: 'row', subUnitId: 's1', start: 5, count: 3 })
    expect(edit(insCols(1, 1))).toEqual({ kind: 'insert', axis: 'col', subUnitId: 's1', start: 1, count: 1 })
    expect(parseStruct(cells({}))).toBeNull()
  })
  it('maps indices through inserts and removals', () => {
    const ins = edit(insRows(3, 2))
    expect([mapIndex(ins, 2), mapIndex(ins, 3), mapIndex(ins, 10)]).toEqual([2, 5, 12])
    const rem = edit(remRows(3, 2))
    expect([mapIndex(rem, 2), mapIndex(rem, 3), mapIndex(rem, 4), mapIndex(rem, 5)]).toEqual([2, null, null, 3])
  })
  it('maps spans: shift, widen, clamp, consume', () => {
    const ins = edit(insRows(3, 2))
    expect(mapSpan(ins, 0, 1)).toEqual([0, 1])
    expect(mapSpan(ins, 4, 6)).toEqual([6, 8])
    expect(mapSpan(ins, 1, 5)).toEqual([1, 7]) // widened: content grew inside
    const rem = edit(remRows(3, 2))
    expect(mapSpan(rem, 5, 8)).toEqual([3, 6])
    expect(mapSpan(rem, 3, 4)).toBeNull()
    expect(mapSpan(rem, 2, 6)).toEqual([2, 4]) // inside chopped
    expect(mapSpan(rem, 0, 3)).toEqual([0, 2]) // trailing chopped
    expect(mapSpan(rem, 4, 8)).toEqual([3, 6]) // leading chopped
  })
})

describe('row and column move permutations', () => {
  it('normalizes only equal-sized, non-overlapping moves', () => {
    expect(move(moveRows(2, 7, 2))).toEqual({ kind: 'move', axis: 'row', subUnitId: 's1', sourceStart: 2, sourceEnd: 3, targetStart: 7, targetEnd: 8 })
    expect(move(moveCols(6, 1, 3))).toMatchObject({ axis: 'col', sourceStart: 6, sourceEnd: 8, targetStart: 1, targetEnd: 3 })
    expect(parseMove({ id: 'sheet.mutation.move-rows', params: { subUnitId: 's1', sourceRange: range(2, 3), targetRange: range(5, 5) } })).toBeNull()
    expect(parseMove(moveRows(2, 3, 2))).toBeNull()
  })

  it('maps indices and disjoint span unions in both directions', () => {
    const down = move(moveRows(2, 7, 2))
    expect([1, 2, 3, 4, 8, 9].map((index) => mapIndexThroughMove(down, index))).toEqual([1, 7, 8, 2, 6, 9])
    expect(mapSpanThroughMove(down, 2, 4)).toEqual([[2, 2], [7, 8]])
    expect(mapSpanThroughMove(down, 2, 8)).toEqual([[2, 8]])
    const up = move(moveRows(7, 2, 2))
    expect([1, 2, 6, 7, 8, 9].map((index) => mapIndexThroughMove(up, index))).toEqual([1, 4, 8, 2, 3, 9])
  })

  it('rebases sparse cell writes exactly through a move', () => {
    const result = transformOps([cells({ 2: { 0: { v: 'moved' } }, 4: { 0: { v: 'shifted' } }, 9: { 0: { v: 'still' } } })], [moveRows(2, 7, 2)])
    expect(result.lossy).toBe(false)
    expect((result.ops[0]!.params as { cellValue: unknown }).cellValue).toEqual({ 7: { 0: { v: 'moved' } }, 2: { 0: { v: 'shifted' } }, 9: { 0: { v: 'still' } } })
  })

  it('splits multi-range mutations and selections instead of widening them', () => {
    const moved = move(moveRows(2, 7, 2))
    const op: OpRecord = { id: 'sheet.mutation.add-worksheet-merge', params: { subUnitId: 's1', ranges: [range(2, 4, 1, 3)] } }
    const result = transformOps([op], [moveRows(2, 7, 2)])
    expect(result.lossy).toBe(false)
    expect((result.ops[0]!.params as { ranges: unknown }).ranges).toEqual([range(2, 2, 1, 3), range(7, 8, 1, 3)])
    const selection = transformSelectionThroughMove({ sheet: 's1', ranges: [[2, 1, 4, 3]], active: [3, 2] }, moved)
    expect(selection.ranges).toEqual([[2, 1, 2, 3], [7, 1, 8, 3]])
    expect(selection.active).toEqual([8, 2])
  })

  it('flags a single-range mutation that becomes disjoint', () => {
    const op: OpRecord = { id: 'sheet.mutation.set-row-height', params: { subUnitId: 's1', range: range(2, 4) } }
    const result = transformOps([op], [moveRows(2, 7, 2)])
    expect(result).toEqual({ ops: [op], lossy: true })
  })

  it('reports valid moves as structural selection edits', () => {
    expect(structEditsOf([moveRows(2, 7, 2)])).toMatchObject({ edits: [], moves: [move(moveRows(2, 7, 2))], removedSheets: [], lossy: false })
  })
})

describe('transformOp', () => {
  it('shifts cell writes below an insert and drops removed rows', () => {
    const w = cells({ '2': { '0': { v: 'a' } }, '5': { '1': { v: 'b' } } })
    const shifted = transformOp(w, edit(insRows(3, 2)))!
    expect((shifted.params as { cellValue: unknown }).cellValue).toEqual({ '2': { '0': { v: 'a' } }, '7': { '1': { v: 'b' } } })
    const afterRemove = transformOp(w, edit(remRows(5, 1)))!
    expect((afterRemove.params as { cellValue: unknown }).cellValue).toEqual({ '2': { '0': { v: 'a' } } })
    expect(transformOp(cells({ '5': { '1': { v: 'b' } } }), edit(remRows(5, 1)))).toBeNull()
  })
  it('shifts columns independently of rows', () => {
    const w = cells({ '2': { '3': { v: 'a' } } })
    const shifted = transformOp(w, edit(insCols(1, 2)))!
    expect((shifted.params as { cellValue: unknown }).cellValue).toEqual({ '2': { '5': { v: 'a' } } })
  })
  it('leaves other sheets alone', () => {
    const w = cells({ '9': { '0': { v: 'x' } } }, 'other')
    expect(transformOp(w, edit(insRows(0, 5)))).toBe(w)
  })
  it('transforms structural vs structural on the same axis', () => {
    // pending insert at 10 after incoming insert of 2 at 3 → anchor 12
    const t1 = transformOp(insRows(10, 1), edit(insRows(3, 2)))!
    expect(edit(t1).start).toBe(12)
    // pending remove [5,8] after incoming remove [6,7] → remove [5,6]
    const t2 = transformOp(remRows(5, 4), edit(remRows(6, 2)))!
    expect(edit(t2)).toMatchObject({ start: 5, count: 2 })
    // pending remove fully consumed
    expect(transformOp(remRows(6, 2), edit(remRows(5, 4)))).toBeNull()
    // pending insert into removed span lands at the cut
    const t3 = transformOp(insRows(6, 1), edit(remRows(5, 4)))!
    expect(edit(t3).start).toBe(5)
    // different axes don't interact
    const colIns = insCols(2, 1)
    expect(transformOp(colIns, edit(insRows(0, 3)))).toBe(colIns)
    const moved = transformOp(moveRows(5, 10, 2), edit(insRows(0, 2)))!
    expect(parseMove(moved)).toMatchObject({ sourceStart: 7, sourceEnd: 8, targetStart: 12, targetEnd: 13 })
    const ambiguous = transformOps([moveRows(5, 10, 2)], [insRows(6, 1)])
    expect(ambiguous.ops).toEqual([moveRows(5, 10, 2)])
    expect(ambiguous.lossy).toBe(true)
  })
  it('shifts generic range-bearing params (merges)', () => {
    const merge: OpRecord = { id: 'sheet.mutation.add-worksheet-merge', params: { unitId: 'wb', subUnitId: 's1', ranges: [range(4, 5, 0, 2), range(1, 1, 0, 0)] } }
    const t = transformOp(merge, edit(insRows(2, 3)))!
    expect((t.params as { ranges: unknown }).ranges).toEqual([range(7, 8, 0, 2), range(1, 1, 0, 0)])
    const gone = transformOp({ id: 'x.merge', params: { subUnitId: 's1', range: range(3, 4) } }, edit(remRows(3, 2)))
    expect(gone).toBeNull()
  })
})

describe('transformOps', () => {
  it('applies every structural edit in order and reports lossiness', () => {
    const pending = [cells({ '10': { '0': { v: 'p' } } })]
    const incoming = [insRows(0, 1), remRows(3, 1), { id: 'sheet.mutation.move-range', params: { unitId: 'wb', subUnitId: 's1', sourceRange: range(0, 0), targetRange: range(9, 9) } }]
    const res = transformOps(pending, incoming)
    // +1 for the insert at 0, -1 for the removal above → net 10
    expect((res.ops[0].params as { cellValue: unknown }).cellValue).toEqual({ '10': { '0': { v: 'p' } } })
    expect(res.lossy).toBe(true)
  })
  it('drops ops on a removed sheet', () => {
    const res = transformOps([cells({ '1': { '1': { v: 1 } } }, 'doomed'), cells({ '1': { '1': { v: 1 } } }, 'safe')], [
      { id: 'sheet.mutation.remove-sheet', params: { unitId: 'wb', subUnitId: 'doomed', subUnitName: 'D' } },
    ])
    expect(res.ops).toHaveLength(1)
    expect((res.ops[0].params as { subUnitId: string }).subUnitId).toBe('safe')
  })

  it('treats valid content moves as non-geometric and fails closed against structure', () => {
    expect(transformOps([cells({ 9: { 0: { v: 'pending' } } })], [moveRange()])).toEqual({ ops: [cells({ 9: { 0: { v: 'pending' } } })], lossy: false })
    expect(structEditsOf([moveRange()])).toMatchObject({ edits: [], moves: [], lossy: false })
    expect(transformOps([moveRange()], [insRows(0, 1)])).toEqual({ ops: [moveRange()], lossy: true })
  })
})

describe('transformSelection', () => {
  it('shifts, clamps and drops ranges; keeps other sheets', () => {
    const sel = { sheet: 's1', ranges: [[5, 0, 8, 2], [0, 0, 1, 1]], active: [5, 0] as [number, number], mode: 'editing' as const, draft: 'live' }
    const shifted = transformSelection(sel, edit(insRows(3, 2)), 's1')
    expect(shifted.ranges).toEqual([[7, 0, 10, 2], [0, 0, 1, 1]])
    expect(shifted.active).toEqual([7, 0])
    expect(shifted.mode).toBe('editing')
    expect(shifted.draft).toBe('live')
    const afterRemove = transformSelection(sel, edit(remRows(5, 4)), 's1')
    expect(afterRemove.ranges).toEqual([[0, 0, 1, 1]])
    expect(afterRemove.active).toBeUndefined()
    expect(transformSelection(sel, edit(insRows(0, 1)), 'other')).toBe(sel)
  })
})
