import { describe, expect, it } from 'vitest'
import { MOVE_RANGE, REORDER_RANGE, SET_RANGE_VALUES, type OpRecord } from './sync'
import { CollaborativeUndoManager } from './undo'

const range = (r0: number, r1: number, c0 = 0, c1 = 10) => ({ startRow: r0, endRow: r1, startColumn: c0, endColumn: c1 })
const cells = (matrix: Record<string, Record<string, unknown>>, sheet = 's1'): OpRecord => ({
  id: SET_RANGE_VALUES,
  params: { unitId: 'wb', subUnitId: sheet, cellValue: matrix },
})
const insertRows = (start: number, count = 1): OpRecord => ({
  id: 'sheet.mutation.insert-row',
  params: { unitId: 'wb', subUnitId: 's1', range: range(start, start + count - 1) },
})
const removeRows = (start: number, count = 1): OpRecord => ({
  id: 'sheet.mutation.remove-rows',
  params: { unitId: 'wb', subUnitId: 's1', range: range(start, start + count - 1) },
})
const moveRows = (from: number, to: number, count = 1): OpRecord => ({
  id: 'sheet.mutation.move-rows',
  params: { unitId: 'wb', subUnitId: 's1', sourceRange: range(from, from + count - 1), targetRange: range(to, to + count - 1) },
})

describe('CollaborativeUndoManager', () => {
  it('plans and commits undo/redo without exposing mutable history state', () => {
    const manager = new CollaborativeUndoManager()
    const undo = cells({ 2: { 1: { v: 'before' } } })
    const redo = cells({ 2: { 1: { v: 'after' } } })
    const id = manager.record({ undoOps: [undo], redoOps: [redo], selectionBefore: { sheet: 's1', ranges: [[2, 1, 2, 1]], active: [2, 1] } })

    const plan = manager.planUndo()
    expect(plan).toMatchObject({ status: 'ready', direction: 'undo', entryId: id, ops: [undo] })
    if (plan.status !== 'ready') throw new Error('expected ready plan')
    ;(plan.ops[0]!.params.cellValue as Record<string, Record<string, { v: string }>>)['2']!['1']!.v = 'tampered'
    expect(manager.commit(plan)).toBe(true)
    expect(manager.undoDepth).toBe(0)
    expect(manager.redoDepth).toBe(1)
    expect(manager.planRedo()).toMatchObject({ status: 'ready', direction: 'redo', ops: [redo] })
  })

  it('invalidates a stale plan when a remote operation arrives', () => {
    const manager = new CollaborativeUndoManager()
    manager.record({ id: 'edit', undoOps: [cells({ 2: { 0: { v: 1 } } })], redoOps: [cells({ 2: { 0: { v: 2 } } })] })
    const stale = manager.planUndo()
    manager.rebaseRemote([insertRows(0)])
    expect(manager.commit(stale)).toBe(false)
    expect(manager.planUndo()).toMatchObject({ status: 'ready', entryId: 'edit' })
  })

  it('rebases cell inverses and selections through ordered inserts, removals, and moves', () => {
    const manager = new CollaborativeUndoManager()
    manager.record({
      undoOps: [cells({ 5: { 0: { v: 'old' } } })],
      redoOps: [cells({ 5: { 0: { v: 'new' } } })],
      selectionBefore: { sheet: 's1', ranges: [[5, 0, 5, 0]], active: [5, 0] },
      selectionAfter: { sheet: 's1', ranges: [[5, 0, 5, 0]], active: [5, 0] },
    })

    manager.rebaseRemote([insertRows(0, 2), removeRows(3), moveRows(6, 10)])
    const plan = manager.planUndo()
    expect(plan).toMatchObject({ status: 'ready', selection: { ranges: [[10, 0, 10, 0]], active: [10, 0] } })
    if (plan.status !== 'ready') throw new Error('expected ready plan')
    expect(plan.ops[0]!.params.cellValue).toEqual({ 10: { 0: { v: 'old' } } })
  })

  it('neutralizes only cells overwritten by a later remote writer', () => {
    const manager = new CollaborativeUndoManager()
    manager.record({
      undoOps: [cells({ 3: { 0: { v: 'a-old' }, 1: { v: 'b-old' } } })],
      redoOps: [cells({ 3: { 0: { v: 'a-new' }, 1: { v: 'b-new' } } })],
    })
    manager.rebaseRemote([cells({ 3: { 0: { v: 'remote' } } })])

    const plan = manager.planUndo()
    expect(plan).toMatchObject({ status: 'ready' })
    if (plan.status !== 'ready') throw new Error('expected ready plan')
    expect(plan.ops[0]!.params.cellValue).toEqual({ 3: { 1: { v: 'b-old' } } })
    expect(manager.commit(plan)).toBe(true)
    const redo = manager.planRedo()
    expect(redo).toMatchObject({ status: 'ready' })
    if (redo.status !== 'ready') throw new Error('expected ready plan')
    expect(redo.ops[0]!.params.cellValue).toEqual({ 3: { 1: { v: 'b-new' } } })
  })

  it('moves pending local history through earlier remote structure without surrendering cell ownership', () => {
    const manager = new CollaborativeUndoManager()
    manager.record({
      id: 'offline-local',
      undoOps: [cells({ 3: { 0: { v: 'before' } } })],
      redoOps: [cells({ 3: { 0: { v: 'local' } } })],
      selectionBefore: { sheet: 's1', ranges: [[3, 0, 3, 0]], active: [3, 0] },
    })
    manager.rebaseRemote([
      cells({ 3: { 0: { v: 'remote-earlier' } } }),
      insertRows(0, 2),
    ], { remoteBeforeEntryIds: ['offline-local'] })

    const plan = manager.planUndo()
    expect(plan).toMatchObject({ status: 'ready', selection: { ranges: [[5, 0, 5, 0]], active: [5, 0] } })
    if (plan.status !== 'ready') throw new Error('expected ready plan')
    expect(plan.ops[0]!.params.cellValue).toEqual({ 5: { 0: { v: 'before' } } })
  })

  it('turns a fully overwritten entry into a deterministic no-op checkpoint', () => {
    const manager = new CollaborativeUndoManager()
    manager.record({ id: 'overwritten', undoOps: [cells({ 1: { 1: { v: 'old' } } })], redoOps: [cells({ 1: { 1: { v: 'new' } } })] })
    expect(manager.rebaseRemote([cells({ 1: { 1: { v: 'remote' } } })])).toEqual({ rebased: 0, neutralized: 1, blocked: 0 })
    const plan = manager.planUndo()
    expect(plan).toMatchObject({ status: 'neutralized', entryId: 'overwritten', ops: [] })
    expect(manager.commit(plan)).toBe(true)
    expect(manager.planRedo()).toMatchObject({ status: 'neutralized', entryId: 'overwritten', ops: [] })
  })

  it('keeps same coordinates on other sheets and drops history for a removed sheet', () => {
    const manager = new CollaborativeUndoManager()
    manager.record({ undoOps: [cells({ 1: { 1: { v: 'old' } } }, 's2')], redoOps: [cells({ 1: { 1: { v: 'new' } } }, 's2')] })
    manager.rebaseRemote([cells({ 1: { 1: { v: 'remote' } } }, 's1')])
    expect(manager.planUndo()).toMatchObject({ status: 'ready' })
    manager.rebaseRemote([{ id: 'sheet.mutation.remove-sheet', params: { unitId: 'wb', subUnitId: 's2' } }])
    expect(manager.planUndo()).toMatchObject({ status: 'neutralized', selection: null })
  })

  it('fails closed when a remote write follows a local structural action', () => {
    const manager = new CollaborativeUndoManager()
    manager.record({ undoOps: [removeRows(4)], redoOps: [insertRows(4)] })
    const report = manager.rebaseRemote([cells({ 4: { 2: { v: 'remote' } } })])
    expect(report.blocked).toBe(1)
    expect(manager.planUndo()).toMatchObject({
      status: 'blocked',
      reasons: [{ code: 'remote-write-after-local-structure', opId: SET_RANGE_VALUES }],
    })
  })

  it('rebases content-range inverses through later cell ownership', () => {
    const manager = new CollaborativeUndoManager()
    const move = (oldValue: string): OpRecord => ({
      id: MOVE_RANGE,
      params: {
        unitId: 'wb',
        fromRange: range(1, 1, 0, 1), toRange: range(5, 5, 0, 1),
        from: { subUnitId: 's1', value: { 1: { 0: { v: oldValue }, 1: null } } },
        to: { subUnitId: 's1', value: { 5: { 0: null, 1: { v: oldValue } } } },
      },
    })
    manager.record({ undoOps: [move('old')], redoOps: [move('new')] })
    manager.rebaseRemote([cells({ 1: { 0: { v: 'remote' } }, 5: { 1: { v: 'remote' } } })])
    const plan = manager.planUndo()
    expect(plan).toMatchObject({ status: 'ready' })
    if (plan.status !== 'ready') throw new Error('expected ready plan')
    expect(plan.ops[0]!.params.from).toMatchObject({ value: { 1: { 1: null } } })
    expect(plan.ops[0]!.params.to).toMatchObject({ value: { 5: { 0: null } } })

    const reorder = new CollaborativeUndoManager()
    reorder.record({
      undoOps: [{ id: REORDER_RANGE, params: { unitId: 'wb', subUnitId: 's1', range: range(0, 1, 0, 2), order: { 0: 1, 1: 0 } } }],
      redoOps: [{ id: REORDER_RANGE, params: { unitId: 'wb', subUnitId: 's1', range: range(0, 1, 0, 2), order: { 0: 1, 1: 0 } } }],
    })
    reorder.rebaseRemote([cells({ 0: { 1: { v: 'remote' } } })])
    const reorderPlan = reorder.planUndo()
    expect(reorderPlan).toMatchObject({ status: 'ready' })
    if (reorderPlan.status !== 'ready') throw new Error('expected ready plan')
    expect(reorderPlan.ops).toHaveLength(2)

    const structural = new CollaborativeUndoManager()
    structural.record({ undoOps: [move('old')], redoOps: [move('new')] })
    structural.rebaseRemote([insertRows(0)])
    expect(structural.planUndo()).toMatchObject({ status: 'blocked', reasons: [{ code: 'lossy-structural-transform' }] })
  })

  it('fails closed for malformed content-range, rich, and floating-object operations', () => {
    const local = new CollaborativeUndoManager()
    local.record({
      undoOps: [{ id: 'sheet.mutation.set-drawing', params: { subUnitId: 's1', drawingId: 'd1' } }],
      redoOps: [{ id: 'sheet.mutation.set-drawing', params: { subUnitId: 's1', drawingId: 'd1' } }],
    })
    expect(local.planUndo()).toMatchObject({ status: 'blocked', reasons: [{ code: 'unsupported-local-operation' }] })

    const remote = new CollaborativeUndoManager()
    remote.record({ undoOps: [cells({ 2: { 0: { v: 1 } } })], redoOps: [cells({ 2: { 0: { v: 2 } } })] })
    remote.rebaseRemote([{ id: 'sheet.mutation.reorder-range', params: { subUnitId: 's1', range: range(0, 5) } }])
    expect(remote.planUndo()).toMatchObject({ status: 'blocked', reasons: [{ code: 'unsupported-remote-operation' }] })
  })

  it('blocks a move that splits a single-range inverse but supports range lists', () => {
    const single = new CollaborativeUndoManager()
    single.record({
      undoOps: [{ id: 'sheet.mutation.set-row-height', params: { subUnitId: 's1', range: range(2, 4) } }],
      redoOps: [{ id: 'sheet.mutation.set-row-height', params: { subUnitId: 's1', range: range(2, 4) } }],
    })
    single.rebaseRemote([moveRows(2, 7, 2)])
    expect(single.planUndo()).toMatchObject({ status: 'blocked', reasons: [{ code: 'lossy-structural-transform' }] })

    const list = new CollaborativeUndoManager()
    list.record({
      undoOps: [{ id: 'sheet.mutation.add-worksheet-merge', params: { subUnitId: 's1', ranges: [range(2, 4)] } }],
      redoOps: [{ id: 'sheet.mutation.add-worksheet-merge', params: { subUnitId: 's1', ranges: [range(2, 4)] } }],
    })
    list.rebaseRemote([moveRows(2, 7, 2)])
    const plan = list.planUndo()
    expect(plan).toMatchObject({ status: 'ready' })
    if (plan.status !== 'ready') throw new Error('expected ready plan')
    expect(plan.ops[0]!.params.ranges).toEqual([range(2, 2), range(7, 8)])
  })

  it('isolates caller objects and clears redo on a new local action', () => {
    const manager = new CollaborativeUndoManager()
    const undo = cells({ 1: { 0: { v: 'old' } } })
    manager.record({ undoOps: [undo], redoOps: [cells({ 1: { 0: { v: 'new' } } })] })
    ;(undo.params.cellValue as Record<string, Record<string, { v: string }>>)['1']!['0']!.v = 'mutated-after-record'
    const plan = manager.planUndo()
    expect(plan).toMatchObject({ status: 'ready' })
    if (plan.status !== 'ready') throw new Error('expected ready plan')
    expect(plan.ops[0]!.params.cellValue).toEqual({ 1: { 0: { v: 'old' } } })
    expect(manager.commit(plan)).toBe(true)
    expect(manager.redoDepth).toBe(1)
    manager.record({ undoOps: [cells({ 9: { 0: { v: 0 } } })], redoOps: [cells({ 9: { 0: { v: 1 } } })] })
    expect(manager.redoDepth).toBe(0)
  })

  it('rejects duplicate identities and malformed sparse matrices', () => {
    const manager = new CollaborativeUndoManager()
    manager.record({ id: 'same', undoOps: [cells({ 1: { 0: 1 } })], redoOps: [cells({ 1: { 0: 2 } })] })
    expect(() => manager.record({ id: 'same', undoOps: [], redoOps: [] })).toThrow('already exists')
    const malformed = new CollaborativeUndoManager()
    malformed.record({ undoOps: [cells({ '-1': { 0: 1 } })], redoOps: [cells({ '-1': { 0: 2 } })] })
    expect(malformed.planUndo()).toMatchObject({ status: 'blocked', reasons: [{ code: 'unsupported-local-operation' }] })
    const malformedStructure = new CollaborativeUndoManager()
    malformedStructure.record({ undoOps: [removeRows(-2)], redoOps: [insertRows(-2)] })
    const malformedPlan = malformedStructure.planUndo()
    expect(malformedPlan.status).toBe('blocked')
    if (malformedPlan.status !== 'blocked') throw new Error('expected blocked plan')
    expect(malformedPlan.reasons).toHaveLength(2)
    expect(malformedPlan.reasons.every((reason) => reason.code === 'unsupported-local-operation')).toBe(true)
  })
})
