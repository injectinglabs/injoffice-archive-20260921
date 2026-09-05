import { describe, expect, it, vi } from 'vitest'
import { SparklineManager } from './manager'
import type { SparklineRangeRef, SparklineSpec } from './types'

const source = (row: number): SparklineRangeRef => ({ sheetId: 'sheet-1', startRow: row, endRow: row, startColumn: 0, endColumn: 2 })
const spec = (id: string, row: number, type: SparklineSpec['type'] = 'line'): SparklineSpec => ({
  id, type, source: source(row), target: { sheetId: 'sheet-1', row, column: 3 },
})

describe('SparklineManager', () => {
  it('creates, updates, removes, and reports immutable snapshots', () => {
    let n = 0
    const manager = new SparklineManager({ idFactory: (kind) => `${kind}-${++n}` })
    const changed = vi.fn()
    manager.onChange(changed)
    const created = manager.create({ type: 'line', source: source(0), target: { sheetId: 'sheet-1', row: 0, column: 3 } })
    expect(created.id).toBe('sparkline-1')
    manager.update(created.id, { options: { showMarkers: true } })
    const snapshot = manager.serialize()
    snapshot.sparklines[0].type = 'column'
    expect(manager.get(created.id)?.type).toBe('line')
    expect(manager.remove(created.id)).toBe(true)
    expect(manager.remove(created.id)).toBe(false)
    expect(changed).toHaveBeenCalledTimes(3)
  })

  it('groups equal types, shares their axis, and dissolves small groups', () => {
    const manager = new SparklineManager()
    manager.add(spec('a', 0))
    manager.add(spec('b', 1))
    manager.add(spec('c', 2))
    manager.group(['a', 'b', 'c'], 'g1')
    const values: Record<number, unknown[][]> = { 0: [[0, 10]], 1: [[-5, 5]], 2: [[20, 40]] }
    const read = (range: SparklineRangeRef) => values[range.startRow]
    expect(manager.render('a', read, { width: 100, height: 20 }).domain).toEqual({ min: -5, max: 40 })
    expect(manager.render('c', read, { width: 100, height: 20 }).domain).toEqual({ min: -5, max: 40 })
    manager.ungroup(['a', 'b'])
    expect(manager.listGroups()).toEqual([])
    expect(manager.list().every((item) => item.groupId === undefined)).toBe(true)
  })

  it('rejects invalid mutations and hydrates atomically', () => {
    const manager = new SparklineManager()
    manager.add(spec('safe', 0))
    expect(() => manager.add({ ...spec('bad', 1), source: { ...source(1), endRow: 2, endColumn: 2 } })).toThrow('exactly one row or one column')
    expect(() => manager.group(['safe', 'missing'], 'g')).toThrow('Unknown sparkline missing')
    expect(() => manager.add({ ...spec('occupied', 1), target: { ...spec('safe', 0).target } })).toThrow('Target cell already contains')
    expect(() => manager.hydrate({
      version: 1,
      sparklines: [{ ...spec('orphan', 2), groupId: 'absent' }],
      groups: [],
    })).toThrow('unknown group absent')
    expect(manager.list().map((item) => item.id)).toEqual(['safe'])
  })

  it('round-trips a valid grouped snapshot', () => {
    const first = new SparklineManager()
    first.add(spec('a', 0, 'column'))
    first.add(spec('b', 1, 'column'))
    first.group(['a', 'b'], 'group-1')
    const second = new SparklineManager()
    second.hydrate(first.serialize())
    expect(second.serialize()).toEqual(first.serialize())
  })
})
