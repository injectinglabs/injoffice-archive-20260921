import { describe, expect, it } from 'vitest'
import { cellAddress, colName, diffGrids } from './gridDiff'
import { diffText } from './textDiff'

describe('colName / cellAddress', () => {
  it('maps columns like a spreadsheet', () => {
    expect(colName(0)).toBe('A')
    expect(colName(25)).toBe('Z')
    expect(colName(26)).toBe('AA')
    expect(colName(51)).toBe('AZ')
    expect(colName(52)).toBe('BA')
    expect(cellAddress(3, 1)).toBe('B4')
  })
})

describe('diffGrids', () => {
  it('finds changed cells with addresses and values', () => {
    const d = diffGrids(
      [
        ['Quarter', 'Revenue'],
        ['Q1', 120],
      ],
      [
        ['Quarter', 'Revenue'],
        ['Q1', 300],
      ],
    )
    expect(d.changes).toEqual([{ row: 1, col: 1, address: 'B2', from: '120', to: '300' }])
    expect(d.changeCount).toBe(1)
    expect(d.rowDelta).toBe(0)
    expect(d.truncated).toBe(false)
  })

  it('reports shape deltas and treats missing cells as blank', () => {
    const d = diffGrids([['a']], [['a', 'x'], ['y']])
    expect(d.rowDelta).toBe(1)
    expect(d.colDelta).toBe(1)
    expect(d.changes).toContainEqual({ row: 0, col: 1, address: 'B1', from: '', to: 'x' })
    expect(d.changes).toContainEqual({ row: 1, col: 0, address: 'A2', from: '', to: 'y' })
  })

  it('identical grids diff empty', () => {
    const g = [['a', 1], ['b', 2]]
    const d = diffGrids(g, g)
    expect(d.changeCount).toBe(0)
    expect(d.changes).toEqual([])
  })

  it('truncates the change list but keeps the exact count', () => {
    const from = [Array.from({ length: 10 }, () => 'x')]
    const to = [Array.from({ length: 10 }, () => 'y')]
    const d = diffGrids(from, to, { maxChanges: 3 })
    expect(d.changes).toHaveLength(3)
    expect(d.changeCount).toBe(10)
    expect(d.truncated).toBe(true)
  })
})

describe('diffText', () => {
  it('aligns around an inserted paragraph', () => {
    const d = diffText('one\ntwo\nthree', 'one\nNEW\ntwo\nthree')
    expect(d.added).toBe(1)
    expect(d.removed).toBe(0)
    expect(d.ops.filter((o) => o.kind === 'same')).toHaveLength(3)
    expect(d.ops.find((o) => o.kind === 'add')?.line).toBe('NEW')
  })

  it('reports replacement as remove+add', () => {
    const d = diffText('hello world', 'hello there')
    expect(d.removed).toBe(1)
    expect(d.added).toBe(1)
  })

  it('empty-to-content is all adds', () => {
    const d = diffText('', 'a\nb')
    // '' splits to [''] — one removed blank line, two added.
    expect(d.added).toBe(2)
    expect(d.ops.filter((o) => o.kind === 'add')).toHaveLength(2)
  })
})
