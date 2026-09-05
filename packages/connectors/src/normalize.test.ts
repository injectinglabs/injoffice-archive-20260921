import { describe, expect, it } from 'vitest'
import { csvToGrid, jsonToGrid, resolvePath } from './normalize'

describe('resolvePath', () => {
  it('walks dot paths and tolerates misses', () => {
    const root = { data: { items: [1, 2] } }
    expect(resolvePath(root, 'data.items')).toEqual([1, 2])
    expect(resolvePath(root, '')).toBe(root)
    expect(resolvePath(root, 'data.nope.deeper')).toBeUndefined()
  })
})

describe('jsonToGrid', () => {
  it('tabulates arrays of objects with first-seen key union', () => {
    const grid = jsonToGrid([
      { name: 'Ada', role: 'Eng' },
      { name: 'Grace', team: 'Core' },
    ])
    expect(grid[0]).toEqual(['name', 'role', 'team'])
    expect(grid[1]).toEqual(['Ada', 'Eng', null])
    expect(grid[2]).toEqual(['Grace', null, 'Core'])
  })
  it('passes arrays of arrays through and wraps scalars', () => {
    expect(jsonToGrid([[1, 'a'], [2, 'b']])).toEqual([[1, 'a'], [2, 'b']])
    expect(jsonToGrid([1, 2])).toEqual([['value'], [1], [2]])
  })
  it('tabulates plain objects as key/value and stringifies nested values', () => {
    const grid = jsonToGrid({ total: 5, meta: { a: 1 } })
    expect(grid[0]).toEqual(['key', 'value'])
    expect(grid).toContainEqual(['total', 5])
    expect(grid).toContainEqual(['meta', '{"a":1}'])
  })
  it('applies the path before tabulating', () => {
    expect(jsonToGrid({ d: { rows: [{ x: 1 }] } }, 'd.rows')).toEqual([['x'], [1]])
  })
  it('non-tabular payloads yield empty', () => {
    expect(jsonToGrid('hello')).toEqual([])
    expect(jsonToGrid(null)).toEqual([])
  })
})

describe('csvToGrid', () => {
  it('parses simple rows with numeric coercion', () => {
    expect(csvToGrid('a,b\n1,x\n2,y\n')).toEqual([
      ['a', 'b'],
      [1, 'x'],
      [2, 'y'],
    ])
  })
  it('handles quoted fields with commas, newlines, and escaped quotes', () => {
    const grid = csvToGrid('name,note\n"Smith, J","said ""hi""\nthen left"')
    expect(grid[1][0]).toBe('Smith, J')
    expect(grid[1][1]).toBe('said "hi"\nthen left')
  })
  it('quoted numerics stay strings; empty unquoted cells are null', () => {
    const grid = csvToGrid('id,zip\n1,"02134"\n2,')
    expect(grid[1][1]).toBe('02134')
    expect(grid[2][1]).toBeNull()
  })
  it('CRLF endings and trailing newline are clean', () => {
    expect(csvToGrid('a\r\n1\r\n')).toEqual([['a'], [1]])
  })
})
