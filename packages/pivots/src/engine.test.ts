import { describe, expect, it } from 'vitest'
import { bakePivot, fieldMembers, fieldValues, sourceFields } from './engine'

// Sales ledger — the canonical pivot demo shape.
const SOURCE = [
  ['Region', 'Product', 'Quarter', 'Sales', 'Units'],
  ['EU', 'Widget', 'Q1', 100, 10],
  ['EU', 'Gadget', 'Q1', 50, 5],
  ['US', 'Widget', 'Q1', 200, 20],
  ['EU', 'Widget', 'Q2', 120, 12],
  ['US', 'Gadget', 'Q2', 80, 8],
  ['US', 'Widget', 'Q2', 220, 22],
]

describe('bakePivot', () => {
  it('groups one row field and sums a value', () => {
    const { grid } = bakePivot(SOURCE, {
      rows: ['Region'],
      columns: [],
      values: [{ field: 'Sales', agg: 'sum' }],
    })
    expect(grid).toEqual([
      ['Region', 'Sum of Sales'],
      ['EU', 270],
      ['US', 500],
      ['Grand Total', 770],
    ])
  })

  it('supports multiple row fields as flat composite groups', () => {
    const { grid } = bakePivot(SOURCE, {
      rows: ['Region', 'Product'],
      columns: [],
      values: [{ field: 'Sales', agg: 'sum' }],
    })
    expect(grid[0]).toEqual(['Region', 'Product', 'Sum of Sales'])
    expect(grid).toContainEqual(['EU', 'Widget', 220])
    expect(grid).toContainEqual(['US', 'Gadget', 80])
    expect(grid[grid.length - 1]).toEqual(['Grand Total', '', 770])
  })

  it('cross-tabs with a column field including row totals', () => {
    const { grid } = bakePivot(SOURCE, {
      rows: ['Region'],
      columns: ['Quarter'],
      values: [{ field: 'Sales', agg: 'sum' }],
    })
    expect(grid[0]).toEqual(['Region', 'Q1', 'Q2', 'Total Sum of Sales'])
    expect(grid[1]).toEqual(['EU', 150, 120, 270])
    expect(grid[2]).toEqual(['US', 200, 300, 500])
    expect(grid[3]).toEqual(['Grand Total', 350, 420, 770])
  })

  it('supports multiple value fields and all aggregations', () => {
    const { grid } = bakePivot(SOURCE, {
      rows: ['Region'],
      columns: [],
      values: [
        { field: 'Sales', agg: 'avg' },
        { field: 'Units', agg: 'max' },
        { field: 'Product', agg: 'count' },
      ],
    })
    expect(grid[0]).toEqual(['Region', 'Average of Sales', 'Max of Units', 'Count of Product'])
    expect(grid[1]).toEqual(['EU', 90, 12, 3])
    expect(grid[2]).toEqual(['US', (200 + 80 + 220) / 3, 22, 3])
  })

  it('applies slicer filters before grouping and totals', () => {
    const { grid } = bakePivot(SOURCE, {
      rows: ['Region'],
      columns: [],
      values: [{ field: 'Sales', agg: 'sum' }],
      filters: { Quarter: ['Q1'] },
    })
    expect(grid).toEqual([
      ['Region', 'Sum of Sales'],
      ['EU', 150],
      ['US', 200],
      ['Grand Total', 350],
    ])
  })

  it('applies page selections and exclusions before aggregation, then sorts labels', () => {
    const { grid } = bakePivot(SOURCE, {
      rows: ['Region'],
      columns: [],
      values: [{ field: 'Sales', agg: 'sum' }],
      pageFields: [{ field: 'Quarter', selectedItem: 'Q2' }],
      memberFilters: [{ field: 'Product', mode: 'exclude', values: ['Gadget'] }],
      sorts: [{ field: 'Region', direction: 'descending' }],
    })
    expect(grid).toEqual([
      ['Region', 'Sum of Sales'],
      ['US', 220],
      ['EU', 120],
      ['Grand Total', 340],
    ])
  })

  it('empty filter selection filters everything', () => {
    const { grid } = bakePivot(SOURCE, {
      rows: ['Region'],
      columns: [],
      values: [{ field: 'Sales', agg: 'sum' }],
      filters: { Region: [] },
    })
    expect(grid).toEqual([['Region', 'Sum of Sales'], ['Grand Total', null]])
  })

  it('no row fields yields a single Total row', () => {
    const { grid } = bakePivot(SOURCE, {
      rows: [],
      columns: [],
      values: [{ field: 'Sales', agg: 'sum' }],
    })
    expect(grid).toEqual([
      ['', 'Sum of Sales'],
      ['Total', 770],
    ])
  })

  it('grandTotals: false omits the total row', () => {
    const { grid } = bakePivot(SOURCE, {
      rows: ['Region'],
      columns: [],
      values: [{ field: 'Sales', agg: 'sum' }],
      grandTotals: false,
    })
    expect(grid[grid.length - 1]).toEqual(['US', 500])
  })

  it('coerces numeric strings and ignores non-numeric cells in sums', () => {
    const src = [
      ['K', 'V'],
      ['a', '1,250'],
      ['a', 'n/a'],
      ['a', 50],
    ]
    const { grid } = bakePivot(src, {
      rows: ['K'],
      columns: [],
      values: [{ field: 'V', agg: 'sum' }],
    })
    expect(grid[1]).toEqual(['a', 1300])
  })

  it('unknown fields are ignored; no values means empty output', () => {
    const out = bakePivot(SOURCE, { rows: ['Nope'], columns: [], values: [{ field: 'Nada', agg: 'sum' }] })
    expect(out.grid).toEqual([])
  })
})

describe('field helpers', () => {
  it('sourceFields lists deduped headers', () => {
    expect(sourceFields(SOURCE)).toEqual(['Region', 'Product', 'Quarter', 'Sales', 'Units'])
  })
  it('fieldValues lists distinct values unfiltered, in order', () => {
    expect(fieldValues(SOURCE, 'Quarter')).toEqual(['Q1', 'Q2'])
    expect(fieldValues(SOURCE, 'Region')).toEqual(['EU', 'US'])
    expect(fieldValues(SOURCE, 'Missing')).toEqual([])
  })
  it('fieldMembers retains OOXML-relevant runtime kinds', () => {
    expect(fieldMembers([
      ['Value'], ['1'], [1], [true], [null], [''], [{}],
    ], 'Value')).toEqual([
      { value: '1', kind: 'string' },
      { value: '1', kind: 'number' },
      { value: 'true', kind: 'boolean' },
      { value: '', kind: 'blank' },
      { value: '[object Object]', kind: 'unsupported' },
    ])
  })
})
