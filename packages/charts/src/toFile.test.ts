import { describe, expect, it } from 'vitest'
import { toWireChartRemove, toWireCharts, toWireChartUpdate, type ToWireContext } from './toFile'
import type { ChartSpec } from './types'

const spec = (over: Partial<ChartSpec> = {}): ChartSpec => ({
  id: 'c1',
  type: 'column',
  range: { sheetId: 's1', startRow: 0, startColumn: 0, endRow: 4, endColumn: 2 },
  ...over,
})

// A1:C5 block: header row + categories column + two numeric series.
const grid = [
  ['Quarter', 'Revenue', 'Costs'],
  ['Q1', 10, 5],
  ['Q2', 20, 6],
  ['Q3', 30, 7],
  ['Q4', 40, 8],
]

const ctx = (over: Partial<ToWireContext> = {}): ToWireContext => ({
  sheetNameOf: (id) => (id === 's1' ? 'Data' : null),
  readRange: () => grid,
  ...over,
})

describe('toWireCharts', () => {
  it('builds refs matching the renderer interpretation (header + category column)', () => {
    const { charts, skipped } = toWireCharts([{ spec: spec({ title: 'Rev' }) }], ctx())
    expect(skipped).toEqual([])
    expect(charts).toHaveLength(1)
    const c = charts[0]
    expect(c).toMatchObject({ sheetName: 'Data', type: 'column', title: 'Rev' })
    expect(c.series).toEqual([
      { name: 'Revenue', nameRef: 'Data!$B$1', categoriesRef: 'Data!$A$2:$A$5', valuesRef: 'Data!$B$2:$B$5' },
      { name: 'Costs', nameRef: 'Data!$C$1', categoriesRef: 'Data!$A$2:$A$5', valuesRef: 'Data!$C$2:$C$5' },
    ])
    // Default anchor: one column right of the data (endColumn 2 → col 4), 8×18.
    expect(c.anchor).toEqual({ fromCol: 4, fromRow: 0, toCol: 12, toRow: 18 })
  })

  it('uses the file anchor when the chart came from the file', () => {
    const { charts } = toWireCharts(
      [{ spec: spec(), cellAnchor: { FromCol: 5, FromRow: 2, ToCol: 13, ToRow: 20 } }],
      ctx(),
    )
    expect(charts[0].anchor).toEqual({ fromCol: 5, fromRow: 2, toCol: 13, toRow: 20 })
  })

  it('normalizes canonical renderer names to the bounded native writer vocabulary', () => {
    const donut = toWireCharts([{ spec: spec({ type: 'Donut' }) }], ctx())
    expect(donut.skipped).toEqual([])
    expect(donut.charts[0].type).toBe('doughnut')

    const stacked = toWireCharts([{ spec: spec({ type: 'ColumnStacked' }) }], ctx())
    expect(stacked.charts).toEqual([])
    expect(stacked.skipped[0]).toContain('no file representation yet')
  })

  it('respects pinned layout flags (no header, no category column)', () => {
    const numeric = [
      [1, 2],
      [3, 4],
    ]
    const { charts } = toWireCharts(
      [{ spec: spec({ range: { sheetId: 's1', startRow: 10, startColumn: 1, endRow: 11, endColumn: 2 }, firstRowIsHeader: false, firstColumnIsCategory: false }) }],
      ctx({ readRange: () => numeric }),
    )
    expect(charts[0].series).toEqual([
      { name: 'Series 1', nameRef: undefined, categoriesRef: undefined, valuesRef: 'Data!$B$11:$B$12' },
      { name: 'Series 2', nameRef: undefined, categoriesRef: undefined, valuesRef: 'Data!$C$11:$C$12' },
    ])
  })

  it('quotes sheet names that need it', () => {
    const { charts } = toWireCharts([{ spec: spec() }], ctx({ sheetNameOf: () => "My 'Q1' Data" }))
    expect(charts[0].series[0].valuesRef).toBe("'My ''Q1'' Data'!$B$2:$B$5")
  })

  it('skips unwritable types, missing sheets and empty data — with reasons', () => {
    const { charts, skipped } = toWireCharts(
      [
        { spec: spec({ id: 'w', type: 'waterfall', title: 'W' }) },
        { spec: spec({ id: 'g', range: { ...spec().range, sheetId: 'gone' } }) },
        { spec: spec({ id: 'e' }) },
      ],
      ctx({ readRange: () => [] }),
    )
    expect(charts).toEqual([])
    expect(skipped).toHaveLength(3)
    expect(skipped[0]).toContain('waterfall')
    expect(skipped[1]).toContain('sheet no longer exists')
    expect(skipped[2]).toContain('empty')
  })

  it('drops non-numeric columns from the series, like the renderer', () => {
    const g = [
      ['Name', 'Note', 'Value'],
      ['a', 'x', 1],
      ['b', 'y', 2],
    ]
    const { charts } = toWireCharts(
      [{ spec: spec({ range: { sheetId: 's1', startRow: 0, startColumn: 0, endRow: 2, endColumn: 2 } }) }],
      ctx({ readRange: () => g }),
    )
    expect(charts[0].series).toHaveLength(1)
    expect(charts[0].series[0].name).toBe('Value')
    expect(charts[0].series[0].valuesRef).toBe('Data!$C$2:$C$3')
  })

  it('refuses to serialize a hydrated chart as a duplicate add', () => {
    const native = spec({ nativeIdentity: { part: 'xl/charts/chart1.xml', drawingPart: 'xl/drawings/drawing1.xml', objectId: 3 } })
    expect(toWireCharts([{ spec: native }], ctx())).toEqual({
      charts: [],
      skipped: ['chart c1: it already has native identity; use an update request'],
    })
  })
})

describe('native lifecycle wire requests', () => {
  const nativeIdentity = { part: 'xl/charts/chart3.xml', drawingPart: 'xl/drawings/drawing2.xml', objectId: 42 }

  it('builds stable identity-bound remove and update requests', () => {
    const native = spec({ nativeIdentity, title: 'Native' })
    expect(toWireChartRemove(native)).toEqual({ request: { operation: 'remove', identity: nativeIdentity }, skipped: [] })
    const anchor = { FromCol: 1, FromRow: 2, ToCol: 9, ToRow: 20 }
    expect(toWireChartUpdate({ spec: native, cellAnchor: anchor }, ctx())).toEqual({
      request: {
        operation: 'update',
        identity: nativeIdentity,
        chart: expect.objectContaining({ sheetName: 'Data', type: 'column', title: 'Native', anchor: { fromCol: 1, fromRow: 2, toCol: 9, toRow: 20 } }),
      },
      skipped: [],
    })
  })

  it('fails closed for browser-created or malformed identities', () => {
    expect(toWireChartRemove(spec())).toEqual({ skipped: ['chart c1: it has no valid hydrated native identity'] })
    expect(toWireChartUpdate({ spec: spec({ nativeIdentity: { part: '../chart.xml', drawingPart: 'xl/drawings/drawing1.xml', objectId: 1 } }) }, ctx())).toEqual({
      skipped: ['chart c1: it has no valid hydrated native identity'],
    })
  })

  it('refuses a native update without its hydrated anchor', () => {
    expect(toWireChartUpdate({ spec: spec({ nativeIdentity }) }, ctx())).toEqual({
      skipped: ['chart c1: it has no hydrated native cell anchor'],
    })
  })
})
