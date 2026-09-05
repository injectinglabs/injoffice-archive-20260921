import { describe, expect, it } from 'vitest'
import { extractChartData } from './extract'
import { buildEChartsOption } from './option'
import type { ChartSpec } from './types'

const spec = (over: Partial<ChartSpec> = {}): ChartSpec => ({
  id: 'c1',
  type: 'column',
  range: { sheetId: 's1', startRow: 0, startColumn: 0, endRow: 3, endColumn: 2 },
  ...over,
})

describe('extractChartData', () => {
  it('detects header row and category column in a classic block', () => {
    const grid = [
      ['Quarter', 'Revenue', 'Costs'],
      ['Q1', 100, 60],
      ['Q2', 130, 70],
      ['Q3', 150, 80],
    ]
    const d = extractChartData(grid, {})
    expect(d.categories).toEqual(['Q1', 'Q2', 'Q3'])
    expect(d.series.map((s) => s.name)).toEqual(['Revenue', 'Costs'])
    expect(d.series[0].values).toEqual([100, 130, 150])
  })

  it('handles a headerless all-numeric block', () => {
    const d = extractChartData(
      [
        [1, 2],
        [3, 4],
      ],
      {},
    )
    expect(d.categories).toEqual(['Row 1', 'Row 2'])
    expect(d.series).toHaveLength(2)
    expect(d.series[0].name).toBe('Series 1')
    expect(d.series[0].values).toEqual([1, 3])
  })

  it('respects explicit pins over heuristics', () => {
    const grid = [
      ['2019', '2020'],
      ['10', '20'],
    ]
    const pinned = extractChartData(grid, { firstRowIsHeader: true, firstColumnIsCategory: false })
    expect(pinned.series.map((s) => s.name)).toEqual(['2019', '2020'])
    expect(pinned.categories).toEqual(['Row 1'])
  })

  it('parses numeric strings, thousands separators, and percentages', () => {
    const d = extractChartData(
      [
        ['Region', 'Share'],
        ['EU', '1,250'],
        ['US', '42%'],
        ['APAC', ''],
      ],
      {},
    )
    expect(d.series[0].values).toEqual([1250, 0.42, null])
  })

  it('drops columns with no numeric content instead of charting nulls', () => {
    const d = extractChartData(
      [
        ['City', 'Note', 'Pop'],
        ['Oslo', 'nice', 700],
        ['Rome', 'warm', 2800],
      ],
      {},
    )
    expect(d.series.map((s) => s.name)).toEqual(['Pop'])
  })

  it('returns empty shapes for an empty grid', () => {
    expect(extractChartData([], {})).toEqual({ categories: [], series: [] })
  })
})

describe('buildEChartsOption', () => {
  const data = {
    categories: ['Q1', 'Q2'],
    series: [
      { name: 'Rev', values: [1, 2] },
      { name: 'Cost', values: [3, 4] },
    ],
  }

  it('maps column to bar series on a category axis', () => {
    const o = buildEChartsOption(spec({ type: 'column' }), data) as any
    expect(o.xAxis.type).toBe('category')
    expect(o.series.map((s: any) => s.type)).toEqual(['bar', 'bar'])
  })

  it('maps bar to horizontal orientation (value x-axis)', () => {
    const o = buildEChartsOption(spec({ type: 'bar' }), data) as any
    expect(o.xAxis.type).toBe('value')
    expect(o.yAxis.type).toBe('category')
  })

  it('maps pie to slice data from the first series', () => {
    const o = buildEChartsOption(spec({ type: 'pie' }), data) as any
    expect(o.series[0].type).toBe('pie')
    expect(o.series[0].data).toEqual([
      { name: 'Q1', value: 1 },
      { name: 'Q2', value: 2 },
    ])
  })

  it('maps two-series scatter to XY pairs', () => {
    const o = buildEChartsOption(spec({ type: 'scatter' }), data) as any
    expect(o.series[0].data).toEqual([
      [1, 3],
      [2, 4],
    ])
  })

  it('shows a legend only for multi-series unless pinned', () => {
    const single = { categories: ['a'], series: [{ name: 's', values: [1] }] }
    expect((buildEChartsOption(spec(), single) as any).legend).toBeUndefined()
    expect((buildEChartsOption(spec(), data) as any).legend).toBeDefined()
    expect((buildEChartsOption(spec({ legend: true }), single) as any).legend).toBeDefined()
  })

  it('area charts carry an areaStyle', () => {
    const o = buildEChartsOption(spec({ type: 'area' }), data) as any
    expect(o.series[0].areaStyle).toBeDefined()
  })
})
