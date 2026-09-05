import { describe, expect, it } from 'vitest'
import { fiveNumberSummary, linearTrend, movingAverage, waterfallSegments } from './analysis'
import { buildEChartsOption } from './option'
import type { ChartSpec } from './types'

const spec = (over: Partial<ChartSpec> = {}): ChartSpec => ({
  id: 'c1',
  type: 'column',
  range: { sheetId: 's1', startRow: 0, startColumn: 0, endRow: 3, endColumn: 2 },
  ...over,
})

describe('linearTrend', () => {
  it('fits a perfect line exactly', () => {
    expect(linearTrend([2, 4, 6, 8])).toEqual([2, 4, 6, 8])
  })
  it('spans nulls with the fitted line', () => {
    const t = linearTrend([2, null, 6, 8])
    expect(t[1]).toBeCloseTo(4, 5)
  })
  it('returns nulls when underdetermined', () => {
    expect(linearTrend([5])).toEqual([null])
  })
})

describe('movingAverage', () => {
  it('computes a trailing window', () => {
    expect(movingAverage([1, 2, 3, 4], 2)).toEqual([null, 1.5, 2.5, 3.5])
  })
  it('nulls windows containing gaps', () => {
    expect(movingAverage([1, null, 3, 5], 2)).toEqual([null, null, null, 4])
  })
})

describe('fiveNumberSummary', () => {
  it('matches QUARTILE.INC on a known set', () => {
    // 1..9: Q1=3, median=5, Q3=7 under R-7 interpolation.
    expect(fiveNumberSummary([1, 2, 3, 4, 5, 6, 7, 8, 9])).toEqual([1, 3, 5, 7, 9])
  })
  it('interpolates on even counts', () => {
    expect(fiveNumberSummary([1, 2, 3, 4])).toEqual([1, 1.75, 2.5, 3.25, 4])
  })
  it('is null below 2 values', () => {
    expect(fiveNumberSummary([7])).toBeNull()
  })
})

describe('waterfallSegments', () => {
  it('stacks rises and falls on a running base', () => {
    const w = waterfallSegments([100, -30, 50])
    expect(w.base).toEqual([0, 70, 70])
    expect(w.rise).toEqual([100, null, 50])
    expect(w.fall).toEqual([null, 30, null])
    expect(w.total).toBe(120)
  })
})

describe('buildEChartsOption — phase 2 features', () => {
  const data = {
    categories: ['Q1', 'Q2', 'Q3'],
    series: [
      { name: 'Rev', values: [10, 20, 30] },
      { name: 'Margin', values: [0.4, 0.42, 0.45] },
    ],
  }

  it('combo: per-series type override renders a line inside a column chart', () => {
    const o = buildEChartsOption(spec({ series: { Margin: { type: 'line' } } }), data) as any
    expect(o.series.map((s: any) => s.type)).toEqual(['bar', 'line'])
  })

  it('secondary axis: adds a right axis and binds the flagged series', () => {
    const o = buildEChartsOption(spec({ series: { Margin: { secondaryAxis: true } } }), data) as any
    expect(o.yAxis).toHaveLength(2)
    expect(o.series[1].yAxisIndex).toBe(1)
    expect(o.series[0].yAxisIndex).toBe(0)
  })

  it('trendline: appends a dashed silent line series', () => {
    const o = buildEChartsOption(spec({ series: { Rev: { trendline: 'linear' } } }), data) as any
    expect(o.series).toHaveLength(3)
    const trend = o.series[1]
    expect(trend.name).toBe('Rev (trend)')
    expect(trend.lineStyle.type).toBe('dashed')
    expect(trend.data[0]).toBeCloseTo(10, 5)
  })

  it('waterfall: invisible base + rise/fall + total bar', () => {
    const wData = { categories: ['Start', 'Down'], series: [{ name: 'Δ', values: [100, -30] }] }
    const o = buildEChartsOption(spec({ type: 'waterfall' }), wData) as any
    expect(o.xAxis.data).toEqual(['Start', 'Down', 'Total'])
    expect(o.series[0].itemStyle.color).toBe('transparent')
    expect(o.series[3].data[2]).toBe(70)
  })

  it('heatmap: matrix cells + visualMap bounds', () => {
    const o = buildEChartsOption(spec({ type: 'heatmap' }), data) as any
    expect(o.series[0].type).toBe('heatmap')
    expect(o.visualMap.min).toBe(0.4)
    expect(o.visualMap.max).toBe(30)
  })

  it('radar: one indicator per category, one polygon per series', () => {
    const o = buildEChartsOption(spec({ type: 'radar' }), data) as any
    expect(o.radar.indicator).toHaveLength(3)
    expect(o.series[0].data).toHaveLength(2)
  })

  it('sankey: parses "Source → Target" categories into links', () => {
    const sData = {
      categories: ['Ads → Signups', 'Signups → Paid', 'Ads -> Churn'],
      series: [{ name: 'Flow', values: [100, 40, 10] }],
    }
    const o = buildEChartsOption(spec({ type: 'sankey' }), sData) as any
    expect(o.series[0].links).toHaveLength(3)
    expect(o.series[0].links[1]).toEqual({ source: 'Signups', target: 'Paid', value: 40 })
    const names = o.series[0].data.map((n: any) => n.name)
    expect(names).toContain('Ads')
    expect(names).toContain('Churn')
  })

  it('boxplot: one five-number box per series', () => {
    const bData = {
      categories: ['a', 'b', 'c', 'd', 'e'],
      series: [{ name: 'S', values: [1, 2, 3, 4, 5] }],
    }
    const o = buildEChartsOption(spec({ type: 'boxplot' }), bData) as any
    expect(o.series[0].data[0]).toEqual([1, 2, 3, 4, 5])
  })

  it('valueFormat percent formats axis labels', () => {
    const o = buildEChartsOption(spec({ valueFormat: 'percent' }), data) as any
    expect(o.yAxis[0].axisLabel.formatter(0.42)).toBe('42%')
  })

  it('theme palette overrides the default color list', () => {
    const o = buildEChartsOption(spec(), data, { palette: ['#111111', '#222222'] }) as any
    expect(o.color).toEqual(['#111111', '#222222'])
  })
})
