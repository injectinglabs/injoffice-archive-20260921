import { describe, expect, it } from 'vitest'
import { buildEChartsOption } from './option'
import {
  CANONICAL_CHART_TYPES,
  LEGACY_CHART_TYPES,
  isChartType,
  normalizeChartType,
} from './types'
import type { CanonicalChartType, ChartData, ChartSpec } from './types'
import { assertValidChartInput, ChartDataValidationError, validateChartInput } from './validation'

const range = { sheetId: 's1', startRow: 0, startColumn: 0, endRow: 3, endColumn: 4 }
const spec = (type: ChartSpec['type'], over: Partial<ChartSpec> = {}): ChartSpec => ({ id: 'c1', type, range, ...over })
const data: ChartData = {
  categories: ['A → B', 'B → C', 'C/D', 'D/E'],
  series: [
    { name: 'Open/X', values: [10, 20, 30, 40] },
    { name: 'Close/Y', values: [12, 18, 35, 38] },
    { name: 'Low/Size', values: [8, 15, 28, 34] },
    { name: 'High', values: [15, 24, 39, 44] },
  ],
}

const expectedSeriesTypes: Record<CanonicalChartType, string[]> = {
  Line: ['line', 'line', 'line', 'line'],
  Column: ['bar', 'bar', 'bar', 'bar'],
  ColumnStacked: ['bar', 'bar', 'bar', 'bar'],
  ColumnPercentStacked: ['bar', 'bar', 'bar', 'bar'],
  Bar: ['bar', 'bar', 'bar', 'bar'],
  BarStacked: ['bar', 'bar', 'bar', 'bar'],
  BarPercentStacked: ['bar', 'bar', 'bar', 'bar'],
  Pie: ['pie'],
  Donut: ['pie'],
  Area: ['line', 'line', 'line', 'line'],
  AreaStacked: ['line', 'line', 'line', 'line'],
  AreaPercentStacked: ['line', 'line', 'line', 'line'],
  Radar: ['radar'],
  Scatter: ['scatter'],
  Combination: ['bar', 'line', 'bar', 'line'],
  WordCloud: ['custom'],
  Funnel: ['funnel'],
  Bubble: ['scatter'],
  Relation: ['graph'],
  Waterfall: ['bar', 'bar', 'bar', 'bar'],
  Pareto: ['bar', 'line'],
  Sankey: ['sankey'],
  Heatmap: ['heatmap'],
  Boxplot: ['boxplot'],
  Candlestick: ['candlestick'],
  Histogram: ['bar'],
  Treemap: ['treemap'],
  Sunburst: ['sunburst'],
  Gauge: ['gauge'],
  Chord: ['graph'],
}

describe('Univer Charts public vocabulary', () => {
  it('contains 30 unique canonical types and maps each to an ECharts series', () => {
    expect(CANONICAL_CHART_TYPES).toHaveLength(30)
    expect(new Set(CANONICAL_CHART_TYPES).size).toBe(30)
    for (const type of CANONICAL_CHART_TYPES) {
      const option = buildEChartsOption(spec(type), data) as any
      expect(option.series.map((series: any) => series.type), type).toEqual(expectedSeriesTypes[type])
    }
  })

  it('keeps every legacy lowercase identifier valid and normalizable', () => {
    expect(LEGACY_CHART_TYPES).toHaveLength(14)
    for (const type of LEGACY_CHART_TYPES) {
      expect(isChartType(type)).toBe(true)
      expect(CANONICAL_CHART_TYPES).toContain(normalizeChartType(type))
      const legacy = buildEChartsOption(spec(type), data) as any
      const canonical = buildEChartsOption(spec(normalizeChartType(type)), data) as any
      expect(JSON.parse(JSON.stringify(legacy))).toEqual(JSON.parse(JSON.stringify(canonical)))
    }
    expect(isChartType('column_stacked')).toBe(false)
    expect(isChartType('__proto__')).toBe(false)
  })

  it('builds real stacked and 100% stacked series', () => {
    const stacked = buildEChartsOption(spec('ColumnStacked'), data) as any
    expect(stacked.series.every((series: any) => series.stack === 'total')).toBe(true)

    const percent = buildEChartsOption(spec('ColumnPercentStacked'), data) as any
    const firstCategoryTotal = percent.series.reduce((sum: number, series: any) => sum + series.data[0], 0)
    expect(firstCategoryTotal).toBeCloseTo(1)
    expect(percent.yAxis.max).toBe(1)
  })

  it('uses documented specialty data conventions', () => {
    const bubble = buildEChartsOption(spec('Bubble'), data) as any
    expect(bubble.series[0].data[0]).toEqual([10, 12, 8, 'A → B'])

    const candle = buildEChartsOption(spec('Candlestick'), data) as any
    expect(candle.series[0].data[0]).toEqual([10, 12, 8, 15])

    const relation = buildEChartsOption(spec('Relation'), data) as any
    expect(relation.series[0].links).toEqual([
      { source: 'A', target: 'B', value: 10 },
      { source: 'B', target: 'C', value: 20 },
    ])

    const pareto = buildEChartsOption(spec('Pareto'), data) as any
    expect(pareto.series[1].data.at(-1)).toBe(1)

    const histogram = buildEChartsOption(spec('Histogram'), data) as any
    expect(histogram.series[0].data.reduce((sum: number, count: number) => sum + count, 0)).toBe(4)
  })
})

describe('type-aware chart input validation', () => {
  it('accepts valid canonical and legacy input', () => {
    expect(validateChartInput('Bubble', data).valid).toBe(true)
    expect(validateChartInput('column', data).valid).toBe(true)
  })

  it('rejects missing bubble and candlestick dimensions', () => {
    const twoSeries = { ...data, series: data.series.slice(0, 2) }
    expect(validateChartInput('Bubble', twoSeries).issues.map((issue) => issue.code)).toContain('chart.bubble.series')
    expect(validateChartInput('Candlestick', twoSeries).issues.map((issue) => issue.code)).toContain('chart.candlestick.series')
  })

  it('rejects ragged series and invalid option values', () => {
    const ragged = { ...data, series: [{ name: 'Only', values: [1] }] }
    const result = validateChartInput('Line', ragged, {
      valueFormat: 'currency',
      currencySymbol: ' ',
      series: { Only: { trendline: 'movingAverage', trendlineWindow: 0 } },
    })
    expect(result.valid).toBe(false)
    expect(result.issues.map((issue) => issue.code)).toEqual(expect.arrayContaining([
      'chart.series.length',
      'chart.currency.symbol',
      'chart.trendline.window',
    ]))
  })

  it('reports skipped link rows and provides a strict boundary helper', () => {
    const result = validateChartInput('Chord', data)
    expect(result.valid).toBe(true)
    expect(result.issues).toContainEqual(expect.objectContaining({ code: 'chart.links.skipped', severity: 'warning' }))

    expect(() => assertValidChartInput('Bubble', { categories: [], series: [] })).toThrow(ChartDataValidationError)
  })

  it('fails closed instead of throwing on malformed runtime payloads', () => {
    expect(validateChartInput('Line', null).issues[0].code).toBe('chart.data.shape')
    const malformed = validateChartInput('Line', { categories: ['A'], series: [null] }, [])
    expect(malformed.valid).toBe(false)
    expect(malformed.issues.map((issue) => issue.code)).toEqual(expect.arrayContaining([
      'chart.series.shape',
      'chart.data.empty',
      'chart.options.shape',
    ]))
  })
})
