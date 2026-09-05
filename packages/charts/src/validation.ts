import { isChartType, normalizeChartType } from './types'
import type { ChartData, ChartSpec } from './types'

export type ChartValidationSeverity = 'error' | 'warning'

export interface ChartValidationIssue {
  code: string
  message: string
  severity: ChartValidationSeverity
}

export interface ChartValidationResult {
  valid: boolean
  issues: ChartValidationIssue[]
}

const LINK_RE = /^(.+?)\s*(?:→|->|>)\s*(.+)$/

function numericCount(values: readonly (number | null)[]): number {
  return values.filter((value) => typeof value === 'number' && Number.isFinite(value)).length
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

/**
 * Validate renderer input without coupling the public model to ECharts.
 * Errors mean the requested plot cannot be meaningfully constructed;
 * warnings identify lossy-but-renderable input.
 */
export function validateChartInput(
  type: unknown,
  data: unknown,
  options: unknown = {},
): ChartValidationResult {
  const issues: ChartValidationIssue[] = []
  if (!isChartType(type)) {
    return {
      valid: false,
      issues: [{ code: 'chart.type.unsupported', message: `Unsupported chart type: ${String(type)}`, severity: 'error' }],
    }
  }

  if (!isRecord(data) || !Array.isArray(data.categories) || !Array.isArray(data.series)) {
    return {
      valid: false,
      issues: [{ code: 'chart.data.shape', message: 'Chart data must contain category and series arrays.', severity: 'error' }],
    }
  }

  const chartData = data as unknown as ChartData
  if (chartData.categories.some((category) => typeof category !== 'string')) {
    issues.push({ code: 'chart.categories.value', message: 'Every chart category must be a string.', severity: 'error' })
  }

  const wellFormedSeries: ChartData['series'] = []
  for (const [index, series] of chartData.series.entries()) {
    if (!series || typeof series.name !== 'string' || !Array.isArray(series.values)) {
      issues.push({ code: 'chart.series.shape', message: `Series ${index + 1} is malformed.`, severity: 'error' })
      continue
    }
    wellFormedSeries.push(series)
    if (series.values.length !== chartData.categories.length) {
      issues.push({
        code: 'chart.series.length',
        message: `Series "${series.name}" has ${series.values.length} values for ${chartData.categories.length} categories.`,
        severity: 'error',
      })
    }
    if (series.values.some((value) => value !== null && (!Number.isFinite(value) || typeof value !== 'number'))) {
      issues.push({ code: 'chart.series.value', message: `Series "${series.name}" contains a non-finite value.`, severity: 'error' })
    }
  }

  const canonical = normalizeChartType(type)
  const populated = wellFormedSeries.filter((series) => numericCount(series.values) > 0)
  if (populated.length === 0) {
    issues.push({ code: 'chart.data.empty', message: `${canonical} requires at least one numeric series.`, severity: 'error' })
  }

  if (canonical === 'Bubble' && populated.length < 3) {
    issues.push({
      code: 'chart.bubble.series',
      message: 'Bubble requires three numeric series in x, y, size order.',
      severity: 'error',
    })
  }
  if (canonical === 'Candlestick' && populated.length < 4) {
    issues.push({
      code: 'chart.candlestick.series',
      message: 'Candlestick requires four numeric series in open, close, low, high order.',
      severity: 'error',
    })
  }
  if (canonical === 'Boxplot' && !wellFormedSeries.some((series) => numericCount(series.values) >= 2)) {
    issues.push({
      code: 'chart.boxplot.observations',
      message: 'Boxplot requires at least two observations in one series.',
      severity: 'error',
    })
  }
  if (canonical === 'Relation' || canonical === 'Sankey' || canonical === 'Chord') {
    const usableLinks = chartData.categories.filter((label, index) => LINK_RE.test(label) && wellFormedSeries[0]?.values[index] != null)
    if (usableLinks.length === 0) {
      issues.push({
        code: 'chart.links.empty',
        message: `${canonical} requires "Source → Target" categories and a numeric first series.`,
        severity: 'error',
      })
    } else if (usableLinks.length < chartData.categories.length) {
      issues.push({
        code: 'chart.links.skipped',
        message: `${chartData.categories.length - usableLinks.length} link row(s) will be skipped.`,
        severity: 'warning',
      })
    }
  }

  const chartOptions = isRecord(options)
    ? options as Partial<Pick<ChartSpec, 'series' | 'valueFormat' | 'currencySymbol'>>
    : {}
  if (!isRecord(options)) {
    issues.push({ code: 'chart.options.shape', message: 'Chart options must be an object.', severity: 'error' })
  }
  if (chartOptions.valueFormat === 'currency' && chartOptions.currencySymbol !== undefined &&
      (typeof chartOptions.currencySymbol !== 'string' || chartOptions.currencySymbol.trim() === '')) {
    issues.push({
      code: 'chart.currency.symbol',
      message: 'currencySymbol must not be empty when currency formatting is selected.',
      severity: 'error',
    })
  }

  const seriesOptions = chartOptions.series === undefined
    ? {}
    : isRecord(chartOptions.series)
      ? chartOptions.series
      : null
  if (seriesOptions === null) {
    issues.push({ code: 'chart.series.options.shape', message: 'Per-series options must be an object.', severity: 'error' })
  }
  const supportsSeriesOptions = new Set(['Line', 'Column', 'Area', 'Combination'])
  if (seriesOptions && Object.keys(seriesOptions).length > 0 && !supportsSeriesOptions.has(canonical)) {
    issues.push({
      code: 'chart.series.options.ignored',
      message: `${canonical} ignores combo, secondary-axis, and trendline series options.`,
      severity: 'warning',
    })
  }
  for (const [name, override] of Object.entries(seriesOptions ?? {})) {
    if (!isRecord(override)) {
      issues.push({ code: 'chart.series.option.shape', message: `Series "${name}" options are malformed.`, severity: 'error' })
      continue
    }
    if (override.trendline === 'movingAverage' && override.trendlineWindow !== undefined) {
      if (typeof override.trendlineWindow !== 'number' || !Number.isInteger(override.trendlineWindow) || override.trendlineWindow < 1) {
        issues.push({
          code: 'chart.trendline.window',
          message: `Series "${name}" has an invalid moving-average window.`,
          severity: 'error',
        })
      }
    }
  }

  return { valid: issues.every((issue) => issue.severity !== 'error'), issues }
}

export class ChartDataValidationError extends Error {
  readonly issues: ChartValidationIssue[]

  constructor(issues: ChartValidationIssue[]) {
    super(issues.map((issue) => issue.message).join(' '))
    this.name = 'ChartDataValidationError'
    this.issues = issues
  }
}

/** Strict boundary helper for RPC/import callers handling untrusted data. */
export function assertValidChartInput(
  type: unknown,
  data: unknown,
  options: unknown = {},
): void {
  const result = validateChartInput(type, data, options)
  if (!result.valid) throw new ChartDataValidationError(result.issues.filter((issue) => issue.severity === 'error'))
}
