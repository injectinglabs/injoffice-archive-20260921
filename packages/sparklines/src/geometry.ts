import type {
  SparklineBar,
  SparklineGeometry,
  SparklineOptions,
  SparklinePath,
  SparklinePoint,
  SparklineType,
  SparklineViewport,
} from './types'

const COLORS = {
  series: '#2f73d9', negative: '#c73939', markers: '#2f73d9', high: '#278447',
  low: '#c73939', first: '#6b4db5', last: '#d27519', axis: '#777777',
}

export interface SparklineCompileInput {
  type: SparklineType
  values: readonly (number | null)[]
  viewport: SparklineViewport
  options?: SparklineOptions
  /** Group-wide bounds. Explicit options.min/max still win. */
  domain?: Partial<{ min: number; max: number }>
}

export function numericExtent(values: readonly (number | null)[]): { min: number; max: number } | null {
  const numeric = values.filter((value): value is number => value !== null && Number.isFinite(value))
  if (numeric.length === 0) return null
  return { min: Math.min(...numeric), max: Math.max(...numeric) }
}

function resolveDomain(values: readonly (number | null)[], input: SparklineCompileInput): { min: number; max: number } {
  const extent = numericExtent(values) ?? { min: 0, max: 0 }
  if (input.type === 'win-loss' && input.options?.min === undefined && input.options?.max === undefined) {
    return { min: -1, max: 1 }
  }
  let min = input.options?.min ?? input.domain?.min ?? extent.min
  let max = input.options?.max ?? input.domain?.max ?? extent.max
  if (input.type !== 'line') {
    min = Math.min(min, 0)
    max = Math.max(max, 0)
  }
  if (min >= max) {
    if (input.options?.min !== undefined && input.options.max === undefined) max = min + (Math.abs(min) || 1)
    else if (input.options?.max !== undefined && input.options.min === undefined) min = max - (Math.abs(max) || 1)
    else {
      const midpoint = (min + max) / 2
      min = midpoint - 0.5
      max = midpoint + 0.5
    }
  }
  return { min, max }
}

function pointRole(index: number, value: number, values: readonly (number | null)[], options?: SparklineOptions): SparklinePoint['role'] {
  const numeric = values.filter((candidate): candidate is number => candidate !== null)
  if (options?.showNegative && value < 0) return 'negative'
  if (options?.showHigh && value === Math.max(...numeric)) return 'high'
  if (options?.showLow && value === Math.min(...numeric)) return 'low'
  if (options?.showFirst && index === values.findIndex((candidate) => candidate !== null)) return 'first'
  if (options?.showLast && index === lastNumericIndex(values)) return 'last'
  return 'normal'
}

function lastNumericIndex(values: readonly (number | null)[]): number {
  for (let index = values.length - 1; index >= 0; index--) if (values[index] !== null) return index
  return -1
}

function roleColor(role: SparklinePoint['role'], options?: SparklineOptions): string {
  const colors = { ...COLORS, ...options?.colors }
  return role === 'normal' ? colors.markers : colors[role]
}

const finiteSize = (value: number): number => Number.isFinite(value) && value > 0 ? value : 1

export function compileSparklineGeometry(input: SparklineCompileInput): SparklineGeometry {
  const width = finiteSize(input.viewport.width)
  const height = finiteSize(input.viewport.height)
  const padding = Math.max(0, Math.min(input.viewport.padding ?? 2, Math.min(width, height) / 2))
  const innerWidth = Math.max(0, width - padding * 2)
  const innerHeight = Math.max(0, height - padding * 2)
  const logicalValues = [...input.values]
  const entries = logicalValues.map((value, index) => ({ value, index }))
  if (input.options?.rightToLeft) entries.reverse()
  const values = entries.map((entry) => entry.value)
  const domain = resolveDomain(values, input)
  const y = (value: number): number => padding + ((domain.max - value) / (domain.max - domain.min)) * innerHeight
  const baselineY = Math.max(padding, Math.min(height - padding, y(0)))
  const step = values.length > 1 ? innerWidth / (values.length - 1) : innerWidth / 2
  const x = (index: number): number => values.length > 1 ? padding + index * step : width / 2
  const colors = { ...COLORS, ...input.options?.colors }
  const markers: SparklinePoint[] = []
  const paths: SparklinePath[] = []
  const bars: SparklineBar[] = []

  if (input.type === 'line') {
    let current: SparklinePath['points'] = []
    const connect = input.options?.emptyCells === 'connect'
    entries.forEach(({ value, index: sourceIndex }, displayIndex) => {
      if (value === null) {
        if (!connect && current.length) { paths.push({ points: current, color: colors.series, width: input.options?.lineWeight ?? 1 }); current = [] }
        return
      }
      const role = pointRole(sourceIndex, value, logicalValues, input.options)
      const point = { index: sourceIndex, value, x: x(displayIndex), y: y(value), role, color: roleColor(role, input.options) }
      current.push({ x: point.x, y: point.y })
      if (input.options?.showMarkers || role !== 'normal') markers.push(point)
    })
    if (current.length) paths.push({ points: current, color: colors.series, width: input.options?.lineWeight ?? 1 })
  } else {
    const slot = values.length ? innerWidth / values.length : innerWidth
    const barWidth = Math.max(1, slot * 0.7)
    entries.forEach(({ value: raw, index: sourceIndex }, displayIndex) => {
      if (raw === null) return
      const value = input.type === 'win-loss' ? Math.sign(raw) : raw
      if (input.type === 'win-loss' && value === 0) return
      const valueY = input.type === 'win-loss'
        ? (value > 0 ? padding : height - padding)
        : y(value)
      const role = pointRole(sourceIndex, raw, logicalValues, input.options)
      bars.push({
        index: sourceIndex,
        value: raw,
        x: padding + displayIndex * slot + (slot - barWidth) / 2,
        y: Math.min(valueY, baselineY),
        width: barWidth,
        height: Math.max(1, Math.abs(valueY - baselineY)),
        color: role === 'normal' ? colors.series : roleColor(role, input.options),
      })
    })
  }

  return { type: input.type, width, height, domain, baselineY, paths, bars, markers }
}

const n = (value: number): string => Number(value.toFixed(3)).toString()
const escapeAttribute = (value: string): string => value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')

/** Optional deterministic SVG serialization for hosts that do not provide a
 * canvas renderer. Identical geometry always produces identical markup. */
export function sparklineGeometryToSvg(geometry: SparklineGeometry, axisColor = COLORS.axis): string {
  const parts = [`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${n(geometry.width)} ${n(geometry.height)}" role="img">`]
  parts.push(`<line x1="0" y1="${n(geometry.baselineY)}" x2="${n(geometry.width)}" y2="${n(geometry.baselineY)}" stroke="${escapeAttribute(axisColor)}" stroke-width="0.5"/>`)
  for (const path of geometry.paths) {
    const points = path.points.map((point) => `${n(point.x)},${n(point.y)}`).join(' ')
    parts.push(`<polyline points="${points}" fill="none" stroke="${escapeAttribute(path.color)}" stroke-width="${n(path.width)}"/>`)
  }
  for (const bar of geometry.bars) parts.push(`<rect x="${n(bar.x)}" y="${n(bar.y)}" width="${n(bar.width)}" height="${n(bar.height)}" fill="${escapeAttribute(bar.color)}"/>`)
  for (const marker of geometry.markers) parts.push(`<circle cx="${n(marker.x)}" cy="${n(marker.y)}" r="1.5" fill="${escapeAttribute(marker.color)}"/>`)
  parts.push('</svg>')
  return parts.join('')
}
