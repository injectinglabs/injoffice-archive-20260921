import type { ChartAnchor, ChartIdentity, ChartTypeV1 } from './chartMutationProtocol.js'
import type { RangeRef } from './mutationProtocol.js'

const SHA = /^sha256:[a-f0-9]{64}$/
const PART = /^xl\/(?:charts|drawings)\/[^/]+\.xml$/

/** Desktop/engine projection of one native XLSX chart that can be mutated. */
export interface NativeEditableChartV1 {
  identity: ChartIdentity
  sheet_id: string
  fingerprint_sha256: string
  chart_type: ChartTypeV1 | 'unsupported'
  title: string
  range: RangeRef
  anchor: ChartAnchor
  editable: boolean
  refusal?: string
  categories: string[]
  series: Array<{ name: string; values: string[] }>
}

export type XlsxNativeChart = NativeEditableChartV1

function fail(): never {
  throw new TypeError('Invalid native editable chart projection')
}

function text(value: unknown, max = 1024): string {
  return typeof value === 'string' && value.length <= max ? value : fail()
}

function integer(value: unknown, min: number, max: number): number {
  if (typeof value !== 'number' || !Number.isInteger(value) || value < min || value > max) return fail()
  return value
}

function decodeIdentity(value: unknown): ChartIdentity {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return fail()
  const object = value as Record<string, unknown>
  if (Object.keys(object).length !== 3) return fail()
  const part = text(object.part)
  const drawingPart = text(object.drawingPart)
  if (!PART.test(part) || !PART.test(drawingPart) || !part.startsWith('xl/charts/') || !drawingPart.startsWith('xl/drawings/')) return fail()
  return { part, drawingPart, objectId: integer(object.objectId, 1, 0xffffffff) }
}

function decodeRange(value: unknown): RangeRef {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return fail()
  const object = value as Record<string, unknown>
  if (Object.keys(object).length !== 4) return fail()
  const row = integer(object.row, 0, 1_048_575)
  const column = integer(object.column, 0, 16_383)
  const end_row = integer(object.end_row, 0, 1_048_575)
  const end_column = integer(object.end_column, 0, 16_383)
  if (end_row < row || end_column < column) return fail()
  return { row, column, end_row, end_column }
}

function decodeAnchor(value: unknown): ChartAnchor {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return fail()
  const object = value as Record<string, unknown>
  if (Object.keys(object).length !== 4) return fail()
  const from_row = integer(object.from_row, 0, 1_048_575)
  const from_column = integer(object.from_column, 0, 16_383)
  const to_row = integer(object.to_row, 0, 1_048_576)
  const to_column = integer(object.to_column, 0, 16_384)
  if (to_row <= from_row || to_column <= from_column) return fail()
  return { from_row, from_column, to_row, to_column }
}

export function decodeNativeEditableChartsV1(value: unknown): NativeEditableChartV1[] {
  if (!Array.isArray(value) || value.length > 32) return fail()
  const seen = new Set<string>()
  return value.map((entry): NativeEditableChartV1 => {
    if (!entry || typeof entry !== 'object' || Array.isArray(entry)) return fail()
    const object = entry as Record<string, unknown>
    const keys = Object.keys(object)
    const optional = object.refusal !== undefined
    const expected = optional ? 11 : 10
    if (keys.length !== expected) return fail()
    const identity = decodeIdentity(object.identity)
    if (seen.has(identity.part)) return fail()
    seen.add(identity.part)
    const chartType = object.chart_type
    if (chartType !== 'column' && chartType !== 'bar' && chartType !== 'line' && chartType !== 'pie' && chartType !== 'unsupported') return fail()
    if (typeof object.editable !== 'boolean') return fail()
    const fingerprint_sha256 = text(object.fingerprint_sha256, 71)
    if (!SHA.test(fingerprint_sha256)) return fail()
    if (!Array.isArray(object.categories) || object.categories.length > 1000) return fail()
    if (!Array.isArray(object.series) || object.series.length > 8) return fail()
    const categories = object.categories.map((item) => text(item, 256))
    const series = object.series.map((item) => {
      if (!item || typeof item !== 'object' || Array.isArray(item)) return fail()
      const seriesObject = item as Record<string, unknown>
      if (Object.keys(seriesObject).length !== 2 || !Array.isArray(seriesObject.values) || seriesObject.values.length !== categories.length) return fail()
      return { name: text(seriesObject.name, 256), values: seriesObject.values.map((cell) => text(cell, 256)) }
    })
    if (object.editable && (chartType === 'unsupported' || categories.length === 0 || series.length === 0)) return fail()
    const decoded: NativeEditableChartV1 = {
      identity,
      sheet_id: text(object.sheet_id, 256),
      fingerprint_sha256,
      chart_type: chartType,
      title: text(object.title, 1024),
      range: decodeRange(object.range),
      anchor: decodeAnchor(object.anchor),
      editable: object.editable,
      categories,
      series,
    }
    if (optional) {
      const refusal = text(object.refusal, 4096)
      if (!refusal) return fail()
      decoded.refusal = refusal
    }
    return decoded
  })
}
