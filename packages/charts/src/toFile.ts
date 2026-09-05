// Save-time OOXML serialization (P10): turn mounted ChartSpecs into the
// wire shape the gateway's xlsxpatch.AddChart consumes, so browser-created
// charts survive Save as REAL workbook charts (and round-trip back through
// the ?charts=json bridge on the next load).
//
// Refs are built with the SAME layout decision the renderer uses
// (interpretLayout) — what you see charted is what gets written. Types
// without an OOXML mapping (waterfall, heatmap, …) are skipped and counted;
// the host can tell the user rather than losing them silently.

import { interpretLayout } from './extract'
import type { FileChartAnchor } from './fromFile'
import type { CellRangeRef, ChartSpec, NativeChartIdentity } from './types'
import { normalizeChartType } from './types'

/** Mirror of the gateway's wireChart (files_apply_specs.go). */
export interface WireChartSeries {
  name: string
  nameRef?: string
  categoriesRef?: string
  valuesRef: string
}

export interface WireChart {
  sheetName: string
  type: string
  title?: string
  series: WireChartSeries[]
  anchor: { fromCol: number; fromRow: number; toCol: number; toRow: number }
}

/** The chart types xlsxpatch.AddChart can write (keep in sync with
 *  writableTypes in chartwrite.go). */
const OOXML_TYPE: Readonly<Partial<Record<ReturnType<typeof normalizeChartType>, string>>> = {
  Column: 'column',
  Bar: 'bar',
  Line: 'line',
  Area: 'area',
  Pie: 'pie',
  Donut: 'doughnut',
  Scatter: 'scatter',
}

/** Accepted ChartSpec type names that have a native writer. The values include
 * legacy identifiers so existing capability checks remain true. */
export const OOXML_WRITABLE: ReadonlySet<string> = new Set([
  ...Object.keys(OOXML_TYPE),
  'column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter',
])

export function colName(c: number): string {
  let s = ''
  let n = c + 1
  while (n > 0) {
    const rem = (n - 1) % 26
    s = String.fromCharCode(65 + rem) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

function quoteSheet(name: string): string {
  return /^[A-Za-z0-9_]+$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`
}

function absRef(sheet: string, r0: number, c0: number, r1: number, c1: number): string {
  const a = `$${colName(c0)}$${r0 + 1}`
  if (r0 === r1 && c0 === c1) return `${quoteSheet(sheet)}!${a}`
  return `${quoteSheet(sheet)}!${a}:$${colName(c1)}$${r1 + 1}`
}

export interface ToWireInput {
  spec: ChartSpec
  cellAnchor?: FileChartAnchor
}

export interface ToWireContext {
  /** Resolve a sheet id to its display name; null when the sheet is gone. */
  sheetNameOf: (sheetId: string) => string | null
  /** Read the raw values of a range (same grid the renderer reads). */
  readRange: (range: CellRangeRef) => unknown[][]
}

export interface ToWireResult {
  charts: WireChart[]
  /** Human-readable reasons for anything not written. */
  skipped: string[]
}

export interface WireChartRemove {
  operation: 'remove'
  identity: NativeChartIdentity
}

export interface WireChartUpdate {
  operation: 'update'
  identity: NativeChartIdentity
  chart: WireChart
}

export interface ChartLifecycleWireResult<T extends WireChartRemove | WireChartUpdate> {
  request?: T
  skipped: string[]
}

function nativeIdentityOf(spec: Pick<ChartSpec, 'id' | 'nativeIdentity'>): NativeChartIdentity | null {
  const identity = spec.nativeIdentity
  if (!identity
    || !/^xl\/charts\/[^/]+\.xml$/.test(identity.part)
    || !/^xl\/drawings\/(?!_rels\/)[^/]+\.xml$/.test(identity.drawingPart)
    || !Number.isSafeInteger(identity.objectId)
    || identity.objectId <= 0
    || identity.objectId > 0xffff_ffff) return null
  return { ...identity }
}

/** Build a fail-closed native delete request. */
export function toWireChartRemove(spec: Pick<ChartSpec, 'id' | 'nativeIdentity'>): ChartLifecycleWireResult<WireChartRemove> {
  const identity = nativeIdentityOf(spec)
  if (!identity) return { skipped: [`chart ${spec.id}: it has no valid hydrated native identity`] }
  return { request: { operation: 'remove', identity }, skipped: [] }
}

/** Build a one-chart native update while retaining hydrated identity. */
export function toWireChartUpdate(input: ToWireInput, ctx: ToWireContext): ChartLifecycleWireResult<WireChartUpdate> {
  const identity = nativeIdentityOf(input.spec)
  if (!identity) return { skipped: [`chart ${input.spec.id}: it has no valid hydrated native identity`] }
  if (!input.cellAnchor) return { skipped: [`chart ${input.spec.id}: it has no hydrated native cell anchor`] }
  const result = convertWireCharts([input], ctx, true)
  if (!result.charts[0]) return { skipped: result.skipped }
  return { request: { operation: 'update', identity, chart: result.charts[0] }, skipped: [] }
}

/** Default placement for a chart that never had a file anchor (created in
 *  the browser): an 8×18-cell block one column right of its data. */
function defaultAnchor(range: CellRangeRef): WireChart['anchor'] {
  const fromCol = range.endColumn + 2
  const fromRow = range.startRow
  return { fromCol, fromRow, toCol: fromCol + 8, toRow: fromRow + 18 }
}

/** Convert every mounted chart to the gateway wire shape. */
function convertWireCharts(inputs: ToWireInput[], ctx: ToWireContext, allowNative: boolean): ToWireResult {
  const charts: WireChart[] = []
  const skipped: string[] = []
  for (const { spec, cellAnchor } of inputs) {
    if (spec.nativeIdentity && !allowNative) {
      skipped.push(`chart ${spec.id}: it already has native identity; use an update request`)
      continue
    }
    const label = spec.title || spec.type
    const nativeType = OOXML_TYPE[normalizeChartType(spec.type)]
    if (!nativeType) {
      skipped.push(`"${label}": ${spec.type} charts have no file representation yet (kept on screen only)`)
      continue
    }
    const sheetName = ctx.sheetNameOf(spec.range.sheetId)
    if (!sheetName) {
      skipped.push(`"${label}": its sheet no longer exists`)
      continue
    }
    const grid = ctx.readRange(spec.range)
    const { hasHeader, hasCategoryCol, width } = interpretLayout(grid, spec)
    if (width === 0) {
      skipped.push(`"${label}": its data range is empty`)
      continue
    }
    const r = spec.range
    const bodyStartRow = r.startRow + (hasHeader ? 1 : 0)
    if (bodyStartRow > r.endRow) {
      skipped.push(`"${label}": its data range has no body rows`)
      continue
    }
    const firstDataCol = r.startColumn + (hasCategoryCol ? 1 : 0)
    const categoriesRef = hasCategoryCol ? absRef(sheetName, bodyStartRow, r.startColumn, r.endRow, r.startColumn) : undefined
    const series: WireChartSeries[] = []
    const body = hasHeader ? grid.slice(1) : grid
    for (let c = firstDataCol; c <= r.endColumn; c++) {
      const gc = c - r.startColumn
      // Mirror extract.ts: a column with no numeric content is not a series.
      const hasNumeric = body.some((row) => {
        const v = row?.[gc]
        if (typeof v === 'number' && Number.isFinite(v)) return true
        if (typeof v === 'string') {
          const t = v.trim().replace(/,/g, '')
          return t !== '' && Number.isFinite(Number(t))
        }
        return false
      })
      if (!hasNumeric) continue
      const headerCell = hasHeader ? grid[0]?.[gc] : null
      const name =
        headerCell != null && String(headerCell).trim() !== ''
          ? String(headerCell).trim()
          : `Series ${series.length + 1}`
      series.push({
        name,
        nameRef: hasHeader ? absRef(sheetName, r.startRow, c, r.startRow, c) : undefined,
        categoriesRef,
        valuesRef: absRef(sheetName, bodyStartRow, c, r.endRow, c),
      })
    }
    if (series.length === 0) {
      skipped.push(`"${label}": no numeric series in its data range`)
      continue
    }
    const anchor = cellAnchor
      ? { fromCol: cellAnchor.FromCol, fromRow: cellAnchor.FromRow, toCol: cellAnchor.ToCol, toRow: cellAnchor.ToRow }
      : defaultAnchor(r)
    charts.push({ sheetName, type: nativeType, title: spec.title || undefined, series, anchor })
  }
  return { charts, skipped }
}

/** Convert fresh mounted charts to native add payloads. Hydrated charts are
 * refused so a host cannot accidentally duplicate them on save. */
export function toWireCharts(inputs: ToWireInput[], ctx: ToWireContext): ToWireResult {
  return convertWireCharts(inputs, ctx, false)
}
