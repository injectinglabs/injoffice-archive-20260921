import type { ChartData, ChartSpec } from './types'

// Pure data extraction: a raw values grid (whatever the host read from the
// range) → categories + numeric series. Pure so it unit-tests without Univer.
//
// Interpretation rules, in order:
//  1. spec.firstRowIsHeader / spec.firstColumnIsCategory pin the layout.
//  2. Otherwise, heuristics: the first row is a header when at least one of
//     its cells (past the first column) is non-numeric while the rows below
//     are mostly numeric; the first column is categories when its body cells
//     are mostly non-numeric.
// These mirror what a person means by "chart this block" often enough for an
// MVP; the chart panel exposes the pins for when they guess wrong.

function toNumber(v: unknown): number | null {
  if (typeof v === 'number' && Number.isFinite(v)) return v
  if (typeof v === 'string') {
    const t = v.trim().replace(/,/g, '')
    if (t === '') return null
    const n = Number(t)
    if (Number.isFinite(n)) return n
    // "42%" → 0.42, matching how the value reads in a chart context.
    const pct = /^(-?\d+(?:\.\d+)?)%$/.exec(t)
    if (pct) return Number(pct[1]) / 100
  }
  return null
}

function toLabel(v: unknown, fallback: string): string {
  if (v == null) return fallback
  const s = String(v).trim()
  return s === '' ? fallback : s
}

function mostlyNonNumeric(cells: unknown[]): boolean {
  const filled = cells.filter((c) => c != null && String(c).trim() !== '')
  if (filled.length === 0) return false
  const nonNumeric = filled.filter((c) => toNumber(c) === null)
  return nonNumeric.length * 2 > filled.length
}

/** The header/category interpretation for a grid — THE decision both the
 *  renderer (extractChartData) and the save-time OOXML ref builder
 *  (toFile.ts) must share, or what you see charted isn't what gets written. */
export function interpretLayout(
  grid: unknown[][],
  spec: Pick<ChartSpec, 'firstRowIsHeader' | 'firstColumnIsCategory'>,
): { hasHeader: boolean; hasCategoryCol: boolean; width: number } {
  if (grid.length === 0 || (grid[0]?.length ?? 0) === 0) return { hasHeader: false, hasCategoryCol: false, width: 0 }
  const width = Math.max(...grid.map((r) => r.length))
  const hasHeader =
    spec.firstRowIsHeader ?? (grid.length > 1 && mostlyNonNumeric(grid[0].slice(1)) && !mostlyNonNumeric(grid.slice(1).flatMap((r) => r.slice(1))))
  const body = hasHeader ? grid.slice(1) : grid
  const hasCategoryCol = spec.firstColumnIsCategory ?? (width > 1 && mostlyNonNumeric(body.map((r) => r[0])))
  return { hasHeader, hasCategoryCol, width }
}

export function extractChartData(
  grid: unknown[][],
  spec: Pick<ChartSpec, 'firstRowIsHeader' | 'firstColumnIsCategory'>,
): ChartData {
  if (grid.length === 0 || (grid[0]?.length ?? 0) === 0) return { categories: [], series: [] }

  const { hasHeader, hasCategoryCol, width } = interpretLayout(grid, spec)
  const body = hasHeader ? grid.slice(1) : grid

  const firstDataCol = hasCategoryCol ? 1 : 0
  const categories = body.map((r, i) =>
    hasCategoryCol ? toLabel(r[0], `Row ${i + 1}`) : `Row ${i + 1}`,
  )

  const series: ChartData['series'] = []
  for (let c = firstDataCol; c < width; c++) {
    const name = hasHeader ? toLabel(grid[0][c], `Series ${c - firstDataCol + 1}`) : `Series ${c - firstDataCol + 1}`
    const values = body.map((r) => toNumber(r[c]))
    // A column with no numeric content at all isn't a series (stray label
    // columns inside the block would otherwise chart as a flat line of nulls).
    if (values.some((v) => v !== null)) series.push({ name, values })
  }
  return { categories, series }
}
