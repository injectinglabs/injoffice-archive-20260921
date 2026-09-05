import type { EmptyCellBehavior, SparklineRangeRef } from './types'

function numericValue(value: unknown): number | null {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null
  if (typeof value !== 'string') return null
  const text = value.trim().replace(/,/g, '')
  if (!text) return null
  const percent = /^([+-]?(?:\d+(?:\.\d*)?|\.\d+))%$/.exec(text)
  if (percent) return Number(percent[1]) / 100
  const number = Number(text)
  return Number.isFinite(number) ? number : null
}

/** Flattens a one-row or one-column source in display order. Invalid and
 * non-numeric cells use the configured empty-cell behavior. `connect` keeps
 * gaps as null so the geometry compiler can bridge them. */
export function extractSparklineValues(
  grid: readonly (readonly unknown[])[],
  range: SparklineRangeRef,
  emptyCells: EmptyCellBehavior = 'gap',
): Array<number | null> {
  const length = range.startRow === range.endRow
    ? range.endColumn - range.startColumn + 1
    : range.endRow - range.startRow + 1
  const source = range.startRow === range.endRow
    ? (grid[0] ?? []).slice(0, length)
    : grid.slice(0, length).map((row) => row[0])
  return source.map((value) => {
    const numeric = numericValue(value)
    if (numeric !== null) return numeric
    return emptyCells === 'zero' ? 0 : null
  })
}
