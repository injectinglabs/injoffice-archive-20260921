// Pure numeric analysis used by option.ts: trendlines and five-number
// summaries. No dependencies, fully unit-tested — these numbers end up in
// front of users, so they are computed here where a test can pin them, not
// inline in renderer glue.

/** Least-squares linear regression over (index, value) pairs, evaluated back
 *  at every index. Null values are excluded from the fit but the returned
 *  array covers every index, so the trendline spans gaps in the data. */
export function linearTrend(values: (number | null)[]): (number | null)[] {
  const points: [number, number][] = []
  values.forEach((v, i) => {
    if (v !== null) points.push([i, v])
  })
  if (points.length < 2) return values.map(() => null)
  const n = points.length
  let sumX = 0
  let sumY = 0
  let sumXY = 0
  let sumXX = 0
  for (const [x, y] of points) {
    sumX += x
    sumY += y
    sumXY += x * y
    sumXX += x * x
  }
  const denom = n * sumXX - sumX * sumX
  if (denom === 0) return values.map(() => null)
  const slope = (n * sumXY - sumX * sumY) / denom
  const intercept = (sumY - slope * sumX) / n
  return values.map((_, i) => slope * i + intercept)
}

/** Trailing moving average with the given window (default 3). Entries with
 *  fewer than `window` non-null trailing values (including the current one)
 *  are null, matching how spreadsheet apps draw MA trendlines. */
export function movingAverage(values: (number | null)[], window = 3): (number | null)[] {
  const w = Math.max(1, Math.floor(window))
  return values.map((_, i) => {
    if (i + 1 < w) return null
    const slice = values.slice(i + 1 - w, i + 1)
    if (slice.some((v) => v === null)) return null
    const nums = slice as number[]
    return nums.reduce((a, b) => a + b, 0) / w
  })
}

/** Five-number summary [min, Q1, median, Q3, max] using the common
 *  linear-interpolation quantile (Excel's QUARTILE.INC / R-7). Returns null
 *  when fewer than 2 numeric values exist. */
export function fiveNumberSummary(values: (number | null)[]): [number, number, number, number, number] | null {
  const nums = values.filter((v): v is number => v !== null).sort((a, b) => a - b)
  if (nums.length < 2) return null
  const q = (p: number): number => {
    const idx = (nums.length - 1) * p
    const lo = Math.floor(idx)
    const hi = Math.ceil(idx)
    if (lo === hi) return nums[lo]
    return nums[lo] + (nums[hi] - nums[lo]) * (idx - lo)
  }
  return [nums[0], q(0.25), q(0.5), q(0.75), nums[nums.length - 1]]
}

/** Waterfall decomposition: each value is a delta; returns the invisible
 *  "base" stack and the visible rise/fall segments ECharts needs, plus a
 *  final cumulative total appended as its own bar. */
export function waterfallSegments(values: (number | null)[]): {
  base: number[]
  rise: (number | null)[]
  fall: (number | null)[]
  total: number
} {
  const base: number[] = []
  const rise: (number | null)[] = []
  const fall: (number | null)[] = []
  let running = 0
  for (const v of values) {
    const d = v ?? 0
    if (d >= 0) {
      base.push(running)
      rise.push(d)
      fall.push(null)
    } else {
      base.push(running + d)
      rise.push(null)
      fall.push(-d)
    }
    running += d
  }
  return { base, rise, fall, total: running }
}
