import { describe, expect, it } from 'vitest'
import { compileSparklineGeometry, sparklineGeometryToSvg } from './geometry'

describe('compileSparklineGeometry', () => {
  it('splits line paths at gaps or bridges them deterministically', () => {
    const base = { type: 'line' as const, values: [1, null, 3], viewport: { width: 100, height: 20, padding: 2 } }
    expect(compileSparklineGeometry(base).paths.map((path) => path.points.length)).toEqual([1, 1])
    const connected = compileSparklineGeometry({ ...base, options: { emptyCells: 'connect', showFirst: true, showLast: true } })
    expect(connected.paths.map((path) => path.points.length)).toEqual([2])
    expect(connected.markers.map((marker) => marker.role)).toEqual(['first', 'last'])
    expect(sparklineGeometryToSvg(connected)).toBe(sparklineGeometryToSvg(connected))
    expect(sparklineGeometryToSvg(connected)).toContain('<polyline points="2,18 98,2"')
  })

  it('compiles columns around zero with negative coloring', () => {
    const geometry = compileSparklineGeometry({
      type: 'column', values: [-2, 0, 4], viewport: { width: 60, height: 24, padding: 2 },
      options: { showNegative: true },
    })
    expect(geometry.domain).toEqual({ min: -2, max: 4 })
    expect(geometry.bars).toHaveLength(3)
    expect(geometry.bars[0].color).toBe('#c73939')
    expect(geometry.bars[2].color).toBe('#2f73d9')
    expect(geometry.bars.every((bar) => bar.height >= 1)).toBe(true)
  })

  it('compiles win/loss bars and omits ties', () => {
    const geometry = compileSparklineGeometry({
      type: 'win-loss', values: [-20, 0, 8], viewport: { width: 48, height: 18 },
      options: { showNegative: true },
    })
    expect(geometry.bars.map((bar) => [bar.index, bar.value])).toEqual([[0, -20], [2, 8]])
    expect(geometry.domain).toEqual({ min: -1, max: 1 })
    expect(geometry.baselineY).toBe(9)
    expect(geometry.bars[0].y).toBe(geometry.baselineY)
    expect(geometry.bars[1].y).toBe(2)
  })

  it('reverses display order without changing source-index marker roles', () => {
    const geometry = compileSparklineGeometry({
      type: 'line', values: [1, 2, 3], viewport: { width: 30, height: 12, padding: 1 },
      options: { rightToLeft: true, showFirst: true, showLast: true },
    })
    expect(geometry.markers.map((marker) => [marker.index, marker.role, marker.x])).toEqual([
      [2, 'last', 1],
      [0, 'first', 29],
    ])
  })

  it('keeps a valid domain when one explicit bound excludes all data', () => {
    expect(compileSparklineGeometry({
      type: 'line', values: [1, 2], viewport: { width: 20, height: 10 }, options: { min: 10 },
    }).domain).toEqual({ min: 10, max: 20 })
    expect(compileSparklineGeometry({
      type: 'line', values: [10, 20], viewport: { width: 20, height: 10 }, options: { max: 5 },
    }).domain).toEqual({ min: 0, max: 5 })
  })
})
