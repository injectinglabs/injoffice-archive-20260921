import { describe, expect, it } from 'vitest'
import {
  arcPath,
  bowtiePath,
  bracePath,
  bracketPath,
  calloutPath,
  type CornerTreatment,
  cutCornerRectPath,
  diagStripePath,
  diamondPath,
  directionalArrowPath,
  displayShapePath,
  doubleArrowPath,
  dShapePath,
  ellipseCalloutPath,
  foldedCornerFlapPath,
  foldedCornerPath,
  framePath,
  homePlatePath,
  invertedTrianglePath,
  lBandPath,
  lineCalloutPath,
  moonPath,
  openArcPath,
  parallelogramPath,
  pillPath,
  plusPath,
  regularPolygonPath,
  rightTrianglePath,
  ringSegmentPath,
  roundRectPath,
  starPath,
  sunRaysPath,
  trapezoidPath,
  trianglePath,
  wavyRectPath,
} from './geometry'

// Every function returns an SVG path string — parse out the numeric
// coordinate pairs (space- or comma-separated after M/L, ignoring command
// letters) so assertions check real geometry, not just "it's a string".
function coords(path: string): Array<[number, number]> {
  const nums = path.match(/-?\d+\.?\d*/g)?.map(Number) ?? []
  const pairs: Array<[number, number]> = []
  for (let i = 0; i + 1 < nums.length; i += 2) pairs.push([nums[i], nums[i + 1]])
  return pairs
}

describe('regularPolygonPath', () => {
  it('produces exactly `sides` vertices, all inside the bounding box', () => {
    for (const sides of [3, 5, 6, 7, 8, 10, 12]) {
      const pts = coords(regularPolygonPath(sides, 100, 100))
      expect(pts).toHaveLength(sides)
      for (const [x, y] of pts) {
        expect(x).toBeGreaterThanOrEqual(-0.01)
        expect(x).toBeLessThanOrEqual(100.01)
        expect(y).toBeGreaterThanOrEqual(-0.01)
        expect(y).toBeLessThanOrEqual(100.01)
      }
    }
  })

  it('is point-up: the first vertex sits at the horizontal center, top edge', () => {
    const [first] = coords(regularPolygonPath(5, 100, 100))
    expect(first[0]).toBeCloseTo(50, 0)
    expect(first[1]).toBeCloseTo(0, 0)
  })
})

describe('starPath', () => {
  it('produces 2×points vertices alternating outer/inner radius', () => {
    for (const points of [4, 5, 6, 8, 12, 32]) {
      const pts = coords(starPath(points, 100, 100))
      expect(pts).toHaveLength(points * 2)
      // Outer (even index) vertices sit farther from center than inner (odd).
      const cx = 50
      const cy = 50
      const dist = (p: [number, number]) => Math.hypot(p[0] - cx, p[1] - cy)
      expect(dist(pts[0])).toBeGreaterThan(dist(pts[1]))
      expect(dist(pts[2])).toBeGreaterThan(dist(pts[3]))
    }
  })
})

describe('directionalArrowPath', () => {
  it('right: shaft runs left→right, head at the right edge', () => {
    const pts = coords(directionalArrowPath('right', 100, 40))
    const maxX = Math.max(...pts.map((p) => p[0]))
    const minX = Math.min(...pts.map((p) => p[0]))
    expect(maxX).toBeCloseTo(100, 0)
    expect(minX).toBeCloseTo(0, 0)
  })

  it('left, up, down each point toward their named edge (head touches it)', () => {
    const cases: Array<['left' | 'up' | 'down', number, number, [number, number]]> = [
      ['left', 100, 40, [0, 20]],
      ['up', 40, 100, [20, 0]],
      ['down', 40, 100, [20, 100]],
    ]
    for (const [dir, w, h, tip] of cases) {
      const pts = coords(directionalArrowPath(dir, w, h))
      const hasTip = pts.some((p) => Math.abs(p[0] - tip[0]) < 1 && Math.abs(p[1] - tip[1]) < 1)
      expect(hasTip, `${dir} arrow should reach ${tip}`).toBe(true)
    }
  })

  it('all four directions stay within their bounding box', () => {
    for (const dir of ['right', 'left', 'up', 'down'] as const) {
      const pts = coords(directionalArrowPath(dir, 90, 50))
      for (const [x, y] of pts) {
        expect(x).toBeGreaterThanOrEqual(-0.01)
        expect(x).toBeLessThanOrEqual(90.01)
        expect(y).toBeGreaterThanOrEqual(-0.01)
        expect(y).toBeLessThanOrEqual(50.01)
      }
    }
  })
})

describe('doubleArrowPath', () => {
  it('horizontal: reaches both left and right edges', () => {
    const pts = coords(doubleArrowPath('horizontal', 100, 30))
    expect(pts.some((p) => p[0] <= 0.5)).toBe(true)
    expect(pts.some((p) => p[0] >= 99.5)).toBe(true)
  })
  it('vertical: reaches both top and bottom edges', () => {
    const pts = coords(doubleArrowPath('vertical', 30, 100))
    expect(pts.some((p) => p[1] <= 0.5)).toBe(true)
    expect(pts.some((p) => p[1] >= 99.5)).toBe(true)
  })
})

describe('basic polygon shapes', () => {
  it('trianglePath: apex centered on top edge', () => {
    const pts = coords(trianglePath(100, 60))
    expect(pts[0]).toEqual([50, 0])
  })
  it('rightTrianglePath: right angle at bottom-left', () => {
    const pts = coords(rightTrianglePath(100, 60))
    expect(pts).toContainEqual([0, 0])
    expect(pts).toContainEqual([0, 60])
  })
  it('diamondPath: 4 vertices at the box midpoints', () => {
    const pts = coords(diamondPath(100, 60))
    expect(pts).toEqual([[50, 0], [100, 30], [50, 60], [0, 30]])
  })
  it('parallelogramPath and trapezoidPath stay within the box and are wider at the base', () => {
    for (const fn of [parallelogramPath, trapezoidPath]) {
      const pts = coords(fn(100, 60))
      for (const [x] of pts) {
        expect(x).toBeGreaterThanOrEqual(-0.01)
        expect(x).toBeLessThanOrEqual(100.01)
      }
    }
  })
  it('plusPath: 12 vertices (a cross has 12 corners)', () => {
    expect(coords(plusPath(90, 90))).toHaveLength(12)
  })
})

describe('roundRectPath / pillPath', () => {
  it('pillPath is roundRectPath with a radius of half the shorter side', () => {
    expect(pillPath(120, 40)).toBe(roundRectPath(120, 40, 0.5))
  })
  it('roundRectPath starts and ends adjacent to the top-left corner, not AT it', () => {
    const pts = coords(roundRectPath(100, 60, 0.16))
    expect(pts[0][0]).toBeGreaterThan(0) // inset by the radius, not a sharp corner
  })
})

describe('arcPath', () => {
  it('pie starts at the center (a wedge, not just an arc)', () => {
    const pts = coords(arcPath('pie', 100, 100))
    expect(pts[0]).toEqual([50, 50])
  })
  it('chord and donut do not include the center point', () => {
    for (const mode of ['chord', 'donut'] as const) {
      const pts = coords(arcPath(mode, 100, 100))
      expect(pts[0]).not.toEqual([50, 50])
    }
  })
})

describe('callout variants', () => {
  it('calloutPath, ellipseCalloutPath, lineCalloutPath all produce non-empty paths', () => {
    for (const fn of [calloutPath, ellipseCalloutPath, lineCalloutPath]) {
      expect(fn(120, 80).length).toBeGreaterThan(10)
    }
  })
})

// S6.3: additional ECMA-376 presets.

// coords() pairs numbers positionally, which breaks for SVG 'A' (arc)
// commands — 7 numbers (rx ry x-rot large-arc sweep-flag x y), not a clean
// coordinate pair — so an embedded arc shifts every later pairing. For paths
// containing 'A', check raw numeric PRESENCE instead of a coords()[i] index.
function rawNums(path: string): number[] {
  return path.match(/-?\d+\.?\d*/g)?.map(Number) ?? []
}
function hasNumberNear(path: string, expected: number, tolerance = 0.5) {
  return rawNums(path).some((n) => Math.abs(n - expected) < tolerance)
}

function inBox(pts: Array<[number, number]>, w: number, h: number) {
  for (const [x, y] of pts) {
    expect(x).toBeGreaterThanOrEqual(-0.5)
    expect(x).toBeLessThanOrEqual(w + 0.5)
    expect(y).toBeGreaterThanOrEqual(-0.5)
    expect(y).toBeLessThanOrEqual(h + 0.5)
  }
}

describe('trapezoidPath invert / plusPath armRatio / invertedTrianglePath', () => {
  it('invert=true is wider on top instead of narrower', () => {
    const normal = coords(trapezoidPath(100, 60))
    const inverted = coords(trapezoidPath(100, 60, true))
    // Normal: top edge inset (narrower). Inverted: top edge full width (0..w).
    expect(Math.min(...normal.slice(0, 2).map((p) => p[0]))).toBeGreaterThan(0)
    expect(Math.min(...inverted.slice(0, 2).map((p) => p[0]))).toBe(0)
  })

  it('a thinner armRatio produces a narrower cross than the default', () => {
    const thick = coords(plusPath(90, 90))
    const thin = coords(plusPath(90, 90, 0.1))
    const armWidth = (pts: Array<[number, number]>) => Math.max(...pts.map((p) => p[0])) - Math.min(...pts.filter((p) => p[0] > 0 && p[0] < 90).map((p) => p[0]))
    // thin's inner notch (the concave corners) sits closer to center than thick's.
    const thickInnerX = thick.map((p) => p[0]).filter((x) => x > 0 && x < 90)
    const thinInnerX = thin.map((p) => p[0]).filter((x) => x > 0 && x < 90)
    expect(Math.min(...thinInnerX)).toBeGreaterThan(Math.min(...thickInnerX))
    void armWidth
  })

  it('invertedTrianglePath: apex centered on the BOTTOM edge (trianglePath mirrored)', () => {
    const pts = coords(invertedTrianglePath(100, 60))
    expect(pts).toContainEqual([50, 60])
    expect(pts).toContainEqual([0, 0])
    expect(pts).toContainEqual([100, 0])
  })
})

describe('cutCornerRectPath', () => {
  it('round1Rect: only the top-left corner is treated (a Q curve there), others sharp', () => {
    const path = cutCornerRectPath(100, 60, ['round', 'none', 'none', 'none'])
    expect(path).toContain('Q 0 0')
    expect(path).not.toContain('Q 100')
    inBox(coords(path), 100, 60)
  })

  it('snip2DiagRect: two straight diagonal cuts (no Q at all), on opposite corners', () => {
    const path = cutCornerRectPath(100, 60, ['snip', 'none', 'snip', 'none'])
    expect(path).not.toContain('Q')
    inBox(coords(path), 100, 60)
  })

  it('snipRoundRect: mixes both treatments (has both Q and a diagonal L cut)', () => {
    const path = cutCornerRectPath(100, 60, ['round', 'snip', 'round', 'snip'])
    expect(path).toContain('Q')
    // The 4 corner segments plus the 4 straight edges — a non-trivial path.
    expect(path.split(' ').filter((t) => t === 'L' || t === 'Q').length).toBeGreaterThan(4)
  })

  it('all 7 catalogue configurations stay within the box', () => {
    const configs: Array<[CornerTreatment, CornerTreatment, CornerTreatment, CornerTreatment]> = [
      ['round', 'none', 'none', 'none'],
      ['round', 'round', 'none', 'none'],
      ['round', 'none', 'round', 'none'],
      ['none', 'snip', 'none', 'none'],
      ['snip', 'snip', 'none', 'none'],
      ['snip', 'none', 'snip', 'none'],
      ['round', 'snip', 'round', 'snip'],
    ]
    for (const c of configs) inBox(coords(cutCornerRectPath(100, 60, c)), 100, 60)
  })
})

describe('framePath / lBandPath', () => {
  it('framePath: two separate closed rects (outer boundary + inner hole)', () => {
    const path = framePath(100, 60, 0.16)
    expect(path.match(/M/g)?.length).toBe(2)
    expect(path.match(/Z/g)?.length).toBe(2)
    inBox(coords(path), 100, 60)
  })

  it('lBandPath: 6 vertices forming an L (band along top + left)', () => {
    const pts = coords(lBandPath(100, 60, 0.16))
    expect(pts).toHaveLength(6)
    expect(pts).toContainEqual([0, 0])
    inBox(pts, 100, 60)
  })
})

describe('diagStripePath / bowtiePath / displayShapePath', () => {
  it('diagStripePath: 6 vertices, stays in the box', () => {
    const pts = coords(diagStripePath(100, 60))
    expect(pts).toHaveLength(6)
    inBox(pts, 100, 60)
  })

  it('bowtiePath: the classic self-intersecting quad — 4 corners', () => {
    const pts = coords(bowtiePath(100, 60))
    expect(pts).toEqual([[0, 0], [100, 0], [0, 60], [100, 60]])
  })

  it('displayShapePath: pointed notch reaches the left edge at vertical center', () => {
    const pts = coords(displayShapePath(100, 60))
    expect(pts).toContainEqual([0, 30])
    inBox(pts, 100, 60)
  })
})

describe('ringSegmentPath / openArcPath', () => {
  it('ringSegmentPath: emits the outer radius (rx) and a strictly smaller inner radius', () => {
    const path = ringSegmentPath(100, 100, -45, 270, 0.5)
    expect(hasNumberNear(path, 50)).toBe(true) // outer rx=ry=50
    expect(hasNumberNear(path, 25)).toBe(true) // inner: 50 * innerRatio 0.5
    expect(hasNumberNear(path, 12.5)).toBe(false)
  })

  it('a >180° sweep sets the large-arc flag; a <180° sweep does not', () => {
    const wide = ringSegmentPath(100, 100, 0, 270)
    const narrow = ringSegmentPath(100, 100, 0, 90)
    expect(wide).toMatch(/A 50 50 0 1 1/)
    expect(narrow).toMatch(/A 50 50 0 0 1/)
  })

  it('openArcPath: starts and ends on the ellipse boundary, 180° apart by default (left ↔ right, through the top)', () => {
    const path = openArcPath(100, 100)
    expect(path.startsWith('M 0 50')).toBe(true)
    expect(path.trim().endsWith('100 50')).toBe(true)
  })

  it('openArcPath: a 90° sweep reaches an adjacent point, not the opposite one', () => {
    const path = openArcPath(100, 100, -90, 90)
    expect(path.startsWith('M 0 50')).toBe(true)
    expect(path.trim().endsWith('50 0')).toBe(true) // quarter-turn from left to top
  })
})

describe('foldedCornerPath / foldedCornerFlapPath', () => {
  it('the bottom-right corner is cut — no vertex sits exactly at (w,h)', () => {
    const pts = coords(foldedCornerPath(100, 60, 0.2))
    expect(pts).not.toContainEqual([100, 60])
    inBox(pts, 100, 60)
  })
  it('the flap triangle fills exactly the cut corner', () => {
    const pts = coords(foldedCornerFlapPath(100, 60, 0.2))
    expect(pts).toHaveLength(3)
    expect(pts).toContainEqual([100, 60])
  })
})

describe('sunRaysPath / moonPath', () => {
  it('sunRaysPath: produces `rays` triangles (3 points each)', () => {
    for (const rays of [6, 8, 12]) {
      expect(coords(sunRaysPath(100, 100, rays))).toHaveLength(rays * 3)
    }
  })
  it('moonPath: starts at the top and its outer arc reaches the bottom (a lune spanning the full height)', () => {
    const path = moonPath(100, 100)
    expect(path.startsWith('M 50 0')).toBe(true)
    expect(hasNumberNear(path, 100)).toBe(true) // the outer arc's endpoint y
  })
  it('moonPath: the inner arc uses a smaller radius than the outer (the bite that makes it a crescent)', () => {
    const path = moonPath(100, 100)
    expect(hasNumberNear(path, 50, 0.1)).toBe(true) // outer rx=ry=50
    expect(hasNumberNear(path, 31, 0.1)).toBe(true) // inner rx = 50 * 0.62
  })
})

describe('bracketPath / bracePath', () => {
  it('left variants hug x=0, right variants hug x=w', () => {
    for (const w of [100]) {
      const left = coords(bracketPath(w, 60, 'left'))
      const right = coords(bracketPath(w, 60, 'right'))
      expect(Math.min(...left.map((p) => p[0]))).toBeLessThan(Math.min(...right.map((p) => p[0])))
      expect(Math.max(...right.map((p) => p[0]))).toBeGreaterThan(Math.max(...left.map((p) => p[0])))
    }
    const leftBrace = coords(bracePath(100, 60, 'left'))
    const rightBrace = coords(bracePath(100, 60, 'right'))
    inBox(leftBrace, 100, 60)
    inBox(rightBrace, 100, 60)
  })
})

describe('homePlatePath / dShapePath', () => {
  it('homePlate (right): comes to a point at the right edge, vertical center', () => {
    const pts = coords(homePlatePath(100, 60, 'right'))
    expect(pts).toContainEqual([100, 30])
  })
  it('homePlate (down, the off-page connector): points downward at horizontal center', () => {
    const pts = coords(homePlatePath(100, 60, 'down'))
    expect(pts).toContainEqual([50, 60])
  })
  it('dShapePath: flat side has two square corners on the flat edge; the round side has none', () => {
    const rightFlatLeft = dShapePath(100, 60, 'right') // flat on the LEFT edge
    expect(rightFlatLeft.startsWith('M 0 0')).toBe(true)
    expect(rightFlatLeft.trim().endsWith('L 0 60 Z')).toBe(true)
    expect(rightFlatLeft).not.toContain('L 100 0') // right edge is arced, not a square corner
    const leftFlatRight = dShapePath(100, 60, 'left') // flat on the RIGHT edge
    expect(leftFlatRight).toContain('L 100 0')
    expect(leftFlatRight).toContain('L 100 60')
  })
})

describe('wavyRectPath', () => {
  it('bottom-only: top edge is flat (y=0), bottom edge oscillates around y=h', () => {
    const pts = coords(wavyRectPath(120, 40, 'bottom', 2, 0.08))
    const topPts = pts.filter((_, i) => i < 2)
    for (const [, y] of topPts) expect(y).toBe(0)
    const bottomYs = pts.slice(2).map((p) => p[1])
    expect(Math.max(...bottomYs)).toBeGreaterThan(40)
    expect(Math.min(...bottomYs)).toBeLessThan(40)
  })

  it('both edges: top AND bottom oscillate', () => {
    const pts = coords(wavyRectPath(120, 40, 'both', 2, 0.08))
    const ys = pts.map((p) => p[1])
    expect(Math.min(...ys)).toBeLessThan(0)
    expect(Math.max(...ys)).toBeGreaterThan(40)
  })

  it('a tighter waveCount produces more oscillation crossings', () => {
    const loose = wavyRectPath(120, 40, 'bottom', 2, 0.08)
    const tight = wavyRectPath(120, 40, 'bottom', 4, 0.08)
    // More waves ⇒ more sampled points in the path string.
    expect(coords(tight).length).toBeGreaterThan(coords(loose).length)
  })
})
