import type { NativeShapePreset,NativeTransform } from '@injoffice/pptx-native'
import type { RenderPathCommand, RenderRect, RenderTransform } from './types.js'

export const IDENTITY_PPM = 1_000_000

export function translationTransform(txEmu: number, tyEmu: number): RenderTransform {
  return { aPpm: IDENTITY_PPM, bPpm: 0, cPpm: 0, dPpm: IDENTITY_PPM, txEmu, tyEmu }
}

/** Clockwise rotation about the original DrawingML frame center, without trig. */
export function quarterTurnTransform(frame:NativeTransform):RenderTransform {
  const {x,y,cx,cy,quarterTurns:q}=frame
  if(!q)return translationTransform(x,y)
  const value=(n:bigint)=>{const v=Number(n);if(!Number.isSafeInteger(v))throw new RangeError('rotation translation exceeds integer precision');return v}
  if(q!==2&&cx%2!==cy%2)throw new RangeError('rotation center requires fractional EMU')
  const sum=BigInt(cx)+BigInt(cy),difference=BigInt(cx)-BigInt(cy)
  if(q===1)return {aPpm:0,bPpm:IDENTITY_PPM,cPpm:-IDENTITY_PPM,dPpm:0,txEmu:value(BigInt(x)+sum/2n),tyEmu:value(BigInt(y)-difference/2n)}
  if(q===2)return {aPpm:-IDENTITY_PPM,bPpm:0,cPpm:0,dPpm:-IDENTITY_PPM,txEmu:value(BigInt(x)+BigInt(cx)),tyEmu:value(BigInt(y)+BigInt(cy))}
  return {aPpm:0,bPpm:-IDENTITY_PPM,cPpm:IDENTITY_PPM,dPpm:0,txEmu:value(BigInt(x)+difference/2n),tyEmu:value(BigInt(y)+sum/2n)}
}

export function localBounds(cx: number, cy: number): RenderRect {
  return { x: 0, y: 0, cx, cy }
}

function scalePpm(value: number, ppm: number): number {
  const product = value * ppm
  if (!Number.isSafeInteger(product)) {
    const quotient = value / IDENTITY_PPM * ppm
    if (!Number.isSafeInteger(Math.round(quotient))) throw new RangeError('preset geometry exceeds integer precision')
    return Math.round(quotient)
  }
  return Math.round(product / IDENTITY_PPM)
}

function polygon(cx: number, cy: number, pointsPpm: readonly (readonly [number, number])[]): readonly RenderPathCommand[] {
  const points = pointsPpm.map(([x, y]) => ({ x: scalePpm(cx, x), y: scalePpm(cy, y) }))
  return [
    { kind: 'moveTo', x: points[0]!.x, y: points[0]!.y },
    ...points.slice(1).map(({ x, y }) => ({ kind: 'lineTo' as const, x, y })),
    { kind: 'close' },
  ]
}

/** DrawingML default pentagon: hf=105146, vf=110557, not an inscribed regular polygon.
 * Derived from the public preset equations (also documented in Apache POI's
 * presetShapeDefinitions.xml). Keep guide precision until the final EMU rounding.
 */
function pentagonGuides(cx: number, cy: number) {
  const halfX = cx / 2
  const scaledX = halfX * 105146 / 100000
  const scaledY = cy / 2 * 110557 / 100000
  const root5 = Math.sqrt(5)
  // cos(18°), cos(54°), sin(18°), sin(54°) in radical form.
  const dx1 = scaledX * Math.sqrt(10 + 2 * root5) / 4
  const dx2 = scaledX * Math.sqrt(10 - 2 * root5) / 4
  const y1 = scaledY * (1 - (root5 - 1) / 4)
  const y2 = scaledY * (1 + (root5 + 1) / 4)
  return { halfX, dx1, dx2, y1, y2 }
}

/** ECMA-376 pentagon text rect: x2, y1*dx2/dx1, x3, y2; before body insets. */
export function defaultPentagonTextRect(cx: number, cy: number): RenderRect {
  const { halfX, dx1, dx2, y1, y2 } = pentagonGuides(cx, cy)
  const x = Math.round(halfX - dx2)
  const y = Math.round(y1 * dx2 / dx1)
  return { x, y, cx: Math.round(halfX + dx2) - x, cy: Math.round(y2) - y }
}

function pentagon(cx: number, cy: number): readonly RenderPathCommand[] {
  const { halfX, dx1, dx2, y1, y2 } = pentagonGuides(cx, cy)
  const points = [[halfX - dx1, y1], [halfX, 0], [halfX + dx1, y1], [halfX + dx2, y2], [halfX - dx2, y2]]
  return [
    ...points.map(([x, y], index) => ({
      kind: index === 0 ? 'moveTo' as const : 'lineTo' as const,
      x: Math.round(x!), y: Math.round(y!),
    })),
    { kind: 'close' },
  ]
}
// DrawingML default guides. ss is the shorter side, not always the width.
function hexagonGuides(cx: number, cy: number) {
  const ss = Math.min(cx, cy)
  const dx = ss / 4
  const dy = cy / 2 * 115470 / 100000 * Math.sqrt(3) / 2
  return { dx, top: cy / 2 - dy, bottom: cy / 2 + dy }
}

function guidePolygon(points: readonly (readonly [number, number])[]): readonly RenderPathCommand[] {
  return [...points.map(([x, y], i) => ({ kind: i === 0 ? 'moveTo' as const : 'lineTo' as const, x: Math.round(x), y: Math.round(y) })), { kind: 'close' }]
}

/** Default DrawingML text rectangles, before a:bodyPr insets. */
export function defaultPresetTextRect(preset: NativeShapePreset, cx: number, cy: number): RenderRect {
  if (preset === 'pentagon') return defaultPentagonTextRect(cx, cy)
  if (preset === 'star5') {
    const { outer, dx1, y1, y3 } = star5Guides(cx, cy)
    const x = Math.round(outer.halfX - dx1), y = Math.round(y1)
    return { x, y, cx: Math.round(outer.halfX + dx1) - x, cy: Math.round(y3) - y }
  }
  let left = 0, top = 0, right = cx, bottom = cy
  if (preset === 'roundRect') {
    left = top = Math.min(cx, cy) * 16667 / 100000 * 29289 / 100000
    right -= left; bottom -= top
  } else if (preset === 'rightArrow') {
    top = cy / 4; bottom = cy * 3 / 4
    right = cx - Math.min(cx, cy) / 4
  } else if (preset === 'hexagon') {
    // Default adj <= maxAdj/2, so q8 = 2 + 2*ss/w.
    const q8 = 2 + 2 * Math.min(cx, cy) / cx
    left = cx * q8 / 24; top = cy * q8 / 24
    right -= left; bottom -= top
  }
  const x = Math.round(left), y = Math.round(top)
  return { x, y, cx: Math.round(right) - x, cy: Math.round(bottom) - y }
}
/** DrawingML default star5 shares the pentagon outer guides, but uses
 * adj=19098 for its five inner vertices (not an inscribed generic star). */
function star5Guides(cx: number, cy: number) {
  const outer = pentagonGuides(cx, cy)
  const scaledY = cy / 2 * 110557 / 100000
  const innerX = cx / 2 * 105146 / 100000 * 19098 / 50000
  const innerY = scaledY * 19098 / 50000
  const root5 = Math.sqrt(5)
  const dx1 = innerX * Math.sqrt(10 + 2 * root5) / 4 // cos(342°)
  const dx2 = innerX * Math.sqrt(10 - 2 * root5) / 4 // cos(54°)
  const y1 = scaledY - innerY * (root5 + 1) / 4 // sin(54°)
  const y2 = scaledY + innerY * (root5 - 1) / 4 // -sin(342°)
  return { outer, dx1, dx2, y1, y2, y3: scaledY + innerY }
}

function star5(cx: number, cy: number): readonly RenderPathCommand[] {
  const { outer: o, dx1, dx2, y1, y2, y3 } = star5Guides(cx, cy)
  return guidePolygon([
    [o.halfX - o.dx1, o.y1], [o.halfX - dx2, y1], [o.halfX, 0],
    [o.halfX + dx2, y1], [o.halfX + o.dx1, o.y1], [o.halfX + dx1, y2],
    [o.halfX + o.dx2, o.y2], [o.halfX, y3], [o.halfX - o.dx2, o.y2],
    [o.halfX - dx1, y2],
  ])
}

/** Preset paths contain only integer EMU; guide results round once at the path boundary. */
export function presetPath(preset: NativeShapePreset, cx: number, cy: number): readonly RenderPathCommand[] {
  const rect = localBounds(cx, cy)
  switch (preset) {
    case 'rect': return [{ kind: 'rect', rect }]
    case 'roundRect': return [{ kind: 'roundRect', rect, radiusEmu: Math.round(Math.min(cx, cy) * 16667 / 100000) }]
    case 'ellipse': return [{ kind: 'ellipse', rect }]
    case 'triangle': return polygon(cx, cy, [[500_000, 0], [1_000_000, 1_000_000], [0, 1_000_000]])
    case 'diamond': return polygon(cx, cy, [[500_000, 0], [1_000_000, 500_000], [500_000, 1_000_000], [0, 500_000]])
    case 'rightArrow': {
      const shoulder = cx - Math.min(cx, cy) / 2
      return guidePolygon([[0, cy / 4], [shoulder, cy / 4], [shoulder, 0], [cx, cy / 2], [shoulder, cy], [shoulder, cy * 3 / 4], [0, cy * 3 / 4]])
    }
    case 'pentagon': return pentagon(cx, cy)
    case 'hexagon': {
      const { dx, top, bottom } = hexagonGuides(cx, cy)
      return guidePolygon([[0, cy / 2], [dx, top], [cx - dx, top], [cx, cy / 2], [cx - dx, bottom], [dx, bottom]])
    }
    case 'star5': return star5(cx, cy)
  }
}

export function connectorPath(cx: number, cy: number, flipH: boolean): readonly RenderPathCommand[] {
  return flipH
    ? [{ kind: 'moveTo', x: cx, y: 0 }, { kind: 'lineTo', x: 0, y: cy }]
    : [{ kind: 'moveTo', x: 0, y: 0 }, { kind: 'lineTo', x: cx, y: cy }]
}
