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
const HEXAGON = [[250_000, 0], [750_000, 0], [1_000_000, 500_000], [750_000, 1_000_000], [250_000, 1_000_000], [0, 500_000]] as const
const STAR5 = [
  [500_000, 0], [617_557, 338_197], [975_528, 345_492], [690_211, 561_803], [793_893, 904_508],
  [500_000, 700_000], [206_107, 904_508], [309_789, 561_803], [24_472, 345_492], [382_443, 338_197],
] as const

/** Preset paths contain only integer EMU; guide results round once at the path boundary. */
export function presetPath(preset: NativeShapePreset, cx: number, cy: number): readonly RenderPathCommand[] {
  const rect = localBounds(cx, cy)
  switch (preset) {
    case 'rect': return [{ kind: 'rect', rect }]
    case 'roundRect': return [{ kind: 'roundRect', rect, radiusEmu: Math.max(1, Math.round(Math.min(cx, cy) / 8)) }]
    case 'ellipse': return [{ kind: 'ellipse', rect }]
    case 'triangle': return polygon(cx, cy, [[500_000, 0], [1_000_000, 1_000_000], [0, 1_000_000]])
    case 'diamond': return polygon(cx, cy, [[500_000, 0], [1_000_000, 500_000], [500_000, 1_000_000], [0, 500_000]])
    case 'rightArrow': return polygon(cx, cy, [[0, 250_000], [625_000, 250_000], [625_000, 0], [1_000_000, 500_000], [625_000, 1_000_000], [625_000, 750_000], [0, 750_000]])
    case 'pentagon': return pentagon(cx, cy)
    case 'hexagon': return polygon(cx, cy, HEXAGON)
    case 'star5': return polygon(cx, cy, STAR5)
  }
}

export function connectorPath(cx: number, cy: number, flipH: boolean): readonly RenderPathCommand[] {
  return flipH
    ? [{ kind: 'moveTo', x: cx, y: 0 }, { kind: 'lineTo', x: 0, y: cy }]
    : [{ kind: 'moveTo', x: 0, y: 0 }, { kind: 'lineTo', x: cx, y: cy }]
}
