import type { NativeShapePreset } from '@injoffice/pptx-native'
import type { RenderPathCommand, RenderRect, RenderTransform } from './types.js'

export const IDENTITY_PPM = 1_000_000

export function translationTransform(txEmu: number, tyEmu: number): RenderTransform {
  return { aPpm: IDENTITY_PPM, bPpm: 0, cPpm: 0, dPpm: IDENTITY_PPM, txEmu, tyEmu }
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

const PENTAGON = [[500_000, 0], [975_528, 345_492], [793_893, 904_508], [206_107, 904_508], [24_472, 345_492]] as const
const HEXAGON = [[250_000, 0], [750_000, 0], [1_000_000, 500_000], [750_000, 1_000_000], [250_000, 1_000_000], [0, 500_000]] as const
const STAR5 = [
  [500_000, 0], [617_557, 338_197], [975_528, 345_492], [690_211, 561_803], [793_893, 904_508],
  [500_000, 700_000], [206_107, 904_508], [309_789, 561_803], [24_472, 345_492], [382_443, 338_197],
] as const

/** Preset paths contain only integer EMU and fixed integer ratios. */
export function presetPath(preset: NativeShapePreset, cx: number, cy: number): readonly RenderPathCommand[] {
  const rect = localBounds(cx, cy)
  switch (preset) {
    case 'rect': return [{ kind: 'rect', rect }]
    case 'roundRect': return [{ kind: 'roundRect', rect, radiusEmu: Math.max(1, Math.round(Math.min(cx, cy) / 8)) }]
    case 'ellipse': return [{ kind: 'ellipse', rect }]
    case 'triangle': return polygon(cx, cy, [[500_000, 0], [1_000_000, 1_000_000], [0, 1_000_000]])
    case 'diamond': return polygon(cx, cy, [[500_000, 0], [1_000_000, 500_000], [500_000, 1_000_000], [0, 500_000]])
    case 'rightArrow': return polygon(cx, cy, [[0, 250_000], [625_000, 250_000], [625_000, 0], [1_000_000, 500_000], [625_000, 1_000_000], [625_000, 750_000], [0, 750_000]])
    case 'pentagon': return polygon(cx, cy, PENTAGON)
    case 'hexagon': return polygon(cx, cy, HEXAGON)
    case 'star5': return polygon(cx, cy, STAR5)
  }
}

export function connectorPath(cx: number, cy: number, flipH: boolean): readonly RenderPathCommand[] {
  return flipH
    ? [{ kind: 'moveTo', x: cx, y: 0 }, { kind: 'lineTo', x: 0, y: cy }]
    : [{ kind: 'moveTo', x: 0, y: 0 }, { kind: 'lineTo', x: cx, y: cy }]
}
