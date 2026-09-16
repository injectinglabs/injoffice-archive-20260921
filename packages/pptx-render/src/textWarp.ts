import type { RenderRect, RenderTextBodyNode, RenderTextRunNode, RenderTransform } from './types.js'

export const MODELED_PRESET_TEXT_WARPS = ['textArchUp', 'textArchDown', 'textDeflate'] as const
export type ModeledPresetTextWarp = (typeof MODELED_PRESET_TEXT_WARPS)[number]

const PPM = 1_000_000
const DEFLATE_DEFAULT_ADJ = 18_750
const DEFLATE_MAX_ADJ = 50_000

export function isModeledPresetTextWarp(preset: string | undefined): preset is ModeledPresetTextWarp {
  return preset === 'textArchUp' || preset === 'textArchDown' || preset === 'textDeflate'
}

export interface PresetTextWarpSpec {
  readonly preset: ModeledPresetTextWarp
  readonly adj?: number
  readonly bounds: RenderRect
}

export function textBodyWarp(textBody: Pick<RenderTextBodyNode, 'bounds' | 'presetTextWarp' | 'presetTextWarpAdj'>): PresetTextWarpSpec | undefined {
  // Warp in the text body's local frame. Body/orientation transforms compose
  // around that local paint; skipping here would advertise warp and paint flat.
  if (!isModeledPresetTextWarp(textBody.presetTextWarp)) return undefined
  if (!Number.isSafeInteger(textBody.bounds.cx) || !Number.isSafeInteger(textBody.bounds.cy) || textBody.bounds.cx <= 0 || textBody.bounds.cy <= 0) return undefined
  const adj = textBody.presetTextWarpAdj
  if (adj !== undefined && (!Number.isSafeInteger(adj) || adj < 0 || adj > 100_000)) return undefined
  return { preset: textBody.presetTextWarp, ...(adj === undefined ? {} : { adj }), bounds: textBody.bounds }
}

/** Integer parabola dip 4*dy*x*(w-x)/w^2 peaking at dy in the middle. */
export function warpParabolaDip(x: bigint, width: bigint, dy: bigint): bigint {
  if (width <= 0n) return 0n
  const clamped = x < 0n ? 0n : x > width ? width : x
  return (4n * dy * clamped * (width - clamped)) / (width * width)
}

export function warpPoint(spec: PresetTextWarpSpec, x: number, y: number): { readonly x: number; readonly y: number; readonly theta: number } {
  const width = BigInt(spec.bounds.cx)
  const height = BigInt(spec.bounds.cy)
  const relX = BigInt(x) - BigInt(spec.bounds.x)
  const relY = BigInt(y) - BigInt(spec.bounds.y)
  const adj = spec.adj ?? (spec.preset === 'textDeflate' ? DEFLATE_DEFAULT_ADJ : 0)
  const bulge = spec.preset === 'textDeflate'
    ? (BigInt(Math.min(DEFLATE_MAX_ADJ, Math.max(0, adj))) * height) / 100000n
    : adj === 0 ? height / 2n : (BigInt(Math.min(DEFLATE_MAX_ADJ, Math.max(0, adj))) * height) / 100000n
  const dip = warpParabolaDip(relX, width, bulge)
  let outY = relY
  let slopeNum = 0n
  if (spec.preset === 'textDeflate') {
    // Envelope: top arch-down, bottom arch-up. v=0.5 is a straight waist.
    const span = height - 2n * dip
    outY = height === 0n ? relY : dip + (relY * span) / height
    slopeNum = height === 0n ? 0n : ((height - 2n * relY) * 4n * bulge * (width - 2n * relX)) / height
  } else if (spec.preset === 'textArchDown') {
    outY = relY + dip
    slopeNum = 4n * bulge * (width - 2n * relX)
  } else {
    outY = relY - dip
    slopeNum = -4n * bulge * (width - 2n * relX)
  }
  const theta = Math.atan2(Number(slopeNum), Number(width * width))
  return { x, y: spec.bounds.y + Number(outY), theta: Number.isFinite(theta) ? theta : 0 }
}

function ppm(value: number): number {
  const rounded = Math.round(value)
  return rounded === 0 ? 0 : rounded
}

export function warpGlyphTransform(fromX: number, fromY: number, toX: number, toY: number, theta: number): RenderTransform | undefined {
  const cos = Math.cos(theta), sin = Math.sin(theta)
  const aPpm = ppm(cos * PPM), bPpm = ppm(sin * PPM), cPpm = ppm(-sin * PPM), dPpm = ppm(cos * PPM)
  const txEmu = ppm(toX - cos * fromX + sin * fromY)
  const tyEmu = ppm(toY - sin * fromX - cos * fromY)
  const values = [aPpm, bPpm, cPpm, dPpm, txEmu, tyEmu]
  if (values.some((value) => !Number.isSafeInteger(value))) return undefined
  if (aPpm === PPM && bPpm === 0 && cPpm === 0 && dPpm === PPM && txEmu === 0 && tyEmu === 0) return undefined
  return { aPpm, bPpm, cPpm, dPpm, txEmu, tyEmu }
}

export function warpGlyphRun(run: RenderTextRunNode, spec: PresetTextWarpSpec | undefined): readonly { readonly transform?: RenderTransform; readonly run: RenderTextRunNode }[] {
  if (!spec || run.status === 'refused' || run.glyphs.length === 0) return [{ run }]
  const painted: { readonly transform?: RenderTransform; readonly run: RenderTextRunNode }[] = []
  for (const glyph of run.glyphs) {
    const fromX = run.x + glyph.xEmu, fromY = run.baselineY + glyph.yEmu
    if (![fromX, fromY].every((value) => Number.isSafeInteger(value))) return [{ run }]
    const warped = warpPoint(spec, fromX, fromY)
    const transform = warpGlyphTransform(fromX, fromY, warped.x, warped.y, warped.theta)
    painted.push({ ...(transform ? { transform } : {}), run: { ...run, glyphs: [glyph] } })
  }
  return painted
}
