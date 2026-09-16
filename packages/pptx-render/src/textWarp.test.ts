import { describe, expect, it } from 'vitest'
import { warpGlyphRun, warpGlyphTransform, warpParabolaDip, warpPoint, type PresetTextWarpSpec } from './textWarp.js'
import type { RenderTextRunNode } from './types.js'

const bounds = { x: 100, y: 200, cx: 1_000, cy: 800 }

function spec(preset: PresetTextWarpSpec['preset'], adj?: number): PresetTextWarpSpec {
  return { preset, bounds, ...(adj === undefined ? {} : { adj }) }
}

describe('preset text warp', () => {
  it('deflates along arch-down top and arch-up bottom with a straight waist', () => {
    const deflate = spec('textDeflate', 37_500)
    const left = warpPoint(deflate, bounds.x, bounds.y)
    const topCenter = warpPoint(deflate, bounds.x + bounds.cx / 2, bounds.y)
    const waist = warpPoint(deflate, bounds.x + bounds.cx / 2, bounds.y + bounds.cy / 2)
    const bottomCenter = warpPoint(deflate, bounds.x + bounds.cx / 2, bounds.y + bounds.cy)
    const right = warpPoint(deflate, bounds.x + bounds.cx, bounds.y)
    expect(left).toMatchObject({ x: bounds.x, y: bounds.y })
    expect(right).toMatchObject({ x: bounds.x + bounds.cx, y: bounds.y })
    expect(left.theta).toBeGreaterThan(0)
    expect(right.theta).toBeLessThan(0)
    expect(topCenter.y).toBe(bounds.y + Number(warpParabolaDip(500n, 1000n, 300n)))
    expect(topCenter.y).toBeGreaterThan(bounds.y)
    expect(bottomCenter.y).toBeLessThan(bounds.y + bounds.cy)
    expect(waist.y).toBe(bounds.y + bounds.cy / 2)
    expect(waist.theta).toBe(0)
    expect(topCenter.theta).toBe(0)
    expect(warpPoint(deflate, bounds.x + 200, bounds.y).theta).toBeGreaterThan(0)
    expect(warpPoint(deflate, bounds.x + 800, bounds.y).theta).toBeLessThan(0)
  })

  it('arches whole lines up or down without inventing unmodeled presets', () => {
    const up = warpPoint(spec('textArchUp', 50_000), bounds.x + bounds.cx / 2, bounds.y + 400)
    const down = warpPoint(spec('textArchDown', 50_000), bounds.x + bounds.cx / 2, bounds.y + 400)
    expect(up.y).toBeLessThan(bounds.y + 400)
    expect(down.y).toBeGreaterThan(bounds.y + 400)
    expect(warpPoint(spec('textArchUp'), bounds.x, bounds.y + 400).y).toBe(bounds.y + 400)
    const identity = warpGlyphTransform(10, 20, 10, 20, 0)
    expect(identity).toBeUndefined()
    const shifted = warpGlyphTransform(10, 20, 10, 40, 0)
    expect(shifted).toMatchObject({ aPpm: 1_000_000, bPpm: 0, cPpm: 0, dPpm: 1_000_000, txEmu: 0, tyEmu: 20 })
  })

  it('wraps each glyph of a modeled run and leaves refused runs flat', () => {
    const run: RenderTextRunNode = {
      kind: 'textRun', sourceElementId: 'el', paragraphIndex: 0, runIndex: 0, startUtf16: 0, endUtf16: 2,
      text: 'AB', direction: 'ltr', fontSizeMilliPoints: 20_000, color: '000000', bold: false, italic: false,
      x: 100, baselineY: 200, advanceInlineEmu: 200, lineHeightEmu: 100, status: 'shaped',
      glyphs: [
        { glyphId: 1, clusterIndex: 0, advanceXMilliPoints: 1_000, advanceYMilliPoints: 0, offsetXMilliPoints: 0, offsetYMilliPoints: 0, xEmu: 0, yEmu: 0, advanceXEmu: 100, advanceYEmu: 0, offsetXEmu: 0, offsetYEmu: 0 },
        { glyphId: 2, clusterIndex: 1, advanceXMilliPoints: 1_000, advanceYMilliPoints: 0, offsetXMilliPoints: 0, offsetYMilliPoints: 0, xEmu: 100, yEmu: 0, advanceXEmu: 100, advanceYEmu: 0, offsetXEmu: 0, offsetYEmu: 0 },
      ],
      clusters: [], decisions: [], attemptedFaceIds: [],
    }
    const warped = warpGlyphRun(run, spec('textDeflate', 37_500))
    expect(warped).toHaveLength(2)
    expect(warped[0]!.transform).toBeDefined()
    expect(warped[1]!.transform).toBeDefined()
    expect(warped[1]!.transform!.tyEmu).not.toBe(0)
    expect(warped[0]!.run.glyphs).toHaveLength(1)
    expect(warped[1]!.run.glyphs[0]!.glyphId).toBe(2)
    expect(warpGlyphRun({ ...run, status: 'refused' }, spec('textDeflate', 37_500))).toEqual([{ run: { ...run, status: 'refused' } }])
    expect(warpGlyphRun(run, undefined)).toEqual([{ run }])
  })
})
