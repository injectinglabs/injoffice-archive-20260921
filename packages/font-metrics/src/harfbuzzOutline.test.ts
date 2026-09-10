import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { describe, it, expect } from 'vitest'
import { createHarfBuzzOutlineProviderV1 } from './harfbuzz.js'

const require = createRequire(import.meta.url)
const bytes = new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
const contentDigest = 'sha256:7da195a74c55bef988d0d48f9508bd5d849425c1770dba5d7bfc6ce9ed848954'
describe('native HarfBuzz outlines', () => {
  it('produces repeatable real contours and preserves empty space glyphs', () => {
    const provider = createHarfBuzzOutlineProviderV1({ bytes, contentDigest })
    const glyph = provider.outline(36) // DejaVu Sans A
    expect(glyph.units_per_em).toBe(524288)
    expect(glyph.path.length).toBeGreaterThan(5)
    expect(glyph.path).toEqual(provider.outline(36).path)
    expect(provider.outline(3).path).toEqual([]) // space
    const curved = provider.outline(68) // a includes half-unit implied points
    expect(curved.units_per_em).toBe(glyph.units_per_em)
    for (const command of curved.path) for (const value of Object.values(command)) if (typeof value === 'number') expect(Number.isSafeInteger(value)).toBe(true)
  })
  it('refuses substituted font bytes and invalid glyph IDs', () => {
    expect(() => createHarfBuzzOutlineProviderV1({ bytes, contentDigest: `sha256:${'0'.repeat(64)}` })).toThrow(/digest/)
    const provider = createHarfBuzzOutlineProviderV1({ bytes, contentDigest })
    for (const id of [-1, 0.5, Infinity, 65536]) expect(() => provider.outline(id)).toThrow(/glyph/)
  })
})
