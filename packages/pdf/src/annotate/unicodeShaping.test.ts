import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import fontkit from '@pdf-lib/fontkit'
import { describe, expect, it, vi } from 'vitest'
import { prepareUnicodeShaper } from './unicodeShaping.js'
const require = createRequire(import.meta.url)
const bytes = new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
const parsed = fontkit.create(bytes)

describe('source-aware HarfBuzz horizontal shaping', () => {
  it('retains original UTF-16 clusters through composition, ligatures, and supplementary text', async () => {
    const shape = await prepareUnicodeShaper(bytes, parsed.unitsPerEm, parsed.numGlyphs)
    const run = shape('e\u0301 ffi 😀')
    expect(run.value).toBe('e\u0301 ffi 😀')
    expect(run.glyphs.map(g => [g.cluster, g.end])).toEqual([[0, 2], [2, 3], [3, 6], [6, 7], [7, 9]])
    expect(run.glyphs[0]!.id).toBe(parsed.glyphForCodePoint(0xe9).id)
    expect(run.glyphs[2]!.id).toBe(parsed.glyphForCodePoint(0xfb03).id)
  })
  it('retains zero advances and positioned marks rather than approximating nominal widths', async () => {
    const shape = await prepareUnicodeShaper(bytes, parsed.unitsPerEm, parsed.numGlyphs)
    const run = shape('x\u0301\u0323')
    expect(run.glyphs).toHaveLength(3)
    expect(run.glyphs.filter(g => g.advance === 0)).toHaveLength(2)
    expect(run.glyphs.some(g => g.y !== 0 || g.x !== run.width)).toBe(true)
    expect(run.glyphs.every(g => g.cluster === 0 && g.end === 3)).toBe(true)
  })
  it('itemizes mixed LTR scripts and pins language without losing global UTF-16 offsets', async () => {
    const hb = await import('harfbuzzjs'), shape = await prepareUnicodeShaper(bytes, parsed.unitsPerEm, parsed.numGlyphs)
    const script = vi.spyOn(hb.Buffer.prototype, 'setScript'), language = vi.spyOn(hb.Buffer.prototype, 'setLanguage')
    try {
      const run = shape(' e\u0301 Ω\u0301 Ж 😀')
      expect(script.mock.calls.map(call => call[0])).toEqual(['Latn', 'Grek', 'Cyrl'])
      expect(language.mock.calls.map(call => call[0])).toEqual(['und', 'und', 'und'])
      expect(run.glyphs[0]!.cluster).toBe(0)
      expect(run.glyphs.at(-1)!.cluster).toBe(9)
      expect(run.glyphs.at(-1)!.end).toBe(11)
      expect(run.glyphs.every(g => run.value.slice(g.cluster, g.end).length > 0)).toBe(true)
      expect(() => shape('x\u05B0')).toThrow('supported LTR')
      expect(() => shape('a\u0483')).toThrow('cross-script mark')
    } finally { script.mockRestore(); language.mockRestore() }
  })

})
