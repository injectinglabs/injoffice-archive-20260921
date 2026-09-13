import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { expect, it } from 'vitest'
import fontkit from '@pdf-lib/fontkit'
import { normalizeOwnedCffFontkit } from './unicodeCffFontkit.js'
import { prepareUnicodeCffFont } from './unicodeCffFont.js'
import { readFontFace } from '../textEdit/fontCmap.js'
const require = createRequire(import.meta.url)
const bytes = new Uint8Array(readFileSync(new URL('../../testdata/fonts/NotoSansJP-CID-subset.otf', import.meta.url)))
function fixture() {
  const font = fontkit.create(bytes.slice()), table = readFontFace(bytes).tables.get('CFF ')!
  const validated = prepareUnicodeCffFont(bytes.subarray(table.offset, table.offset + table.length), font.numGlyphs, font.unitsPerEm)
  const cff = (font as unknown as { 'CFF ': any })['CFF ']
  return { font, validated, cff }
}
it('pins the private parser contract and reproduces the exact FD range-boundary error', () => {
  expect(require('@pdf-lib/fontkit/package.json').version).toBe('1.1.1')
  const { font, cff, validated } = fixture(), originalMethod = Object.getPrototypeOf(cff).fdForGlyph, originalBytes = bytes.slice()
  const glyph = font.glyphForCodePoint(0x4fae)
  expect(cff.topDict.FDSelect.ranges.some((r: { first: number }) => r.first === glyph.id)).toBe(true)
  expect(cff.fdForGlyph(glyph.id)).not.toBe(validated.sourceGlyphFDs![glyph.id])
  expect(() => glyph.bbox).toThrow('Unknown op: 0')
  // Production uses a fresh instance before any glyph path is attempted.
  const owned = fixture(); normalizeOwnedCffFontkit(owned.font, owned.validated)
  expect(owned.cff.fdForGlyph(glyph.id)).toBe(owned.validated.sourceGlyphFDs![glyph.id])
  expect(owned.font.getGlyph(glyph.id).bbox).toMatchObject({ minX: 16, maxX: 973, maxY: 841 })
  expect(Object.getPrototypeOf(cff).fdForGlyph).toBe(originalMethod)
  expect(cff.topDict.FDSelect.version).toBe(3)
  expect(bytes).toEqual(originalBytes)
})
it('uses source GID→FD selection at every range boundary and matches independent HarfBuzz ink', async () => {
  const { font, cff, validated } = fixture(), original = bytes.slice()
  expect(validated.glyphCids.some((cid, gid) => cid !== gid)).toBe(true)
  normalizeOwnedCffFontkit(font, validated)
  const hb = await import('harfbuzzjs'), hbFont = new hb.Font(new hb.Face(new hb.Blob(bytes), 0)); hbFont.setScale(font.unitsPerEm, font.unitsPerEm)
  for (let gid = 0; gid < font.numGlyphs; gid++) {
    expect(cff.fdForGlyph(gid)).toBe(validated.sourceGlyphFDs![gid])
    const glyph = font.getGlyph(gid), path = glyph.path as unknown as { commands: unknown[] }, extents = hbFont.glyphExtents(gid)!
    if (!path.commands.length) { expect(extents.width).toBe(0); continue }
    const box = glyph.bbox
    // HB computes integer extents with rounded curves; fontkit uses floating
    // curve extrema. Their independent bounds agree within two design units.
    expect(Math.abs(box.minX - extents.xBearing)).toBeLessThanOrEqual(2)
    expect(Math.abs(box.maxX - extents.xBearing - extents.width)).toBeLessThanOrEqual(2)
    expect(Math.abs(box.maxY - extents.yBearing)).toBeLessThanOrEqual(2)
    expect(Math.abs(box.minY - extents.yBearing - extents.height)).toBeLessThanOrEqual(2)
  }
  expect(bytes).toEqual(original)
})
it('refuses unexpected private parser shape or disagreement with validated source selection', () => {
  for (const mutate of [
    (cff: any) => { cff.version = 2 },
    (cff: any) => { cff.topDict.FDSelect.version = 4 },
    (cff: any) => { cff.topDict.FDSelect.ranges[0].fd = 255 },
    (cff: any) => { cff.topDict.FDSelect.ranges[1].first = 0 },
    (cff: any) => { cff.topDict.FDSelect.sentinel-- },
    (cff: any) => { cff.fdForGlyph = undefined },
  ]) { const { font, cff, validated } = fixture(); mutate(cff); expect(() => normalizeOwnedCffFontkit(font, validated)).toThrow(/parser shape/) }
})
