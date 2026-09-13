import type { PDFDocument, PDFFont } from 'pdf-lib'
import { readFontFace } from '../textEdit/fontCmap.js'

/** Caller-supplied, fixed TrueType face. This profile does not discover fonts. */
export interface EmbeddedTextAppearanceFont {
  fontBytes: Uint8Array
}

const features = { kern: false, liga: false, clig: false, calt: false }
const maxBytes = 16 * 1024 * 1024

export async function prepareEmbeddedTextFont(doc: PDFDocument, options: EmbeddedTextAppearanceFont): Promise<{
  embed(): Promise<PDFFont>
  qualify(value: string): void
}> {
  if (!(options.fontBytes instanceof Uint8Array) || options.fontBytes.length < 12 || options.fontBytes.length > maxBytes) {
    throw new RangeError('embedded appearance font must contain 12..16777216 bytes')
  }
  const bytes = options.fontBytes.slice()
  const face = readFontFace(bytes)
  if (face.offset !== 0 || !face.tables.has('glyf') || ['fvar', 'COLR', 'SVG ', 'sbix', 'CBDT', 'CBLC'].some(tag => face.tables.has(tag))) {
    throw new Error('embedded appearances require a standalone fixed TrueType outline font')
  }
  // The package's Node entry exposes named exports, while its browser ESM
  // entry exposes a default object despite the shared named-export typings.
  const imported = await import('@pdf-lib/fontkit')
  const fontkit = imported.default ?? imported
  const parsed = fontkit.create(bytes)
  if (!Number.isFinite(parsed.unitsPerEm) || parsed.unitsPerEm <= 0 || !Number.isInteger(parsed.numGlyphs) || parsed.numGlyphs < 1 || parsed.numGlyphs > 65535
    || ![parsed.ascent, parsed.descent, parsed.italicAngle, parsed.bbox.minX, parsed.bbox.minY, parsed.bbox.maxX, parsed.bbox.maxY].every(Number.isFinite)
    || parsed.ascent <= parsed.descent || parsed.bbox.minX > parsed.bbox.maxX || parsed.bbox.minY >= parsed.bbox.maxY
    || (parsed.capHeight != null && !Number.isFinite(parsed.capHeight)) || (parsed.xHeight != null && !Number.isFinite(parsed.xHeight))) {
    throw new Error('invalid embedded appearance font metrics')
  }
  doc.registerFontkit(fontkit)
  // PDF subset encoding uses one CID per glyph. Aliased scalars cannot share
  // that CID without losing the original text in its ToUnicode mapping.
  const scalarByGlyph = new Map<number, number>()
  return {
    embed: () => doc.embedFont(bytes, { subset: true, features }),
    qualify(value) {
      if (value.length > 4096) throw new Error('embedded appearance text exceeds 4096 UTF-16 units')
      // Only independent horizontal glyphs are qualified. Marks, bidi controls,
      // contextual shaping and surrogate errors must not silently lose positioning.
      if (![...value].every(character => /^[\p{Script=Latin}\p{Script=Greek}\p{Script=Cyrillic}\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]$/u.test(character)
        || (/^\p{Script=Common}$/u.test(character) && /^[0-9\p{Punctuation}\p{Symbol} ]$/u.test(character)))
        || /[\p{Mark}\p{Control}\p{Format}\p{Surrogate}]/u.test(value)) {
        throw new Error('embedded appearances support independent horizontal Unicode glyphs only')
      }
      const scalars = [...value].map(character => character.codePointAt(0)!)
      if (scalars.some(scalar => !parsed.hasGlyphForCodePoint(scalar))) throw new Error('embedded appearance font is missing a requested glyph')
      const run = parsed.layout(value, features)
      const pending = new Map(scalarByGlyph)
      for (let index = 0; index < run.glyphs.length; index++) {
        const glyph = run.glyphs[index]!
        const scalar = scalars[index]!
        if ((pending.has(glyph.id) && pending.get(glyph.id) !== scalar)
          || glyph.codePoints.length !== 1 || glyph.codePoints[0] !== scalar) {
          throw new Error('embedded appearance glyph aliases cannot preserve Unicode text')
        }
        pending.set(glyph.id, scalar)
      }
      if (run.glyphs.length !== scalars.length || run.positions.length !== scalars.length || run.glyphs.some((glyph, index) => {
        const position = run.positions[index]!
        return glyph.id === 0 || glyph.id !== parsed.glyphForCodePoint(scalars[index]!).id
          || !Number.isFinite(glyph.advanceWidth) || glyph.advanceWidth < 0
          || position.xAdvance !== glyph.advanceWidth || position.yAdvance !== 0 || position.xOffset !== 0 || position.yOffset !== 0
      })) throw new Error('embedded appearance text requires unsupported glyph shaping or positioning')
      for (const [glyph, scalar] of pending) scalarByGlyph.set(glyph, scalar)
    },
  }
}
