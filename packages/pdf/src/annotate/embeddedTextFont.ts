import type { AppearanceProviderFor, PDFDocument, PDFFont, PDFTextField } from 'pdf-lib'
import { readFontFace } from '../textEdit/fontCmap.js'
import { prepareUnicodeShaper, type UnicodeShapingOptions } from './unicodeShaping.js'
import { UnicodeFontResource, type EncodedUnicodeRun } from './unicodeFontResource.js'
import { unicodeTextAppearance } from './unicodeTextAppearance.js'

/** Caller-supplied, fixed TrueType face. This profile does not discover fonts. */
export interface EmbeddedTextAppearanceFont extends UnicodeShapingOptions {
  fontBytes: Uint8Array
}

const maxBytes = 16 * 1024 * 1024

export async function prepareEmbeddedTextFont(doc: PDFDocument, options: EmbeddedTextAppearanceFont): Promise<{
  embed(): Promise<PDFFont>
  qualify(value: string): void
  appearanceProvider: AppearanceProviderFor<PDFTextField>
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
  const shape = await prepareUnicodeShaper(bytes, parsed.unitsPerEm, parsed.numGlyphs, options)
  const resource = new UnicodeFontResource(parsed, bytes, doc)
  let qualified: EncodedUnicodeRun | undefined
  return {
    embed: () => resource.embed(),
    appearanceProvider(field, widget, font) {
      if (!qualified) throw new Error('embedded appearance has no qualified text run')
      resource.commit()
      return unicodeTextAppearance(field, widget, font, qualified, qualified.run.glyphs.map(g => parsed.getGlyph(g.id).advanceWidth))
    },
    qualify(value) {
      qualified = resource.qualify(shape(value))
    },
  }
}
