import type { AppearanceProviderFor, PDFDocument, PDFFont, PDFTextField } from 'pdf-lib'
import { readFontFace } from '../textEdit/fontCmap.js'
import { prepareUnicodeShaper, type UnicodeShapingOptions } from './unicodeShaping.js'
import { UnicodeFontResource, type EncodedUnicodeRun } from './unicodeFontResource.js'
import { unicodeTextAppearance } from './unicodeTextAppearance.js'
import { standaloneUnicodeFontFace } from './unicodeFontFace.js'
import { prepareUnicodeCffFont } from './unicodeCffFont.js'
import { normalizeOwnedCffFontkit } from './unicodeCffFontkit.js'

/** Caller-supplied fixed TrueType or CFF1 face, optionally selected from a collection. */
export interface EmbeddedTextAppearanceFont extends UnicodeShapingOptions {
  fontBytes: Uint8Array
  /** Zero-based collection face; standalone fonts require zero. Defaults to zero. */
  faceIndex?: number
}

const maxBytes = 16 * 1024 * 1024

export async function prepareEmbeddedTextFont(doc: PDFDocument, options: EmbeddedTextAppearanceFont): Promise<{
  embed(): Promise<PDFFont>
  qualify(value: string): void
  attachAppearanceResources(field: PDFTextField, font: PDFFont): void
  appearanceProvider: AppearanceProviderFor<PDFTextField>
}> {
  if (!(options.fontBytes instanceof Uint8Array) || options.fontBytes.length < 12 || options.fontBytes.length > maxBytes) {
    throw new RangeError('embedded appearance font must contain 12..16777216 bytes')
  }
  const bytes = standaloneUnicodeFontFace(options.fontBytes, options.faceIndex)
  const face = readFontFace(bytes)
  const cffTable = face.tables.get('CFF ')
  if (face.offset !== 0 || face.tables.has('glyf') === Boolean(cffTable) || ['CFF2', 'fvar', 'COLR', 'SVG ', 'sbix', 'CBDT', 'CBLC'].some(tag => face.tables.has(tag))) {
    throw new Error('embedded appearances require fixed TrueType or CFF1 outlines')
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
  const cff = cffTable ? prepareUnicodeCffFont(bytes.subarray(cffTable.offset, cffTable.offset + cffTable.length), parsed.numGlyphs, parsed.unitsPerEm) : undefined
  if (cff) normalizeOwnedCffFontkit(parsed, cff)
  const shape = await prepareUnicodeShaper(bytes, parsed.unitsPerEm, parsed.numGlyphs, options)
  const resource = new UnicodeFontResource(parsed, bytes, doc, cff)
  let qualified: EncodedUnicodeRun | undefined
  return {
    embed: () => resource.embed(),
    attachAppearanceResources: (field, font) => resource.attachAppearanceResources(field, font),
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
