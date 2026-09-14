import { readFileSync } from 'node:fs'
import { PDFDict, PDFDocument, PDFName, PDFRawStream, decodePDFRawStream } from 'pdf-lib'
import fontkit from '@pdf-lib/fontkit'
import { expect, it } from 'vitest'
import { applyFormValues } from './forms.js'
import { prepareEmbeddedTextFont } from './embeddedTextFont.js'
import { UnicodeFontResource } from './unicodeFontResource.js'
import { prepareUnicodeCffFont } from './unicodeCffFont.js'
import { normalizeOwnedCffFontkit } from './unicodeCffFontkit.js'
import { readFontFace } from '../textEdit/fontCmap.js'
const bytes = new Uint8Array(readFileSync(new URL('../../testdata/fonts/NotoSansKR-CID-subset.otf', import.meta.url)))
const decode = (stream: PDFRawStream) => Buffer.from(decodePDFRawStream(stream).decode()).toString()
it.each([false, true])('paints both CFF aliases through injective views, including later-field refresh (reverse=%s)', async reverse => {
  const values = reverse ? ['한글', '한글', '한글 한글'] : ['한글', '한글', '한글 한글']
  const doc = await PDFDocument.create(), page = doc.addPage([500, 400])
  values.forEach((_, i) => { const field = doc.getForm().createTextField(`text${i}`); field.setText('BEFORE'); field.addToPage(page, { x: 20, y: 300 - i * 100, width: 450, height: 70 }); field.setFontSize(24) })
  const source = await doc.save(), copy = source.slice(), fontCopy = bytes.slice()
  const result = await applyFormValues(source, values.map((value, i) => ({ name: `text${i}`, kind: 'text' as const, value })), { textAppearance: { fontBytes: bytes } })
  expect(result.applied).toBe(3); expect(result.skipped).toEqual([]); expect(source).toEqual(copy); expect(bytes).toEqual(fontCopy)
  const saved = await PDFDocument.load(result.bytes), descendants = new Set<string>(), encodings = new Set<string>(), unicode = new Set<string>()
  for (let i = 0; i < values.length; i++) {
    const field = saved.getForm().getTextField(`text${i}`); expect(field.getText()).toBe(values[i])
    const ap = saved.context.lookup(field.acroField.getWidgets()[0]!.getNormalAppearance()) as PDFRawStream
    const fonts = ap.dict.lookup(PDFName.of('Resources'), PDFDict).lookup(PDFName.of('Font'), PDFDict)
    for (const name of fonts.keys()) {
      const font = fonts.lookup(name, PDFDict), encoding = font.get(PDFName.of('Encoding'))!, cmap = font.get(PDFName.of('ToUnicode'))!
      descendants.add(font.get(PDFName.of('DescendantFonts'))!.toString()); encodings.add(encoding.toString()); unicode.add(cmap.toString())
      const destinations = [...decode(saved.context.lookup(encoding) as PDFRawStream).matchAll(/<[0-9A-F]{4}> (\d+)$/gm)].map(m => m[1])
      expect(new Set(destinations).size).toBe(destinations.length)
    }
  }
  expect(descendants.size).toBe(1); expect(encodings.size).toBe(2); expect(unicode.size).toBe(2)
  expect(saved.context.enumerateIndirectObjects().filter(([, object]) => object instanceof PDFRawStream && object.dict.get(PDFName.of('Subtype')) === PDFName.of('CIDFontType0C'))).toHaveLength(1)
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs'), task = pdfjs.getDocument({ data: result.bytes.slice() })
  try {
    const page = await (await task.promise).getPage(1), operators = await page.getOperatorList()
    const glyphs = operators.fnArray.flatMap((op, i) => op === pdfjs.OPS.showText ? operators.argsArray[i]![0] : []).filter(g => typeof g === 'object')
    expect(glyphs).toHaveLength(9)
    // This is the renderer's actual resolved font mapping, not just ToUnicode
    // extraction: before views, the decomposed aliases were isInFont=false.
    expect(glyphs.every(g => g.isInFont)).toBe(true)
    expect(glyphs.map(g => g.unicode).join('')).toBe(values.join(''))
  } finally { await task.destroy() }
})
it('bounds CFF views before committing resources and retains the last accepted view', async () => {
  const font = fontkit.create(bytes), table = readFontFace(bytes).tables.get('CFF ')!, cff = prepareUnicodeCffFont(bytes.subarray(table.offset, table.offset + table.length), font.numGlyphs, font.unitsPerEm)
  normalizeOwnedCffFontkit(font, cff)
  const doc = await PDFDocument.create(), resource = new UnicodeFontResource(font, bytes, doc, cff), gid = font.glyphForCodePoint(0xd55c).id
  // Synthetic qualified source aliases exercise the resource budget without
  // relying on a font containing 257 naturally equivalent Unicode sequences.
  const run = (i: number) => ({ value: String.fromCodePoint(0xe000 + i), unitsPerEm: font.unitsPerEm, width: 1000, glyphs: [{ id: gid, cluster: 0, end: 1, x: 0, y: 0, advance: 1000 }] })
  // Exercise every real commit: each refresh rewrites all existing view CMaps.
  // Keep the full boundary coverage, with a test-local allowance for shared CI CPUs.
  for (let i = 0; i < 256; i++) { const encoded = resource.qualify(run(i)); expect(encoded.fontViews).toEqual([i]); if (i === 0) await resource.embed(); else resource.commit() }
  const count = doc.context.enumerateIndirectObjects().length
  expect(() => resource.qualify(run(256))).toThrow('256 semantic font views')
  expect(doc.context.enumerateIndirectObjects()).toHaveLength(count)
  expect(resource.qualify(run(255)).fontViews).toEqual([255])
}, 30_000)

it('attaches views only to a fresh matching primary font and refuses resource-name collisions', async () => {
  const doc = await PDFDocument.create(), field = doc.getForm().createTextField('text'), value = '한글 한글'
  field.setText('BEFORE'); field.addToPage(doc.addPage([500, 200]), { x: 20, y: 40, width: 450, height: 80 })
  const prepared = await prepareEmbeddedTextFont(doc, { fontBytes: bytes }); prepared.qualify(value)
  const font = await prepared.embed(); field.setText(value); field.updateAppearances(font, prepared.appearanceProvider)
  const ap = doc.context.lookup(field.acroField.getWidgets()[0]!.getNormalAppearance()) as PDFRawStream
  const fonts = ap.dict.lookup(PDFName.of('Resources'), PDFDict).lookup(PDFName.of('Font'), PDFDict), aliasName = PDFName.of(`${font.name}-InjofficeView1`)
  fonts.set(aliasName, font.ref)
  expect(() => prepared.attachAppearanceResources(field, font)).toThrow('resource collision')
  fonts.delete(aliasName); fonts.set(PDFName.of(font.name), doc.context.nextRef())
  expect(() => prepared.attachAppearanceResources(field, font)).toThrow('primary font')
  fonts.set(PDFName.of(font.name), font.ref); prepared.attachAppearanceResources(field, font)
  expect(fonts.keys()).toHaveLength(2)
})
