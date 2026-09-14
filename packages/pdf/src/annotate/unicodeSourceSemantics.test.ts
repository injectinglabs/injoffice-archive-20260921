import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { createHash } from 'node:crypto'
import { decodePDFRawStream, PDFDict, PDFDocument, PDFHexString, PDFName, PDFRawStream } from 'pdf-lib'
import fontkit from '@pdf-lib/fontkit'
import { expect, it } from 'vitest'
import { applyFormValues } from './forms.js'
import { prepareUnicodeShaper } from './unicodeShaping.js'
const require = createRequire(import.meta.url)
const dejavu = new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
const fontFixture = (name: string) => new Uint8Array(readFileSync(new URL(`../../testdata/fonts/${name}`, import.meta.url)))
const decoded = (stream: PDFRawStream) => Buffer.from(decodePDFRawStream(stream).decode()).toString('latin1')
async function saved(value: string, fontBytes: Uint8Array) {
  const doc = await PDFDocument.create(), field = doc.getForm().createTextField('text')
  field.setText('BEFORE'); field.addToPage(doc.addPage([600, 180]), { x: 20, y: 50, width: 560, height: 80 }); field.setFontSize(24)
  const source = await doc.save(), sourceCopy = source.slice(), fontCopy = fontBytes.slice()
  const result = await applyFormValues(source, [{ name: 'text', kind: 'text', value }], { textAppearance: { fontBytes } })
  expect(result.applied).toBe(1); expect(result.skipped).toEqual([])
  expect(source).toEqual(sourceCopy); expect(fontBytes).toEqual(fontCopy)
  const output = await PDFDocument.load(result.bytes), text = output.getForm().getTextField('text')
  expect(text.getText() ?? '').toBe(value)
  const ap = output.context.lookup(text.acroField.getWidgets()[0]!.getNormalAppearance()) as PDFRawStream
  return { doc: output, ap, content: decoded(ap) }
}
async function extracted(doc: PDFDocument) {
  doc.getForm().flatten({ updateFieldAppearances: false })
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const task = pdfjs.getDocument({ data: await doc.save() })
  try { const page = await (await task.promise).getPage(1); return (await page.getTextContent({ disableNormalization: true })).items.map(item => 'str' in item ? item.str : '').join('') }
  finally { await task.destroy() }
}
it.each(['\u200B', '\u00AD', '\uFE0F', '\u034F', '\u{E0100}', '\u200D', '\u200C', '\u200B\u00AD\u{E0100}'])('preserves exact zero-glyph replacement source %j without invented text glyphs', async value => {
  const parsed = fontkit.create(dejavu), shape = await prepareUnicodeShaper(dejavu, parsed.unitsPerEm, parsed.numGlyphs)
  expect(shape(value).glyphs).toEqual([]); expect(shape(value).requiresActualText).toBe(true)
  const { doc, content } = await saved(value, dejavu)
  expect(content).toContain(`/ActualText ${PDFHexString.fromText(value).toString()}`)
  expect(content).toContain('<> Tj'); expect(content).not.toMatch(/<[0-9A-F]+> Tj/)
  // PDF.js ignores an empty text-show operator, even with ActualText. The form
  // value and standards-valid replacement text above remain exact authority.
  expect(await extracted(doc)).toBe('')
})
it.each(['', ' ', '\u200Bhello\u00AD', 'a\uFE0F', 'a\u{E0100}'])('keeps empty/space/mixed ignorable source %j', async value => {
  const { content } = await saved(value, dejavu)
  expect(content).not.toContain('FFFC')
})
it.each([['侮', '\uFE00'], ['倦', '\u{E0100}']])('selects real nondefault cmap14 glyph for %s + %j and saves exact semantics', async (base, selector) => {
  const bytes = fontFixture('NotoSansJP-CID-subset.otf'), parsed = fontkit.create(bytes), shape = await prepareUnicodeShaper(bytes, parsed.unitsPerEm, parsed.numGlyphs)
  const plain = shape(base), selected = shape(base + selector)
  expect(plain.glyphs).toHaveLength(1); expect(selected.glyphs).toHaveLength(1)
  expect(selected.glyphs[0]!.id).not.toBe(plain.glyphs[0]!.id)
  const { doc, ap } = await saved(base + selector, bytes)
  const fonts = ap.dict.lookup(PDFName.of('Resources'), PDFDict).lookup(PDFName.of('Font'), PDFDict), font = fonts.lookup(fonts.keys()[0]!, PDFDict)
  expect(decoded(font.lookup(PDFName.of('ToUnicode')) as PDFRawStream)).toContain(PDFHexString.fromText(base + selector).toString().replace('FEFF', ''))
  expect(await extracted(doc)).toBe(base + selector)
})
it('retains composed and decomposed Hangul sources with identical correctly shaped glyphs', async () => {
  const bytes = fontFixture('NotoSansKR-CID-subset.otf'), parsed = fontkit.create(bytes), shape = await prepareUnicodeShaper(bytes, parsed.unitsPerEm, parsed.numGlyphs)
  const manifest = JSON.parse(readFileSync(new URL('../../testdata/fonts/NotoSansKR-CID-subset.otf.json', import.meta.url), 'utf8'))
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(manifest.sha256)
  expect(createHash('sha256').update(fontFixture(manifest.license)).digest('hex')).toBe(manifest.licenseSha256)
  const composed = '한글', decomposed = '한글'
  expect(shape(composed).glyphs.map(g => g.id)).toEqual(shape(decomposed).glyphs.map(g => g.id))
  expect(shape(composed).glyphs).toHaveLength(2)
  expect(shape(composed).glyphs.map(g => g.id)).toEqual([...composed].map(c => parsed.glyphForCodePoint(c.codePointAt(0)!).id))
  for (const value of [composed, decomposed]) { const { doc } = await saved(value, bytes); expect(await extracted(doc)).toBe(value) }
})
it('saves neighboring base and nondefault variants in one mixed CFF run', async () => {
  await saved('Aé Ω 日本語 侮侮\uFE00 倦倦\u{E0100}', fontFixture('NotoSansJP-CID-subset.otf'))
})
