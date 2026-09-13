import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import fontkit from '@pdf-lib/fontkit'
import { decodePDFRawStream, PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFRawStream } from 'pdf-lib'
import { describe, expect, it, vi } from 'vitest'
import { prepareUnicodeShaper } from './unicodeShaping.js'
import { applyFormValues } from './forms.js'
const require = createRequire(import.meta.url)
const bytes = new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
const parsed = fontkit.create(bytes)
const decoded = (stream: PDFRawStream) => Buffer.from(decodePDFRawStream(stream).decode()).toString('latin1')
async function source() {
  const doc = await PDFDocument.create(), field = doc.getForm().createTextField('text')
  field.setText('BEFORE'); field.addToPage(doc.addPage([500, 200]), { x: 20, y: 60, width: 450, height: 60 }); field.setFontSize(24)
  return doc.save()
}
function artwork(doc: PDFDocument) {
  const field = doc.getForm().getTextField('text'), widget = field.acroField.getWidgets()[0]!
  const ap = doc.context.lookup(widget.getNormalAppearance()) as PDFRawStream
  const fonts = ap.dict.lookup(PDFName.of('Resources'), PDFDict).lookup(PDFName.of('Font'), PDFDict)
  return { ap, font: fonts.lookup(fonts.keys()[0]!, PDFDict) }
}

describe('contextual RTL form appearances', () => {
  it('retains Arabic contextual glyph selection and lam-alef source clusters', async () => {
    const shape = await prepareUnicodeShaper(bytes, parsed.unitsPerEm, parsed.numGlyphs)
    const run = shape('السلام عليكم')
    expect(run.requiresActualText).toBe(true)
    const clusters = [...new Set(run.glyphs.map(g => g.cluster))].map(start => { const glyph = run.glyphs.find(g => g.cluster === start)!; return run.value.slice(start, glyph.end) })
    expect(clusters).toContain('لا')
    expect(run.glyphs.some(g => g.id !== parsed.glyphForCodePoint(run.value.codePointAt(g.cluster)!).id)).toBe(true)
    expect(run.glyphs[0]!.cluster).toBeGreaterThan(run.glyphs.at(-1)!.cluster)
    expect(run.glyphs[0]!.x).toBe(0)
  })
  it('passes explicit base direction and language without rewriting source', async () => {
    const hb = await import('harfbuzzjs'), spy = vi.spyOn(hb.Buffer.prototype, 'setLanguage')
    try {
      const shape = await prepareUnicodeShaper(bytes, parsed.unitsPerEm, parsed.numGlyphs, { direction: 'rtl', language: 'ar' })
      const run = shape('abc !')
      expect(run.value).toBe('abc !'); expect(run.requiresActualText).toBe(true)
      expect(spy.mock.calls.every(args => args[0] === 'ar')).toBe(true)
      expect(run.glyphs[0]!.cluster).toBe(4)
    } finally { spy.mockRestore() }
  })
  it('uses joiners as shaping controls while preserving exact source spans', async () => {
    const shape = await prepareUnicodeShaper(bytes, parsed.unitsPerEm, parsed.numGlyphs)
    const joined = shape('لا'), separated = shape('ل\u200Cا')
    expect(joined.glyphs).toHaveLength(1); expect(separated.glyphs).toHaveLength(2)
    expect([...new Set(separated.glyphs.map(g => g.cluster))].sort((a, b) => a - b).map(start => { const glyph = separated.glyphs.find(g => g.cluster === start)!; return separated.value.slice(start, glyph.end) }).join('')).toBe('ل\u200Cا')
  })
  it('mirrors a bracket glyph in RTL without changing its logical source mapping', async () => {
    const shape = await prepareUnicodeShaper(bytes, parsed.unitsPerEm, parsed.numGlyphs, { direction: 'rtl' })
    const run = shape('אבג (')
    expect(run.glyphs.find(g => g.cluster === 4)?.id).toBe(parsed.glyphForCodePoint(')'.codePointAt(0)!).id)
    expect(run.value).toBe('אבג (')
  })
  it('snapshots shaping options before asynchronous loading', async () => {
    const options = { direction: 'rtl' as const, language: 'ar' }
    const pending = prepareUnicodeShaper(bytes, parsed.unitsPerEm, parsed.numGlyphs, options)
    Object.assign(options, { direction: 'ltr', language: 'en' })
    expect((await pending)('abc !').glyphs[0]!.cluster).toBe(4)
  })
  it.each(['مرحبا', 'مَرْحَبًا', 'السلام عليكم', 'abc (مرحبا 123) xyz', 'abc \u2067مرحبا 123\u2069 xyz', 'אבג (123) xyz', '\u2067\u2069', '\u200F'])('saves exact RTL/formatting source and standard replacement semantics for %j', async value => {
    const original = await source(), copy = original.slice(), fontCopy = bytes.slice()
    const result = await applyFormValues(original, [{ name: 'text', kind: 'text', value }], { textAppearance: { fontBytes: bytes } })
    expect(result.applied).toBe(1); expect(result.skipped).toEqual([])
    const doc = await PDFDocument.load(result.bytes), { ap, font } = artwork(doc)
    expect(doc.getForm().getTextField('text').getText()).toBe(value)
    expect(decoded(ap)).toContain(`/ActualText ${PDFHexString.fromText(value).toString()}`)
    expect(decoded(ap)).not.toMatch(/NaN|Infinity/)
    const cid = font.lookup(PDFName.of('DescendantFonts'), PDFArray).lookup(0, PDFDict)
    expect(cid.lookup(PDFName.of('CIDToGIDMap'))).toBeInstanceOf(PDFRawStream)
    if (value === 'السلام عليكم') expect(decoded(font.lookup(PDFName.of('ToUnicode')) as PDFRawStream)).toContain('<06440627>')
    expect(original).toEqual(copy); expect(bytes).toEqual(fontCopy)
  })
  it('distinguishes exact canonical values/mappings from PDF.js bidi extraction heuristics', async () => {
    const value = 'السلام عليكم'
    const result = await applyFormValues(await source(), [{ name: 'text', kind: 'text', value }], { textAppearance: { fontBytes: bytes } })
    const doc = await PDFDocument.load(result.bytes)
    expect(doc.getForm().getTextField('text').getText()).toBe(value)
    doc.getForm().flatten({ updateFieldAppearances: false })
    const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs'), task = pdfjs.getDocument({ data: await doc.save() })
    try {
      const content = await (await (await task.promise).getPage(1)).getTextContent({ disableNormalization: true })
      // PDF.js ignores ActualText and reverses the internal lam-alef mapping.
      // Do not reverse ToUnicode's source span to hide this reader behavior.
      expect(content.items.map(item => 'str' in item ? item.str : '').join('')).toBe('السالم عليكم')
    } finally { await task.destroy() }
  })
  it.each([{ direction: 'sideways' }, { language: 'ar /Bad' }, { language: 'x'.repeat(64) }])('rejects invalid shaping options before returning modified bytes: %j', async options => {
    const original = await source(), copy = original.slice()
    await expect(applyFormValues(original, [{ name: 'text', kind: 'text', value: 'مرحبا' }], { textAppearance: { fontBytes: bytes, ...options } as never })).rejects.toThrow()
    expect(original).toEqual(copy)
  })
})
