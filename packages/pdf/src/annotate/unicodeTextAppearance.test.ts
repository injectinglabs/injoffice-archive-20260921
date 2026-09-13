import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { decodePDFRawStream, degrees, PDFDict, PDFDocument, PDFName, PDFRawStream, TextAlignment } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import fontkit from '@pdf-lib/fontkit'
import { prepareUnicodeShaper } from './unicodeShaping.js'
import { applyFormValues } from './forms.js'
const require = createRequire(import.meta.url)
const fontBytes = new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
async function fixture(alignment = TextAlignment.Left, rotation = 0, size = 24) {
  const doc = await PDFDocument.create(), page = doc.addPage([400, 400]), field = doc.getForm().createTextField('text')
  field.addToPage(page, { x: 100, y: 100, width: 220, height: 40, rotate: degrees(rotation) })
  field.setAlignment(alignment); field.setFontSize(size)
  return doc.save({ updateFieldAppearances: false })
}
function appearance(doc: PDFDocument) {
  const widget = doc.getForm().getTextField('text').acroField.getWidgets()[0]!
  return doc.context.lookup(widget.getNormalAppearance()) as PDFRawStream
}
function decoded(stream: PDFRawStream) { return Buffer.from(decodePDFRawStream(stream).decode()).toString('latin1') }
async function extracted(doc: PDFDocument) {
  // Test the actual saved widget appearance placed into page content. Extraction
  // of an unflattened form's canonical value is tested separately above.
  doc.getForm().flatten({ updateFieldAppearances: false })
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const task = pdfjs.getDocument({ data: await doc.save() })
  try {
    const page = await (await task.promise).getPage(1)
    const content = await page.getTextContent({ disableNormalization: true })
    return content.items.map(item => 'str' in item ? item.str : '').join('')
  } finally { await task.destroy() }
}

describe('saved cluster appearances', () => {
  it.each([0, 90, 180, 270].flatMap(rotation => [TextAlignment.Left, TextAlignment.Center, TextAlignment.Right].map(alignment => ({ rotation, alignment }))))('fits plain italic negative bearings in automatic size $alignment/$rotation', async ({ rotation, alignment }) => {
    const bytes = new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans-Oblique.ttf'))), parsed = fontkit.create(bytes)
    const box = parsed.glyphForCodePoint(106).bbox
    expect(box.minX).toBeLessThan(0)
    const result = await applyFormValues(await fixture(alignment, rotation, 0), [{ name: 'text', kind: 'text', value: 'j' }], { textAppearance: { fontBytes: bytes } })
    expect(result.applied).toBe(1)
    const doc = await PDFDocument.load(result.bytes), content = decoded(appearance(doc))
    const size = Number(content.match(/\/[^\s]+ ([\d.]+) Tf/)![1]), scale = size / parsed.unitsPerEm
    const matrix = [...content.matchAll(/1 0 0 1 ([\d.-]+) ([\d.-]+) Tm/g)][1]!
    const x = Number(matrix[1]), y = Number(matrix[2])
    // Read the actual clipping polygon in the same local coordinate system as
    // the text matrices; the provider rotates both together outside this scope.
    const clip = content.match(/([\d.-]+) ([\d.-]+) m\n[\d.-]+ [\d.-]+ l\n([\d.-]+) ([\d.-]+) l\n[\d.-]+ [\d.-]+ l\nh\nW/)!
    const [left, bottom, right, top] = clip.slice(1).map(Number)
    expect(x + box.minX * scale).toBeGreaterThanOrEqual(left! - 1e-7)
    expect(x + box.maxX * scale).toBeLessThanOrEqual(right! + 1e-7)
    expect(y + box.minY * scale).toBeGreaterThanOrEqual(bottom! - 1e-7)
    expect(y + box.maxY * scale).toBeLessThanOrEqual(top! + 1e-7)
    expect(doc.getForm().getTextField('text').getText()).toBe('j')
  })
  it('retains authored fixed-size italic positioning and clipping', async () => {
    const bytes = new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans-Oblique.ttf')))
    const result = await applyFormValues(await fixture(TextAlignment.Left, 0, 55), [{ name: 'text', kind: 'text', value: 'j' }], { textAppearance: { fontBytes: bytes } })
    const content = decoded(appearance(await PDFDocument.load(result.bytes)))
    expect(content).toContain('/DejaVuSans-Oblique 55 Tf')
    expect([...content.matchAll(/ Tm/g)]).toHaveLength(1)
    expect(content).toMatch(/1 0 0 1 2 [\d.-]+ Tm/)
  })
  it.each(['e\u0301', 'a\u0301\u0323', 'x\u0301\u0323', 'ffi', '\uFB03', 'e\u0301 é ffi 😀', '  ffi  '])('preserves exact source and reader text for %j', async value => {
    const source = await fixture(), copy = source.slice(), fontCopy = fontBytes.slice()
    const result = await applyFormValues(source, [{ name: 'text', kind: 'text', value }], { textAppearance: { fontBytes } })
    expect(result.applied).toBe(1); expect(result.skipped).toEqual([])
    const doc = await PDFDocument.load(result.bytes)
    expect(doc.getForm().getTextField('text').getText()).toBe(value)
    expect(source).toEqual(copy); expect(fontBytes).toEqual(fontCopy)
    // Reader whitespace heuristics may drop field-edge spaces; all interior
    // Unicode including the exact decomposed order must be retained.
    expect(await extracted(doc)).toBe(value.trim())
  })
  it('uses native outlines only for unpartitionable continuation glyphs, with exact replacement text', async () => {
    const value = 'a\u0301\u0323'
    const result = await applyFormValues(await fixture(), [{ name: 'text', kind: 'text', value }], { textAppearance: { fontBytes } })
    const doc = await PDFDocument.load(result.bytes), ap = appearance(doc), content = decoded(ap)
    expect(content).toContain('/ActualText <FEFF006103010323>')
    expect([...content.matchAll(/ Tj/g)]).toHaveLength(1)
    expect(content).toContain(' l\n'); expect(content).toContain('f\n')
    const fonts = ap.dict.lookup(PDFName.of('Resources'), PDFDict).lookup(PDFName.of('Font'), PDFDict)
    const font = fonts.lookup(fonts.keys()[0]!, PDFDict)
    const cmap = decoded(font.lookup(PDFName.of('ToUnicode')) as PDFRawStream)
    expect(cmap).toContain('<0001> <006103010323>')
    expect(cmap).not.toContain('<0002>')
    expect(await extracted(doc)).toBe(value)
  })
  it.each([0, 90, 180, 270])('keeps positioned marks in owned %s-degree widgets with autosizing', async rotation => {
    const value = 'x\u0301\u0323 ffi'
    const result = await applyFormValues(await fixture(TextAlignment.Center, rotation, 0), [{ name: 'text', kind: 'text', value }], { textAppearance: { fontBytes } })
    expect(result.applied).toBe(1)
    const doc = await PDFDocument.load(result.bytes), content = decoded(appearance(doc))
    const size = Number(content.match(/\/(?:[^\s]+) ([\d.]+) Tf/)![1])
    expect(size).toBeGreaterThan(0)
    expect(content).not.toMatch(/NaN|Infinity/)
    expect(doc.getForm().getTextField('text').getText()).toBe(value)
  })
  it.each([TextAlignment.Left, TextAlignment.Center, TextAlignment.Right])('fits tall stacked ink and negative bearings for alignment %s', async alignment => {
    const value = 'j' + '\u0302'.repeat(12) + '\u0323'
    const parsed = fontkit.create(fontBytes), shape = await prepareUnicodeShaper(fontBytes, parsed.unitsPerEm, parsed.numGlyphs), run = shape(value)
    const ink = run.glyphs.map(g => { const b = parsed.getGlyph(g.id).bbox; return { minX: g.x + b.minX, maxX: g.x + b.maxX, minY: g.y + b.minY, maxY: g.y + b.maxY } })
    expect(Math.max(...ink.map(b => b.maxY))).toBeGreaterThan(parsed.ascent)
    expect(Math.min(...ink.map(b => b.minX))).toBeLessThan(0)
    const result = await applyFormValues(await fixture(alignment, 0, 0), [{ name: 'text', kind: 'text', value }], { textAppearance: { fontBytes } })
    expect(result.applied).toBe(1)
    const doc = await PDFDocument.load(result.bytes), content = decoded(appearance(doc)), widget = doc.getForm().getTextField('text').acroField.getWidgets()[0]!
    const size = Number(content.match(/\/(?:[^\s]+) ([\d.]+) Tf/)![1]), scale = size / parsed.unitsPerEm
    const matrices = [...content.matchAll(/1 0 0 1 ([\d.-]+) ([\d.-]+) Tm/g)]
    const first = matrices[1]!, originX = Number(first[1]) - run.glyphs[0]!.x * scale, originY = Number(first[2]) - run.glyphs[0]!.y * scale
    const inset = (widget.getBorderStyle()?.getWidth() ?? 0) + 1, rectangle = widget.getRectangle()
    for (const box of ink) {
      expect(originX + box.minX * scale).toBeGreaterThanOrEqual(inset - 1e-7)
      expect(originX + box.maxX * scale).toBeLessThanOrEqual(rectangle.width - inset + 1e-7)
      expect(originY + box.minY * scale).toBeGreaterThanOrEqual(inset - 1e-7)
      expect(originY + box.maxY * scale).toBeLessThanOrEqual(rectangle.height - inset + 1e-7)
    }
  })

})
