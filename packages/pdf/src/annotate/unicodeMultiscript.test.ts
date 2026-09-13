import { readFileSync } from 'node:fs'
import { createHash } from 'node:crypto'
import { resolve } from 'node:path'
import { createRequire } from 'node:module'
import fontkit from '@pdf-lib/fontkit'
import { decodePDFRawStream, PDFDocument, PDFHexString, PDFRawStream } from 'pdf-lib'
import { expect, it } from 'vitest'
import { prepareUnicodeShaper } from './unicodeShaping.js'
import { applyFormValues } from './forms.js'
const directory = resolve(import.meta.dirname, '../../testdata/fonts')
const cases = [
  ['NotoSansDevanagari', 'क्षि नमस्ते'], ['NotoSansBengali', 'ক্ষি বাংলা'],
  ['NotoSansThai', 'น้ำ ภาษาไทย'], ['NotoSansKhmer', 'ខ្មែរ'],
  ['NotoSansMyanmar', 'မြန်မာ'], ['NotoSansSyriac', 'ܫܠܡܐ'],
] as const
it('retains exact licensed font fixture provenance', () => {
  const manifest = JSON.parse(readFileSync(resolve(directory, 'PROVENANCE.json'), 'utf8'))
  for (const [name, hash] of Object.entries(manifest.files)) expect(createHash('sha256').update(readFileSync(resolve(directory, name))).digest('hex')).toBe(hash)
})
it('matches explicit script runs around paired mixed-script punctuation', async () => {
  const bytes = new Uint8Array(readFileSync(createRequire(import.meta.url).resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf'))), parsed = fontkit.create(bytes)
  const value = 'abc (Ω\u0301) xyz', shape = await prepareUnicodeShaper(bytes, parsed.unitsPerEm, parsed.numGlyphs)
  const hb = await import('harfbuzzjs'), font = new hb.Font(new hb.Face(new hb.Blob(bytes.slice().buffer), 0))
  font.setScale(parsed.unitsPerEm, parsed.unitsPerEm)
  const expected: string[] = []
  let pen = 0
  for (const [start, end, script] of [[0, 5, 'Latn'], [5, 7, 'Grek'], [7, value.length, 'Latn']] as const) {
    const buffer = new hb.Buffer(); buffer.addText(value, start, end - start); buffer.setDirection(hb.Direction.LTR); buffer.setScript(script); buffer.setLanguage('und'); buffer.setClusterLevel(0); buffer.setFlags(hb.BufferFlag.REMOVE_DEFAULT_IGNORABLES); hb.shape(font, buffer)
    const positions = buffer.getGlyphPositions()
    buffer.getGlyphInfos().forEach((glyph, index) => {
      const position = positions[index]!
      expected.push(JSON.stringify({ id: glyph.codepoint, cluster: glyph.cluster, x: pen + position.xOffset, y: position.yOffset }))
      pen += position.xAdvance
    })
  }
  expect(shape(value).glyphs.map(({ id, cluster, x, y }) => JSON.stringify({ id, cluster, x, y })).sort()).toEqual(expected.sort())
})
it('keeps both Arabic bases joined RTL after a differently directed Prepend sign', async () => {
  const bytes = new Uint8Array(readFileSync(resolve(directory, 'NotoSansArabic-Regular.ttf'))), parsed = fontkit.create(bytes)
  const value = '\u0600بب', shape = await prepareUnicodeShaper(bytes, parsed.unitsPerEm, parsed.numGlyphs)
  const run = shape(value), hb = await import('harfbuzzjs'), font = new hb.Font(new hb.Face(new hb.Blob(bytes.slice().buffer), 0))
  font.setScale(parsed.unitsPerEm, parsed.unitsPerEm)
  const joined = new hb.Buffer(); joined.addText(value, 1, 2); joined.setDirection(hb.Direction.RTL); joined.setScript('Arab'); joined.setLanguage('und'); hb.shape(font, joined)
  expect(run.glyphs.filter(glyph => glyph.cluster >= 1).map(glyph => glyph.id)).toEqual(joined.getGlyphInfos().map(glyph => glyph.codepoint))
  expect(run.glyphs.map(glyph => glyph.cluster)).toEqual([2, 1, 0])
  const doc = await PDFDocument.create(), field = doc.getForm().createTextField('text')
  field.setText('BEFORE'); field.addToPage(doc.addPage([500, 200]), { x: 20, y: 80, width: 460, height: 60 })
  const source = await doc.save(), copy = source.slice()
  const result = await applyFormValues(source, [{ name: 'text', kind: 'text', value }], { textAppearance: { fontBytes: bytes } })
  expect(result.applied).toBe(1); expect(result.skipped).toEqual([])
  expect((await PDFDocument.load(result.bytes)).getForm().getTextField('text').getText()).toBe(value)
  expect(source).toEqual(copy)
})
it.each(cases)('shapes and saves contextual %s source %s', async (name, value) => {
  const bytes = new Uint8Array(readFileSync(resolve(directory, `${name}-Regular.ttf`))), originalFont = bytes.slice()
  const parsed = fontkit.create(bytes)
  const shape = await prepareUnicodeShaper(bytes, parsed.unitsPerEm, parsed.numGlyphs)
  const run = shape(value)
  expect(run.glyphs.length).toBeGreaterThan(0)
  expect(run.glyphs.every(glyph => glyph.id > 0)).toBe(true)
  // Compare the itemized result against a direct single-script HarfBuzz run.
  const hb = await import('harfbuzzjs'), font = new hb.Font(new hb.Face(new hb.Blob(bytes.slice().buffer), 0)), buffer = new hb.Buffer()
  font.setScale(parsed.unitsPerEm, parsed.unitsPerEm); buffer.addText(value); buffer.guessSegmentProperties(); buffer.setLanguage('und'); buffer.setFlags(hb.BufferFlag.REMOVE_DEFAULT_IGNORABLES); buffer.setClusterLevel(0); hb.shape(font, buffer)
  const direct = buffer.getGlyphInfos().map((info, index) => ({ id: info.codepoint, cluster: info.cluster, ...buffer.getGlyphPositions()[index]! }))
  expect(run.glyphs.map(glyph => glyph.id).sort((a, b) => a - b)).toEqual(direct.map(glyph => glyph.id).sort((a, b) => a - b))
  expect(run.width).toBe(direct.reduce((sum, glyph) => sum + glyph.xAdvance, 0))
  let pen = 0
  const positions = direct.map(glyph => {
    const result = { id: glyph.id, cluster: glyph.cluster, x: pen + glyph.xOffset, y: glyph.yOffset }
    pen += glyph.xAdvance
    return JSON.stringify(result)
  }).sort()
  expect(run.glyphs.map(({ id, cluster, x, y }) => JSON.stringify({ id, cluster, x, y })).sort()).toEqual(positions)
  const doc = await PDFDocument.create(), field = doc.getForm().createTextField('text')
  field.setText('BEFORE'); field.addToPage(doc.addPage([500, 200]), { x: 20, y: 80, width: 460, height: 60 })
  const source = await doc.save(), copy = source.slice()
  const result = await applyFormValues(source, [{ name: 'text', kind: 'text', value }], { textAppearance: { fontBytes: bytes } })
  expect(result.applied).toBe(1); expect(result.skipped).toEqual([])
  const saved = await PDFDocument.load(result.bytes), savedField = saved.getForm().getTextField('text')
  expect(savedField.getText()).toBe(value)
  const ap = saved.context.lookup(savedField.acroField.getWidgets()[0]!.getNormalAppearance()) as PDFRawStream
  const content = Buffer.from(decodePDFRawStream(ap).decode()).toString('latin1')
  expect(content).not.toMatch(/NaN|Infinity/)
  if (content.includes('ActualText')) expect(content).toContain(`/ActualText ${PDFHexString.fromText(value)}`)
  expect(source).toEqual(copy); expect(bytes).toEqual(originalFont)
  saved.getForm().flatten({ updateFieldAppearances: false })
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs'), task = pdfjs.getDocument({ data: await saved.save() })
  try {
    const content = await (await (await task.promise).getPage(1)).getTextContent({ disableNormalization: true })
    const extracted = content.items.map(item => 'str' in item ? item.str : '').join('')
    // Geometry-based reader spacing differs from canonical form semantics.
    const readerText = name === 'NotoSansDevanagari' ? 'क्षि नमस् ते' : name === 'NotoSansMyanmar' ? 'မြ န်မာ' : value
    expect(extracted).toBe(readerText)
  } finally { await task.destroy() }
})
