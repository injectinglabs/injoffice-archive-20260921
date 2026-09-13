import { createRequire } from 'node:module'
import { readFileSync } from 'node:fs'
import { decodePDFRawStream, PDFDict, PDFDocument, PDFHexString, PDFName, PDFRawStream, TextAlignment } from 'pdf-lib'
import { describe, expect, it, vi } from 'vitest'
import fontkit from '@pdf-lib/fontkit'
import { applyFormValues } from './forms.js'
import { embeddedTextAppearance } from './embeddedTextAppearance.js'
import { prepareEmbeddedTextFont } from './embeddedTextFont.js'

const require = createRequire(import.meta.url)
const bytes = new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
const options = { textAppearance: { fontBytes: bytes } }
async function fixture({ alignment = TextAlignment.Left, size = 20, width = 100, rotated = false } = {}) {
  const doc = await PDFDocument.create(), field = doc.getForm().createTextField('text')
  field.setText('BEFORE')
  field.addToPage(doc.addPage([300, 300]), { x: 10, y: 100, width, height: 100 })
  field.setAlignment(alignment); field.setFontSize(size)
  const widget = field.acroField.getWidgets()[0]!
  widget.getBorderStyle()!.setWidth(0)
  widget.setRectangle({ x: 10, y: 100, width: rotated ? 100 : width, height: rotated ? width : 100 })
  if (rotated) widget.dict.set(PDFName.of('MK'), doc.context.obj({ R: 90 }))
  return doc.save({ updateFieldAppearances: false })
}
function appearance(doc: PDFDocument) {
  const widget = doc.getForm().getTextField('text').acroField.getWidgets()[0]!
  const ap = doc.context.lookup(widget.getNormalAppearance()) as PDFRawStream
  return { ap, content: Buffer.from(decodePDFRawStream(ap).decode()).toString('latin1') }
}
function fontDictionary(ap: PDFRawStream) {
  const fonts = ap.dict.lookup(PDFName.of('Resources'), PDFDict).lookup(PDFName.of('Font'), PDFDict)
  return fonts.lookup(fonts.keys()[0]!, PDFDict)
}

describe('positioned embedded horizontal appearances', () => {
  it.each([TextAlignment.Left, TextAlignment.Center, TextAlignment.Right])('uses actual AV advance for alignment %s and saves exact TJ correction', async alignment => {
    const source = await fixture({ alignment }), copy = source.slice()
    const result = await applyFormValues(source, [{ name: 'text', kind: 'text', value: 'AV' }], options)
    expect(result.applied).toBe(1)
    const saved = await PDFDocument.load(result.bytes), { content, ap } = appearance(saved)
    // DejaVuSans 2.37: A/V nominal widths1401 each, AV pair adjustment-131,
    // unitsPerEm2048. These pinned values are independent of the new adapter.
    expect(content).toContain('[ <0001> 63.96484375 <0002> ] TJ')
    const width = 2671 / 2048 * 20
    const x = Number(content.match(/1 0 0 1 ([\d.]+) [\d.]+ Tm/)![1])
    expect(x).toBeCloseTo(alignment === TextAlignment.Left ? 1 : alignment === TextAlignment.Center ? 1 + (98 - width) / 2 : 99 - width, 8)
    const cmap = fontDictionary(ap).lookup(PDFName.of('ToUnicode')) as PDFRawStream
    expect(Buffer.from(decodePDFRawStream(cmap).decode()).toString('latin1')).toContain('<0001> <0041>')
    expect(saved.getForm().getTextField('text').getText()).toBe('AV')
    expect(content).not.toContain('FFFC'); expect(source).toEqual(copy)
  })

  it('sizes the complete kerned run automatically and preserves rotation', async () => {
    const result = await applyFormValues(await fixture({ size: 0, width: 50, rotated: true }), [{ name: 'text', kind: 'text', value: 'AV' }], options)
    const { content } = appearance(await PDFDocument.load(result.bytes))
    // width48/advance1.30419921875 permits36pt; nominal unkerned width permits35.
    expect(content).toContain('/DejaVuSans 36 Tf')
    expect(content).toContain('63.96484375')
    expect(content).toMatch(/ 1 -1 .* 0 0 cm/)
  })

  it.each(['AV To', 'AV  To', ' AV To ', ''])('keeps spaces/empty value exact: %j', async value => {
    const source = await fixture({ size: 0 })
    const result = await applyFormValues(source, [{ name: 'text', kind: 'text', value }], options)
    const saved = await PDFDocument.load(result.bytes), { content, ap } = appearance(saved)
    expect(saved.getForm().getTextField('text').getText() ?? '').toBe(value)
    if (value) {
      expect(content).toContain('63.96484375'); expect(content).toContain('169.921875')
      const cmap = Buffer.from(decodePDFRawStream(fontDictionary(ap).lookup(PDFName.of('ToUnicode')) as PDFRawStream).decode()).toString('latin1')
      const mapping = new Map([...cmap.matchAll(/<([\dA-F]{4})> <([\dA-F]+)>/g)].map(match => [match[1], match[2]]))
      const encoded = content.match(/\[(.*?)\] TJ/)![1]!
      const scalars = [...encoded.matchAll(/<([\dA-F]{4})>/g)].map(match => String.fromCodePoint(parseInt(mapping.get(match[1])!, 16))).join('')
      expect(scalars).toBe(value)
    } else expect(content).toContain('<> Tj')
  })

  it('measures one whole run rather than separately losing cross-space advances', async () => {
    const doc = await PDFDocument.load(await fixture({ size: 0, width: 50 })), field = doc.getForm().getTextField('text')
    const prepared = await prepareEmbeddedTextFont(doc, options.textAppearance)
    prepared.qualify('A V'); const font = await prepared.embed(); field.setText('A V')
    // Synthetic qualified vector exercises the adapter with adjustments across
    // a space: the full positioned width is1em, while nominal width is3453/2048.
    const run = { value: 'A V', unitsPerEm: 2048, nominalAdvances: [1401, 651, 1401], advances: [1000, 48, 1000] }
    field.updateAppearances(font, (source, widget, selected) => embeddedTextAppearance(source, widget, selected, run))
    const saved = await PDFDocument.load(await doc.save({ updateFieldAppearances: false }))
    expect(appearance(saved).content).toContain('/DejaVuSans 47 Tf')
    expect(saved.getForm().getTextField('text').getText()).toBe('A V')
  })

  it.each(['enableMultiline', 'enablePassword', 'enableFileSelection', 'enableRichFormatting', 'enableCombing'] as const)('does not let the layout token mask %s', async flag => {
    const doc = await PDFDocument.load(await fixture()), field = doc.getForm().getTextField('text')
    field.setMaxLength(8); field[flag]()
    const source = await doc.save({ updateFieldAppearances: false })
    const result = await applyFormValues(source, [{ name: 'text', kind: 'text', value: 'AV' }], options)
    expect(result.applied).toBe(0); expect(result.bytes).toBe(source)
  })

  it('leaves font methods and source text accessor intact, refusing a stale qualified value', async () => {
    const doc = await PDFDocument.load(await fixture()), field = doc.getForm().getTextField('text')
    const prepared = await prepareEmbeddedTextFont(doc, options.textAppearance)
    prepared.qualify('AV'); const font = await prepared.embed(); field.setText('AV')
    const getText = field.getText, encode = font.encodeText, width = font.widthOfTextAtSize
    field.updateAppearances(font, prepared.appearanceProvider)
    expect(field.getText).toBe(getText); expect(font.encodeText).toBe(encode); expect(font.widthOfTextAtSize).toBe(width)
    field.setText('To')
    expect(() => field.updateAppearances(font, prepared.appearanceProvider)).toThrow('qualified text run')
  })

  it('rejects an encoding whose CID count differs from the qualified run', async () => {
    const doc = await PDFDocument.load(await fixture()), field = doc.getForm().getTextField('text')
    const prepared = await prepareEmbeddedTextFont(doc, options.textAppearance)
    prepared.qualify('AV'); const font = await prepared.embed(); field.setText('AV')
    const encode = vi.spyOn(font, 'encodeText').mockReturnValue(PDFHexString.of('0001'))
    expect(() => field.updateAppearances(font, prepared.appearanceProvider)).toThrow('encoding does not match')
    encode.mockRestore()
  })
  it.each([
    ['xAdvance', NaN], ['xAdvance', -1], ['xOffset', 10], ['yOffset', 10], ['yAdvance', 10],
  ] as const)('refuses unsupported %s=%s before mutating source', async (property, value) => {
    const parsed = fontkit.create(bytes)
    const layout = parsed.layout.bind(parsed)
    const positioned = vi.spyOn(parsed, 'layout').mockImplementation((...args) => {
      const run = layout(...args)
      run.positions[0]![property] = value
      return run
    })
    const create = vi.spyOn(fontkit, 'create').mockReturnValue(parsed)
    try {
      const source = await fixture(), copy = source.slice()
      const result = await applyFormValues(source, [{ name: 'text', kind: 'text', value: 'AV' }], options)
      expect(result.applied).toBe(0); expect(result.bytes).toBe(source); expect(source).toEqual(copy)
      expect(result.skipped[0]?.reason).toContain('unsupported glyph shaping or positioning')
    } finally { create.mockRestore(); positioned.mockRestore() }
  })

})
