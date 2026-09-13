import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { decodePDFRawStream, degrees, PDFArray, PDFDict, PDFDocument, PDFName, PDFRawStream } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { applyFormValues } from './forms.js'
import { readFontFace } from '../textEdit/fontCmap.js'

const require = createRequire(import.meta.url)
const fontBytes = new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
const options = { textAppearance: { fontBytes } }

async function fixture(rotated = false): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const page = doc.addPage([400, 400])
  for (const name of ['first', 'second']) {
    const field = doc.getForm().createTextField(name)
    field.setText('BEFORE')
    field.addToPage(page, { x: 30, y: name === 'first' ? 200 : 100, width: 200, height: 32 })
    if (rotated && name === 'first') field.addToPage(doc.addPage([400, 400]), {
      x: 100, y: 100, width: 200, height: 32, rotate: degrees(90),
    })
  }
  return doc.save()
}

function streams(doc: PDFDocument, name: string): PDFRawStream[] {
  return doc.getForm().getTextField(name).acroField.getWidgets().map(widget => {
    const stream = doc.context.lookup(widget.getNormalAppearance())
    if (!(stream instanceof PDFRawStream)) throw new Error('missing saved normal appearance')
    return stream
  })
}

function decoded(stream: PDFRawStream): string {
  return Buffer.from(decodePDFRawStream(stream).decode()).toString('latin1')
}

function lookupStream(dict: PDFDict, key: string): PDFRawStream {
  const stream = dict.lookup(PDFName.of(key))
  if (!(stream instanceof PDFRawStream)) throw new Error(`missing ${key} stream`)
  return stream
}

function embeddedFont(stream: PDFRawStream): PDFDict {
  const fonts = stream.dict.lookup(PDFName.of('Resources'), PDFDict).lookup(PDFName.of('Font'), PDFDict)
  const font = fonts.lookup(fonts.keys()[0]!, PDFDict)
  expect(font.lookup(PDFName.of('Subtype'), PDFName).toString()).toBe('/Type0')
  expect(font.lookup(PDFName.of('Encoding'), PDFName).toString()).toBe('/Identity-H')
  const cid = font.lookup(PDFName.of('DescendantFonts'), PDFArray).lookup(0, PDFDict)
  const file = lookupStream(cid.lookup(PDFName.of('FontDescriptor'), PDFDict), 'FontFile2')
  const bytes = decodePDFRawStream(file).decode()
  expect(bytes.length).toBeGreaterThan(100)
  expect(bytes.length).toBeLessThan(fontBytes.length)
  expect(readFontFace(bytes).tables.has('glyf')).toBe(true)
  return font
}

// Replace only the fixture cmap with two scalars sharing one real outline.
// This reproduces ordinary font aliases without depending on a system font.
function aliasedFont(): Uint8Array {
  const bytes = new Uint8Array(fontBytes.length + 52)
  bytes.set(fontBytes)
  const view = new DataView(bytes.buffer)
  let record = 12
  while (String.fromCharCode(...bytes.subarray(record, record + 4)) !== 'cmap') record += 16
  const base = fontBytes.length
  view.setUint32(record + 8, base)
  view.setUint32(record + 12, 52)
  view.setUint16(base + 2, 1)
  view.setUint16(base + 4, 3)
  view.setUint16(base + 6, 10)
  view.setUint32(base + 8, 12)
  view.setUint16(base + 12, 12)
  view.setUint32(base + 16, 40)
  view.setUint32(base + 24, 2)
  for (const [index, scalar] of [0x00e9, 0x03a9].entries()) {
    const offset = base + 28 + index * 12
    view.setUint32(offset, scalar)
    view.setUint32(offset + 4, scalar)
    view.setUint32(offset + 8, 171) // DejaVuSans eacute outline
  }
  return bytes
}

describe('embedded Unicode form appearances', () => {
  it('saves subset outlines and exact BMP/supplementary ToUnicode mappings in every owned rotated widget', async () => {
    const source = await fixture(true)
    const sourceCopy = source.slice()
    const fontCopy = fontBytes.slice()
    const text = 'éΩЖ😀'
    const result = await applyFormValues(source, [{ name: 'first', kind: 'text', value: text }], options)
    expect(result).toMatchObject({ applied: 1, skipped: [], appearances: [{ name: 'first', status: 'generated', widgets: 2 }] })
    const loaded = await PDFDocument.load(result.bytes)
    expect(loaded.getForm().getTextField('first').getText()).toBe(text)
    for (const stream of streams(loaded, 'first')) {
      const font = embeddedFont(stream)
      const cmap = decoded(lookupStream(font, 'ToUnicode'))
      const mappings = new Map([...cmap.matchAll(/<([0-9A-Fa-f]{4})>\s+<([0-9A-Fa-f]+)>/g)].map(match => [match[1]!.toUpperCase(), match[2]!.toUpperCase()]))
      const encoded = decoded(stream).match(/<([0-9A-Fa-f]+)> Tj/)?.[1]
      expect(encoded).toBeDefined()
      const unicode = encoded!.match(/.{4}/g)!.map(cid => mappings.get(cid.toUpperCase())).join('')
      expect(unicode).toBe('00E903A90416D83DDE00')
    }
    const rotation = decoded(streams(loaded, 'first')[1]!).split('\n')[1]!.split(' ').slice(0, 6).map(Number)
    for (const [index, expected] of [0, 1, -1, 0, 0, 0].entries()) expect(rotation[index]).toBeCloseTo(expected)
    expect(streams(loaded, 'second').map(decoded)).toEqual(streams(await PDFDocument.load(source), 'second').map(decoded))
    expect(source).toEqual(sourceCopy)
    expect(fontBytes).toEqual(fontCopy)
  })

  it('refuses glyph aliases across separate fields instead of corrupting ToUnicode', async () => {
    const source = await fixture()
    const result = await applyFormValues(source, [
      { name: 'first', kind: 'text', value: 'é' },
      { name: 'second', kind: 'text', value: 'Ω' },
    ], { textAppearance: { fontBytes: aliasedFont() } })
    expect(result.applied).toBe(1)
    expect(result.skipped).toHaveLength(1)
    expect(result.skipped[0]?.name).toBe('second')
    const loaded = await PDFDocument.load(result.bytes)
    expect(loaded.getForm().getTextField('second').getText()).toBe('BEFORE')
    expect(streams(loaded, 'second').map(decoded)).toEqual(streams(await PDFDocument.load(source), 'second').map(decoded))
    expect(decoded(lookupStream(embeddedFont(streams(loaded, 'first')[0]!), 'ToUnicode'))).toMatch(/<0001> <00E9>/i)
  })

  it.each(['e\u0301', 'مرحبا', 'א', 'A١', '\u200f', '\ud800', '漢', 'a'.repeat(4097)])('preserves source and appearance when Unicode qualification refuses %s', async value => {
    const source = await fixture()
    const copy = source.slice()
    const result = await applyFormValues(source, [{ name: 'first', kind: 'text', value }], options)
    expect(result.applied).toBe(0)
    expect(result.skipped).toHaveLength(1)
    expect(result.bytes).toBe(source)
    expect(source).toEqual(copy)
  })

  it('saves an empty value using the embedded profile', async () => {
    const result = await applyFormValues(await fixture(), [{ name: 'first', kind: 'text', value: '' }], options)
    expect(result.applied).toBe(1)
    expect(result.skipped).toEqual([])
    const loaded = await PDFDocument.load(result.bytes)
    expect(loaded.getForm().getTextField('first').getText() ?? '').toBe('')
    expect(decoded(streams(loaded, 'first')[0]!)).toContain('<> Tj')
  })

  it.each([new Uint8Array(), new Uint8Array(16 * 1024 * 1024 + 1), new Uint8Array(12)])('rejects malformed or oversized font bytes before returning a modified PDF', async supplied => {
    const source = await fixture()
    const copy = source.slice()
    await expect(applyFormValues(source, [{ name: 'first', kind: 'text', value: 'é' }], { textAppearance: { fontBytes: supplied } })).rejects.toThrow()
    expect(source).toEqual(copy)
  })
})
