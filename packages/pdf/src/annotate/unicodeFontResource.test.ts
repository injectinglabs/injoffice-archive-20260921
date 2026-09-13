import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import fontkit from '@pdf-lib/fontkit'
import { decodePDFRawStream, PDFArray, PDFDict, PDFDocument, PDFName, PDFRawStream } from 'pdf-lib'
import { describe, expect, it, vi } from 'vitest'
import { UnicodeFontResource } from './unicodeFontResource.js'
import { prepareUnicodeShaper } from './unicodeShaping.js'
const require = createRequire(import.meta.url)
const bytes = new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
const decoded = (stream: PDFRawStream) => Buffer.from(decodePDFRawStream(stream).decode()).toString('latin1')

describe('semantic CID resources', () => {
  it('does not publish a discarded qualification into existing resources', async () => {
    const doc = await PDFDocument.create(), parsed = fontkit.create(bytes)
    const shape = await prepareUnicodeShaper(bytes, parsed.unitsPerEm, parsed.numGlyphs)
    const resource = new UnicodeFontResource(parsed, bytes, doc)
    resource.qualify(shape('A'))
    const font = await resource.embed(), root = doc.context.lookup(font.ref, PDFDict)
    const before = decoded(root.lookup(PDFName.of('ToUnicode')) as PDFRawStream)
    resource.qualify(shape('Ω')) // models a later field's setText/MaxLen refusal
    expect(decoded(root.lookup(PDFName.of('ToUnicode')) as PDFRawStream)).toBe(before)
    resource.qualify(shape('Ж')); resource.commit()
    const after = decoded(root.lookup(PDFName.of('ToUnicode')) as PDFRawStream)
    expect(after).toContain('<0001> <0041>'); expect(after).toContain('<0002> <0416>'); expect(after).not.toContain('03A9')
    const cid = root.lookup(PDFName.of('DescendantFonts'), PDFArray).lookup(0, PDFDict)
    const map = decodePDFRawStream(cid.lookup(PDFName.of('CIDToGIDMap')) as PDFRawStream).decode()
    expect(new DataView(map.buffer, map.byteOffset, map.byteLength).getUint16(4)).toBe(parsed.glyphForCodePoint(0x416).id)
  })
  it('naturally partitions independent base/mark glyphs without an outline continuation', async () => {
    const doc = await PDFDocument.create(), parsed = fontkit.create(bytes)
    const shape = await prepareUnicodeShaper(bytes, parsed.unitsPerEm, parsed.numGlyphs)
    const resource = new UnicodeFontResource(parsed, bytes, doc)
    const result = resource.qualify(shape('x\u0301'))
    expect(result.cids).toEqual([1, 2]); expect(result.outlines).toEqual([undefined, undefined])
  })
  it('bounds untrusted continuation paths and leaves prior resources unchanged', async () => {
    const doc = await PDFDocument.create(), parsed = fontkit.create(bytes)
    const shape = await prepareUnicodeShaper(bytes, parsed.unitsPerEm, parsed.numGlyphs)
    const resource = new UnicodeFontResource(parsed, bytes, doc)
    resource.qualify(shape('A')); const font = await resource.embed(), root = doc.context.lookup(font.ref, PDFDict)
    const before = decoded(root.lookup(PDFName.of('ToUnicode')) as PDFRawStream)
    const run = shape('a\u0301\u0323'), gid = run.glyphs[1]!.id, glyph = parsed.getGlyph(gid)
    const get = parsed.getGlyph.bind(parsed)
    const spy = vi.spyOn(parsed, 'getGlyph').mockImplementation(id => id === gid ? Object.defineProperty(Object.create(glyph), 'path', { value: { commands: [{ command: 'lineTo', args: [NaN, 0] }] } }) : get(id))
    try { expect(() => resource.qualify(run)).toThrow('invalid embedded continuation outline') } finally { spy.mockRestore() }
    expect(decoded(root.lookup(PDFName.of('ToUnicode')) as PDFRawStream)).toBe(before)
  })
  it('enforces the PDF ToUnicode 512-byte destination limit before allocating resources', async () => {
    const doc = await PDFDocument.create(), parsed = fontkit.create(bytes), resource = new UnicodeFontResource(parsed, bytes, doc)
    const glyph = { id: parsed.glyphForCodePoint(65).id, cluster: 0, end: 257, x: 0, y: 0, advance: 1000 }
    expect(() => resource.qualify({ value: 'a'.repeat(257), width: 1000, unitsPerEm: parsed.unitsPerEm, glyphs: [glyph] })).toThrow('512-byte ToUnicode')
    expect(resource.qualify({ value: 'a'.repeat(256), width: 1000, unitsPerEm: parsed.unitsPerEm, glyphs: [{ ...glyph, end: 256 }] }).cids).toEqual([1])
  })

})
