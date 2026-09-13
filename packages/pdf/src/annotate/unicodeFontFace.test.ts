import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { execFileSync } from 'node:child_process'
import fontkit from '@pdf-lib/fontkit'
import { decodePDFRawStream, PDFArray, PDFDict, PDFDocument, PDFName, PDFRawStream } from 'pdf-lib'
import { expect, it } from 'vitest'
import { fontChecksum, standaloneUnicodeFontFace } from './unicodeFontFace.js'
import { applyFormValues } from './forms.js'
const directory = resolve(import.meta.dirname, '../../testdata/fonts')
const read = (name: string) => new Uint8Array(readFileSync(resolve(directory, name)))
const collection = read('NotoSans-Devanagari-Bengali.ttc')
const tag = (view: DataView, offset: number) => String.fromCharCode(...new Uint8Array(view.buffer, view.byteOffset + offset, 4))
const tableRecords = (bytes: Uint8Array, offset = 0) => {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  return Array.from({ length: view.getUint16(offset + 4) }, (_, i) => {
    const record = offset + 12 + i * 16
    return { name: tag(view, record), checksum: view.getUint32(record + 4), offset: view.getUint32(record + 8), length: view.getUint32(record + 12), record }
  })
}
it('reproduces the two licensed font collection offline', () => {
  expect(execFileSync(process.execPath, [resolve(import.meta.dirname, '../../../../scripts/generate-pdf-font-collection.mjs'), '--check'], { encoding: 'utf8' })).toContain('hashes match')
})
it.each([0, 1])('extracts selected face %s with independent valid SFNT checksums and unchanged source', index => {
  const copy = collection.slice(), face = standaloneUnicodeFontFace(collection, index), font = fontkit.create(face)
  expect(font.familyName).toBe(index === 0 ? 'Noto Sans Devanagari' : 'Noto Sans Bengali')
  expect(fontChecksum(face)).toBe(0xb1b0afba)
  const records = tableRecords(face), view = new DataView(face.buffer)
  expect(records.map(table => table.name)).toEqual(records.map(table => table.name).sort())
  expect(records.some(table => table.name === 'DSIG')).toBe(false)
  for (const table of records) {
    expect(table.offset % 4).toBe(0)
    expect(fontChecksum(face.subarray(table.offset, table.offset + table.length), table.name === 'head')).toBe(table.checksum)
  }
  const power = Math.floor(Math.log2(records.length))
  expect(view.getUint16(6)).toBe(2 ** power * 16); expect(view.getUint16(8)).toBe(power); expect(view.getUint16(10)).toBe(records.length * 16 - 2 ** power * 16)
  expect(collection).toEqual(copy)
})
it('retains original standalone bytes and disallows a nonzero standalone face', () => {
  const original = read('NotoSansDevanagari-Regular.ttf')
  expect(standaloneUnicodeFontFace(original)).toEqual(original)
  expect(() => standaloneUnicodeFontFace(original, 1)).toThrow(/face index/)
})
it('accepts TTC2 headers and omits a bounded collection signature from the extracted face', () => {
  const input = new DataView(collection.buffer), output = new Uint8Array(collection.length + 20), view = new DataView(output.buffer)
  output.set(collection.subarray(0, 20)); output.set(collection.subarray(20), 32)
  view.setUint32(4, 0x20000); view.setUint32(20, 0x44534947); view.setUint32(24, 8); view.setUint32(28, output.length - 8)
  for (let index = 0; index < 2; index++) {
    const previous = input.getUint32(12 + index * 4), next = previous + 12
    view.setUint32(12 + index * 4, next)
    for (const table of tableRecords(collection, previous)) view.setUint32(table.record + 12 + 8, table.offset + 12)
  }
  expect(standaloneUnicodeFontFace(output, 1)).toEqual(standaloneUnicodeFontFace(collection, 1))
  view.setUint32(28, 0)
  expect(() => standaloneUnicodeFontFace(output, 1)).toThrow(/signature overlaps/)
})
it('rejects malformed collection versions, counts and directory references', () => {
  for (const [offset, value] of [[4, 0x30000], [8, 0], [8, 65], [12, collection.length], [12, 0], [16, 24]]) {
    const bytes = collection.slice(); new DataView(bytes.buffer).setUint32(offset!, value!)
    expect(() => standaloneUnicodeFontFace(bytes)).toThrow()
  }
})
it.each([-1, 2, 0.5, NaN, Infinity])('rejects invalid collection face index %s', index => {
  expect(() => standaloneUnicodeFontFace(collection, index)).toThrow()
})
it('rejects duplicate, overlapping, unaligned and metadata-overlapping selected tables', () => {
  for (const mutation of ['duplicate', 'overlap', 'zeroMask', 'directory', 'unaligned', 'outside', 'shortHead']) {
    const bytes = collection.slice(), view = new DataView(bytes.buffer), base = view.getUint32(12), records = tableRecords(bytes, base), a = records[0]!, b = records[1]!
    if (mutation === 'duplicate') view.setUint32(b.record, view.getUint32(a.record))
    if (mutation === 'overlap') view.setUint32(b.record + 8, a.offset)
    if (mutation === 'zeroMask') {
      view.setUint32(a.record + 12, 32)
      view.setUint32(b.record + 8, a.offset); view.setUint32(b.record + 12, 0)
      view.setUint32(records[2]!.record + 8, a.offset + 4); view.setUint32(records[2]!.record + 12, 4)
    }
    if (mutation === 'directory') view.setUint32(a.record + 8, 0)
    if (mutation === 'unaligned') view.setUint32(a.record + 8, a.offset + 1)
    if (mutation === 'outside') view.setUint32(a.record + 12, bytes.length)
    if (mutation === 'shortHead') view.setUint32(records.find(table => table.name === 'head')!.record + 12, 12)
    expect(() => standaloneUnicodeFontFace(bytes)).toThrow()
  }
})
it('saves the second selected collection face and preserves exact source values/bytes', async () => {
  const value = 'ক্ষি বাংলা', doc = await PDFDocument.create(), field = doc.getForm().createTextField('text')
  field.setText('BEFORE'); field.addToPage(doc.addPage([500, 200]), { x: 20, y: 60, width: 460, height: 80 }); field.setFontSize(24)
  const source = await doc.save(), copy = source.slice(), fontCopy = collection.slice()
  const result = await applyFormValues(source, [{ name: 'text', kind: 'text', value }], { textAppearance: { fontBytes: collection, faceIndex: 1 } })
  expect(result.applied).toBe(1); expect(result.skipped).toEqual([])
  const saved = await PDFDocument.load(result.bytes), savedField = saved.getForm().getTextField('text')
  expect(savedField.getText()).toBe(value)
  const ap = saved.context.lookup(savedField.acroField.getWidgets()[0]!.getNormalAppearance()) as PDFRawStream
  const fonts = ap.dict.lookup(PDFName.of('Resources'), PDFDict).lookup(PDFName.of('Font'), PDFDict), font = fonts.lookup(fonts.keys()[0]!, PDFDict)
  const cid = font.lookup(PDFName.of('DescendantFonts'), PDFArray).lookup(0, PDFDict), stream = cid.lookup(PDFName.of('FontDescriptor'), PDFDict).lookup(PDFName.of('FontFile2')) as PDFRawStream
  expect(decodePDFRawStream(stream).decode()).toEqual(standaloneUnicodeFontFace(collection, 1))
  expect(source).toEqual(copy); expect(collection).toEqual(fontCopy)
  await expect(applyFormValues(source, [{ name: 'text', kind: 'text', value }], { textAppearance: { fontBytes: collection, faceIndex: 2 } })).rejects.toThrow(/face index/)
  expect(source).toEqual(copy); expect(collection).toEqual(fontCopy)
})
