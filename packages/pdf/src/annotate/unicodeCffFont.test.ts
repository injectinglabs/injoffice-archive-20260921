import { readFileSync } from 'node:fs'
import { createHash } from 'node:crypto'
import { resolve } from 'node:path'
import fontkit from '@pdf-lib/fontkit'
import { decodePDFRawStream, PDFArray, PDFDict, PDFDocument, PDFName, PDFRawStream } from 'pdf-lib'
import { expect, it } from 'vitest'
import { readFontFace } from '../textEdit/fontCmap.js'
import { prepareUnicodeCffFont } from './unicodeCffFont.js'
import { validateCffPrograms } from './unicodeCffProgram.js'
import { applyFormValues } from './forms.js'
import { prepareUnicodeShaper } from './unicodeShaping.js'

const directory = resolve(import.meta.dirname, '../../testdata/fonts')
const bytes = new Uint8Array(readFileSync(resolve(directory, 'NotoSansDevanagari-Regular.otf')))
const parsed = fontkit.create(bytes), table = readFontFace(bytes).tables.get('CFF ')!
const raw = bytes.subarray(table.offset, table.offset + table.length)
const cat = (...parts: (readonly number[] | Uint8Array)[]) => Uint8Array.from(parts.flatMap(part => [...part]))
const num = (n: number) => [29, n >>> 24 & 255, n >>> 16 & 255, n >>> 8 & 255, n & 255]
const idx = (parts: Uint8Array[]) => parts.length ? cat([0, parts.length, 1], Array.from({ length: parts.length + 1 }, (_, i) => 1 + parts.slice(0, i).reduce((sum, part) => sum + part.length, 0)), ...parts) : new Uint8Array(2)
function tinyNameFont(sid = 1, supplementalSid?: number, extra: number[] = []): Uint8Array {
  const names = idx([cat([84, 101, 115, 116])])
  const makeTop = (charset: number, glyphs: number, encoding: number) => idx([cat(num(charset), [15], num(glyphs), [17], ...(supplementalSid === undefined ? [] : [num(encoding), [16]]), extra)])
  const prefix = 4 + names.length + makeTop(0, 0, 0).length + 4
  const charset = cat([0, sid >>> 8, sid & 255]), glyphs = idx([cat([14]), cat([14])])
  return cat([1, 0, 4, 1], names, makeTop(prefix, prefix + charset.length, prefix + charset.length + glyphs.length), [0, 0, 0, 0], charset, glyphs, supplementalSid === undefined ? [] : [128, 1, 65, 1, 66, supplementalSid >>> 8, supplementalSid & 255])
}
function tinyCidFont(aliasRange = false, fdMatrix = false): Uint8Array {
  const names = idx([cat([84, 101, 115, 116])]), strings = idx([new TextEncoder().encode('Adobe'), new TextEncoder().encode('Identity')])
  const makeTop = (charset: number, glyphs: number, fds: number, select: number) => idx([cat(num(391), num(392), num(0), [12, 30], num(2), [12, 34], num(charset), [15], num(glyphs), [17], num(fds), [12, 36], num(select), [12, 37])])
  const prefix = 4 + names.length + makeTop(0, 0, 0, 0).length + strings.length + 2
  const glyphs = idx([cat([14]), cat([14])]), fds = idx([cat(num(0), num(0), [18], ...(fdMatrix ? [[...num(2), ...num(0), ...num(0), ...num(1), ...num(0), ...num(0), 12, 7]] : [])), cat(num(0), num(0), [18])])
  const fdsOffset = prefix + 3 + glyphs.length, selectOffset = fdsOffset + fds.length
  return cat([1, 0, 4, 1], names, makeTop(prefix, prefix + 3, fdsOffset, aliasRange ? prefix : selectOffset), strings, [0, 0], [0, 0, 1], glyphs, fds, [0, 0, 1])
}
function manyIndexEntries(shared: boolean): Uint8Array {
  const largeIndex = (glyphs: boolean) => {
    const output = new Uint8Array(3 + 65536 * 4 + (glyphs ? 65535 : 0)), view = new DataView(output.buffer)
    view.setUint16(0, 65535); output[2] = 4
    for (let i = 0; i <= 65535; i++) view.setUint32(3 + i * 4, glyphs ? i + 1 : 1)
    if (glyphs) output.fill(14, 3 + 65536 * 4)
    return output
  }
  const names = idx([cat([84, 101, 115, 116])]), strings = idx([new TextEncoder().encode('Adobe'), new TextEncoder().encode('Identity')]), globals = largeIndex(false), glyphs = largeIndex(true), locals = largeIndex(false)
  const top = (charset: number, chars: number, fds: number, select: number) => idx([cat(num(391), num(392), num(0), [12, 30], num(65535), [12, 34], num(charset), [15], num(chars), [17], num(fds), [12, 36], num(select), [12, 37])])
  const fds = (first: number, second: number) => idx([cat(num(6), num(first), [18]), cat(num(6), num(second), [18])])
  const charsetOffset = 4 + names.length + top(0, 0, 0, 0).length + strings.length + globals.length
  const glyphOffset = charsetOffset + 5, fdOffset = glyphOffset + glyphs.length, selectOffset = fdOffset + fds(0, 0).length
  const privateOffset = selectOffset + 8, secondPrivate = privateOffset + 6 + locals.length
  return cat([1, 0, 4, 4], names, top(charsetOffset, glyphOffset, fdOffset, selectOffset), strings, globals, [2, 0, 1, 255, 253], glyphs, fds(privateOffset, shared ? privateOffset : secondPrivate), [3, 0, 1, 0, 0, 0, 255, 255], num(6), [19], locals, ...(shared ? [] : [num(6), [19], locals]))
}
it('retains pinned fixed CFF font provenance', () => {
  const provenance = JSON.parse(readFileSync(resolve(directory, 'NotoSansDevanagari-Regular.otf.json'), 'utf8'))
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(provenance.sha256)
  expect(readFileSync(resolve(directory, provenance.license), 'utf8')).toContain('SIL OPEN FONT LICENSE')
})
it('wraps complete name-keyed programs with identity CIDs and revalidates as CID-keyed', () => {
  const copy = raw.slice(), converted = prepareUnicodeCffFont(raw, parsed.numGlyphs, parsed.unitsPerEm)
  expect([...converted.bytes.subarray(0, 4)]).toEqual([1, 0, 4, 4])
  expect(converted.bytes.subarray(converted.bytes.length - raw.length)).toEqual(raw)
  expect(converted.glyphCids).toEqual(Array.from({ length: parsed.numGlyphs }, (_, gid) => gid))
  expect(converted.registry).toBe('Adobe'); expect(converted.ordering).toBe('Identity')
  expect(prepareUnicodeCffFont(converted.bytes, parsed.numGlyphs, parsed.unitsPerEm)).toEqual(converted)
  expect(raw).toEqual(copy)
})
it('rejects malformed headers, counts and matrix mismatches before resources exist', () => {
  for (const index of [0, 1, 2, 3]) { const corrupt = raw.slice(); corrupt[index] = 255; expect(() => prepareUnicodeCffFont(corrupt, parsed.numGlyphs, parsed.unitsPerEm)).toThrow() }
  expect(() => prepareUnicodeCffFont(raw, parsed.numGlyphs - 1, parsed.unitsPerEm)).toThrow(/glyph counts/)
  expect(() => prepareUnicodeCffFont(raw, parsed.numGlyphs, 2048)).toThrow(/FontMatrix/)
})
it('rejects nonexistent charset SIDs and supplements absent from the encoded charset', () => {
  expect(() => prepareUnicodeCffFont(tinyNameFont(), 2, 1000)).not.toThrow()
  expect(() => prepareUnicodeCffFont(tinyNameFont(64000), 2, 1000)).toThrow(/SID/)
  expect(() => prepareUnicodeCffFont(tinyNameFont(1, 64000), 2, 1000)).toThrow(/supplement/)
  expect(() => prepareUnicodeCffFont(tinyNameFont(1, 2), 2, 1000)).toThrow(/supplement/)
  expect(() => prepareUnicodeCffFont(tinyNameFont(1, 1), 2, 1000)).not.toThrow()
})
it('rejects same-byte charset/FDSelect aliases, wrong DICT cardinalities and FD matrices', () => {
  expect(() => prepareUnicodeCffFont(tinyCidFont(), 2, 1000)).not.toThrow()
  expect(() => prepareUnicodeCffFont(tinyCidFont(true), 2, 1000)).toThrow(/overlapping/)
  expect(() => prepareUnicodeCffFont(tinyCidFont(false, true), 2, 1000)).toThrow(/FontMatrix/)
  expect(() => prepareUnicodeCffFont(tinyNameFont(1, undefined, [...num(1), 5]), 2, 1000)).toThrow(/DICT/)
  expect(() => prepareUnicodeCffFont(tinyNameFont(1, undefined, [...num(64000), 0]), 2, 1000)).toThrow(/SID/)
  expect(() => prepareUnicodeCffFont(tinyNameFont(1, undefined, [...num(0), 12, 39]), 2, 1000)).toThrow(/DICT/)
  const corrupt = tinyNameFont(); corrupt[7] = 0
  expect(() => prepareUnicodeCffFont(corrupt, 2, 1000)).toThrow(/INDEX/)
})
it('shares repeated subroutine INDEX objects and bounds total unique INDEX allocations', () => {
  const shared = manyIndexEntries(true)
  expect(prepareUnicodeCffFont(shared, 65535, 1000).bytes).toEqual(shared)
  expect(() => prepareUnicodeCffFont(manyIndexEntries(false), 65535, 1000)).toThrow(/INDEX entry limit/)
})
it('preserves real sparse CID mappings and all eight source font dictionaries', () => {
  const filename = 'NotoSansJP-CID-subset.otf', fixture = new Uint8Array(readFileSync(resolve(directory, filename)))
  const manifest = JSON.parse(readFileSync(resolve(directory, filename + '.json'), 'utf8'))
  expect(createHash('sha256').update(fixture).digest('hex')).toBe(manifest.sha256)
  expect(createHash('sha256').update(readFileSync(resolve(directory, manifest.license))).digest('hex')).toBe(manifest.licenseSha256)
  const font = fontkit.create(fixture), t = readFontFace(fixture).tables.get('CFF ')!, cff = fixture.subarray(t.offset, t.offset + t.length)
  const result = prepareUnicodeCffFont(cff, font.numGlyphs, font.unitsPerEm)
  expect(result.bytes).toEqual(cff)
  expect(result.glyphCids[font.glyphForCodePoint(65).id]).toBe(34)
  expect(result.glyphCids[font.glyphForCodePoint(233).id]).toBe(167)
  expect(result.glyphCids.some((cid, gid) => cid !== gid)).toBe(true)
  const internal = (font as unknown as { 'CFF ': { topDict: { FDArray: unknown[]; FDSelect: { ranges: { fd: number }[] } } } })['CFF ']
  expect(internal.topDict.FDArray).toHaveLength(8)
  expect(new Set(internal.topDict.FDSelect.ranges.map(range => range.fd)).size).toBe(8)
})
it('rejects name-keyed composites, malformed subroutines, random and stack underflow', () => {
  for (const glyph of [[139, 139, 139, 139, 14], [139, 10, 14], [12, 23, 14], [12, 18, 14], [139, 139, 21], [139, 12, 21, 22, 14], [139, 149, 12, 29, 21, 14], [139, 140, 1, 140, 140, 5, 14], [19, 14], [139, 139, 21, 139, 140, 1, 14], [1, 14], [139, 140, 1, 139, 139, 21, 20, 128, 14], [139, 140, 1, 19, 128, 139, 140, 19, 192, 14]]) {
    expect(() => validateCffPrograms([new Uint8Array(glyph)], [], [[]], [0])).toThrow()
  }
  expect(() => validateCffPrograms([new Uint8Array([139, 139, 21, 14])], [], [[]], [0])).not.toThrow()
  expect(() => validateCffPrograms([new Uint8Array([139, 140, 1, 19, 128, 139, 139, 21, 19, 128, 14])], [], [[]], [0])).not.toThrow()
  expect(() => validateCffPrograms([new Uint8Array([139, 138, 1, 139, 139, 21, 14])], [], [[]], [0])).toThrow(/negative/)
  expect(() => validateCffPrograms([new Uint8Array([139, 119, 1, 139, 139, 21, 14])], [], [[]], [0])).not.toThrow()
  expect(() => validateCffPrograms([new Uint8Array([139, 140, 3, 139, 140, 1, 139, 139, 21, 14])], [], [[]], [0])).toThrow(/phase/)
})
it.each([['NotoSansDevanagari-Regular.otf', 'क़', 'क़', '0958', '0915093C'], ['NotoSansJP-CID-subset.otf', 'é', 'é', '00E9', '00650301']])('maps %s aliases to one CFF CID without merging ToUnicode values', async (filename, first, second, firstHex, secondHex) => {
  const bytes = new Uint8Array(readFileSync(resolve(directory, filename!))), parsed = fontkit.create(bytes), table = readFontFace(bytes).tables.get('CFF ')!
  const cff = prepareUnicodeCffFont(bytes.subarray(table.offset, table.offset + table.length), parsed.numGlyphs, parsed.unitsPerEm)
  const shape = await prepareUnicodeShaper(bytes, parsed.unitsPerEm, parsed.numGlyphs)
  const composed = shape(first!), decomposed = shape(second!)
  expect(composed.glyphs).toHaveLength(1)
  expect(composed.glyphs.map(g => g.id)).toEqual(decomposed.glyphs.map(g => g.id))
  const doc = await PDFDocument.create(), page = doc.addPage([500, 200])
  for (const [name, y] of [['first', 110], ['second', 40]] as const) { const field = doc.getForm().createTextField(name); field.setText('BEFORE'); field.addToPage(page, { x: 20, y, width: 460, height: 60 }) }
  const source = await doc.save()
  const result = await applyFormValues(source, [{ name: 'first', kind: 'text', value: first! }, { name: 'second', kind: 'text', value: second! }], { textAppearance: { fontBytes: bytes } })
  expect(result.applied).toBe(2)
  const saved = await PDFDocument.load(result.bytes)
  expect(saved.getForm().getTextField('first').getText()).toBe(first)
  expect(saved.getForm().getTextField('second').getText()).toBe(second)
  const ap = saved.context.lookup(saved.getForm().getTextField('first').acroField.getWidgets()[0]!.getNormalAppearance()) as PDFRawStream
  const fonts = ap.dict.lookup(PDFName.of('Resources'), PDFDict).lookup(PDFName.of('Font'), PDFDict), font = fonts.lookup(fonts.keys()[0]!, PDFDict)
  const decoded = (name: string) => Buffer.from(decodePDFRawStream(font.lookup(PDFName.of(name)) as PDFRawStream).decode()).toString()
  const gid = cff.glyphCids[composed.glyphs[0]!.id]!
  expect(decoded('Encoding')).toContain(`<0001> ${gid}\n<0002> ${gid}`)
  expect(decoded('ToUnicode')).toContain(`<0001> <${firstHex}>\n<0002> <${secondHex}>`)
  saved.getForm().flatten({ updateFieldAppearances: false })
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs'), task = pdfjs.getDocument({ data: await saved.save() })
  try {
    const content = await (await (await task.promise).getPage(1)).getTextContent({ disableNormalization: true })
    expect(content.items.map(item => 'str' in item ? item.str : '').join('')).toBe(first! + second!)
  } finally { await task.destroy() }
})
it.each([['NotoSansDevanagari-Regular.otf', 'क्षि नमस्ते'], ['NotoSansJP-CID-subset.otf', 'Aé Ω 日本語かなカナ']])('saves %s Unicode with explicit raw CID resources and exact source values', async (filename, value) => {
  const bytes = new Uint8Array(readFileSync(resolve(directory, filename!))), parsed = fontkit.create(bytes), table = readFontFace(bytes).tables.get('CFF ')!, raw = bytes.subarray(table.offset, table.offset + table.length)
  const doc = await PDFDocument.create(), field = doc.getForm().createTextField('text')
  field.setText('BEFORE'); field.addToPage(doc.addPage([500, 200]), { x: 20, y: 60, width: 460, height: 80 }); field.setFontSize(24)
  const source = await doc.save(), copy = source.slice(), fontCopy = bytes.slice()
  const result = await applyFormValues(source, [{ name: 'text', kind: 'text', value }], { textAppearance: { fontBytes: bytes } })
  expect(result.applied).toBe(1); expect(result.skipped).toEqual([])
  const saved = await PDFDocument.load(result.bytes), savedField = saved.getForm().getTextField('text')
  expect(savedField.getText()).toBe(value)
  const ap = saved.context.lookup(savedField.acroField.getWidgets()[0]!.getNormalAppearance()) as PDFRawStream
  const fonts = ap.dict.lookup(PDFName.of('Resources'), PDFDict).lookup(PDFName.of('Font'), PDFDict), font = fonts.lookup(fonts.keys()[0]!, PDFDict)
  const cid = font.lookup(PDFName.of('DescendantFonts'), PDFArray).lookup(0, PDFDict), descriptor = cid.lookup(PDFName.of('FontDescriptor'), PDFDict)
  expect(String(cid.get(PDFName.of('Subtype')))).toBe('/CIDFontType0')
  expect(cid.has(PDFName.of('CIDToGIDMap'))).toBe(false); expect(descriptor.has(PDFName.of('FontFile2'))).toBe(false)
  const stream = descriptor.lookup(PDFName.of('FontFile3')) as PDFRawStream
  expect(String(stream.dict.get(PDFName.of('Subtype')))).toBe('/CIDFontType0C')
  expect(decodePDFRawStream(stream).decode()).toEqual(prepareUnicodeCffFont(raw, parsed.numGlyphs, parsed.unitsPerEm).bytes)
  const encoding = Buffer.from(decodePDFRawStream(font.lookup(PDFName.of('Encoding')) as PDFRawStream).decode()).toString()
  expect(encoding).toContain('begincidchar')
  if (filename!.includes('JP')) expect(encoding).toContain('<0001> 34')
  expect(source).toEqual(copy); expect(bytes).toEqual(fontCopy)
  saved.getForm().flatten({ updateFieldAppearances: false })
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs'), task = pdfjs.getDocument({ data: await saved.save() })
  try {
    const content = await (await (await task.promise).getPage(1)).getTextContent({ disableNormalization: true })
    expect(content.items.map(item => 'str' in item ? item.str : '').join('')).toBe(filename!.includes('Devanagari') ? 'क्षि नमस् ते' : value)
  } finally { await task.destroy() }
})
