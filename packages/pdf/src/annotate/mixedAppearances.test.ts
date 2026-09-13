import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { resolve } from 'node:path'
import { decodePDFRawStream, PDFArray, PDFBool, PDFDict, PDFDocument, PDFHexString, PDFName, PDFRawStream } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { applyFormValues, type FormValuesOptions } from './forms.js'
import type { FormValueSpec } from './types.js'

const require = createRequire(import.meta.url)
const fontBytes = new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
const options = { textAppearance: { fontBytes }, choiceAppearance: { font: 'Courier' as const } }
const text: FormValueSpec = { name: 'text', kind: 'text', value: 'AV café Ω Ж 😀' }
const choice: FormValueSpec = { name: 'choice', kind: 'choice', value: 'b' }
async function fixture(unicodeChoice = false, priorRequest = false) {
  const doc = await PDFDocument.create(), page = doc.addPage([400, 400]), form = doc.getForm()
  for (const name of ['text', 'untouched']) {
    const field = form.createTextField(name)
    field.setText('BEFORE')
    field.addToPage(page, { x: 10, y: name === 'text' ? 300 : 250, width: 300, height: 30 })
  }
  const dropdown = form.createDropdown('choice')
  dropdown.acroField.setOptions(['a', 'b'].map(value => ({ value: PDFHexString.fromText(value), display: PDFHexString.fromText(unicodeChoice ? 'Ω' + value : value.toUpperCase()) })))
  // Author the initial ASCII artwork before optionally installing Unicode labels.
  const authored = dropdown.acroField.getOptions()
  dropdown.acroField.setOptions(['a', 'b'].map(value => ({ value: PDFHexString.fromText(value), display: PDFHexString.fromText(value.toUpperCase()) })))
  dropdown.addToPage(page, { x: 10, y: 180, width: 200, height: 40 })
  dropdown.acroField.setOptions(authored)
  if (priorRequest) form.acroForm.dict.set(PDFName.of('NeedAppearances'), PDFBool.True)
  return doc.save({ updateFieldAppearances: false })
}
function stream(doc: PDFDocument, name: string) {
  const value = doc.context.lookup(doc.getForm().getField(name).acroField.getWidgets()[0]!.getNormalAppearance())
  if (!(value instanceof PDFRawStream)) throw new Error('missing saved appearance')
  return value
}
const decoded = (value: PDFRawStream) => Buffer.from(decodePDFRawStream(value).decode()).toString('latin1')
function font(doc: PDFDocument, name: string) {
  const fonts = stream(doc, name).dict.lookup(PDFName.of('Resources'), PDFDict).lookup(PDFName.of('Font'), PDFDict)
  return fonts.lookup(fonts.keys()[0]!, PDFDict)
}

describe('independent embedded text and standard choice appearances', () => {
  it.each([false, true])('keeps CFF, standard choice and missing-glyph refusal independent, choice first=%s', async choiceFirst => {
    const bytes = new Uint8Array(readFileSync(resolve(import.meta.dirname, '../../testdata/fonts/NotoSansDevanagari-Regular.otf'))), fontCopy = bytes.slice()
    const cffText: FormValueSpec = { ...text, value: 'क्षि नमस्ते' }, refused: FormValueSpec = { name: 'untouched', kind: 'text', value: '\u{10FFFF}' }
    const source = await fixture(false, true), copy = source.slice(), before = await PDFDocument.load(source)
    const result = await applyFormValues(source, choiceFirst ? [choice, refused, cffText] : [cffText, refused, choice], { textAppearance: { fontBytes: bytes }, choiceAppearance: options.choiceAppearance })
    expect(result.applied).toBe(2); expect(result.skipped.map(item => item.name)).toEqual(['untouched'])
    const saved = await PDFDocument.load(result.bytes), cid = font(saved, 'text').lookup(PDFName.of('DescendantFonts'), PDFArray).lookup(0, PDFDict)
    expect(saved.getForm().getTextField('text').getText()).toBe(cffText.value)
    expect(String(cid.get(PDFName.of('Subtype')))).toBe('/CIDFontType0')
    expect(String(font(saved, 'choice').get(PDFName.of('BaseFont')))).toBe('/Courier')
    expect(decoded(stream(saved, 'untouched'))).toBe(decoded(stream(before, 'untouched')))
    expect(saved.getForm().getTextField('untouched').getText()).toBe('BEFORE')
    expect(saved.getForm().acroForm.dict.get(PDFName.of('NeedAppearances'))).toBe(PDFBool.True)
    expect(source).toEqual(copy); expect(bytes).toEqual(fontCopy)
  })
  it.each([false, true])('keeps RTL resources independent with choice first=%s', async choiceFirst => {
    const rtl: FormValueSpec = { ...text, value: 'abc \u2067مَرْحَبًا 123\u2069 xyz' }
    const source = await fixture(false, true), copy = source.slice()
    const result = await applyFormValues(source, choiceFirst ? [choice, rtl] : [rtl, choice], options)
    expect(result.applied).toBe(2); expect(result.skipped).toEqual([])
    const saved = await PDFDocument.load(result.bytes)
    expect(saved.getForm().getTextField('text').getText()).toBe(rtl.value)
    expect(decoded(stream(saved, 'text'))).toContain(`/ActualText ${PDFHexString.fromText(rtl.value as string)}`)
    expect(font(saved, 'text').get(PDFName.of('Subtype'))?.toString()).toBe('/Type0')
    expect(font(saved, 'choice').get(PDFName.of('BaseFont'))?.toString()).toBe('/Courier')
    expect(saved.getForm().acroForm.dict.get(PDFName.of('NeedAppearances'))).toBe(PDFBool.True)
    expect(source).toEqual(copy)
  })
  it.each([false, true])('generates both profiles with choice first=%s and distinct font resources', async choiceFirst => {
    const source = await fixture(), copy = source.slice(), before = await PDFDocument.load(source)
    const result = await applyFormValues(source, choiceFirst ? [choice, text] : [text, choice], options)
    expect(result.applied).toBe(2); expect(result.skipped).toEqual([])
    expect(result.appearances?.map(item => item.status)).toEqual(['generated', 'generated'])
    const saved = await PDFDocument.load(result.bytes), form = saved.getForm()
    expect(form.getTextField('text').getText()).toBe(text.value)
    expect(form.getDropdown('choice').acroField.getValues().map(value => value.decodeText())).toEqual(['b'])
    expect(form.getDropdown('choice').acroField.dict.get(PDFName.of('Opt'))?.toString()).toBe(before.getForm().getDropdown('choice').acroField.dict.get(PDFName.of('Opt'))?.toString())
    const textFont = font(saved, 'text'), choiceFont = font(saved, 'choice')
    expect(textFont.get(PDFName.of('Subtype'))?.toString()).toBe('/Type0')
    expect(decoded(stream(saved, 'text'))).toContain('63.96484375')
    expect(decoded(stream(saved, 'text'))).toContain('] TJ')
    const cid = textFont.lookup(PDFName.of('DescendantFonts'), PDFArray).lookup(0, PDFDict)
    expect(cid.lookup(PDFName.of('FontDescriptor'), PDFDict).lookup(PDFName.of('FontFile2'))).toBeInstanceOf(PDFRawStream)
    expect(decoded(textFont.lookup(PDFName.of('ToUnicode')) as PDFRawStream)).toContain('<D83DDE00>')
    expect(choiceFont.get(PDFName.of('BaseFont'))?.toString()).toBe('/Courier')
    expect(decoded(stream(saved, 'choice'))).toContain('<42> Tj')
    expect(decoded(stream(saved, 'choice'))).not.toContain('<62> Tj')
    expect(decoded(stream(saved, 'untouched'))).toBe(decoded(stream(before, 'untouched')))
    expect(form.acroForm.dict.has(PDFName.of('NeedAppearances'))).toBe(false)
    expect(source).toEqual(copy)
  })

  it.each(['text', 'choice'] as const)('does not relax the rejected %s profile when the other succeeds', async rejected => {
    const source = await fixture(rejected === 'choice'), before = await PDFDocument.load(source)
    const result = await applyFormValues(source, [rejected === 'text' ? { ...text, value: 'अ' } : text, choice], options)
    expect(result.applied).toBe(1); expect(result.skipped.map(item => item.name)).toEqual([rejected])
    const saved = await PDFDocument.load(result.bytes)
    expect(decoded(stream(saved, rejected))).toBe(decoded(stream(before, rejected)))
    if (rejected === 'text') {
      expect(saved.getForm().getTextField('text').getText()).toBe('BEFORE')
      expect(saved.getForm().getDropdown('choice').acroField.getValues()[0]!.decodeText()).toBe('b')
    } else {
      expect(saved.getForm().getTextField('text').getText()).toBe(text.value)
      expect(saved.getForm().getDropdown('choice').acroField.getValues()).toEqual([])
    }
  })

  it.each([
    { textAppearance: options.textAppearance },
    { choiceAppearance: options.choiceAppearance },
  ] satisfies FormValuesOptions[])('keeps unrequested profiles viewer-required: %j', async selected => {
    const result = await applyFormValues(await fixture(), [text, choice], selected)
    expect(result.applied).toBe(2)
    expect(result.appearances?.map(item => item.status)).toEqual('textAppearance' in selected ? ['generated', 'viewer-required'] : ['viewer-required', 'generated'])
    expect((await PDFDocument.load(result.bytes)).getForm().acroForm.dict.get(PDFName.of('NeedAppearances'))).toBe(PDFBool.True)
  })

  it('preserves a prior viewer request when both profiles generate', async () => {
    const result = await applyFormValues(await fixture(false, true), [text, choice], options)
    expect(result.appearances?.every(item => item.status === 'generated')).toBe(true)
    expect((await PDFDocument.load(result.bytes)).getForm().acroForm.dict.get(PDFName.of('NeedAppearances'))).toBe(PDFBool.True)
  })

  it('refuses mixed XFA before touching source or font data', async () => {
    const doc = await PDFDocument.load(await fixture())
    doc.getForm().acroForm.dict.set(PDFName.of('XFA'), doc.context.obj('preserve'))
    const source = await doc.save({ updateFieldAppearances: false })
    const result = await applyFormValues(source, [text, choice], options)
    expect(result.applied).toBe(0); expect(result.skipped).toHaveLength(2); expect(result.bytes).toBe(source)
  })
})
