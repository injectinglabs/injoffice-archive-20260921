import { describe, expect, it } from 'vitest'
import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFHexString,
  PDFName,
  PDFStream,
  PDFRef,
  PDFString,
  degrees,
} from 'pdf-lib'
import { PNG } from 'pngjs'
import { applyAnnotDeletes } from './annotDelete.js'
import { applyDrawings, applyNoteEdits } from './drawing.js'
import { applyFormValues } from './forms.js'
import { applyMarkups } from './markup.js'
import { applySignatureStamps, applyStamps, VISUAL_SIGNATURE_CONTENT_PREFIX } from './stamp.js'

const key = (name: string) => PDFName.of(name)

async function blankPdf(pageCount = 1): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  for (let index = 0; index < pageCount; index += 1) doc.addPage([300, 300])
  return doc.save()
}

async function annotationEntries(bytes: Uint8Array, pageIndex = 0): Promise<Array<{ ref: PDFRef; dict: PDFDict }>> {
  const doc = await PDFDocument.load(bytes)
  const annots = doc.getPage(pageIndex).node.Annots()
  if (!annots) return []
  const entries: Array<{ ref: PDFRef; dict: PDFDict }> = []
  for (let index = 0; index < annots.size(); index += 1) {
    const ref = annots.get(index)
    if (ref instanceof PDFRef) entries.push({ ref, dict: doc.context.lookup(ref, PDFDict) })
  }
  return entries
}

function decodedText(dict: PDFDict, name: string): string | undefined {
  return dict.lookupMaybe(key(name), PDFString, PDFHexString)?.decodeText()
}

function solidPngBase64(red: number, green: number, blue: number): string {
  const png = new PNG({ width: 2, height: 2 })
  for (let index = 0; index < png.data.length; index += 4) {
    png.data[index] = red
    png.data[index + 1] = green
    png.data[index + 2] = blue
    png.data[index + 3] = 255
  }
  return PNG.sync.write(png).toString('base64')
}

describe('independent PDF annotation replacement', () => {
  it('authors ISO text markup dictionaries and normal appearances', async () => {
    const output = await applyMarkups(await blankPdf(), [
      { page: 1, type: 'highlight', color: [1, 0.8, 0], quads: [[20, 80, 120, 80, 20, 65, 120, 65]] },
      { page: 1, type: 'underline', color: [0, 0, 1], quads: [[20, 50, 120, 50, 20, 40, 120, 40]] },
    ])
    const entries = await annotationEntries(output)
    expect(entries.map(({ dict }) => dict.lookup(key('Subtype'), PDFName).decodeText())).toEqual(['Highlight', 'Underline'])
    for (const { dict } of entries) {
      expect(dict.lookup(key('QuadPoints'), PDFArray).size()).toBe(8)
      const appearance = dict.lookup(key('AP'), PDFDict).lookup(key('N'), PDFStream)
      expect(appearance.getContentsSize()).toBeGreaterThan(10)
    }
  })

  it('authors drawing appearances and threads same-batch notes', async () => {
    const output = await applyDrawings(await blankPdf(), [
      { page: 1, kind: 'ink', color: [1, 0, 0], width: 2, paths: [[10, 10, 20, 20, 30, 10]] },
      { page: 1, kind: 'rect', color: [0, 1, 0], width: 1, rect: [40, 40, 90, 80] },
      { page: 1, kind: 'ellipse', color: [0, 0, 1], width: 1, rect: [100, 40, 150, 80] },
      { page: 1, kind: 'arrow', color: [0, 0, 0], width: 2, from: [20, 100], to: [100, 130] },
      { page: 1, kind: 'note', color: [1, 1, 0], at: [30, 160], contents: 'root', localId: 'root' },
      { page: 1, kind: 'note', color: [1, 1, 0], at: [60, 160], contents: 'reply', replyToLocalId: 'root' },
    ])
    const entries = await annotationEntries(output)
    expect(entries.map(({ dict }) => dict.lookup(key('Subtype'), PDFName).decodeText())).toEqual([
      'Ink', 'Square', 'Circle', 'Line', 'Text', 'Text',
    ])
    const replyTarget = entries[5].dict.get(key('IRT'))
    expect(replyTarget).toBeInstanceOf(PDFRef)
    expect((replyTarget as PDFRef).objectNumber).toBe(entries[4].ref.objectNumber)
  })

  it('edits a note in place and deletes only a confirmed identity', async () => {
    const created = await applyDrawings(await blankPdf(), [
      { page: 1, kind: 'note', color: [1, 1, 0], at: [30, 40], contents: 'before' },
    ])
    const [{ ref }] = await annotationEntries(created)
    const edited = await applyNoteEdits(created, [{
      page: 1, objNum: ref.objectNumber, rect: [30, 40, 50, 60], oldContents: 'before', contents: 'after',
    }])
    expect(decodedText((await annotationEntries(edited))[0].dict, 'Contents')).toBe('after')

    const staleDelete = await applyAnnotDeletes(edited, [{
      page: 1, objNum: ref.objectNumber, subtype: 'note', rect: [30, 40, 50, 60], contents: 'before',
    }])
    expect(await annotationEntries(staleDelete)).toHaveLength(1)
    const deleted = await applyAnnotDeletes(staleDelete, [{
      page: 1, objNum: ref.objectNumber, subtype: 'note', rect: [30, 40, 50, 60], contents: 'after',
    }])
    expect(await annotationEntries(deleted)).toHaveLength(0)
  })

  it('fills supported fields and reports missing or wrong-type fields per item', async () => {
    const doc = await PDFDocument.create()
    const page = doc.addPage([300, 300])
    const form = doc.getForm()
    form.createTextField('customer.name').addToPage(page, { x: 10, y: 250, width: 120, height: 20 })
    form.createCheckBox('customer.active').addToPage(page, { x: 10, y: 220, width: 15, height: 15 })
    const radio = form.createRadioGroup('customer.plan')
    radio.addOptionToPage('basic', page, { x: 10, y: 190, width: 15, height: 15 })
    const choice = form.createDropdown('customer.region')
    choice.setOptions(['west', 'east'])
    choice.addToPage(page, { x: 10, y: 160, width: 100, height: 20 })

    const result = await applyFormValues(await doc.save(), [
      { name: 'customer.name', kind: 'text', value: 'Ada' },
      { name: 'customer.active', kind: 'checkbox', checked: true },
      { name: 'customer.plan', kind: 'radio', value: 'basic' },
      { name: 'customer.region', kind: 'choice', value: 'west' },
      { name: 'customer.missing', kind: 'text', value: 'x' },
      { name: 'customer.active', kind: 'text', value: 'wrong' },
    ])
    expect(result.applied).toBe(4)
    expect(result.skipped.map(({ name }) => name)).toEqual(['customer.missing', 'customer.active'])
    const saved = (await PDFDocument.load(result.bytes)).getForm()
    expect(saved.getTextField('customer.name').getText()).toBe('Ada')
    expect(saved.getCheckBox('customer.active').isChecked()).toBe(true)
    expect(saved.getRadioGroup('customer.plan').getSelected()).toBe('basic')
    expect(saved.getDropdown('customer.region').getSelected()).toEqual(['west'])
  })

  it('draws content stamps and creates rotated signature stamp appearances', async () => {
    const sourceDoc = await PDFDocument.create()
    sourceDoc.addPage([300, 300]).setRotation(degrees(90))
    const png = solidPngBase64(20, 100, 220)
    const contentStamped = await applyStamps(await sourceDoc.save(), [{ page: 1, image: png, rect: [10, 10, 40, 30], opacity: 0.5 }])
    const signed = await applySignatureStamps(contentStamped, [{ page: 1, image: png, rect: [50, 50, 120, 90], formFieldName: 'approval' }])
    const entries = await annotationEntries(signed)
    expect(entries).toHaveLength(1)
    expect(entries[0].dict.lookup(key('Subtype'), PDFName).decodeText()).toBe('Stamp')
    expect(decodedText(entries[0].dict, 'Contents')).toBe(`${VISUAL_SIGNATURE_CONTENT_PREFIX}approval`)
    expect(entries[0].dict.lookup(key('AP'), PDFDict).lookup(key('N'), PDFStream).getContentsString()).toContain('cm')
  })
})
