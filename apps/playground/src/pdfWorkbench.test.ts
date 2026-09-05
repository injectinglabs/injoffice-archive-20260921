import { PDFArray, PDFDocument, PDFName, StandardFonts } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { VISUAL_SIGNATURE_CONTENT_PREFIX } from '../../../packages/pdf/src/annotate/stamp'
import {
  applyPdfAnnotDelete,
  applyPdfDrawing,
  applyPdfFormValues,
  applyPdfMarkup,
  applyPdfNote,
  applyPdfNoteEdit,
  applyPdfPageOp,
  applyPdfStamp,
  listPdfAnnots,
  listPdfFormFields,
  mergePdfBytes,
  movePageOrder,
  pdfGeometry,
  splitPdfPrefix,
} from './pdfWorkbench'

async function makeDoc(labels: string[]): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const font = await doc.embedFont(StandardFonts.Helvetica)
  for (const label of labels) {
    const page = doc.addPage([400, 300])
    page.drawText(label, { x: 24, y: 250, size: 18, font })
  }
  return doc.save()
}

async function makeFormDoc(): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const page = doc.addPage([400, 300])
  const form = doc.getForm()
  form.createTextField('memo').addToPage(page, { x: 20, y: 200, width: 160, height: 20 })
  form.createCheckBox('agree').addToPage(page, { x: 20, y: 160, width: 16, height: 16 })
  const region = form.createDropdown('region')
  region.addOptions(['US', 'UK'])
  region.addToPage(page, { x: 20, y: 120, width: 120, height: 20 })
  return doc.save()
}

describe('playground PDF workbench helpers', () => {
  it('rotates, inserts, deletes, merges, splits, and n-ups through applyPageOps', async () => {
    const two = await makeDoc(['one', 'two'])
    const rotated = await applyPdfPageOp(two, { type: 'rotate', pages: [1], degrees: 90 })
    expect((await pdfGeometry(rotated)).pages[0]?.rotation).toBe(90)

    const inserted = await applyPdfPageOp(two, { type: 'insertBlank', at: 2 })
    expect((await pdfGeometry(inserted)).pageCount).toBe(3)

    const deleted = await applyPdfPageOp(two, { type: 'delete', pages: [1] })
    expect((await pdfGeometry(deleted)).pageCount).toBe(1)

    const reordered = await applyPdfPageOp(two, { type: 'reorder', order: movePageOrder(2, 1, 1)! })
    expect((await pdfGeometry(reordered)).pageCount).toBe(2)

    const merged = await mergePdfBytes(two, await makeDoc(['three']))
    expect((await pdfGeometry(merged)).pageCount).toBe(3)

    const split = await splitPdfPrefix(two, 1)
    expect((await pdfGeometry(split)).pageCount).toBe(1)

    const nUp = await applyPdfPageOp(two, { type: 'nUp', n: 2 })
    expect((await pdfGeometry(nUp)).pageCount).toBe(1)
  })

  it('adds markup, drawings, notes, stamps, and form values in the browser', async () => {
    const bytes = await makeFormDoc()
    const highlighted = await applyPdfMarkup(bytes, 1, 'highlight')
    const underlined = await applyPdfMarkup(highlighted, 1, 'underline')
    const struck = await applyPdfMarkup(underlined, 1, 'strikeout')
    const drawn = await applyPdfDrawing(struck, 1, 'rect')
    const inked = await applyPdfDrawing(drawn, 1, 'ink')
    const noted = await applyPdfNote(inked, 1, 'shared note')
    const stamped = await applyPdfStamp(noted, 1, false)
    const signed = await applyPdfStamp(stamped, 1, true)

    const annots = await listPdfAnnots(signed)
    expect(annots.map((annot) => annot.subtype).sort()).toEqual(['highlight', 'note', 'strikeout', 'underline'])

    const note = annots.find((annot) => annot.subtype === 'note')!
    const edited = await applyPdfNoteEdit(signed, note, 'updated note')
    expect((await listPdfAnnots(edited)).find((annot) => annot.subtype === 'note')?.contents).toBe('updated note')

    const highlight = (await listPdfAnnots(edited)).find((annot) => annot.subtype === 'highlight')!
    const withoutHighlight = await applyPdfAnnotDelete(edited, highlight)
    expect((await listPdfAnnots(withoutHighlight)).some((annot) => annot.subtype === 'highlight')).toBe(false)

    const fields = await listPdfFormFields(bytes)
    expect(fields.map((field) => field.kind).sort()).toEqual(['checkbox', 'choice', 'text'])
    const filled = await applyPdfFormValues(bytes, [
      { name: 'memo', kind: 'text', value: 'InjOffice' },
      { name: 'agree', kind: 'checkbox', checked: true },
      { name: 'region', kind: 'choice', value: 'UK' },
    ])
    const nextFields = await listPdfFormFields(filled)
    expect(nextFields).toEqual([
      { name: 'memo', kind: 'text', value: 'InjOffice' },
      { name: 'agree', kind: 'checkbox', checked: true },
      { name: 'region', kind: 'choice', value: 'UK' },
    ])

    const signedDoc = await PDFDocument.load(signed)
    const annotArray = signedDoc.getPages()[0]!.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
    expect(annotArray?.size()).toBeGreaterThan(annots.length)
    expect(VISUAL_SIGNATURE_CONTENT_PREFIX.length).toBeGreaterThan(0)
  })

  it('computes adjacent page reorderings', () => {
    expect(movePageOrder(3, 2, -1)).toEqual([2, 1, 3])
    expect(movePageOrder(3, 2, 1)).toEqual([1, 3, 2])
    expect(movePageOrder(1, 1, 1)).toBeNull()
    expect(movePageOrder(3, 0, 1)).toBeNull()
  })
})
