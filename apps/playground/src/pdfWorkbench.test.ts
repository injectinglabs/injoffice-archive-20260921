import { PDFArray, PDFDocument, PDFName, StandardFonts } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { VISUAL_SIGNATURE_CONTENT_PREFIX } from '../../../packages/pdf/src/annotate/stamp'
import {
  applyPdfAnnotDelete,
  applyPdfDrawing,
  applyPdfFormValues,
  pdfFormResultMessage,
  restorePdfSkippedDrafts,
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
    const nextFields = await listPdfFormFields(filled.bytes)
    expect(filled.applied).toBe(3)
    expect(filled.skipped).toEqual([])
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

  it('preserves partial and all-skipped form results without claiming success', async () => {
    const doc = await PDFDocument.create();doc.addPage()
    doc.getForm().createTextField('memo')
    const bytes = await doc.save()
    const partial = await applyPdfFormValues(bytes, [{ name: 'memo', kind: 'text', value: 'accepted' }, { name: 'missing', kind: 'text', value: 'not accepted' }])
    expect(partial.applied).toBe(1)
    expect(partial.skipped).toEqual([{ name: 'missing', reason: 'field not found' }])
    expect(pdfFormResultMessage(partial)).toBe('Applied 1 form value. Skipped 1: missing: field not found')
    const skipped = await applyPdfFormValues(bytes, [{ name: 'memo', kind: 'checkbox', checked: true }])
    expect(skipped.bytes).toBe(bytes)
    expect(skipped.applied).toBe(0)
    expect(pdfFormResultMessage(skipped)).toContain('No form values applied. Skipped 1: memo: field is not a checkbox')
    expect(pdfFormResultMessage({ applied: 2, skipped: [] })).toBe('Applied 2 form values.')
    await expect(applyPdfFormValues(new Uint8Array([1, 2]), [{ name: 'memo', kind: 'text', value: 'x' }])).rejects.toThrow()
  })

  it('computes adjacent page reorderings', () => {
    expect(movePageOrder(3, 2, -1)).toEqual([2, 1, 3])
    expect(movePageOrder(3, 2, 1)).toEqual([1, 3, 2])
    expect(movePageOrder(1, 1, 1)).toBeNull()
    expect(movePageOrder(3, 0, 1)).toBeNull()
  })

  it('retains only matching skipped drafts after a mixed form update', () => {
    const canonical = [{ name: 'memo', kind: 'text' as const, value: 'BEFORE' }, { name: 'agree', kind: 'checkbox' as const, checked: true }]
    const drafts = [{ name: 'memo', kind: 'text' as const, value: '你好' }, { name: 'agree', kind: 'checkbox' as const, checked: false }]
    expect(restorePdfSkippedDrafts(canonical, drafts, [{ name: 'memo' }])).toEqual([drafts[0], canonical[1]])
    expect(restorePdfSkippedDrafts(canonical, drafts, [])).toEqual(canonical)
    expect(restorePdfSkippedDrafts(canonical, [...drafts, drafts[0]!], [{ name: 'memo' }])).toEqual(canonical)
    expect(canonical[0]!.value).toBe('BEFORE')
  })

  it('forwards the explicit form font policy and reports generated appearances separately', async () => {
    const doc = await PDFDocument.create()
    const page = doc.addPage()
    const field = doc.getForm().createTextField('memo')
    field.addToPage(page, { x: 30, y: 30, width: 180, height: 30 })
    const bytes = await doc.save()
    const result = await applyPdfFormValues(bytes, [{ name: 'memo', kind: 'text', value: 'AFTER' }], { textAppearance: { font: 'Courier' } })
    expect(result.appearances).toEqual([{ name: 'memo', status: 'generated', widgets: 1 }])
    expect(pdfFormResultMessage(result)).toBe('Applied 1 form value. Generated 1 widget appearance with the selected font.')
    expect((await PDFDocument.load(result.bytes)).getForm().getTextField('memo').getText()).toBe('AFTER')
    const skipped = await applyPdfFormValues(bytes, [{ name: 'memo', kind: 'text', value: '你好' }], { textAppearance: { font: 'Helvetica' } })
    expect(skipped.applied).toBe(0)
    expect(skipped.bytes).toBe(bytes)
    expect(pdfFormResultMessage(skipped)).not.toContain('Generated')
    expect(pdfFormResultMessage({ applied: 2, skipped: [], appearances: [{ name: 'memo', status: 'generated', widgets: 2 }, { name: 'choice', status: 'viewer-required', widgets: 1 }] })).toContain('Generated 2 widget appearances with the selected font. 1 field still requires viewer-generated appearances.')
  })
})
