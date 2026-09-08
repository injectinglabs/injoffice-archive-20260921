import { describe, expect, it } from 'vitest'
import { PDFDocument, degrees } from 'pdf-lib'
import { PdfViewerDocument } from '../../../packages/pdf/src/viewer'
import { annotationEntries, numberArray } from '../../../packages/pdf/src/annotate/pdfObjects'
import { applyPdfPageOp, applyPdfMarkup, applyPdfNote, applyPdfPlacedDrawing, listPdfAnnots } from './pdfWorkbench'
import { pointerToPdf, rangesOverlap, recordPdfEdit, travelPdfHistory, type PdfSnapshot } from './pdfInteraction'

describe('PDF user interaction', () => {
  it('maps pointer coordinates through zoom, crop origin and each page rotation', async () => {
    for (const rotation of [0, 90, 180, 270]) {
      const pdf = await PDFDocument.create()
      const page = pdf.addPage([400, 300])
      page.setCropBox(20, 30, 300, 200)
      page.setRotation(degrees(rotation))
      const viewer = await PdfViewerDocument.load(await pdf.save())
      try {
        const viewport = (await viewer.getPage(1)).getViewport({ scale: 1.5 })
        const [x, y] = viewport.convertToViewportPoint(100, 120)
        // Page appears at half its viewport size in this client rectangle.
        expect(pointerToPdf([50 + x! / 2, 60 + y! / 2], { left: 50, top: 60, width: viewport.width / 2, height: viewport.height / 2 }, viewport)).toEqual([100, 120])
      } finally { await viewer.destroy() }
    }
  })

  it('applies the selected passage geometry and chosen note position', async () => {
    const pdf = await PDFDocument.create()
    pdf.addPage([400, 300])
    const marked = await applyPdfMarkup(await pdf.save(), 1, 'highlight', [[100, 160, 180, 160, 100, 140, 180, 140]])
    const noted = await applyPdfNote(marked, 1, 'Review here', [230, 80])
    const annots = await listPdfAnnots(noted)
    expect(annots.find(a => a.subtype === 'highlight')?.rect).toEqual([100, 140, 180, 160])
    expect(annots.find(a => a.subtype === 'note')?.rect.slice(0, 2)).toEqual([230, 80])
    for (const kind of ['rect', 'ellipse', 'line', 'arrow', 'ink'] as const) {
      const result = await applyPdfPlacedDrawing(noted, 1, kind, [[210, 120], [180, 150], [130, 90]])
      const doc = await PDFDocument.load(result)
      const drawing = annotationEntries(doc.getPage(0)).at(-1)!.dict
      if (kind === 'rect' || kind === 'ellipse') expect(numberArray(drawing, 'Rect')).toEqual([130, 90, 210, 120])
      if (kind === 'line' || kind === 'arrow') expect(numberArray(drawing, 'L')).toEqual([210, 120, 130, 90])
      if (kind === 'ink') expect(numberArray(drawing, 'Rect')).toEqual([129.25, 89.25, 210.75, 150.75])
    }
  })

  it('restores deleted page bytes and page selection, redoes, then discards redo on a new edit', async () => {
    const pdf = await PDFDocument.create()
    pdf.addPage(); pdf.addPage()
    const original: PdfSnapshot = { bytes: await pdf.save(), page: 2, edited: false }
    const deleted: PdfSnapshot = { bytes: await applyPdfPageOp(original.bytes, { type: 'delete', pages: [2] }), page: 1, edited: true }
    const history = recordPdfEdit({ past: [], future: [] }, original)
    const undone = travelPdfHistory(history, deleted, 'undo')!
    expect(undone.snapshot).toEqual(original)
    expect((await PDFDocument.load(undone.snapshot.bytes)).getPageCount()).toBe(2)
    const redone = travelPdfHistory(undone.history, undone.snapshot, 'redo')!
    expect(redone.snapshot).toEqual(deleted)
    expect(recordPdfEdit(undone.history, original).future).toEqual([])
  })

  it('highlights matching passages across item boundaries without marking neighboring text', () => {
    expect(rangesOverlap(0, 8, 6, 7)).toBe(true)
    expect(rangesOverlap(9, 8, 6, 7)).toBe(true)
    expect(rangesOverlap(18, 6, 6, 7)).toBe(false)
  })
})
