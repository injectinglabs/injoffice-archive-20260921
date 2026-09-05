import { PDFDocument, StandardFonts } from 'pdf-lib'
import * as pdfjsLib from 'pdfjs-dist/legacy/build/pdf.mjs'
import { describe, expect, it } from 'vitest'
import { applyPageOps, readInfo } from '../../../packages/pdf/src/pageOps'
import { PDF_NODE_PREFIX } from './pdfHostRoutes'
import { createPdfDemoFixture, PDF_DEMO_PAGE_TITLES } from './pdfDemoFixture'
import { clampPdfPage, pdfDownloadName } from './pages/PdfPage'

describe('playground pdf page-ops proof', () => {
  it('configures the bundled pdf.js worker asset', () => {
    expect(pdfjsLib.GlobalWorkerOptions.workerSrc).toMatch(/pdf\.worker\.min\.mjs/)
  })

  it('rotates a generated PDF through applyPageOps', async () => {
    const doc = await PDFDocument.create()
    const font = await doc.embedFont(StandardFonts.Helvetica)
    const page = doc.addPage([200, 300])
    page.drawText('proof', { x: 20, y: 260, size: 12, font })
    const bytes = await doc.save()
    const rotated = await applyPageOps(bytes, [{ type: 'rotate', pages: 'all', degrees: 90 }])
    const info = await readInfo(rotated)
    expect(info.pageCount).toBe(1)
    expect(info.pages[0]?.rotation).toBe(90)
  })

  it('starts with a multipage operating review and usable approval fields', async () => {
    const bytes = await createPdfDemoFixture()
    const doc = await PDFDocument.load(bytes)
    expect(doc.getPageCount()).toBe(PDF_DEMO_PAGE_TITLES.length)
    expect(doc.getForm().getFields().map((field) => field.getName())).toEqual([
      'release.decisionOwner',
      'release.rolloutTier',
      'release.decisionNotes',
      'release.approved',
    ])
  })

  it('keeps navigation within the document page range', () => {
    expect(clampPdfPage(0, 2)).toBe(1)
    expect(clampPdfPage(2, 2)).toBe(2)
    expect(clampPdfPage(3, 2)).toBe(2)
    expect(clampPdfPage(12, 0)).toBe(1)
  })

  it('uses an explicit edited download name after a page operation', () => {
    expect(pdfDownloadName('contract.pdf', false)).toBe('contract.pdf')
    expect(pdfDownloadName('contract.PDF', true)).toBe('contract-edited.pdf')
    expect(pdfDownloadName('../bad:name.pdf', true)).toBe('-bad-name-edited.pdf')
  })

  it('keeps the Node PDFium host on a dedicated Vite prefix', () => {
    expect(PDF_NODE_PREFIX).toBe('/v1/pdf')
  })
})
