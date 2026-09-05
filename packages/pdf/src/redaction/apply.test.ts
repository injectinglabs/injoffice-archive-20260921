import { PDFDocument, StandardFonts } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { PdfViewerDocument } from '../viewer.js'
import { applyWholeTextRedactions } from './apply.js'

async function fixture(): Promise<Uint8Array> {
  const doc = await PDFDocument.create(); const font = await doc.embedFont(StandardFonts.Helvetica); const page = doc.addPage([400, 300])
  page.drawText('SECRET-729', { x: 40, y: 240, size: 20, font }); page.drawText('PUBLIC', { x: 40, y: 180, size: 20, font })
  return doc.save()
}
async function rect(bytes: Uint8Array, needle: string): Promise<[number, number, number, number]> {
  const doc = await PdfViewerDocument.load(bytes)
  try { const item = (await (await doc.getPage(1)).getTextContent()).items.find(x => 'str' in x && x.str === needle) as { transform: number[]; width: number; height: number } | undefined; if (!item) throw new Error('fixture text missing'); const [, , , , x, y] = item.transform; return [x - 2, y - 2, x + item.width + 2, y + item.height + 2] } finally { await doc.destroy() }
}
describe('applyWholeTextRedactions', () => {
  it('removes underlying whole text and proves fresh extraction/renderability', async () => {
    const input = await fixture(), out = await applyWholeTextRedactions(input, [{ page: 1, rect: await rect(input, 'SECRET-729') }])
    const doc = await PdfViewerDocument.load(out.bytes); try { expect(await doc.getPageText(1)).not.toContain('SECRET-729'); expect(await doc.getPageText(1)).toContain('PUBLIC') } finally { await doc.destroy() }
    expect(out.proofs[0]).toEqual({ removedText: ['SECRET-729'], rendered: true })
  })
  it('refuses a rectangle that cuts an object and leaves caller bytes untouched', async () => {
    const input = await fixture(), [x, y] = await rect(input, 'SECRET-729')
    await expect(applyWholeTextRedactions(input, [{ page: 1, rect: [x, y, x + 10, y + 10] }])).rejects.toThrow('intersects an object boundary')
  })
})
