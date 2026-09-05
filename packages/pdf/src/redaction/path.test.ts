import { PDFDocument, rgb } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { renderPageToPng } from '../ocr/renderPage.js'
import { applyWholePathRedactions, listPagePaths } from './path.js'

async function fixture(): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const page = doc.addPage([200, 200])
  // pdf-lib emits a real content-stream vector path here, rather than an annotation
  // or raster substitute. Its footprint is intentionally isolated on the page.
  page.drawRectangle({ x: 30, y: 40, width: 60, height: 50, color: rgb(1, 0, 0) })
  return doc.save()
}

async function overlappingTextFixture(): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const page = doc.addPage([200, 200])
  page.drawRectangle({ x: 30, y: 40, width: 60, height: 50, color: rgb(1, 0, 0) })
  // The text's object bounds overlap the path rectangle. The redactor must refuse
  // rather than deleting the path while silently leaving sensitive text behind.
  page.drawText('KEEP', { x: 45, y: 60, size: 12 })
  return doc.save()
}

describe('applyWholePathRedactions', () => {
  it('physically removes a real, isolated vector path and proves fresh output opens/renders', async () => {
    const input = await fixture()
    const before = await listPagePaths(input)
    expect(before).toHaveLength(1)
    expect(before[0]!.axisAligned).toBe(true)

    const result = await applyWholePathRedactions(input, [{ page: 1, rect: before[0]!.rect }])
    expect(await listPagePaths(result.bytes)).toEqual([])
    expect(result.proofs[0]!.renderedBytes).toBeGreaterThan(0)
    expect((await renderPageToPng(result.bytes, 1)).png.length).toBeGreaterThan(0)
  })

  it('refuses a partial path footprint rather than masking or deleting extra content', async () => {
    const input = await fixture()
    const [x1, y1, x2, y2] = (await listPagePaths(input))[0]!.rect
    await expect(applyWholePathRedactions(input, [{ page: 1, rect: [x1, y1, (x1 + x2) / 2, y2] }]))
      .rejects.toThrow('intersects an object boundary')
  })

  it('refuses a selection that also covers a non-path page object', async () => {
    const input = await overlappingTextFixture()
    const path = (await listPagePaths(input))[0]!
    await expect(applyWholePathRedactions(input, [{ page: 1, rect: path.rect }]))
      .rejects.toThrow('non-path object')
  })
})
