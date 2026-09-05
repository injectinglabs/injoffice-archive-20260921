import { PDFDocument } from 'pdf-lib'
import { PNG } from 'pngjs'
import { describe, expect, it } from 'vitest'
import { applyImageEdits, listPageImages } from '../image/imageEdit.js'
import { renderPageToPng } from '../ocr/renderPage.js'
import { applyWholeImageRedactions } from './image.js'

function png(): string { const p = new PNG({ width: 4, height: 4 }); p.data.fill(255); return PNG.sync.write(p).toString('base64') }

describe('applyWholeImageRedactions', () => {
  it('physically removes a real listed image and reopens/rasterizes fresh output', async () => {
    const doc = await PDFDocument.create(); doc.addPage([200, 200]);
    const seeded = await applyImageEdits(await doc.save(), [{ kind: 'insertImage', page: 1, image: png(), rect: [20, 20, 80, 80], layer: 'aboveText' }])
    const image = (await listPageImages(seeded.bytes))[0]!
    const result = await applyWholeImageRedactions(seeded.bytes, [{ page: 1, rect: image.rect }])
    expect(await listPageImages(result.bytes)).toEqual([])
    expect(result.proofs[0]!.renderedBytes).toBeGreaterThan(0)
    expect((await renderPageToPng(result.bytes, 1)).png.length).toBeGreaterThan(0)
  })

  it('refuses a stale or partial footprint instead of masking it', async () => {
    const doc = await PDFDocument.create(); doc.addPage([200, 200]);
    const seeded = await applyImageEdits(await doc.save(), [{ kind: 'insertImage', page: 1, image: png(), rect: [20, 20, 80, 80], layer: 'aboveText' }])
    await expect(applyWholeImageRedactions(seeded.bytes, [{ page: 1, rect: [20, 20, 50, 80] }])).rejects.toThrow('ambiguous')
  })
})
