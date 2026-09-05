import { describe, expect, it } from 'vitest'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import { PNG } from 'pngjs'
import { applyImageEdits, listPageImages, verifyImageEdits } from './imageEdit.js'

function solidPng(red: number, green: number, blue: number): string {
  const png = new PNG({ width: 3, height: 2 })
  for (let index = 0; index < png.data.length; index += 4) {
    png.data[index] = red
    png.data[index + 1] = green
    png.data[index + 2] = blue
    png.data[index + 3] = 255
  }
  return PNG.sync.write(png).toString('base64')
}

async function sourcePdf(): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const page = doc.addPage([300, 300])
  const font = await doc.embedFont(StandardFonts.Helvetica)
  page.drawText('layer marker', { x: 20, y: 250, font, size: 12 })
  const image = await doc.embedPng(solidPng(220, 40, 30))
  page.drawImage(image, { x: 20, y: 20, width: 60, height: 40 })
  return doc.save()
}

describe('independent PDF content-image replacement', () => {
  it('lists image geometry and its paint band relative to text', async () => {
    const images = await listPageImages(await sourcePdf())
    expect(images).toHaveLength(1)
    expect(images[0].page).toBe(1)
    expect(images[0].rect).toEqual([20, 20, 80, 60])
    expect(images[0].aboveText).toBe(true)
  })

  it('inserts below text, transforms, replaces, verifies, and deletes', async () => {
    const source = await sourcePdf()
    const inserted = await applyImageEdits(source, [{
      kind: 'insertImage', page: 1, image: solidPng(20, 180, 40), rect: [100, 30, 140, 80], layer: 'belowText',
    }])
    expect(inserted.skipped).toEqual([])
    expect((await listPageImages(inserted.bytes)).find(({ rect }) => rect[0] === 100)?.aboveText).toBe(false)

    const transformed = await applyImageEdits(inserted.bytes, [{
      kind: 'transformImage', page: 1, oldRect: [100, 30, 140, 80], rect: [110, 40, 170, 70], quarterTurns: 1, layer: 'aboveText',
    }])
    expect(transformed.skipped).toEqual([])
    expect(await verifyImageEdits(transformed.bytes, [{ page: 1, rect: [110, 40, 170, 70] }])).toEqual([])

    const replaced = await applyImageEdits(transformed.bytes, [{
      kind: 'replaceImage', page: 1, oldRect: [110, 40, 170, 70], rect: [120, 50, 180, 90], image: solidPng(10, 20, 240),
    }])
    expect(replaced.skipped).toEqual([])
    expect(await verifyImageEdits(replaced.bytes, [{ page: 1, rect: [120, 50, 180, 90] }])).toEqual([])

    const deleted = await applyImageEdits(replaced.bytes, [{ kind: 'deleteImage', page: 1, oldRect: [120, 50, 180, 90] }])
    expect(deleted.skipped).toEqual([])
    expect(await verifyImageEdits(deleted.bytes, [{ page: 1, rect: [120, 50, 180, 90] }])).toEqual([
      { page: 1, reason: 'image missing from saved output' },
    ])
  })

  it('fails softly for missing pages and stale image identities', async () => {
    const source = await sourcePdf()
    const result = await applyImageEdits(source, [
      { kind: 'deleteImage', page: 9, oldRect: [20, 20, 80, 60] },
      { kind: 'transformImage', page: 1, oldRect: [1, 1, 2, 2], rect: [10, 10, 20, 20] },
    ])
    expect(result.bytes).toBe(source)
    expect(result.skipped).toEqual([
      { editIndex: 0, page: 9, reason: 'page does not exist' },
      { editIndex: 1, page: 1, reason: 'image could not be located at oldRect' },
    ])
  })
})
