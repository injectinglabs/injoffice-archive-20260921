import { PDFDict, PDFDocument, PDFName, concatTransformationMatrix, drawObject, popGraphicsState, pushGraphicsState } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { renderPageToPng } from '../ocr/renderPage.js'
import { PdfViewerDocument } from '../viewer.js'
import { applyWholeXObjectRedactions, listPageXObjects } from './xobject.js'

async function fixture(shared = false): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const page = doc.addPage([200, 200])
  // A real Form XObject invoked by a page content stream. The literal text gives
  // the regression proof something recoverable if this is merely painted over.
  const form = doc.context.stream('BT /F1 12 Tf 0 6 Td (XOBJECT-SECRET) Tj ET', {
    Type: 'XObject', Subtype: 'Form', BBox: [0, 0, 100, 20],
    Resources: { Font: { F1: doc.context.obj({ Type: 'Font', Subtype: 'Type1', BaseFont: 'Helvetica' }) } },
  })
  const ref = doc.context.register(form)
  const resources = page.node.Resources()!
  let xobjects = resources.lookupMaybe(PDFName.of('XObject'), PDFDict)
  if (!xobjects) { xobjects = doc.context.obj({}); resources.set(PDFName.of('XObject'), xobjects) }
  xobjects.set(PDFName.of('SecretForm'), ref)
  page.pushOperators(pushGraphicsState(), concatTransformationMatrix(1, 0, 0, 1, 30, 40), drawObject('SecretForm'), popGraphicsState())
  if (shared) {
    const second = doc.addPage([200, 200])
    const secondResources = second.node.Resources()!
    const secondXObjects = doc.context.obj({ SecretForm: ref })
    secondResources.set(PDFName.of('XObject'), secondXObjects)
    second.pushOperators(pushGraphicsState(), concatTransformationMatrix(1, 0, 0, 1, 30, 40), drawObject('SecretForm'), popGraphicsState())
  }
  return doc.save()
}

describe('applyWholeXObjectRedactions', () => {
  it('physically removes a real isolated Form XObject, its resource, and its stream', async () => {
    const input = await fixture()
    const [target] = await listPageXObjects(input)
    expect(target).toMatchObject({ page: 1, name: 'SecretForm', axisAligned: true })
    expect((await renderPageToPng(input, 1)).png.length).toBeGreaterThan(0)

    const result = await applyWholeXObjectRedactions(input, [{ page: 1, name: target!.name, rect: target!.rect, wholeXObject: true }])
    expect(await listPageXObjects(result.bytes)).toEqual([])
    expect(result.proofs[0]!.removedObjectNumber).toBe(target!.objectNumber)
    expect(result.proofs[0]!.renderedBytes).toBeGreaterThan(0)
    const fresh = await PDFDocument.load(result.bytes)
    expect(fresh.getPages()[0]!.node.Resources()?.lookupMaybe(PDFName.of('XObject'), PDFDict)?.get(PDFName.of('SecretForm'))).toBeUndefined()
    expect(fresh.context.enumerateIndirectObjects().some(([ref]) => ref.objectNumber === target!.objectNumber)).toBe(false)
    const viewer = await PdfViewerDocument.load(result.bytes)
    try { expect(await viewer.getPageText(1)).not.toContain('XOBJECT-SECRET') } finally { await viewer.destroy() }
  })

  it('refuses partial or ambiguous/shared Form XObject selection rather than masking', async () => {
    const input = await fixture()
    const [target] = await listPageXObjects(input)
    const [x1, y1, x2, y2] = target!.rect
    await expect(applyWholeXObjectRedactions(input, [{ page: 1, name: target!.name, rect: [x1, y1, (x1 + x2) / 2, y2], wholeXObject: true }]))
      .rejects.toThrow('intersects an object boundary')
    await expect(listPageXObjects(await fixture(true))).resolves.toEqual([])
    const sharedDoc = await fixture(true)
    await expect(applyWholeXObjectRedactions(sharedDoc, [{ page: 1, name: 'SecretForm', rect: target!.rect, wholeXObject: true }]))
      .rejects.toThrow('shared or has unsupported ownership')
  })
})
