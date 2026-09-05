import { PDFArray, PDFDocument, PDFName, PDFRef } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { applyDrawings } from '../annotate/drawing.js'
import { applyMarkups } from '../annotate/markup.js'
import { PdfViewerDocument } from '../viewer.js'
import { applyWholeAnnotationRedactions, listWholeAnnotationRedactionTargets } from './annotation.js'
import type { AnnotDeleteSpec } from '../annotate/types.js'

async function fixture() { const d = await PDFDocument.create(); d.addPage([200, 200]); return applyMarkups(await d.save(), [{ page: 1, type: 'highlight', color: [1, 1, 0], quads: [[10, 90, 100, 90, 10, 80, 100, 80]] }]) }
async function spec(bytes: Uint8Array): Promise<AnnotDeleteSpec> { const d = await PDFDocument.load(bytes); const a = d.getPages()[0]!.node.lookup(PDFName.of('Annots'), PDFArray); const ref = a.get(0) as PDFRef; return { page: 1, objNum: ref.objectNumber, subtype: 'highlight', rect: [10, 80, 100, 90] } }
describe('applyWholeAnnotationRedactions', () => {
  it('lists complete exact identities and withholds unsupported annotation types', async () => {
    const seeded = await fixture()
    const withUnsupported = await applyDrawings(seeded, [{ kind: 'ink', page: 1, color: [1, 0, 0], width: 2, paths: [[20, 20, 40, 30]] }])
    const expected = await spec(withUnsupported)
    const targets = await listWholeAnnotationRedactionTargets(withUnsupported)
    expect(targets).toEqual([{ ...expected, contents: '' }])
  })

  it('physically deletes an exact annotation and fresh-opens the result', async () => { const seeded = await fixture(); const out = await applyWholeAnnotationRedactions(seeded, [await spec(seeded)]); const d = await PDFDocument.load(out); expect(d.getPages()[0]!.node.lookup(PDFName.of('Annots'), PDFArray).size()).toBe(0); const v = await PdfViewerDocument.load(out); try { await v.getPage(1) } finally { await v.destroy() } })
  it('rejects stale identity', async () => { const seeded = await fixture(), s = await spec(seeded); await expect(applyWholeAnnotationRedactions(seeded, [{ ...s, objNum: s.objNum + 99, rect: [101, 80, 190, 90] }])).rejects.toThrow('did not match') })
})
