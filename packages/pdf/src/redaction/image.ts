import { applyImageEdits, listPageImages } from '../image/index.js'
import { renderPageToPng } from '../ocr/renderPage.js'

export interface WholeImageRedaction { page: number; rect: [number, number, number, number] }
export interface ImageRedactionProof { page: number; rect: [number, number, number, number]; renderedBytes: number }
const close = (a: number[], b: number[]) => a.every((v, i) => Math.abs(v - b[i]!) <= 2)

/** Physically removes one previously enumerated, whole image object. The exact
 * footprint requirement intentionally rejects partial, nested, or ambiguous
 * selections; this is never a paint-over operation. */
export async function applyWholeImageRedactions(bytes: Uint8Array, redactions: readonly WholeImageRedaction[]): Promise<{ bytes: Uint8Array; proofs: ImageRedactionProof[] }> {
  let out = bytes; const proofs: ImageRedactionProof[] = []
  for (const r of redactions) {
    const before = (await listPageImages(out)).filter(i => i.page === r.page && close(i.rect, r.rect))
    if (before.length !== 1) throw new Error('redaction image is ambiguous, nested, or no longer matches its listed footprint')
    const result = await applyImageEdits(out, [{ kind: 'deleteImage', page: r.page, oldRect: r.rect }])
    if (result.skipped.length) throw new Error(`image removal failed: ${result.skipped[0]!.reason}`)
    out = result.bytes
    const after = (await listPageImages(out)).filter(i => i.page === r.page && close(i.rect, r.rect))
    if (after.length) throw new Error('image-object proof failed; output discarded')
    const rendered = await renderPageToPng(out, r.page)
    if (!rendered.png.length) throw new Error('raster proof failed; output discarded')
    proofs.push({ page: r.page, rect: r.rect, renderedBytes: rendered.png.length })
  }
  return { bytes: out, proofs }
}
