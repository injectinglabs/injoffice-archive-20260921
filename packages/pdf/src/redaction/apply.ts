import { PdfViewerDocument } from '../viewer.js'
import { chainPdfium, loadPdfium, saveDoc, withDocument, FPDF_PAGEOBJ_TEXT, type Pdfium } from '../textEdit/pdfium.js'

export type RedactionRect = readonly [number, number, number, number]
export interface WholeTextRedaction { page: number; rect: RedactionRect }
export interface RedactionProof { removedText: string[]; rendered: boolean }
export interface RedactionResult { bytes: Uint8Array; proofs: RedactionProof[] }
type Bounds = RedactionRect
const norm = ([x1, y1, x2, y2]: Bounds): Bounds => [Math.min(x1, x2), Math.min(y1, y2), Math.max(x1, x2), Math.max(y1, y2)]
const overlaps = (a: Bounds, b: Bounds) => Math.min(a[2], b[2]) > Math.max(a[0], b[0]) && Math.min(a[3], b[3]) > Math.max(a[1], b[1])
const contains = (a: Bounds, b: Bounds) => b[0] >= a[0] && b[1] >= a[1] && b[2] <= a[2] && b[3] <= a[3]
function textOf(m: Pdfium, obj: number, textPage: number): string { const len = m._FPDFTextObj_GetText(obj, textPage, 0, 0); if (len <= 2) return ''; const p = m._malloc(len); try { m._FPDFTextObj_GetText(obj, textPage, p, len); return Buffer.from(m.HEAPU8.buffer, p, len - 2).toString('utf16le') } finally { m._free(p) } }

/** Remove only complete, non-empty TEXT objects. Boundary intersections and
 * all non-text content reject the entire request; no paint-over fallback. */
export async function applyWholeTextRedactions(bytes: Uint8Array, redactions: readonly WholeTextRedaction[]): Promise<RedactionResult> {
  if (!redactions.length) return { bytes, proofs: [] }
  const m = await loadPdfium()
  return chainPdfium(() => withDocument(m, bytes, async doc => {
    const proofs: RedactionProof[] = []
    for (const redaction of redactions) {
      const page = m._FPDF_LoadPage(doc, redaction.page - 1); if (!page) throw new Error(`redaction page ${redaction.page} does not exist`)
      const textPage = m._FPDFText_LoadPage(page), target = norm(redaction.rect), ptrs = [m._malloc(4), m._malloc(4), m._malloc(4), m._malloc(4)]
      try {
        const remove: Array<{ obj: number; text: string }> = []
        for (let i = 0; i < m._FPDFPage_CountObjects(page); i++) {
          const obj = m._FPDFPage_GetObject(page, i)
          if (!m._FPDFPageObj_GetBounds(obj, ptrs[0]!, ptrs[1]!, ptrs[2]!, ptrs[3]!)) throw new Error('cannot prove page-object bounds')
          const bounds: Bounds = [m.HEAPF32[ptrs[0]! >> 2]!, m.HEAPF32[ptrs[1]! >> 2]!, m.HEAPF32[ptrs[2]! >> 2]!, m.HEAPF32[ptrs[3]! >> 2]!]
          if (!overlaps(bounds, target)) continue
          if (!contains(target, norm(bounds))) throw new Error('redaction intersects an object boundary; no bytes changed')
          if (m._FPDFPageObj_GetType(obj) !== FPDF_PAGEOBJ_TEXT) throw new Error('redaction covers a non-text object; no bytes changed')
          const text = textOf(m, obj, textPage); if (!text) throw new Error('redaction covers ambiguous empty text; no bytes changed')
          remove.push({ obj, text })
        }
        if (!remove.length) throw new Error('redaction contains no eligible text object; no bytes changed')
        for (const item of remove) { if (!m._FPDFPage_RemoveObject(page, item.obj)) throw new Error('PDFium could not remove redacted text'); m._FPDFPageObj_Destroy(item.obj) }
        if (!m._FPDFPage_GenerateContent(page)) throw new Error('PDFium could not regenerate redacted page content')
        proofs.push({ removedText: remove.map(x => x.text), rendered: false })
      } finally { ptrs.forEach(p => m._free(p)); m._FPDFText_ClosePage(textPage); m._FPDF_ClosePage(page) }
    }
    const out = saveDoc(m, doc), viewer = await PdfViewerDocument.load(out)
    try { for (let i = 0; i < redactions.length; i++) { const text = await viewer.getPageText(redactions[i]!.page); for (const removed of proofs[i]!.removedText) if (text.includes(removed)) throw new Error('redaction extraction proof failed; output discarded'); await viewer.getPage(redactions[i]!.page); proofs[i]!.rendered = true } } finally { await viewer.destroy() }
    return { bytes: out, proofs }
  }))
}
