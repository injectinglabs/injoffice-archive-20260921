import { renderPageToPng } from '../ocr/renderPage.js'
import { chainPdfium, FPDF_PAGEOBJ_PATH, loadPdfium, saveDoc, withDocument, type Pdfium } from '../textEdit/pdfium.js'

export type PdfRect = readonly [number, number, number, number]

/** A redaction rectangle is in PDF user space, not viewport/CSS coordinates. */
export interface WholePathRedaction {
  page: number
  rect: PdfRect
}

/** A listed content-stream path. It is intentionally only an inspection key;
 * callers must submit a rectangle enclosing its entire footprint to delete it. */
export interface PagePathRef {
  page: number
  rect: [number, number, number, number]
  axisAligned: boolean
}

export interface PathRedactionProof {
  page: number
  removedRect: [number, number, number, number]
  renderedBytes: number
}

export interface PathRedactionResult {
  bytes: Uint8Array
  proofs: PathRedactionProof[]
}

type Rect = [number, number, number, number]

const EPSILON = 1e-4

function normalized(rect: PdfRect): Rect {
  const [x1, y1, x2, y2] = rect
  if (![x1, y1, x2, y2].every(Number.isFinite)) throw new Error('redaction rectangle must contain finite coordinates')
  const out: Rect = [Math.min(x1, x2), Math.min(y1, y2), Math.max(x1, x2), Math.max(y1, y2)]
  if (out[0] >= out[2] || out[1] >= out[3]) throw new Error('redaction rectangle must have positive area')
  return out
}

function overlaps(a: Rect, b: Rect): boolean {
  return Math.min(a[2], b[2]) > Math.max(a[0], b[0]) + EPSILON
    && Math.min(a[3], b[3]) > Math.max(a[1], b[1]) + EPSILON
}

function contains(outer: Rect, inner: Rect): boolean {
  return inner[0] >= outer[0] - EPSILON
    && inner[1] >= outer[1] - EPSILON
    && inner[2] <= outer[2] + EPSILON
    && inner[3] <= outer[3] + EPSILON
}

function boundsOf(m: Pdfium, obj: number, ptrs: readonly number[]): Rect {
  if (!m._FPDFPageObj_GetBounds(obj, ptrs[0]!, ptrs[1]!, ptrs[2]!, ptrs[3]!)) {
    throw new Error('cannot prove page-object bounds; no bytes changed')
  }
  return normalized([
    m.HEAPF32[ptrs[0]! >> 2]!,
    m.HEAPF32[ptrs[1]! >> 2]!,
    m.HEAPF32[ptrs[2]! >> 2]!,
    m.HEAPF32[ptrs[3]! >> 2]!,
  ])
}

/** A path with rotation or shear has only an axis-aligned *bounding box*; deleting it
 * from a rectangle would therefore overstate what was selected. Refuse it. */
function isAxisAligned(m: Pdfium, obj: number): boolean {
  const matrix = m._malloc(24)
  try {
    if (!m._FPDFPageObj_GetMatrix(obj, matrix)) return false
    const a = m.HEAPF32[matrix >> 2]!
    const b = m.HEAPF32[(matrix >> 2) + 1]!
    const c = m.HEAPF32[(matrix >> 2) + 2]!
    const d = m.HEAPF32[(matrix >> 2) + 3]!
    return [a, b, c, d].every(Number.isFinite) && Math.abs(a) > EPSILON && Math.abs(d) > EPSILON
      && Math.abs(b) <= EPSILON && Math.abs(c) <= EPSILON
  } finally {
    m._free(matrix)
  }
}

/** Enumerate content-stream vector paths without exposing object handles. */
export function listPagePaths(bytes: Uint8Array): Promise<PagePathRef[]> {
  return chainPdfium(async () => {
    const m = await loadPdfium()
    return withDocument(m, bytes, async doc => {
      const out: PagePathRef[] = []
      const ptrs = [m._malloc(4), m._malloc(4), m._malloc(4), m._malloc(4)]
      try {
        for (let pageIndex = 0; pageIndex < m._FPDF_GetPageCount(doc); pageIndex++) {
          const page = m._FPDF_LoadPage(doc, pageIndex)
          if (!page) throw new Error(`could not load page ${pageIndex + 1}`)
          try {
            for (let i = 0; i < m._FPDFPage_CountObjects(page); i++) {
              const obj = m._FPDFPage_GetObject(page, i)
              if (m._FPDFPageObj_GetType(obj) !== FPDF_PAGEOBJ_PATH) continue
              out.push({ page: pageIndex + 1, rect: boundsOf(m, obj, ptrs), axisAligned: isAxisAligned(m, obj) })
            }
          } finally {
            m._FPDF_ClosePage(page)
          }
        }
      } finally {
        ptrs.forEach(ptr => m._free(ptr))
      }
      return out
    })
  })
}

/** Physically remove whole, axis-aligned content-stream paths.
 *
 * This is intentionally not a paint-over API. Every overlapping object must be the
 * one selected path, and all other object classes, partial intersections, rotation,
 * missing geometry, and ambiguous selections fail closed before serialization.
 */
export async function applyWholePathRedactions(bytes: Uint8Array, redactions: readonly WholePathRedaction[]): Promise<PathRedactionResult> {
  if (!redactions.length) return { bytes, proofs: [] }
  const planned = redactions.map(redaction => ({ ...redaction, target: normalized(redaction.rect) }))
  const removed = await chainPdfium(async () => {
    const m = await loadPdfium()
    return withDocument(m, bytes, async doc => {
      const removedRects: Rect[] = []
      for (const redaction of planned) {
        const pageIndex = redaction.page - 1
        if (!Number.isInteger(redaction.page) || pageIndex < 0 || pageIndex >= m._FPDF_GetPageCount(doc)) {
          throw new Error(`redaction page ${redaction.page} does not exist; no bytes changed`)
        }
        const page = m._FPDF_LoadPage(doc, pageIndex)
        if (!page) throw new Error(`could not load redaction page ${redaction.page}; no bytes changed`)
        const ptrs = [m._malloc(4), m._malloc(4), m._malloc(4), m._malloc(4)]
        try {
          const matches: Array<{ obj: number; bounds: Rect }> = []
          for (let i = 0; i < m._FPDFPage_CountObjects(page); i++) {
            const obj = m._FPDFPage_GetObject(page, i)
            const bounds = boundsOf(m, obj, ptrs)
            if (!overlaps(bounds, redaction.target)) continue
            if (!contains(redaction.target, bounds)) {
              throw new Error('redaction intersects an object boundary; no bytes changed')
            }
            if (m._FPDFPageObj_GetType(obj) !== FPDF_PAGEOBJ_PATH) {
              throw new Error('redaction covers a non-path object; no bytes changed')
            }
            if (!isAxisAligned(m, obj)) {
              throw new Error('redaction covers a rotated or sheared path; no bytes changed')
            }
            matches.push({ obj, bounds })
          }
          if (matches.length !== 1) {
            throw new Error(matches.length ? 'redaction is ambiguous: it contains multiple paths; no bytes changed' : 'redaction contains no eligible path object; no bytes changed')
          }
          const match = matches[0]!
          if (!m._FPDFPage_RemoveObject(page, match.obj)) throw new Error('PDFium could not remove redacted path')
          m._FPDFPageObj_Destroy(match.obj)
          if (!m._FPDFPage_GenerateContent(page)) throw new Error('PDFium could not regenerate redacted page content')
          removedRects.push(match.bounds)
        } finally {
          ptrs.forEach(ptr => m._free(ptr))
          m._FPDF_ClosePage(page)
        }
      }
      return { bytes: saveDoc(m, doc), removedRects }
    })
  })

  // Fresh document opens are deliberately outside the mutation chain. They prove
  // both that serialization succeeded and that the chosen object is not retained in
  // the saved content stream. A page render then proves fresh output remains usable.
  const remaining = await listPagePaths(removed.bytes)
  const proofs: PathRedactionProof[] = []
  for (const [index, redaction] of planned.entries()) {
    const target = redaction.target
    if (remaining.some(path => path.page === redaction.page && contains(target, path.rect))) {
      throw new Error('path-object proof failed; output discarded')
    }
    const rendered = await renderPageToPng(removed.bytes, redaction.page)
    if (!rendered.png.length) throw new Error('raster proof failed; output discarded')
    proofs.push({ page: redaction.page, removedRect: removed.removedRects[index]!, renderedBytes: rendered.png.length })
  }
  return { bytes: removed.bytes, proofs }
}
