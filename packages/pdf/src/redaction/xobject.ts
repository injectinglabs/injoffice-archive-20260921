import { PDFArray, PDFDict, PDFDocument, PDFObject, PDFName, PDFRawStream, PDFRef, PDFStream } from 'pdf-lib'
import { renderPageToPng } from '../ocr/renderPage.js'
import { PdfViewerDocument } from '../viewer.js'
import { chainPdfium, FPDF_PAGEOBJ_FORM, loadPdfium, saveDoc, withDocument, type Pdfium } from '../textEdit/pdfium.js'
import type { PdfRect } from './path.js'

/** A Form XObject addressed by its exact page resource name.  This is not a
 * visual overlay API: callers must explicitly acknowledge deletion of the
 * complete Form XObject and provide its complete listed footprint. */
export interface WholeXObjectRedaction {
  page: number
  name: string
  rect: PdfRect
  wholeXObject: true
}

export interface PageXObjectRef {
  page: number
  name: string
  objectNumber: number
  rect: [number, number, number, number]
  axisAligned: boolean
}

export interface XObjectRedactionProof {
  page: number
  name: string
  removedObjectNumber: number
  renderedBytes: number
  extractedCharacters: number
}

export interface XObjectRedactionResult { bytes: Uint8Array; proofs: XObjectRedactionProof[] }

type Rect = [number, number, number, number]
type Planned = { page: number; name: string; rect: Rect; ref: PDFRef; objectNumber: number; streamSignature: string }
const EPSILON = 1e-4
const key = (ref: PDFRef): string => `${ref.objectNumber}:${ref.generationNumber}`

function normalized(rect: PdfRect): Rect {
  const [x1, y1, x2, y2] = rect
  if (![x1, y1, x2, y2].every(Number.isFinite)) throw new Error('redaction rectangle must contain finite coordinates')
  const out: Rect = [Math.min(x1, x2), Math.min(y1, y2), Math.max(x1, x2), Math.max(y1, y2)]
  if (out[0] >= out[2] || out[1] >= out[3]) throw new Error('redaction rectangle must have positive area')
  return out
}
function overlaps(a: Rect, b: Rect): boolean { return Math.min(a[2], b[2]) > Math.max(a[0], b[0]) + EPSILON && Math.min(a[3], b[3]) > Math.max(a[1], b[1]) + EPSILON }
function contains(outer: Rect, inner: Rect): boolean { return inner[0] >= outer[0] - EPSILON && inner[1] >= outer[1] - EPSILON && inner[2] <= outer[2] + EPSILON && inner[3] <= outer[3] + EPSILON }
function sameRect(a: Rect, b: Rect): boolean { return a.every((value, index) => Math.abs(value - b[index]!) <= EPSILON) }

function boundsOf(m: Pdfium, obj: number, ptrs: readonly number[]): Rect {
  if (!m._FPDFPageObj_GetBounds(obj, ptrs[0]!, ptrs[1]!, ptrs[2]!, ptrs[3]!)) throw new Error('cannot prove Form XObject bounds; no bytes changed')
  return normalized([m.HEAPF32[ptrs[0]! >> 2]!, m.HEAPF32[ptrs[1]! >> 2]!, m.HEAPF32[ptrs[2]! >> 2]!, m.HEAPF32[ptrs[3]! >> 2]!])
}
function isAxisAligned(m: Pdfium, obj: number): boolean {
  const matrix = m._malloc(24)
  try {
    if (!m._FPDFPageObj_GetMatrix(obj, matrix)) return false
    const [a, b, c, d] = [0, 1, 2, 3].map(i => m.HEAPF32[(matrix >> 2) + i]!)
    return [a, b, c, d].every(Number.isFinite) && Math.abs(a) > EPSILON && Math.abs(d) > EPSILON && Math.abs(b) <= EPSILON && Math.abs(c) <= EPSILON
  } finally { m._free(matrix) }
}

function xobjectsForPage(doc: PDFDocument, pageNumber: number): PDFDict {
  const page = doc.getPages()[pageNumber - 1]
  if (!page) throw new Error(`redaction page ${pageNumber} does not exist; no bytes changed`)
  const resources = page.node.Resources()
  if (!resources) throw new Error(`page ${pageNumber} has no resources; no bytes changed`)
  const xobjects = resources.lookupMaybe(PDFName.of('XObject'), PDFDict)
  if (!xobjects) throw new Error(`page ${pageNumber} has no XObject resources; no bytes changed`)
  return xobjects
}

function formResource(doc: PDFDocument, page: number, name: string): PDFRef {
  if (!name || name.startsWith('/') || !/^[A-Za-z0-9_.-]+$/.test(name)) throw new Error('XObject name must be an exact non-empty PDF resource name')
  const value = xobjectsForPage(doc, page).get(PDFName.of(name))
  if (!(value instanceof PDFRef)) throw new Error(`XObject ${JSON.stringify(name)} must be an indirect resource; no bytes changed`)
  const stream = doc.context.lookup(value) as unknown
  if (!(stream instanceof PDFRawStream) || stream.dict.lookupMaybe(PDFName.of('Subtype'), PDFName)?.toString() !== '/Form') throw new Error(`XObject ${JSON.stringify(name)} is not an eligible Form XObject; no bytes changed`)
  return value
}

function rawStreamSignature(stream: PDFRawStream): string { return Buffer.from(stream.getContents()).toString('base64') }

/** A redacted resource may not have another direct owner.  Keeping a shared Form
 * stream would leave recoverable content; deleting it would corrupt the other
 * owner. Both cases fail closed. */
function assertUnshared(doc: PDFDocument, target: PDFRef): void {
  let directOwners = 0
  for (const [, object] of doc.context.enumerateIndirectObjects()) {
    directOwners += directReferences(object, target, new Set<PDFObject>())
  }
  if (directOwners !== 1) throw new Error(`Form XObject ${target.toString()} is shared or has unsupported ownership; no bytes changed`)
}

/** Count edges to a ref without following refs. A page's /Resources is often a
 * direct dictionary nested inside the indirect page dictionary, so looking only
 * at top-level values would incorrectly treat the target as ownerless. */
function directReferences(object: PDFObject, target: PDFRef, seen: Set<PDFObject>): number {
  if (seen.has(object)) return 0
  seen.add(object)
  if (object instanceof PDFRef) return key(object) === key(target) ? 1 : 0
  if (object instanceof PDFArray) {
    let count = 0
    for (let i = 0; i < object.size(); i++) count += directReferences(object.get(i), target, seen)
    return count
  }
  if (object instanceof PDFDict) {
    let count = 0
    for (const [, value] of object.entries()) count += directReferences(value, target, seen)
    return count
  }
  if (object instanceof PDFStream) return directReferences(object.dict, target, seen)
  return 0
}

async function planAsync(bytes: Uint8Array, specs: readonly WholeXObjectRedaction[]): Promise<Planned[]> {
  const doc = await PDFDocument.load(bytes)
  const seen = new Set<string>()
  return specs.map(spec => {
    if (spec.wholeXObject !== true) throw new Error('XObject redaction must explicitly acknowledge wholeXObject: true')
    if (!Number.isInteger(spec.page) || spec.page < 1) throw new Error(`redaction page ${spec.page} does not exist; no bytes changed`)
    const ref = formResource(doc, spec.page, spec.name)
    assertUnshared(doc, ref)
    const stream = doc.context.lookup(ref) as unknown as PDFRawStream
    const streamSignature = rawStreamSignature(stream)
    const identicalStreams = doc.context.enumerateIndirectObjects().filter(([, object]) => object instanceof PDFRawStream && rawStreamSignature(object) === streamSignature)
    if (identicalStreams.length !== 1) throw new Error(`Form XObject ${JSON.stringify(spec.name)} has a duplicate stream; no bytes changed`)
    const identity = `${spec.page}:${spec.name}`
    if (seen.has(identity)) throw new Error(`XObject ${JSON.stringify(spec.name)} was selected more than once`)
    seen.add(identity)
    return { page: spec.page, name: spec.name, rect: normalized(spec.rect), ref, objectNumber: ref.objectNumber, streamSignature }
  })
}

async function removeOnePdfium(bytes: Uint8Array, target: Planned): Promise<Uint8Array> {
  return chainPdfium(async () => {
    const m = await loadPdfium()
    return withDocument(m, bytes, async doc => {
      if (target.page > m._FPDF_GetPageCount(doc)) throw new Error(`redaction page ${target.page} does not exist; no bytes changed`)
      const page = m._FPDF_LoadPage(doc, target.page - 1)
      if (!page) throw new Error(`could not load redaction page ${target.page}; no bytes changed`)
      const ptrs = [m._malloc(4), m._malloc(4), m._malloc(4), m._malloc(4)]
      try {
        const matches: Array<{ obj: number; bounds: Rect }> = []
        for (let i = 0; i < m._FPDFPage_CountObjects(page); i++) {
          const obj = m._FPDFPage_GetObject(page, i)
          const bounds = boundsOf(m, obj, ptrs)
          if (!overlaps(bounds, target.rect)) continue
          if (!contains(target.rect, bounds)) throw new Error('redaction intersects an object boundary; no bytes changed')
          if (m._FPDFPageObj_GetType(obj) !== FPDF_PAGEOBJ_FORM) throw new Error('redaction covers a non-Form object; no bytes changed')
          if (!isAxisAligned(m, obj) || !sameRect(bounds, target.rect)) throw new Error('redaction must exactly match one axis-aligned Form XObject footprint; no bytes changed')
          matches.push({ obj, bounds })
        }
        if (matches.length !== 1) throw new Error(matches.length ? 'redaction is ambiguous: it contains multiple Form XObjects; no bytes changed' : 'redaction contains no eligible Form XObject; no bytes changed')
        if (!m._FPDFPage_RemoveObject(page, matches[0]!.obj)) throw new Error('PDFium could not remove redacted Form XObject')
        m._FPDFPageObj_Destroy(matches[0]!.obj)
        if (!m._FPDFPage_GenerateContent(page)) throw new Error('PDFium could not regenerate redacted page content')
        return saveDoc(m, doc)
      } finally { ptrs.forEach(ptr => m._free(ptr)); m._FPDF_ClosePage(page) }
    })
  })
}

async function removeResourceAndVerify(bytes: Uint8Array, target: Planned): Promise<{ bytes: Uint8Array; proof: XObjectRedactionProof }> {
  const doc = await PDFDocument.load(bytes)
  const xobjects = xobjectsForPage(doc, target.page)
  const freshRef = xobjects.get(PDFName.of(target.name))
  if (freshRef !== undefined && !(freshRef instanceof PDFRef)) throw new Error('verification failed: XObject resource is no longer indirect')
  if (freshRef instanceof PDFRef) {
    assertUnshared(doc, freshRef)
    xobjects.delete(PDFName.of(target.name))
    doc.context.delete(freshRef)
  }
  const out = await doc.save()
  const reloaded = await PDFDocument.load(out)
  const remaining = reloaded.getPages()[target.page - 1]!.node.Resources()?.lookupMaybe(PDFName.of('XObject'), PDFDict)?.get(PDFName.of(target.name))
  if (remaining) throw new Error('verification failed: redacted XObject resource remains')
  if (reloaded.context.enumerateIndirectObjects().some(([, object]) => object instanceof PDFRawStream && rawStreamSignature(object) === target.streamSignature)) throw new Error('verification failed: redacted XObject stream remains serialized')
  const viewer = await PdfViewerDocument.load(out)
  try {
    await viewer.getPage(target.page)
    const extracted = await viewer.getPageText(target.page)
    const rendered = await renderPageToPng(out, target.page)
    if (!rendered.png.length) throw new Error('verification failed: redacted page did not rasterize')
    return { bytes: out, proof: { page: target.page, name: target.name, removedObjectNumber: target.objectNumber, renderedBytes: rendered.png.length, extractedCharacters: extracted.length } }
  } finally { await viewer.destroy() }
}

/** List only Form XObjects that have a unique, unshared resource owner and a
 * single axis-aligned PDFium footprint. Ambiguous or unsupported resources are
 * intentionally omitted so they cannot be selected for destructive redaction. */
export async function listPageXObjects(bytes: Uint8Array): Promise<PageXObjectRef[]> {
  const doc = await PDFDocument.load(bytes)
  const candidates: Array<{ page: number; name: string; ref: PDFRef }> = []
  for (let page = 1; page <= doc.getPageCount(); page++) {
    let resources: PDFDict
    try { resources = xobjectsForPage(doc, page) } catch { continue }
    for (const [name, value] of resources.entries()) {
      if (!(value instanceof PDFRef)) continue
      const stream = doc.context.lookup(value) as unknown
      if (!(stream instanceof PDFRawStream) || stream.dict.lookupMaybe(PDFName.of('Subtype'), PDFName)?.toString() !== '/Form') continue
      try { assertUnshared(doc, value) } catch { continue }
      candidates.push({ page, name: name.toString().slice(1), ref: value })
    }
  }
  return chainPdfium(async () => {
    const m = await loadPdfium()
    return withDocument(m, bytes, async handle => {
      const out: PageXObjectRef[] = []
      for (const candidate of candidates) {
        const page = m._FPDF_LoadPage(handle, candidate.page - 1)
        if (!page) continue
        const ptrs = [m._malloc(4), m._malloc(4), m._malloc(4), m._malloc(4)]
        try {
          const forms: Rect[] = []
          for (let i = 0; i < m._FPDFPage_CountObjects(page); i++) {
            const obj = m._FPDFPage_GetObject(page, i)
            if (m._FPDFPageObj_GetType(obj) === FPDF_PAGEOBJ_FORM && isAxisAligned(m, obj)) forms.push(boundsOf(m, obj, ptrs))
          }
          if (forms.length === 1) out.push({ page: candidate.page, name: candidate.name, objectNumber: candidate.ref.objectNumber, rect: forms[0]!, axisAligned: true })
        } finally { ptrs.forEach(ptr => m._free(ptr)); m._FPDF_ClosePage(page) }
      }
      return out
    })
  })
}

/** Physically removes exactly one fully selected, unshared Form XObject at a
 * time, then removes its resource and indirect stream. Every stage fresh-opens,
 * structurally checks, extracts, and rasterizes output before returning bytes. */
export async function applyWholeXObjectRedactions(bytes: Uint8Array, redactions: readonly WholeXObjectRedaction[]): Promise<XObjectRedactionResult> {
  if (!redactions.length) return { bytes, proofs: [] }
  let out = bytes
  const proofs: XObjectRedactionProof[] = []
  for (const spec of redactions) {
    const [target] = await planAsync(out, [spec])
    out = await removeOnePdfium(out, target!)
    const proven = await removeResourceAndVerify(out, target!)
    out = proven.bytes
    proofs.push(proven.proof)
  }
  return { bytes: out, proofs }
}
