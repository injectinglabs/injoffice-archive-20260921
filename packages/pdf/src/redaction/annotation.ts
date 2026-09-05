import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFNumber, PDFRef, PDFString } from 'pdf-lib'
import { applyAnnotDeletes, type AnnotDeleteSpec } from '../annotate/index.js'
import { PdfViewerDocument } from '../viewer.js'

const N = {
  Annots: PDFName.of('Annots'),
  Contents: PDFName.of('Contents'),
  Rect: PDFName.of('Rect'),
  Subtype: PDFName.of('Subtype'),
} as const

const eligibleSubtypes = new Set<AnnotDeleteSpec['subtype']>(['highlight', 'underline', 'strikeout', 'note'])

/** Exact saved identity accepted by applyWholeAnnotationRedactions. Contents
 * is always present, including as an empty string, because a missing contents
 * constraint can widen matching for comments in one note thread. */
export interface WholeAnnotationRedactionTarget {
  page: number
  objNum: number
  subtype: AnnotDeleteSpec['subtype']
  rect: [number, number, number, number]
  contents: string
}

function rectOf(dict: PDFDict): [number, number, number, number] | null {
  const rect = dict.lookupMaybe(N.Rect, PDFArray)
  if (!rect || rect.size() !== 4) return null
  const values: number[] = []
  for (let i = 0; i < 4; i++) {
    const value = rect.lookup(i, PDFNumber)
    if (!value || !Number.isFinite(value.asNumber())) return null
    values.push(value.asNumber())
  }
  const [x1, y1, x2, y2] = values as [number, number, number, number]
  if (Math.min(x1, x2) >= Math.max(x1, x2) || Math.min(y1, y2) >= Math.max(y1, y2)) return null
  return [x1, y1, x2, y2]
}

function contentsOf(dict: PDFDict): string | null {
  const contents = dict.get(N.Contents)
  if (!contents) return ''
  if (contents instanceof PDFString || contents instanceof PDFHexString) return contents.decodeText()
  return null
}

/**
 * Enumerates exact, currently supported saved annotation identities without
 * mutating bytes. Only indirect highlight/underline/strikeout/Text-note
 * annotations with finite positive /Rect and decodable /Contents are returned.
 * Unsupported annotation types and direct/malformed entries are withheld: no
 * rectangle-only or contentless fallback can reach a destructive operation.
 */
export async function listWholeAnnotationRedactionTargets(bytes: Uint8Array): Promise<WholeAnnotationRedactionTarget[]> {
  const doc = await PDFDocument.load(bytes)
  const out: WholeAnnotationRedactionTarget[] = []
  for (const [pageIndex, page] of doc.getPages().entries()) {
    const annots = page.node.lookupMaybe(N.Annots, PDFArray)
    if (!annots) continue
    for (let i = 0; i < annots.size(); i++) {
      const ref = annots.get(i)
      if (!(ref instanceof PDFRef)) continue
      const dict = doc.context.lookupMaybe(ref, PDFDict)
      const rawSubtype = dict?.lookupMaybe(N.Subtype, PDFName)?.asString().replace(/^\//, '').toLowerCase()
      const subtype = rawSubtype === 'text' ? 'note' : rawSubtype
      if (!dict || !subtype || !eligibleSubtypes.has(subtype as AnnotDeleteSpec['subtype'])) continue
      const rect = rectOf(dict)
      const contents = contentsOf(dict)
      if (!rect || contents === null) continue
      out.push({ page: pageIndex + 1, objNum: ref.objectNumber, subtype: subtype as AnnotDeleteSpec['subtype'], rect, contents })
    }
  }
  return out
}

/** Exact saved annotation identity only. A no-op delete is a hard failure. */
export async function applyWholeAnnotationRedactions(bytes: Uint8Array, annotations: readonly AnnotDeleteSpec[]): Promise<Uint8Array> {
  let out = bytes
  for (const annotation of annotations) {
    const next = await applyAnnotDeletes(out, [annotation])
    if (next === out) throw new Error('annotation identity did not match; no bytes changed')
    // Fresh identity scan: a second exact delete must be a no-op. If it can
    // still delete, identity was ambiguous and the first output is discarded.
    if (await applyAnnotDeletes(next, [annotation]) !== next) throw new Error('annotation identity is ambiguous; output discarded')
    const viewer = await PdfViewerDocument.load(next)
    try { await viewer.getPage(annotation.page) } finally { await viewer.destroy() }
    out = next
  }
  return out
}
