import { PDFDocument } from 'pdf-lib'
import type { AnnotDeleteSpec } from './types.js'
import {
  annotationEntries,
  canonicalRect,
  nameValue,
  numberArray,
  sameNumbers,
  textValue,
} from './pdfObjects.js'

const subtypes = { highlight: 'Highlight', underline: 'Underline', strikeout: 'StrikeOut', note: 'Text' } as const

function confirmsIdentity(entry: ReturnType<typeof annotationEntries>[number], spec: AnnotDeleteSpec): boolean {
  if (nameValue(entry.dict, 'Subtype') !== subtypes[spec.subtype]) return false
  if (!sameNumbers(numberArray(entry.dict, 'Rect'), canonicalRect(spec.rect))) return false
  if (spec.subtype === 'note') return spec.contents !== undefined && textValue(entry.dict, 'Contents') === spec.contents
  return true
}

export async function applyAnnotDeletes(bytes: Uint8Array, deletes: AnnotDeleteSpec[]): Promise<Uint8Array> {
  if (deletes.length === 0) return bytes
  const doc = await PDFDocument.load(bytes)
  let changed = false
  for (const spec of deletes) {
    if (!Number.isInteger(spec.page) || spec.page < 1 || spec.page > doc.getPageCount()) continue
    const page = doc.getPage(spec.page - 1)
    const entries = annotationEntries(page)
    const hinted = entries.find((entry) => entry.ref?.objectNumber === spec.objNum && confirmsIdentity(entry, spec))
    const match = hinted ?? entries.find((entry) => confirmsIdentity(entry, spec))
    if (!match) continue
    page.node.Annots()?.remove(match.index)
    if (match.ref) doc.context.delete(match.ref)
    changed = true
  }
  return changed ? doc.save() : bytes
}
