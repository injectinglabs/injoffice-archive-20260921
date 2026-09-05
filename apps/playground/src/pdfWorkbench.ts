import { PDFArray, PDFCheckBox, PDFDict, PDFDocument, PDFDropdown, PDFHexString, PDFName, PDFNumber, PDFOptionList, PDFRadioGroup, PDFRef, PDFString, PDFTextField } from 'pdf-lib'
import { applyAnnotDeletes } from '../../../packages/pdf/src/annotate/annotDelete'
import { applyDrawings, applyNoteEdits } from '../../../packages/pdf/src/annotate/drawing'
import { applyFormValues } from '../../../packages/pdf/src/annotate/forms'
import { applyMarkups } from '../../../packages/pdf/src/annotate/markup'
import { applySignatureStamps, applyStamps, VISUAL_SIGNATURE_CONTENT_PREFIX } from '../../../packages/pdf/src/annotate/stamp'
import type { DrawingSpec, FormValueSpec, MarkupType } from '../../../packages/pdf/src/annotate/types'
import { applyPageOps, mergeDocuments, readInfo, splitDocument } from '../../../packages/pdf/src/pageOps'
import type { PageOpSpec, PdfDocumentInfo } from '../../../packages/pdf/src/types'

/** Minimal 1×1 PNG used for stamp/signature proofs. */
export const STAMP_PNG_B64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='

export type PdfAnnot = {
  page: number
  objNum: number
  subtype: MarkupType | 'note'
  rect: [number, number, number, number]
  contents?: string
}

export type PdfFormField = FormValueSpec

export async function pdfGeometry(bytes: Uint8Array): Promise<PdfDocumentInfo> {
  return readInfo(bytes)
}

export async function applyPdfPageOp(bytes: Uint8Array, op: PageOpSpec): Promise<Uint8Array> {
  return applyPageOps(bytes, [op])
}

export async function mergePdfBytes(base: Uint8Array, extra: Uint8Array): Promise<Uint8Array> {
  return mergeDocuments([base, extra])
}

export async function splitPdfPrefix(bytes: Uint8Array, toPage: number): Promise<Uint8Array> {
  const parts = await splitDocument(bytes, [{ from: 1, to: toPage }])
  const first = parts[0]
  if (!first) throw new Error('split produced no document')
  return first
}

function defaultQuad(width: number, height: number): number[] {
  const left = Math.min(72, width * 0.12)
  const right = Math.min(width - 36, left + Math.min(280, width * 0.55))
  const top = height - 72
  const bottom = top - 18
  return [left, top, right, top, left, bottom, right, bottom]
}

export async function applyPdfMarkup(bytes: Uint8Array, page: number, type: MarkupType): Promise<Uint8Array> {
  const info = await readInfo(bytes)
  const geometry = info.pages[page - 1]
  if (!geometry) throw new Error(`page ${page} is out of range`)
  return applyMarkups(bytes, [{ page, type, color: type === 'highlight' ? [1, 0.92, 0.2] : [0.15, 0.2, 0.7], quads: [defaultQuad(geometry.width, geometry.height)] }])
}

export async function applyPdfDrawing(bytes: Uint8Array, page: number, kind: Exclude<DrawingSpec['kind'], 'note'>): Promise<Uint8Array> {
  const info = await readInfo(bytes)
  const geometry = info.pages[page - 1]
  if (!geometry) throw new Error(`page ${page} is out of range`)
  const color: [number, number, number] = [0.12, 0.35, 0.72]
  const width = 1.5
  const pad = 48
  const rect: [number, number, number, number] = [pad, pad, Math.min(geometry.width - pad, pad + 180), Math.min(geometry.height - pad, pad + 90)]
  const drawing: DrawingSpec =
    kind === 'rect' ? { kind, page, color, width, rect }
    : kind === 'ellipse' ? { kind, page, color, width, rect }
    : kind === 'line' ? { kind, page, color, width, from: [rect[0], rect[1]], to: [rect[2], rect[3]] }
    : kind === 'arrow' ? { kind, page, color, width, from: [rect[0], rect[3]], to: [rect[2], rect[1]] }
    : { kind: 'ink', page, color, width, paths: [[rect[0], rect[1], (rect[0] + rect[2]) / 2, rect[3], rect[2], rect[1]]] }
  return applyDrawings(bytes, [drawing])
}

export async function applyPdfNote(bytes: Uint8Array, page: number, contents: string): Promise<Uint8Array> {
  const info = await readInfo(bytes)
  const geometry = info.pages[page - 1]
  if (!geometry) throw new Error(`page ${page} is out of range`)
  return applyDrawings(bytes, [{
    kind: 'note',
    page,
    color: [1, 0.85, 0.2],
    at: [Math.max(36, geometry.width - 72), Math.max(72, geometry.height - 72)],
    contents,
    author: 'InjOffice playground',
  }])
}

export async function applyPdfStamp(bytes: Uint8Array, page: number, signature: boolean): Promise<Uint8Array> {
  const rect: [number, number, number, number] = [36, 36, 108, 72]
  if (signature) return applySignatureStamps(bytes, [{ page, image: STAMP_PNG_B64, rect }])
  return applyStamps(bytes, [{ page, image: STAMP_PNG_B64, rect, opacity: 0.85 }])
}

export { VISUAL_SIGNATURE_CONTENT_PREFIX }

export async function applyPdfFormValues(bytes: Uint8Array, values: FormValueSpec[]): Promise<Uint8Array> {
  const result = await applyFormValues(bytes, values)
  return result.bytes
}

export async function applyPdfAnnotDelete(bytes: Uint8Array, annot: PdfAnnot): Promise<Uint8Array> {
  return applyAnnotDeletes(bytes, [{
    page: annot.page,
    objNum: annot.objNum,
    subtype: annot.subtype,
    rect: annot.rect,
    contents: annot.contents,
  }])
}

export async function applyPdfNoteEdit(bytes: Uint8Array, annot: PdfAnnot, contents: string): Promise<Uint8Array> {
  return applyNoteEdits(bytes, [{
    page: annot.page,
    objNum: annot.objNum,
    rect: annot.rect,
    oldContents: annot.contents ?? '',
    contents,
  }])
}

function rectOf(dict: PDFDict): [number, number, number, number] | null {
  const rect = dict.lookupMaybe(PDFName.of('Rect'), PDFArray)
  if (!rect || rect.size() !== 4) return null
  const values = [0, 1, 2, 3].map((index) => rect.lookupMaybe(index, PDFNumber)?.asNumber() ?? NaN)
  if (values.some((value) => !Number.isFinite(value))) return null
  return values as [number, number, number, number]
}

function contentsOf(dict: PDFDict): string | undefined {
  const value = dict.lookup(PDFName.of('Contents'))
  if (value instanceof PDFString || value instanceof PDFHexString) return value.decodeText()
  return undefined
}

export async function listPdfAnnots(bytes: Uint8Array): Promise<PdfAnnot[]> {
  const doc = await PDFDocument.load(bytes)
  const annots: PdfAnnot[] = []
  const subtypeMap: Record<string, PdfAnnot['subtype']> = {
    Highlight: 'highlight',
    Underline: 'underline',
    StrikeOut: 'strikeout',
    Text: 'note',
  }
  doc.getPages().forEach((page, pageIndex) => {
    const arr = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
    if (!arr) return
    for (let i = 0; i < arr.size(); i++) {
      const ref = arr.get(i)
      if (!(ref instanceof PDFRef)) continue
      const dict = doc.context.lookupMaybe(ref, PDFDict)
      if (!dict) continue
      const subtype = dict.lookupMaybe(PDFName.of('Subtype'), PDFName)
      const name = subtype?.asString().replace(/^\//, '') ?? ''
      const mapped = subtypeMap[name]
      const rect = rectOf(dict)
      if (!mapped || !rect) continue
      annots.push({ page: pageIndex + 1, objNum: ref.objectNumber, subtype: mapped, rect, contents: contentsOf(dict) })
    }
  })
  return annots
}

export async function listPdfFormFields(bytes: Uint8Array): Promise<PdfFormField[]> {
  const doc = await PDFDocument.load(bytes)
  const fields: PdfFormField[] = []
  for (const field of doc.getForm().getFields()) {
    const name = field.getName()
    if (field instanceof PDFTextField) fields.push({ name, kind: 'text', value: field.getText() ?? '' })
    else if (field instanceof PDFCheckBox) fields.push({ name, kind: 'checkbox', checked: field.isChecked() })
    else if (field instanceof PDFRadioGroup) fields.push({ name, kind: 'radio', value: field.getSelected() ?? '' })
    else if (field instanceof PDFDropdown || field instanceof PDFOptionList) fields.push({ name, kind: 'choice', value: field.getSelected()?.[0] ?? '' })
  }
  return fields
}

export function movePageOrder(pageCount: number, page: number, delta: -1 | 1): number[] | null {
  if (page < 1 || page > pageCount) return null
  const order = Array.from({ length: pageCount }, (_, index) => index + 1)
  const index = page - 1
  const next = index + delta
  if (next < 0 || next >= pageCount) return null
  ;[order[index], order[next]] = [order[next]!, order[index]!]
  return order
}
