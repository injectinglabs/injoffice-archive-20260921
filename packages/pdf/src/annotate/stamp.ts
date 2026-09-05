import { PDFDocument } from 'pdf-lib'
import type { SignatureStampSpec, StampSpec } from './types.js'
import {
  addAnnotation,
  appearanceStream,
  canonicalRect,
  checkedPage,
  pdfDate,
  pdfText,
  setNormalAppearance,
  type Rect,
} from './pdfObjects.js'

export const VISUAL_SIGNATURE_CONTENT_PREFIX = 'InjOffice visual signature field: '

const format = (value: number): string => Number(value.toFixed(5)).toString()

function imageMatrix(rect: Rect, counterRotation: number): [number, number, number, number, number, number] {
  const width = rect[2] - rect[0]
  const height = rect[3] - rect[1]
  const turns = ((Math.round(counterRotation / 90) % 4) + 4) % 4
  if (turns === 1) return [0, -height, width, 0, 0, height]
  if (turns === 2) return [-width, 0, 0, -height, width, height]
  if (turns === 3) return [0, height, -width, 0, width, 0]
  return [width, 0, 0, height, 0, 0]
}

export async function applyStamps(bytes: Uint8Array, stamps: StampSpec[]): Promise<Uint8Array> {
  if (stamps.length === 0) return bytes
  const doc = await PDFDocument.load(bytes)
  for (const spec of stamps) {
    const page = checkedPage(doc, spec.page)
    const rect = canonicalRect(spec.rect)
    const image = await doc.embedPng(spec.image)
    page.drawImage(image, {
      x: rect[0],
      y: rect[1],
      width: rect[2] - rect[0],
      height: rect[3] - rect[1],
      opacity: spec.opacity === undefined ? undefined : Math.max(0, Math.min(1, spec.opacity)),
    })
  }
  return doc.save()
}

export async function applySignatureStamps(bytes: Uint8Array, stamps: SignatureStampSpec[]): Promise<Uint8Array> {
  if (stamps.length === 0) return bytes
  const doc = await PDFDocument.load(bytes)
  for (const spec of stamps) {
    const page = checkedPage(doc, spec.page)
    const rect = canonicalRect(spec.rect)
    const image = await doc.embedPng(spec.image)
    const resources = doc.context.obj({ XObject: { Im0: image.ref } })
    const matrix = imageMatrix(rect, -page.getRotation().angle)
    const appearance = appearanceStream(doc, rect, `q\n${matrix.map(format).join(' ')} cm\n/Im0 Do\nQ`, resources)
    const annotation = addAnnotation(doc, page, {
      Subtype: 'Stamp',
      Rect: rect,
      Name: 'Approved',
      Contents: spec.formFieldName ? pdfText(`${VISUAL_SIGNATURE_CONTENT_PREFIX}${spec.formFieldName}`) : undefined,
      InjOfficeFormField: spec.formFieldName ? pdfText(spec.formFieldName) : undefined,
      F: 4,
      M: pdfDate(),
    })
    setNormalAppearance(doc, annotation, appearance)
  }
  return doc.save()
}
