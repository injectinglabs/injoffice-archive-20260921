import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFHexString,
  PDFName,
  PDFNumber,
  PDFRef,
  PDFString,
  type PDFPage,
} from 'pdf-lib'

export type Rect = [number, number, number, number]

export const pdfName = (value: string): PDFName => PDFName.of(value)

export function checkedPage(doc: PDFDocument, pageNumber: number): PDFPage {
  if (!Number.isInteger(pageNumber) || pageNumber < 1 || pageNumber > doc.getPageCount()) {
    throw new RangeError(`page ${pageNumber} does not exist`)
  }
  return doc.getPage(pageNumber - 1)
}

export function canonicalRect(rect: readonly number[]): Rect {
  if (rect.length !== 4 || rect.some((value) => !Number.isFinite(value))) {
    throw new TypeError('rectangle must contain four finite numbers')
  }
  return [
    Math.min(rect[0], rect[2]),
    Math.min(rect[1], rect[3]),
    Math.max(rect[0], rect[2]),
    Math.max(rect[1], rect[3]),
  ]
}

export function rectFromPoints(points: readonly number[], padding = 0): Rect {
  if (points.length < 2 || points.length % 2 !== 0 || points.some((value) => !Number.isFinite(value))) {
    throw new TypeError('point list must contain finite x/y pairs')
  }
  const xs: number[] = []
  const ys: number[] = []
  for (let index = 0; index < points.length; index += 2) {
    xs.push(points[index])
    ys.push(points[index + 1])
  }
  return [Math.min(...xs) - padding, Math.min(...ys) - padding, Math.max(...xs) + padding, Math.max(...ys) + padding]
}

export function colorTuple(color: readonly number[]): [number, number, number] {
  if (color.length !== 3 || color.some((value) => !Number.isFinite(value))) {
    throw new TypeError('color must contain three finite components')
  }
  return color.map((value) => Math.max(0, Math.min(1, value))) as [number, number, number]
}

export function addAnnotation(doc: PDFDocument, page: PDFPage, values: Record<string, unknown>): PDFRef {
  const definedValues = Object.fromEntries(Object.entries(values).filter(([, value]) => value !== undefined))
  const dict = doc.context.obj({ Type: 'Annot', ...definedValues })
  const ref = doc.context.register(dict)
  page.node.addAnnot(ref)
  return ref
}

export function appearanceStream(
  doc: PDFDocument,
  rect: Rect,
  content: string,
  resources?: PDFDict,
): PDFRef {
  const width = Math.max(0.001, rect[2] - rect[0])
  const height = Math.max(0.001, rect[3] - rect[1])
  const stream = doc.context.stream(content, {
    Type: 'XObject',
    Subtype: 'Form',
    FormType: 1,
    BBox: [0, 0, width, height],
    Resources: resources ?? doc.context.obj({}),
  })
  return doc.context.register(stream)
}

export function setNormalAppearance(doc: PDFDocument, annotation: PDFRef, appearance: PDFRef): void {
  const dict = doc.context.lookup(annotation, PDFDict)
  dict.set(pdfName('AP'), doc.context.obj({ N: appearance }))
}

export function annotationEntries(page: PDFPage): Array<{ index: number; ref?: PDFRef; dict: PDFDict }> {
  const annots = page.node.Annots()
  if (!annots) return []
  const entries: Array<{ index: number; ref?: PDFRef; dict: PDFDict }> = []
  for (let index = 0; index < annots.size(); index += 1) {
    const raw = annots.get(index)
    const dict = page.doc.context.lookupMaybe(raw, PDFDict)
    if (dict) entries.push({ index, ref: raw instanceof PDFRef ? raw : undefined, dict })
  }
  return entries
}

export function numberArray(dict: PDFDict, key: string): number[] | undefined {
  const value = dict.lookupMaybe(pdfName(key), PDFArray)
  if (!value) return undefined
  const output: number[] = []
  for (let index = 0; index < value.size(); index += 1) {
    const number = value.lookupMaybe(index, PDFNumber)
    if (!number) return undefined
    output.push(number.asNumber())
  }
  return output
}

export function nameValue(dict: PDFDict, key: string): string | undefined {
  return dict.lookupMaybe(pdfName(key), PDFName)?.decodeText()
}

export function textValue(dict: PDFDict, key: string): string | undefined {
  return dict.lookupMaybe(pdfName(key), PDFString, PDFHexString)?.decodeText()
}

export function sameNumbers(left: readonly number[] | undefined, right: readonly number[], tolerance = 0.01): boolean {
  return Boolean(left && left.length === right.length && left.every((value, index) => Math.abs(value - right[index]) <= tolerance))
}

export function pdfDate(ms = Date.now()): PDFString {
  return PDFString.fromDate(new Date(ms))
}

export function pdfText(value: string): PDFHexString {
  return PDFHexString.fromText(value)
}
