import { PDFDocument } from 'pdf-lib'
import type { MarkupSpec } from './types.js'
import {
  addAnnotation,
  appearanceStream,
  canonicalRect,
  checkedPage,
  colorTuple,
  pdfDate,
  rectFromPoints,
  setNormalAppearance,
} from './pdfObjects.js'

const subtype = { highlight: 'Highlight', underline: 'Underline', strikeout: 'StrikeOut' } as const
const format = (value: number): string => Number(value.toFixed(5)).toString()

function markupAppearance(doc: PDFDocument, spec: MarkupSpec, rect: [number, number, number, number]): ReturnType<typeof appearanceStream> {
  const color = colorTuple(spec.color)
  const commands: string[] = ['q']
  if (spec.type === 'highlight') commands.push('/GS gs', `${color.map(format).join(' ')} rg`)
  else commands.push(`${color.map(format).join(' ')} RG`, '1 w')

  for (const quad of spec.quads) {
    if (quad.length !== 8 || quad.some((value) => !Number.isFinite(value))) {
      throw new TypeError('each markup quad must contain eight finite numbers')
    }
    const p = quad.map((value, index) => value - rect[index % 2])
    if (spec.type === 'highlight') {
      const xs = [p[0], p[2], p[4], p[6]]
      const ys = [p[1], p[3], p[5], p[7]]
      const minX = Math.min(...xs)
      const maxX = Math.max(...xs)
      const minY = Math.min(...ys)
      const maxY = Math.max(...ys)
      const axisAligned = xs.every((value) => value === minX || value === maxX) && ys.every((value) => value === minY || value === maxY)
      if (axisAligned) commands.push(`${format(minX)} ${format(minY)} ${format(maxX - minX)} ${format(maxY - minY)} re f`)
      else commands.push(
        `${format(p[0])} ${format(p[1])} m`,
        `${format(p[2])} ${format(p[3])} l`,
        `${format(p[6])} ${format(p[7])} l`,
        `${format(p[4])} ${format(p[5])} l h f`,
      )
    } else {
      const startX = spec.type === 'underline' ? p[4] : (p[0] + p[4]) / 2
      const startY = spec.type === 'underline' ? p[5] : (p[1] + p[5]) / 2
      const endX = spec.type === 'underline' ? p[6] : (p[2] + p[6]) / 2
      const endY = spec.type === 'underline' ? p[7] : (p[3] + p[7]) / 2
      commands.push(`${format(startX)} ${format(startY)} m ${format(endX)} ${format(endY)} l S`)
    }
  }
  commands.push('Q')

  const resources = spec.type === 'highlight'
    ? doc.context.obj({ ExtGState: { GS: { Type: 'ExtGState', ca: 0.35, BM: 'Multiply' } } })
    : undefined
  return appearanceStream(doc, rect, commands.join('\n'), resources)
}

export async function applyMarkups(bytes: Uint8Array, markups: MarkupSpec[]): Promise<Uint8Array> {
  if (markups.length === 0) return bytes
  const doc = await PDFDocument.load(bytes)
  for (const spec of markups) {
    const page = checkedPage(doc, spec.page)
    if (spec.quads.length === 0) throw new TypeError('markup must contain at least one quad')
    const flattened = spec.quads.flat()
    const rect = canonicalRect(rectFromPoints(flattened))
    const annotation = addAnnotation(doc, page, {
      Subtype: subtype[spec.type],
      Rect: rect,
      QuadPoints: flattened,
      C: colorTuple(spec.color),
      F: 4,
      M: pdfDate(),
    })
    setNormalAppearance(doc, annotation, markupAppearance(doc, spec, rect))
  }
  return doc.save()
}
