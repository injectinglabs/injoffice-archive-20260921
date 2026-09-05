import { PDFDict, PDFDocument, PDFName, type PDFRef } from 'pdf-lib'
import type { DrawingSpec, NoteEditSpec, NoteReplyTarget } from './types.js'
import {
  addAnnotation,
  annotationEntries,
  appearanceStream,
  canonicalRect,
  checkedPage,
  colorTuple,
  nameValue,
  numberArray,
  pdfDate,
  pdfName,
  pdfText,
  rectFromPoints,
  sameNumbers,
  setNormalAppearance,
  textValue,
  type Rect,
} from './pdfObjects.js'

const format = (value: number): string => Number(value.toFixed(5)).toString()
const stroke = (color: readonly number[], width: number): string => `${color.map(format).join(' ')} RG\n${format(Math.max(0.1, width))} w\n1 J 1 j`

function locateSavedNote(doc: PDFDocument, target: NoteReplyTarget): PDFRef | undefined {
  let identityMatch: PDFRef | undefined
  for (const page of doc.getPages()) {
    for (const entry of annotationEntries(page)) {
      if (
        entry.ref &&
        nameValue(entry.dict, 'Subtype') === 'Text' &&
        sameNumbers(numberArray(entry.dict, 'Rect'), canonicalRect(target.rect)) &&
        textValue(entry.dict, 'Contents') === target.contents
      ) {
        if (entry.ref.objectNumber === target.objNum) return entry.ref
        identityMatch ??= entry.ref
      }
    }
  }
  return identityMatch
}

function lineAppearance(
  doc: PDFDocument,
  rect: Rect,
  from: [number, number],
  to: [number, number],
  color: [number, number, number],
  width: number,
  arrow: boolean,
): PDFRef {
  const x1 = from[0] - rect[0]
  const y1 = from[1] - rect[1]
  const x2 = to[0] - rect[0]
  const y2 = to[1] - rect[1]
  const commands = [stroke(color, width), `${format(x1)} ${format(y1)} m ${format(x2)} ${format(y2)} l S`]
  if (arrow) {
    const angle = Math.atan2(y2 - y1, x2 - x1)
    const length = Math.max(6, width * 4)
    for (const offset of [Math.PI * 0.82, -Math.PI * 0.82]) {
      commands.push(`${format(x2)} ${format(y2)} m ${format(x2 + Math.cos(angle + offset) * length)} ${format(y2 + Math.sin(angle + offset) * length)} l S`)
    }
  }
  return appearanceStream(doc, rect, commands.join('\n'))
}

function shapeAppearance(doc: PDFDocument, spec: Extract<DrawingSpec, { kind: 'rect' | 'ellipse' }>, rect: Rect): PDFRef {
  const [x1, y1, x2, y2] = canonicalRect(spec.rect)
  const inset = Math.max(0.1, spec.width) / 2
  const width = x2 - x1
  const height = y2 - y1
  const commands = [stroke(colorTuple(spec.color), spec.width)]
  if (spec.kind === 'rect') {
    commands.push(`${format(x1 - rect[0] + inset)} ${format(y1 - rect[1] + inset)} ${format(Math.max(0, width - inset * 2))} ${format(Math.max(0, height - inset * 2))} re S`)
  } else {
    const cx = (x1 + x2) / 2 - rect[0]
    const cy = (y1 + y2) / 2 - rect[1]
    const rx = Math.max(0, width / 2 - inset)
    const ry = Math.max(0, height / 2 - inset)
    const k = 0.552284749831
    commands.push(
      `${format(cx + rx)} ${format(cy)} m`,
      `${format(cx + rx)} ${format(cy + k * ry)} ${format(cx + k * rx)} ${format(cy + ry)} ${format(cx)} ${format(cy + ry)} c`,
      `${format(cx - k * rx)} ${format(cy + ry)} ${format(cx - rx)} ${format(cy + k * ry)} ${format(cx - rx)} ${format(cy)} c`,
      `${format(cx - rx)} ${format(cy - k * ry)} ${format(cx - k * rx)} ${format(cy - ry)} ${format(cx)} ${format(cy - ry)} c`,
      `${format(cx + k * rx)} ${format(cy - ry)} ${format(cx + rx)} ${format(cy - k * ry)} ${format(cx + rx)} ${format(cy)} c S`,
    )
  }
  return appearanceStream(doc, rect, commands.join('\n'))
}

function inkAppearance(doc: PDFDocument, spec: Extract<DrawingSpec, { kind: 'ink' }>, rect: Rect): PDFRef {
  const commands = [stroke(colorTuple(spec.color), spec.width)]
  for (const path of spec.paths) {
    if (path.length < 2 || path.length % 2 !== 0 || path.some((value) => !Number.isFinite(value))) {
      throw new TypeError('ink paths must contain finite x/y pairs')
    }
    commands.push(`${format(path[0] - rect[0])} ${format(path[1] - rect[1])} m`)
    for (let index = 2; index < path.length; index += 2) {
      commands.push(`${format(path[index] - rect[0])} ${format(path[index + 1] - rect[1])} l`)
    }
    commands.push('S')
  }
  return appearanceStream(doc, rect, commands.join('\n'))
}

export async function applyDrawings(bytes: Uint8Array, drawings: DrawingSpec[]): Promise<Uint8Array> {
  if (drawings.length === 0) return bytes
  const doc = await PDFDocument.load(bytes)
  const localNotes = new Map<string, PDFRef>()

  for (const spec of drawings) {
    const page = checkedPage(doc, spec.page)
    const color = colorTuple(spec.color)

    if (spec.kind === 'note') {
      const rect: Rect = [spec.at[0], spec.at[1], spec.at[0] + 20, spec.at[1] + 20]
      let reply: PDFRef | undefined
      if (spec.replyToLocalId) {
        reply = localNotes.get(spec.replyToLocalId)
      } else if (spec.replyToSaved) {
        reply = locateSavedNote(doc, spec.replyToSaved)
      }
      const ref = addAnnotation(doc, page, {
        Subtype: 'Text', Rect: rect, C: color, Contents: pdfText(spec.contents),
        T: spec.author ? pdfText(spec.author) : undefined,
        CreationDate: pdfDate(spec.createdMs), M: pdfDate(),
        IRT: reply, RT: reply ? 'R' : undefined,
      })
      if (spec.localId) {
        if (localNotes.has(spec.localId)) throw new Error(`duplicate note localId: ${spec.localId}`)
        localNotes.set(spec.localId, ref)
      }
      continue
    }

    if (spec.kind === 'ink') {
      const allPoints = spec.paths.flat()
      if (allPoints.length === 0) throw new TypeError('ink annotation must contain a path')
      const rect = rectFromPoints(allPoints, Math.max(0.1, spec.width) / 2)
      const inkList = spec.paths.map((path) => path.slice())
      const ref = addAnnotation(doc, page, { Subtype: 'Ink', Rect: rect, InkList: inkList, C: color, BS: { W: spec.width }, F: 4, M: pdfDate() })
      setNormalAppearance(doc, ref, inkAppearance(doc, spec, rect))
      continue
    }

    if (spec.kind === 'rect' || spec.kind === 'ellipse') {
      const rect = canonicalRect(spec.rect)
      const ref = addAnnotation(doc, page, { Subtype: spec.kind === 'rect' ? 'Square' : 'Circle', Rect: rect, C: color, BS: { W: spec.width }, F: 4, M: pdfDate() })
      setNormalAppearance(doc, ref, shapeAppearance(doc, spec, rect))
      continue
    }

    const padding = Math.max(3, spec.width * 4)
    const rect = rectFromPoints([...spec.from, ...spec.to], padding)
    const isArrow = spec.kind === 'arrow'
    const ref = addAnnotation(doc, page, {
      Subtype: 'Line', Rect: rect, L: [...spec.from, ...spec.to], C: color,
      BS: { W: spec.width }, LE: ['None', isArrow ? 'OpenArrow' : 'None'], F: 4, M: pdfDate(),
    })
    setNormalAppearance(doc, ref, lineAppearance(doc, rect, spec.from, spec.to, color, spec.width, isArrow))
  }
  return doc.save()
}

export async function applyNoteEdits(bytes: Uint8Array, edits: NoteEditSpec[]): Promise<Uint8Array> {
  if (edits.length === 0) return bytes
  const doc = await PDFDocument.load(bytes)
  let changed = false
  for (const edit of edits) {
    if (!Number.isInteger(edit.page) || edit.page < 1 || edit.page > doc.getPageCount()) continue
    const page = doc.getPage(edit.page - 1)
    const match = annotationEntries(page).find((entry) =>
      entry.ref?.objectNumber === edit.objNum &&
      nameValue(entry.dict, 'Subtype') === 'Text' &&
      sameNumbers(numberArray(entry.dict, 'Rect'), canonicalRect(edit.rect)) &&
      textValue(entry.dict, 'Contents') === edit.oldContents,
    )
    if (!match) continue
    match.dict.set(pdfName('Contents'), pdfText(edit.contents))
    match.dict.set(pdfName('M'), pdfDate())
    changed = true
  }
  return changed ? doc.save() : bytes
}
